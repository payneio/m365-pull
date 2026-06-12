import type { PublicClientApplication } from "@azure/msal-browser"

// Teams meeting recording transcripts via SharePoint REST API v2.1.
//
// Why this exists: /me/onlineMeetings and its transcript endpoints are
// hard-gated to meetings the signed-in user organized. Microsoft Graph
// has no delegated path for non-organizer transcripts. But the actual
// transcript files live on the *recording video file* in SharePoint
// (or OneDrive), and anyone with access to the recording can read its
// `media/transcripts` metadata.
//
// Strategy mirrors the Chrome extension `bkrabach/teams-transcript-md`:
//   1. Resolve sharing URL -> driveId, itemId, SP hostname via Graph
//      /shares/{encoded-url}/driveItem (Graph token).
//   2. Hit SharePoint REST v2.1 with a SharePoint-resource bearer token
//      for `?$expand=media/transcripts`. (SP REST, not Graph -- Graph
//      doesn't expose media/transcripts as an expansion.)
//   3. Rewrite each transcript's temporaryDownloadUrl to the
//      `streamContent?is=1&applymediaedits=false` form to get raw VTT.
//
// Requires AllSites.Read (or Sites.Read.All) on the Office 365 SharePoint
// Online API in the Entra app registration for SP resource-token acquisition.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const GRAPH_SCOPES = ["Files.Read.All"]

/** Encode a URL into Microsoft Graph's sharing-URL token format. */
function urlToShareId(url: string): string {
  // base64-encode the URL, then make URL-safe and strip padding
  const b64 = btoa(unescape(encodeURIComponent(url)))
  return "u!" + b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-")
}

function spHostFromUrl(url: string): string {
  try {
    return new URL(url).host
  } catch {
    throw new Error(`Not a valid URL: ${url}`)
  }
}

async function getGraphToken(
  msal: PublicClientApplication,
  scopes: string[],
): Promise<string> {
  const account = msal.getActiveAccount()
  if (!account) throw new Error("No active account — sign in first.")
  const r = await msal.acquireTokenSilent({ account, scopes })
  return r.accessToken
}

async function getSpToken(
  msal: PublicClientApplication,
  spHost: string,
): Promise<string> {
  const account = msal.getActiveAccount()
  if (!account) throw new Error("No active account — sign in first.")
  // .default asks for whatever permissions are already consented to for
  // this resource. If the app registration has no SharePoint API perms
  // (e.g. AllSites.Read on Office 365 SharePoint Online), this will fail.
  const scopes = [`https://${spHost}/.default`]
  try {
    const r = await msal.acquireTokenSilent({ account, scopes })
    return r.accessToken
  } catch (err) {
    console.warn("Silent SP token failed; falling back to redirect:", err)
    await msal.acquireTokenRedirect({ scopes })
    throw new Error("Redirecting for SharePoint scope…")
  }
}

export interface RecordingInfo {
  driveId: string
  itemId: string
  name: string
  spHost: string
  webUrl: string
}

/** Thrown when a recording lives in another org's SharePoint (cross-tenant
 * meeting). Graph's /shares/{id}/driveItem returns 400 "Invalid hostname for
 * this tenancy" because this account can't resolve a foreign tenant's host.
 * The transcript simply isn't accessible via this account — it's not a real
 * failure, so callers should label and count it separately. */
export class CrossTenantRecordingError extends Error {
  readonly crossTenant = true as const
  constructor(
    message = "Recording stored in another tenant — transcript not accessible via this account",
  ) {
    super(message)
    this.name = "CrossTenantRecordingError"
  }
}


/** Resolve a SharePoint sharing URL (or stream.aspx URL) to the underlying
 * drive/item via Microsoft Graph. Requires Files.Read.All. */
export async function resolveRecordingFromUrl(
  msal: PublicClientApplication,
  spUrl: string,
): Promise<RecordingInfo> {
  const spHost = spHostFromUrl(spUrl)
  const shareId = urlToShareId(spUrl)
  const token = await getGraphToken(msal, GRAPH_SCOPES)

  const resp = await fetch(`${GRAPH_BASE}/shares/${shareId}/driveItem`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!resp.ok) {
    const body = await resp.text()
    // Cross-tenant recording: the .mp4 lives in another org's SharePoint, which
    // this account can't resolve. Graph returns 400 "Invalid hostname for this
    // tenancy" / invalidRequest. Surface as a typed, non-failure outcome.
    if (resp.status === 400 && /hostname|invalidRequest/i.test(body)) {
      throw new CrossTenantRecordingError()
    }
    throw new Error(
      `Resolve share: ${resp.status} ${resp.statusText} — ${body.slice(0, 300)}`,
    )
  }
  const item = (await resp.json()) as {
    id: string
    name: string
    webUrl: string
    parentReference?: { driveId?: string }
  }
  if (!item.parentReference?.driveId) {
    throw new Error("Resolved item is missing parentReference.driveId.")
  }
  return {
    driveId: item.parentReference.driveId,
    itemId: item.id,
    name: item.name,
    spHost,
    webUrl: item.webUrl,
  }
}

export interface RecordingTranscriptPayload {
  source: "teams.recordings"
  recordingName: string
  recordingUrl: string
  driveId: string
  itemId: string
  spHost: string
  fetchedAt: string
  transcriptCount: number
  transcripts: Array<{
    id: string
    contentType: string
    content: string
  }>
}

/** Fetch all transcripts attached to a recording via SharePoint REST. */
export async function fetchRecordingTranscripts(
  msal: PublicClientApplication,
  recording: RecordingInfo,
  options: {
    onProgress?: (progress: { stage: string; count?: number; total?: number }) => void
  } = {},
): Promise<RecordingTranscriptPayload> {
  options.onProgress?.({ stage: "Acquiring SharePoint token…" })
  const spToken = await getSpToken(msal, recording.spHost)

  options.onProgress?.({ stage: "Reading transcript metadata…" })
  const metaUrl =
    `https://${recording.spHost}/_api/v2.1/drives/${recording.driveId}/items/${recording.itemId}` +
    `?select=name,media/transcripts&$expand=media/transcripts`
  const metaResp = await fetch(metaUrl, {
    headers: {
      Authorization: `Bearer ${spToken}`,
      Accept: "application/json",
    },
  })
  if (!metaResp.ok) {
    const body = await metaResp.text()
    throw new Error(
      `SharePoint metadata: ${metaResp.status} ${metaResp.statusText} — ${body.slice(0, 300)}`,
    )
  }
  const meta = (await metaResp.json()) as {
    name?: string
    media?: {
      transcripts?: Array<{ id: string; temporaryDownloadUrl?: string }>
    }
  }
  const transcriptMetas = meta.media?.transcripts ?? []

  const base: RecordingTranscriptPayload = {
    source: "teams.recordings",
    recordingName: meta.name ?? recording.name,
    recordingUrl: recording.webUrl,
    driveId: recording.driveId,
    itemId: recording.itemId,
    spHost: recording.spHost,
    fetchedAt: new Date().toISOString(),
    transcriptCount: 0,
    transcripts: [],
  }

  if (transcriptMetas.length === 0) {
    return base
  }

  const transcripts: RecordingTranscriptPayload["transcripts"] = []
  for (let i = 0; i < transcriptMetas.length; i++) {
    const t = transcriptMetas[i]
    options.onProgress?.({
      stage: `Fetching transcript ${i + 1} of ${transcriptMetas.length}`,
      count: i,
      total: transcriptMetas.length,
    })
    let dl = t.temporaryDownloadUrl
    if (!dl) continue
    // Rewrite to raw VTT (mirroring extension capture.js)
    if (dl.includes("/content")) {
      dl = dl.replace(
        /\/content(\?.*)?$/,
        "/streamContent?is=1&applymediaedits=false",
      )
    } else if (dl.includes("/streamContent?")) {
      dl = dl.replace(
        /\/streamContent\?.*$/,
        "/streamContent?is=1&applymediaedits=false",
      )
    }

    // Bearer first -- in cross-origin SPA context, SP requires the SP token.
    // (The Chrome extension can get away with cookies; we cannot.)
    let vttResp = await fetch(dl, {
      headers: { Authorization: `Bearer ${spToken}` },
    })
    if (!vttResp.ok && vttResp.status === 401) {
      vttResp = await fetch(dl)
    }
    if (!vttResp.ok) {
      throw new Error(
        `VTT fetch: ${vttResp.status} ${vttResp.statusText}`,
      )
    }
    const vtt = await vttResp.text()
    transcripts.push({
      id: t.id,
      contentType: "text/vtt",
      content: vtt,
    })
  }
  options.onProgress?.({
    stage: "Done.",
    count: transcripts.length,
    total: transcripts.length,
  })

  return { ...base, transcriptCount: transcripts.length, transcripts }
}
