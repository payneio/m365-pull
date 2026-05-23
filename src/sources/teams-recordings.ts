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
// Two unknowns surface here at runtime:
//   - SP-resource token acquisition may need an extra app-registration
//     API permission (e.g. AllSites.Read delegated on Office 365 SharePoint).
//   - SharePoint CORS may reject our cross-origin SPA call. The extension
//     dodges CORS via Manifest V3 host_permissions; we cannot.

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

interface DriveItemHit {
  name?: string
  webUrl?: string
  id?: string
  parentReference?: { driveId?: string }
  remoteItem?: {
    id?: string
    webUrl?: string
    name?: string
    parentReference?: { driveId?: string }
  }
}

interface SearchResponse {
  value?: Array<{
    hitsContainers?: Array<{
      hits?: Array<{
        resource?: DriveItemHit
      }>
    }>
  }>
}

function yyyymmdd(iso: string): string {
  if (!iso) return ""
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ""
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, "0")
  const day = String(d.getUTCDate()).padStart(2, "0")
  return `${y}${m}${day}`
}

function infoFromHit(hit: DriveItemHit): RecordingInfo | null {
  const r = hit.remoteItem ?? hit
  const id = r.id
  const driveId = r.parentReference?.driveId
  const webUrl = r.webUrl
  const name = r.name
  if (!id || !driveId || !webUrl || !name) return null
  let spHost: string
  try {
    spHost = new URL(webUrl).host
  } catch {
    return null
  }
  return { driveId, itemId: id, name, spHost, webUrl }
}

/** Search the user's accessible content for the recording matching a meeting.
 *
 * Teams names recordings predictably:
 *   `<subject>-<YYYYMMDD>_<HHMMSS>UTC-Meeting Recording.mp4`
 *
 * The recording lives in the organizer's `/Recordings/` folder. For meetings
 * the user attended, the recording is shared with them — visible via Graph
 * search across drive items they have access to.
 */
/** Extract a matchable prefix from a meeting subject — strips recurrence-style
 * suffixes Teams appends (e.g., " - 99/1523"), illegal filename chars, and
 * anything after the first separator. */
function subjectPrefix(subject: string): string {
  return subject
    .split(/\s+[-|]\s+|\//)[0] // before " - …", " | …", or "/"
    .replace(/["<>:|?*\\]/g, "") // strip illegal filename chars
    .trim()
    .toLowerCase()
}

export async function findRecordingForMeeting(
  msal: PublicClientApplication,
  subject: string,
  startDateTime: string,
): Promise<RecordingInfo | null> {
  const datePart = yyyymmdd(startDateTime)
  if (!datePart) return null
  const prefix = subjectPrefix(subject).slice(0, 60)
  const token = await getGraphToken(msal, GRAPH_SCOPES)

  // Search broadly: date + "Meeting Recording" catches all recordings from
  // that day across user's accessible content. Subject filtering happens
  // client-side so we don't get bitten by Teams rewriting the subject when
  // it becomes a filename.
  const queryString = `"Meeting Recording" ${datePart}`

  const resp = await fetch(`${GRAPH_BASE}/search/query`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ["driveItem"],
          query: { queryString },
          size: 50,
        },
      ],
    }),
  })
  if (!resp.ok) {
    const body = await resp.text()
    throw new Error(
      `Graph /search/query: ${resp.status} ${resp.statusText} — ${body.slice(0, 200)}`,
    )
  }
  const data = (await resp.json()) as SearchResponse
  const hits = data.value?.[0]?.hitsContainers?.[0]?.hits ?? []

  // Two-pass match. Pass 1: strict — name contains date + "Meeting Recording"
  // + the subject prefix. Pass 2: relaxed — date + "Meeting Recording" only,
  // no subject. If there's only one recording from that day, pass 2 finds it
  // even when the calendar subject diverges from the filename.
  const candidates = hits
    .map((h) => h.resource)
    .filter((r): r is DriveItemHit => Boolean(r))
    .filter((r) => {
      const name = (r.remoteItem?.name ?? r.name ?? "").toLowerCase()
      return (
        name.endsWith(".mp4") &&
        name.includes("meeting recording") &&
        name.includes(datePart)
      )
    })

  console.log(
    "[m365-pull] findRecordingForMeeting",
    JSON.stringify({
      subject,
      subjectPrefix: prefix,
      datePart,
      hits: hits.length,
      candidates: candidates.length,
      names: candidates.map((r) => r.remoteItem?.name ?? r.name),
    }),
  )

  if (candidates.length === 0) return null

  // Strict subject-prefix match
  if (prefix) {
    for (const r of candidates) {
      const name = (r.remoteItem?.name ?? r.name ?? "").toLowerCase()
      if (name.includes(prefix)) {
        const info = infoFromHit(r)
        if (info) return info
      }
    }
  }
  // Relaxed: only one candidate on this date, take it
  if (candidates.length === 1) {
    return infoFromHit(candidates[0])
  }
  // Multiple candidates, none matched the subject prefix — ambiguous.
  // Return the first one anyway and let the user verify the saved filename
  // (recordings are uncommon enough on most users' days that this is rare).
  return infoFromHit(candidates[0])
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
