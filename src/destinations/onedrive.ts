import type { PublicClientApplication } from "@azure/msal-browser"

// Save downloads to a path in the user's OneDrive root via Graph.
// Uses the broad Files.ReadWrite scope, NOT AppFolder — the AppFolder
// is reserved for app state (marks, chatPrefs). Downloads land in a
// location the user actually expects to find them and can sync to
// their local OneDrive client / WSL.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const SCOPES = ["Files.ReadWrite"]

export interface OneDriveSaveResult {
  saved: boolean
  reason?: string
  path?: string
  webUrl?: string
}


async function getToken(
  msal: PublicClientApplication,
  scopes: string[],
): Promise<string> {
  const account = msal.getActiveAccount()
  if (!account) throw new Error("No active account — sign in first.")
  try {
    const r = await msal.acquireTokenSilent({ account, scopes })
    return r.accessToken
  } catch (err) {
    console.warn("Silent token acquisition failed; redirecting:", err)
    await msal.acquireTokenRedirect({ scopes })
    throw new Error("Redirecting for consent…")
  }
}

/** Sanitize each path segment without touching the slash separators. */
function sanitizePath(rawPath: string): string {
  const segments = rawPath.split("/").filter(Boolean)
  const cleaned = segments
    .map((seg) =>
      seg.replace(/[<>:"\\|?*\x00-\x1F]/g, "-").replace(/-+/g, "-").trim(),
    )
    .filter(Boolean)
  return "/" + cleaned.join("/")
}

/** Resolve the destination folder's Graph driveItem `webUrl` — the canonical
 * OneDrive-on-the-web link to open that folder. We read the driveItem metadata
 * (GET /me/drive/root:/{path}) rather than hand-constructing a SharePoint URL,
 * which is tenant-specific and fragile.
 *
 * Returns null when the folder doesn't exist yet (404 — nothing downloaded
 * there), when not signed in, or on any error, so the caller hides the link.
 * This is a PASSIVE convenience read: it acquires the token silently only and
 * NEVER triggers an interactive redirect (unlike saveBytesToOneDrive, which
 * needs the scope to do real work). */
export async function getOneDriveFolderWebUrl(
  msal: PublicClientApplication,
  folderPath: string,
): Promise<string | null> {
  const account = msal.getActiveAccount()
  if (!account) return null
  let token: string
  try {
    const r = await msal.acquireTokenSilent({ account, scopes: SCOPES })
    token = r.accessToken
  } catch {
    // Passive link — never force a consent redirect just to populate a URL.
    return null
  }
  const safePath = sanitizePath(folderPath)
  const encoded = safePath.split("/").filter(Boolean).map(encodeURIComponent).join("/")
  const url = encoded
    ? `${GRAPH_BASE}/me/drive/root:/${encoded}`
    : `${GRAPH_BASE}/me/drive/root`
  try {
    const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } })
    if (!resp.ok) return null // 404 = folder not created yet; anything else = unavailable
    const item = (await resp.json()) as { webUrl?: string }
    return item.webUrl ?? null
  } catch {
    return null
  }
}

/** PUT a raw string (e.g. markdown, text) to a path in the user's OneDrive. */
export async function saveTextToOneDrive(
  msal: PublicClientApplication,
  path: string,
  content: string,
  mimeType: string,
): Promise<OneDriveSaveResult> {
  return saveBytesToOneDrive(msal, path, content, mimeType)
}

async function saveBytesToOneDrive(
  msal: PublicClientApplication,
  path: string,
  body: string,
  contentType: string,
): Promise<OneDriveSaveResult> {
  const safePath = sanitizePath(path)
  const encoded = safePath
    .split("/")
    .filter(Boolean)
    .map(encodeURIComponent)
    .join("/")
  const url = `${GRAPH_BASE}/me/drive/root:/${encoded}:/content`

  try {
    const token = await getToken(msal, SCOPES)
    const response = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": contentType,
      },
      body,
    })
    if (!response.ok) {
      const body = await response.text()
      return {
        saved: false,
        reason: `${response.status} ${response.statusText} — ${body.slice(0, 300)}`,
      }
    }
    const result = (await response.json()) as {
      name?: string
      parentReference?: { path?: string }
      webUrl?: string
    }
    const reportedPath = result.parentReference?.path
      ? `${result.parentReference.path.replace(/^\/drive\/root:/, "")}/${result.name ?? ""}`
      : safePath
    return {
      saved: true,
      path: reportedPath,
      webUrl: result.webUrl,
    }
  } catch (err) {
    return { saved: false, reason: (err as Error).message }
  }
}
