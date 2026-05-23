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

export interface OneDriveLoadResult<T> {
  found: boolean
  data?: T
  reason?: string
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

/** Load a JSON value from a path in the user's OneDrive.
 *
 * Returns `{ found: false }` on 404 (file doesn't exist yet).
 * Throws-into-result on other errors.
 */
export async function loadFromOneDrive<T = unknown>(
  msal: PublicClientApplication,
  path: string,
): Promise<OneDriveLoadResult<T>> {
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
      headers: { Authorization: `Bearer ${token}` },
    })
    if (response.status === 404) return { found: false }
    if (!response.ok) {
      const body = await response.text()
      return {
        found: false,
        reason: `${response.status} ${response.statusText} — ${body.slice(0, 300)}`,
      }
    }
    const data = (await response.json()) as T
    return { found: true, data }
  } catch (err) {
    return { found: false, reason: (err as Error).message }
  }
}

/** Save a JS value as JSON to a path in the user's OneDrive.
 *
 * Graph PUT to /me/drive/root:/path/to/file:/content auto-creates any
 * missing parent directories along the way.
 */
export async function saveToOneDrive(
  msal: PublicClientApplication,
  path: string,
  data: unknown,
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
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data, null, 2),
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
