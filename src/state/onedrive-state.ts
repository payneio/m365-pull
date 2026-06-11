import type { PublicClientApplication } from "@azure/msal-browser"

// State lives at /Apps/m365-pull/state.json in the user's OneDrive,
// accessible via the AppFolder special endpoint. The narrow
// Files.ReadWrite.AppFolder scope can ONLY read/write inside this folder —
// the rest of the user's OneDrive remains out of reach.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const STATE_PATH = "/me/drive/special/approot:/state.json:/content"
const SCOPES = ["Files.ReadWrite.AppFolder"]

export interface ChatPrefs {
  /** ISO timestamp of the last successful chat sync for this chat. */
  lastSync?: string
}

export interface RecordingPrefs {
  /** Last successful transcript download ISO timestamp. */
  lastSync?: string
}

export type Destination = "browser" | "onedrive"

export interface RecordingRange {
  kind: "this-week" | "last-7d" | "last-30d" | "since-last-download" | "custom"
  /** yyyy-mm-dd; only meaningful when kind === "custom" */
  customFrom?: string
  /** yyyy-mm-dd; only meaningful when kind === "custom" */
  customTo?: string
}

export interface ChatRange {
  kind: "this-week" | "last-7d" | "last-30d" | "since-last-download" | "custom"
  /** yyyy-mm-dd; only meaningful when kind === "custom" */
  customFrom?: string
  /** yyyy-mm-dd; only meaningful when kind === "custom" */
  customTo?: string
}

export interface UserPrefs {
  destination: Destination
  oneDriveFolder: string
  /** Last-used recording date range; synced across devices. Defaults to last-7d. */
  recordingRange?: RecordingRange
  /** Last-used chat list date range; synced across devices. Defaults to last-7d. */
  chatRange?: ChatRange
  /** When true, always include marked chats in the list regardless of range. Defaults to true. */
  markedInclude?: boolean
  /** When true, hide recordings that have already been downloaded. */
  hideDownloaded?: boolean
  /** When true, always include marked recording containers in the list regardless of range. Defaults to true. */
  recordingMarkedInclude?: boolean
}

export interface AppState {
  version: 1
  updatedAt: string
  updatedBy: string
  /** Item IDs the user has marked. Holds chat IDs and recording IDs
   * (composite callId::filename) -- distinct keyspaces, same Set. */
  marks: string[]
  chatPrefs?: Record<string, ChatPrefs>
  recordingPrefs?: Record<string, RecordingPrefs>
  /** User-level preferences; sync across devices via this state blob. */
  userPrefs?: UserPrefs
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

/** Load state from OneDrive. Returns null on 404 (first run). */
export async function loadOneDriveState(
  msal: PublicClientApplication,
): Promise<AppState | null> {
  const token = await getToken(msal, SCOPES)
  const response = await fetch(`${GRAPH_BASE}${STATE_PATH}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (response.status === 404) return null
  if (!response.ok) {
    const body = await response.text()
    throw new Error(
      `Load state: ${response.status} ${response.statusText} — ${body.slice(0, 300)}`,
    )
  }
  const data = (await response.json()) as AppState
  if (data.version !== 1) {
    console.warn(
      "Unknown state schema version; treating as empty:",
      data.version,
    )
    return null
  }
  return data
}

/** Save state to OneDrive. Overwrites unconditionally (last-writer-wins). */
export async function saveOneDriveState(
  msal: PublicClientApplication,
  state: AppState,
): Promise<void> {
  const token = await getToken(msal, SCOPES)
  const response = await fetch(`${GRAPH_BASE}${STATE_PATH}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(state, null, 2),
  })
  if (!response.ok) {
    const body = await response.text()
    throw new Error(
      `Save state: ${response.status} ${response.statusText} — ${body.slice(0, 300)}`,
    )
  }
}

/** Merge two states. Marks are unioned; chatPrefs prefer the newer entry. */
export function mergeStates(
  local: AppState,
  remote: AppState | null,
): AppState {
  if (!remote) return local
  // Marks: union — they're additive by nature
  const marks = new Set<string>([...local.marks, ...remote.marks])
  // chatPrefs: prefer whichever side has a more recent updatedAt
  const localTime = Date.parse(local.updatedAt) || 0
  const remoteTime = Date.parse(remote.updatedAt) || 0
  const newer = remoteTime > localTime ? remote : local
  const older = newer === local ? remote : local
  const chatPrefs: Record<string, ChatPrefs> = {
    ...(older.chatPrefs ?? {}),
    ...(newer.chatPrefs ?? {}),
  }
  // userPrefs: newer wins (single-row config, not additive)
  const userPrefs = newer.userPrefs ?? older.userPrefs
  // recordingPrefs: per-key spread, same shape as chatPrefs
  const recordingPrefs: Record<string, RecordingPrefs> = {
    ...(older.recordingPrefs ?? {}),
    ...(newer.recordingPrefs ?? {}),
  }
  return {
    version: 1,
    updatedAt: new Date().toISOString(),
    updatedBy: local.updatedBy || remote.updatedBy,
    marks: [...marks].sort(),
    chatPrefs: Object.keys(chatPrefs).length > 0 ? chatPrefs : undefined,
    recordingPrefs:
      Object.keys(recordingPrefs).length > 0 ? recordingPrefs : undefined,
    userPrefs,
  }
}

/** Short, recognizable device identifier for the updatedBy field. */
export function deviceIdentifier(): string {
  const ua = navigator.userAgent
  const platform =
    /Windows/.test(ua) ? "Windows" :
    /Mac/.test(ua) ? "macOS" :
    /Linux/.test(ua) ? "Linux" :
    "Unknown"
  const browser =
    /Edg\//.test(ua) ? "Edge" :
    /Chrome\//.test(ua) ? "Chrome" :
    /Firefox\//.test(ua) ? "Firefox" :
    /Safari\//.test(ua) ? "Safari" :
    "Browser"
  return `${platform}/${browser}`
}
