// localStorage for UI preferences (filter state, dropdown selections, current
// source). Kept per-user-key. Not synced cross-device on purpose -- different
// devices may want different views (mobile vs desktop, etc.). If we ever want
// it cross-device, fold into the OneDrive state blob.

type SortKey = "marked-first" | "recent" | "name"

export interface UIState {
  currentSource?: string
  /** Chats source: lookback dropdown value */
  lookback?: string
  chatFilter?: {
    search: string
    enabledTypes: string[]
    sortKey: SortKey
    markedOnly: boolean
  }
  recordingFilter?: {
    search: string
    enabledKinds: string[]
    sortKey: SortKey
    markedOnly: boolean
  }
}

const KEY_PREFIX = "m365-pull.uiState.v1."

function keyFor(userKey: string): string {
  return `${KEY_PREFIX}${userKey}`
}

export function loadUIState(userKey: string): UIState {
  try {
    const raw = localStorage.getItem(keyFor(userKey))
    if (!raw) return {}
    return JSON.parse(raw) as UIState
  } catch {
    return {}
  }
}

export function saveUIState(userKey: string, state: UIState): void {
  try {
    localStorage.setItem(keyFor(userKey), JSON.stringify(state))
  } catch (err) {
    console.warn("Failed to save UI state:", err)
  }
}
