import type { ChatPrefs } from "../state/onedrive-state"

// Local cache of per-chat preferences (currently just lastSync timestamps).
// Mirrors what's in OneDrive state.chatPrefs so the UI can render
// "downloaded X" labels and resolve "since-last-download" lookbacks
// without waiting on a Graph round-trip.

const KEY_PREFIX = "m365-pull.chatPrefs.v1."

function keyFor(userKey: string): string {
  return `${KEY_PREFIX}${userKey}`
}

export function loadChatPrefs(userKey: string): Record<string, ChatPrefs> {
  try {
    const raw = localStorage.getItem(keyFor(userKey))
    if (!raw) return {}
    const data = JSON.parse(raw) as Record<string, ChatPrefs>
    return data || {}
  } catch {
    return {}
  }
}

export function saveChatPrefs(
  userKey: string,
  prefs: Record<string, ChatPrefs>,
): void {
  try {
    localStorage.setItem(keyFor(userKey), JSON.stringify(prefs))
  } catch (err) {
    console.warn("Failed to save chatPrefs:", err)
  }
}
