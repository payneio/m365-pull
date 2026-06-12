// Persistent set of "ignored" chat IDs.
//
// Ignored chats are hidden from the default list view; they appear only when
// the user toggles "Show ignored" in the filter bar. Ignoring is always
// reversible — "Show ignored" reveals them with an un-ignore affordance.
//
// Stored in localStorage for instant reads on load and synced to
// OneDrive state.json (the source of truth) for cross-device persistence.
// The OneDrive sync merges ignored IDs additively (same set-union strategy as
// marks) so no ignores are lost across devices.

const KEY_PREFIX = "m365-pull.ignored.v1."

interface IgnoredData {
  chats: string[]
}

function keyFor(userKey: string): string {
  return `${KEY_PREFIX}${userKey}`
}

export function loadIgnored(userKey: string): Set<string> {
  try {
    const raw = localStorage.getItem(keyFor(userKey))
    if (!raw) return new Set()
    const data = JSON.parse(raw) as IgnoredData
    return new Set(data.chats || [])
  } catch {
    return new Set()
  }
}

export function saveIgnored(userKey: string, ids: Set<string>): void {
  try {
    const data: IgnoredData = { chats: [...ids] }
    localStorage.setItem(keyFor(userKey), JSON.stringify(data))
  } catch (err) {
    console.warn("Failed to save ignored:", err)
  }
}
