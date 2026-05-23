// Persistent set of "marked" chat IDs (the in-app favorite/pin replacement
// for Teams sidebar organization, which Graph doesn't expose).
//
// v0 lives in localStorage keyed by user. When P5 (OneDrive AppFolder state)
// lands, this becomes part of the per-user state.json blob in OneDrive so
// marks sync across devices. The MarksData shape will grow to include
// per-chat lookback overrides and lastSync timestamps at that point.

const KEY_PREFIX = "m365-pull.marks.v1."

interface MarksData {
  chats: string[]
}

function keyFor(userKey: string): string {
  return `${KEY_PREFIX}${userKey}`
}

export function loadMarks(userKey: string): Set<string> {
  try {
    const raw = localStorage.getItem(keyFor(userKey))
    if (!raw) return new Set()
    const data = JSON.parse(raw) as MarksData
    return new Set(data.chats || [])
  } catch {
    return new Set()
  }
}

export function saveMarks(userKey: string, ids: Set<string>): void {
  try {
    const data: MarksData = { chats: [...ids] }
    localStorage.setItem(keyFor(userKey), JSON.stringify(data))
  } catch (err) {
    console.warn("Failed to save marks:", err)
  }
}
