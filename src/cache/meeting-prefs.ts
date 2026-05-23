import type { MeetingPrefs } from "../state/onedrive-state"

// localStorage warm-cache for per-meeting preferences. Mirrors what's in
// the OneDrive state.meetingPrefs so the UI can render "downloaded X" tags
// without waiting on a Graph round-trip.

const KEY_PREFIX = "m365-pull.meetingPrefs.v1."

function keyFor(userKey: string): string {
  return `${KEY_PREFIX}${userKey}`
}

export function loadMeetingPrefs(
  userKey: string,
): Record<string, MeetingPrefs> {
  try {
    const raw = localStorage.getItem(keyFor(userKey))
    if (!raw) return {}
    const data = JSON.parse(raw) as Record<string, MeetingPrefs>
    return data || {}
  } catch {
    return {}
  }
}

export function saveMeetingPrefs(
  userKey: string,
  prefs: Record<string, MeetingPrefs>,
): void {
  try {
    localStorage.setItem(keyFor(userKey), JSON.stringify(prefs))
  } catch (err) {
    console.warn("Failed to save meetingPrefs:", err)
  }
}
