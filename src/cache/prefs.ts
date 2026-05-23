// localStorage warm-cache for user-level preferences. The OneDrive state.json
// (under AppFolder) is the source of truth for cross-device sync; this cache
// just provides instant access on page load before the OneDrive pull lands.

import type { UserPrefs } from "../state/onedrive-state"
export type { UserPrefs, Destination } from "../state/onedrive-state"

const KEY_PREFIX = "m365-pull.userPrefs.v1."

const DEFAULTS: UserPrefs = {
  destination: "browser",
  oneDriveFolder: "/m365-pull/teams-chats",
}

function keyFor(userKey: string): string {
  return `${KEY_PREFIX}${userKey}`
}

export function loadUserPrefs(userKey: string): UserPrefs {
  try {
    const raw = localStorage.getItem(keyFor(userKey))
    if (!raw) return { ...DEFAULTS }
    const data = JSON.parse(raw) as Partial<UserPrefs>
    return {
      destination: data.destination ?? DEFAULTS.destination,
      oneDriveFolder: data.oneDriveFolder ?? DEFAULTS.oneDriveFolder,
    }
  } catch {
    return { ...DEFAULTS }
  }
}

export function saveUserPrefs(userKey: string, prefs: UserPrefs): void {
  try {
    localStorage.setItem(keyFor(userKey), JSON.stringify(prefs))
  } catch (err) {
    console.warn("Failed to save user prefs:", err)
  }
}
