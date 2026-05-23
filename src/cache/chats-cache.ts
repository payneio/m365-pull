import type { TeamsChatItem } from "../sources/teams-chats"

// Bumped when the cache shape changes — old entries are ignored.
// v2 adds `nextCursor` to support partial / progressive loading.
const KEY_PREFIX = "m365-pull.chats.v2."

export interface ChatsCache {
  fetchedAt: string
  chats: TeamsChatItem[]
  /** Opaque Graph cursor for the next page; null when fully loaded. */
  nextCursor: string | null
}

function keyFor(userKey: string): string {
  return `${KEY_PREFIX}${userKey}`
}

export function loadCachedChats(userKey: string): ChatsCache | null {
  try {
    const raw = localStorage.getItem(keyFor(userKey))
    if (!raw) return null
    const parsed = JSON.parse(raw) as ChatsCache
    if (!parsed.fetchedAt || !Array.isArray(parsed.chats)) return null
    return {
      fetchedAt: parsed.fetchedAt,
      chats: parsed.chats,
      nextCursor: parsed.nextCursor ?? null,
    }
  } catch {
    return null
  }
}

export function saveCachedChats(
  userKey: string,
  chats: TeamsChatItem[],
  nextCursor: string | null,
): void {
  try {
    const data: ChatsCache = {
      fetchedAt: new Date().toISOString(),
      chats,
      nextCursor,
    }
    localStorage.setItem(keyFor(userKey), JSON.stringify(data))
  } catch (err) {
    // Quota exceeded or storage unavailable — non-fatal.
    console.warn("Failed to cache chats:", err)
  }
}

export function clearCachedChats(userKey: string): void {
  try {
    localStorage.removeItem(keyFor(userKey))
  } catch {
    /* non-fatal */
  }
}

export function ageMs(cache: ChatsCache): number {
  return Date.now() - new Date(cache.fetchedAt).getTime()
}

export function formatAge(ms: number): string {
  if (ms < 60_000) return "just now"
  const mins = Math.floor(ms / 60_000)
  if (mins < 60) return `${mins} minute${mins === 1 ? "" : "s"} ago`
  const hours = Math.floor(mins / 60)
  if (hours < 24) return `${hours} hour${hours === 1 ? "" : "s"} ago`
  const days = Math.floor(hours / 24)
  return `${days} day${days === 1 ? "" : "s"} ago`
}
