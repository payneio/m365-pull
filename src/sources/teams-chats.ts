import type { PublicClientApplication } from "@azure/msal-browser"

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"

export interface TeamsChatMember {
  displayName?: string
  userId?: string
}

export interface TeamsChatItem {
  id: string
  topic: string | null
  chatType: "oneOnOne" | "group" | "meeting" | string
  createdDateTime: string
  lastUpdatedDateTime: string
  webUrl: string
  members?: TeamsChatMember[]
}

export interface TeamsChatMessage {
  id: string
  createdDateTime: string
  from?: { user?: { displayName?: string; id?: string } } | null
  body?: { content?: string; contentType?: string } | null
  attachments?: unknown[]
  importance?: string
}

async function getToken(
  msal: PublicClientApplication,
  scopes: string[],
): Promise<string> {
  const account = msal.getActiveAccount()
  if (!account) throw new Error("No active account — sign in first.")
  try {
    const result = await msal.acquireTokenSilent({ account, scopes })
    return result.accessToken
  } catch (err) {
    // Silent acquisition can fail when a scope hasn't been consented yet,
    // or the refresh token is stale. Fall back to interactive redirect.
    console.warn("Silent token acquisition failed; falling back to redirect:", err)
    await msal.acquireTokenRedirect({ scopes })
    // The redirect navigates away; throw so callers stop.
    throw new Error("Redirecting for consent…")
  }
}

async function graphGet<T>(
  msal: PublicClientApplication,
  path: string,
  scopes: string[],
): Promise<T> {
  const token = await getToken(msal, scopes)
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!response.ok) {
    const body = await response.text()
    throw new Error(
      `Graph ${path}: ${response.status} ${response.statusText} — ${body.slice(0, 500)}`,
    )
  }
  return response.json() as Promise<T>
}

export interface ListChatsPageResult {
  chats: TeamsChatItem[]
  /** Opaque cursor for the next page; null when there are no more pages. */
  nextCursor: string | null
}

/** Fetch a single page of the user's Teams chats.
 *
 * Pass `cursor=null` for the first page; pass the previous result's
 * `nextCursor` for subsequent pages. When `nextCursor` is null, the list is
 * exhausted.
 */
export async function listChatsPage(
  msal: PublicClientApplication,
  cursor: string | null = null,
): Promise<ListChatsPageResult> {
  interface ChatsApiPage {
    value: TeamsChatItem[]
    "@odata.nextLink"?: string
  }
  const path = cursor ?? "/me/chats?$top=50&$expand=members"
  const page = await graphGet<ChatsApiPage>(msal, path, ["Chat.Read"])
  return {
    chats: page.value,
    nextCursor: page["@odata.nextLink"] ?? null,
  }
}

/** List the user's Teams chats. Pages through all results.
 *
 * Convenience wrapper over `listChatsPage` for callers that want the whole
 * list at once. The `onProgress` callback fires after each page lands, with
 * the running total.
 */
export async function listChats(
  msal: PublicClientApplication,
  onProgress?: (count: number) => void,
): Promise<TeamsChatItem[]> {
  const out: TeamsChatItem[] = []
  let cursor: string | null = null
  do {
    const page = await listChatsPage(msal, cursor)
    out.push(...page.chats)
    onProgress?.(out.length)
    cursor = page.nextCursor
  } while (cursor)
  return out
}

export interface FetchProgress {
  count: number
  /** Oldest createdDateTime seen so far across all fetched pages. */
  oldestSeen: Date | null
}

export interface FetchMessagesOptions {
  since?: Date
  maxMessages?: number
  /** Called after each page lands, with the running total and oldest seen. */
  onProgress?: (progress: FetchProgress) => void
}

/** Fetch messages from a specific chat.
 *
 * Graph does not support `$filter` on chat-message dates, so we page back in
 * descending order until we either hit `since` or run out of messages, then
 * filter client-side.
 */
export async function fetchChatMessages(
  msal: PublicClientApplication,
  chatId: string,
  options: FetchMessagesOptions = {},
): Promise<TeamsChatMessage[]> {
  interface MsgsPage {
    value: TeamsChatMessage[]
    "@odata.nextLink"?: string
  }
  const cap = options.maxMessages ?? 5000
  const out: TeamsChatMessage[] = []
  let oldestSeen: Date | null = null
  let path: string | null = `/me/chats/${encodeURIComponent(chatId)}/messages?$top=50`
  while (path && out.length < cap) {
    const page: MsgsPage = await graphGet<MsgsPage>(msal, path, ["Chat.Read"])
    out.push(...page.value)
    // Track oldest message seen so the UI can show "back to <date>".
    if (page.value.length > 0) {
      const last = page.value[page.value.length - 1]
      if (last.createdDateTime) {
        const lastDate = new Date(last.createdDateTime)
        if (!oldestSeen || lastDate < oldestSeen) oldestSeen = lastDate
      }
    }
    options.onProgress?.({ count: out.length, oldestSeen })
    path = page["@odata.nextLink"] ?? null
    // Optimization: if a `since` cutoff is set and the OLDEST message in this
    // page is already older than the cutoff, stop paging early.
    if (options.since && oldestSeen && oldestSeen < options.since) break
  }
  if (options.since) {
    const cutoff = options.since.getTime()
    return out.filter(
      (m) => m.createdDateTime && new Date(m.createdDateTime).getTime() >= cutoff,
    )
  }
  return out.slice(0, cap)
}

/** Best-effort display name for a chat. */
/** Build a human-readable, conflict-resistant filename for a chat archive.
 *
 * - `withTimestamp: true` -> includes <YYYY-MM-DD>-<HHMM> prefix. Used for
 *   browser saves (one file per download, never overwrites).
 * - `withTimestamp: false` -> stable per chat. Used for OneDrive saves (the
 *   cumulative-archive pattern relies on the same filename across downloads).
 *
 * The chatId is hashed to 8 hex chars (djb2, deterministic) for uniqueness
 * without the 80+ chars of dashes-stripped chatId.
 *
 * Examples:
 *   2026-05-27-1843-Marc-Goodner-a3f12b8c.json   (browser, 1:1)
 *   Team-Pulse-Workstream-5d2e9f01.json          (OneDrive, named group)
 */
export function buildChatArchiveFilename(
  chatId: string,
  displayName: string,
  options: { withTimestamp: boolean; extension: string },
): string {
  const slug =
    (displayName || "chat")
      .replace(/[^a-zA-Z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 60)
      .replace(/-+$/g, "") || "chat"

  // djb2 hash of chatId -> 8 hex chars. Same input -> same output, so
  // OneDrive's archive lookup stays stable across downloads.
  let h = 5381
  for (let i = 0; i < chatId.length; i++) {
    h = ((h << 5) + h + chatId.charCodeAt(i)) | 0
  }
  const shortId = (h >>> 0).toString(16).padStart(8, "0")

  const ext = options.extension.startsWith(".")
    ? options.extension
    : `.${options.extension}`

  if (options.withTimestamp) {
    const d = new Date()
    const pad = (n: number) => String(n).padStart(2, "0")
    const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
    const time = `${pad(d.getHours())}${pad(d.getMinutes())}`
    return `${date}-${time}-${slug}-${shortId}${ext}`
  }
  return `${slug}-${shortId}${ext}`
}

/** Enrich each chat's `lastUpdatedDateTime` with the actual latest message
 * timestamp from `/me/chats/{id}/messages`. Graph's chat-level field is
 * known-stale: it doesn't bump for every new message (especially in 1:1s),
 * so the list can show months-old dates for chats that were active yesterday.
 *
 * One $top=1 message fetch per chat. Run with a concurrency cap so we don't
 * fire 50 requests at once. Failures (closed/archived chats) leave the
 * original value in place rather than blocking the whole load.
 */
export async function enrichChatsWithLatestActivity(
  msal: PublicClientApplication,
  chats: TeamsChatItem[],
  options: { concurrency?: number; onProgress?: (note: string) => void } = {},
): Promise<TeamsChatItem[]> {
  const concurrency = Math.max(1, options.concurrency ?? 5)
  const out: TeamsChatItem[] = chats.map((c) => c)
  let cursor = 0
  let done = 0

  async function worker(): Promise<void> {
    while (true) {
      const idx = cursor++
      if (idx >= chats.length) return
      const chat = chats[idx]
      try {
        const path = `/me/chats/${encodeURIComponent(chat.id)}/messages?$top=1&$select=createdDateTime`
        const resp = await graphGet<{ value: { createdDateTime?: string }[] }>(
          msal,
          path,
          ["Chat.Read"],
        )
        const latest = resp.value?.[0]?.createdDateTime
        if (latest && latest > chat.lastUpdatedDateTime) {
          out[idx] = { ...chat, lastUpdatedDateTime: latest }
        }
      } catch {
        // Leave the chat as-is on lookup failure.
      }
      done++
      options.onProgress?.(`Verifying activity (${done}/${chats.length})\u2026`)
    }
  }

  await Promise.all(Array.from({ length: concurrency }, () => worker()))
  return out
}

export function chatDisplayName(chat: TeamsChatItem): string {
  if (chat.topic) return chat.topic
  const members = (chat.members ?? [])
    .map((m) => m.displayName)
    .filter((n): n is string => Boolean(n))
  if (chat.chatType === "oneOnOne") {
    return members.join(", ") || "(1:1 chat)"
  }
  if (chat.chatType === "group") {
    return members.join(", ") || "(group chat)"
  }
  if (chat.chatType === "meeting") return "(meeting chat)"
  return "(unnamed chat)"
}
