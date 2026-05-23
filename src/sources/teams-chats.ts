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
