import type { PublicClientApplication } from "@azure/msal-browser"
import { sanitizeFilenameName, formatPulledStamp } from "./filename-format"

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"

export interface TeamsChatMember {
  displayName?: string
  userId?: string
}

/** Minimal shape of the Graph chatMessageInfo resource returned by $expand=lastMessagePreview. */
export interface ChatMessagePreview {
  id?: string
  createdDateTime?: string | null
  /** "message" = a real human message; "systemEventMessage" = membership/system event; others exist. */
  messageType?: string | null
  isDeleted?: boolean
  /** Present on systemEventMessage previews. Its "@odata.type" discriminates the
   * event kind — e.g. #microsoft.graph.callRecordingEventMessageDetail (a real
   * meeting/call event) vs #microsoft.graph.membersDeletedEventMessageDetail
   * (org/roster churn — the phantom-activity inflator). Per the Graph v1.0
   * chatMessageInfo schema, eventDetail is returned with $expand=lastMessagePreview. */
  eventDetail?: { "@odata.type"?: string } | null
}

export interface TeamsChatItem {
  id: string
  topic: string | null
  chatType: "oneOnOne" | "group" | "meeting" | string
  createdDateTime: string
  lastUpdatedDateTime: string
  webUrl: string
  members?: TeamsChatMember[]
  /** Last message preview — present when the list call uses $expand=lastMessagePreview. */
  lastMessagePreview?: ChatMessagePreview | null
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

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

/** Graph API GET with bounded retry for transient errors.
 *
 * Retries on 429 (throttle), 502, 503, 504 (gateway/transient) up to 5
 * attempts. Honors the `Retry-After` response header (seconds, clamped to
 * MAX_RETRY_AFTER_S). Falls back to exponential backoff with jitter when the
 * header is absent. All other non-OK statuses throw immediately.
 */
async function graphGet<T>(
  msal: PublicClientApplication,
  path: string,
  scopes: string[],
): Promise<T> {
  const RETRY_STATUSES = new Set([429, 502, 503, 504])
  const MAX_ATTEMPTS = 5
  const BACKOFF_MS = [1000, 2000, 4000, 8000, 16000]
  const MAX_RETRY_AFTER_S = 60

  let lastStatus = 0
  let lastText = ""

  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    const token = await getToken(msal, scopes)
    const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    })

    if (response.ok) {
      return response.json() as Promise<T>
    }

    lastStatus = response.status
    lastText = await response.text()

    if (!RETRY_STATUSES.has(response.status)) {
      // Non-retryable (400, 401, 403, 404, …) — fail immediately.
      throw new Error(
        `Graph ${path}: ${response.status} ${response.statusText} — ${lastText.slice(0, 500)}`,
      )
    }

    if (attempt === MAX_ATTEMPTS - 1) break // exhausted, throw below

    // Compute wait: honor Retry-After header if present, else exponential backoff.
    let waitMs: number
    const retryAfterHeader = response.headers.get("Retry-After")
    if (retryAfterHeader) {
      const parsed = parseInt(retryAfterHeader, 10)
      waitMs = (isNaN(parsed) ? 1 : Math.min(parsed, MAX_RETRY_AFTER_S)) * 1000
    } else {
      waitMs = BACKOFF_MS[Math.min(attempt, BACKOFF_MS.length - 1)]
    }
    waitMs += Math.random() * 250 // jitter 0–250 ms

    await sleep(waitMs)
  }

  throw new Error(
    `Graph ${path}: ${lastStatus} after ${MAX_ATTEMPTS} attempts — ${lastText.slice(0, 500)}`,
  )
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
  const path = cursor ?? "/me/chats?$top=50&$expand=members,lastMessagePreview"
  const page = await graphGet<ChatsApiPage>(msal, path, ["Chat.Read"])
  return {
    // Slim members to displayName-only at the parse boundary. chatDisplayName()
    // only reads displayName; userId is not consumed in the chats flow.
    // Keeps both the in-memory model and the localStorage cache small.
    chats: page.value.map((chat) => ({
      ...chat,
      members: chat.members?.map((m) => ({ displayName: m.displayName })),
    })),
    nextCursor: page["@odata.nextLink"] ?? null,
  }
}


/** True iff `id` is a valid Teams chat MRI (`19:…@thread.v2` / `@unq.gbl.spaces`).
 *
 * Teams chat ids are thread MRIs that ALWAYS start with `19:`. Ids that look
 * like Outlook/Exchange item ids (base64-ish, starting `AAMk…`) are NOT chat
 * MRIs — they're calendar/meeting-derived ids (or stale marks from an earlier
 * build). Calling `/me/chats/{id}` with one of those 400s with
 * "Invalid MRI, should start with digits and colon". We guard against that
 * here so the marked-include enrichment never fires a doomed request. */
export function isTeamsChatMri(id: string): boolean {
  return id.startsWith("19:")
}

/** Fetch a single chat by ID, expanding slim members (displayName-only) and
 * the last message preview (used to derive a real activity date).
 *
 * Returns null on 404 (deleted / no longer accessible) OR when `chatId` is not
 * a valid Teams chat MRI (no network call made) so callers can fail soft
 * without aborting the overall load. All other errors are re-thrown.
 */
export async function fetchChatById(
  msal: PublicClientApplication,
  chatId: string,
): Promise<TeamsChatItem | null> {
  // Skip ids that aren't Teams chat MRIs — they'd 400 ("Invalid MRI"). These
  // are orphaned/stale marks (e.g. AAMk… Outlook ids); treat as un-fetchable.
  if (!isTeamsChatMri(chatId)) return null
  try {
    const chat = await graphGet<TeamsChatItem>(
      msal,
      `/me/chats/${encodeURIComponent(chatId)}?$expand=members,lastMessagePreview`,
      ["Chat.Read"],
    )
    return {
      ...chat,
      members: chat.members?.map((m) => ({ displayName: m.displayName })),
    }
  } catch (err) {
    // 404 = deleted or no longer accessible; return null so caller can skip.
    if ((err as Error).message.includes("404")) return null
    throw err
  }
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

/** Build a versioned, sort-by-name-friendly filename for a chat archive
 * (Phase 3). Format:
 *
 *   <Name>__chat__pulled-<YYYY-MM-DD-HHMM>__<rangeStart>_to_<rangeEnd>.<ext>
 *
 * - <Name>          = sanitized chat display name (keeps human-readable spaces)
 * - pulled-<...>     = when THIS version was downloaded (primary sort/version key)
 * - <rangeStart>/<rangeEnd> = the real YYYY-MM-DD span of the messages INCLUDED
 *   in this download (resolved bounds, not the literal "all"/"since" words)
 *
 * The same name is used for BOTH the browser and OneDrive destinations, so every
 * pull is its own dated file and sort-by-name reveals the version history.
 *
 * Examples:
 *   Marc Goodner__chat__pulled-2026-06-17-1843__2026-05-18_to_2026-06-17.md
 *   Team Pulse Workstream__chat__pulled-2026-06-17-0902__2026-06-10_to_2026-06-17.md
 */
export function buildChatArchiveFilename(
  displayName: string,
  options: {
    pulledAt: Date
    /** ISO YYYY-MM-DD of the earliest included message (resolved). */
    rangeStart: string
    /** ISO YYYY-MM-DD of the latest included message (resolved). */
    rangeEnd: string
    extension: string
  },
): string {
  const name = sanitizeFilenameName(displayName)
  const pulled = formatPulledStamp(options.pulledAt)
  const ext = options.extension.startsWith(".")
    ? options.extension
    : `.${options.extension}`
  return `${name}__chat__pulled-${pulled}__${options.rangeStart}_to_${options.rangeEnd}${ext}`
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
