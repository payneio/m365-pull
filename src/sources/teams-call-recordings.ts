// Call recordings source -- discovers recordings by walking /me/chats and
// extracting callRecordingEventMessageDetail events.
//
// Deep message scan: each qualifying chat's messages are paginated until the
// oldest message predates the window (createdDateTime < fromMs), not capped at
// 50. A probe confirmed the old top-50 fetch silently missed 11 of 38
// in-window recordings -- whole chats were invisible. Deep scan is mandatory
// for "sync ALL" to be true.
//
// Per-chat page cap: 30 pages (~1500 msgs). If a chat hits the cap before
// reaching the window edge, container.truncated is set to true and surfaced
// in the UI -- we never silently stop.
//
// Concurrency: ~4 chat scans run in parallel, keeping total latency reasonable.
//
// Transcript fetching uses the SharePoint REST v2.1 path in teams-recordings.ts
// (Graph /shares -> driveItem -> _api/v2.1 with media/transcripts expansion).

import type { PublicClientApplication } from "@azure/msal-browser"

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const SCOPES = ["Chat.Read"]
// Members listing requires ChatMember.Read (subset of Chat.Read in some tenants;
// kept explicit so the token request is clear).

export type ChatType = "oneOnOne" | "group" | "meeting"

export interface Participant {
  /** Display name (may be empty for some bot/app participants). */
  displayName: string
  /** Entra oid for users; app id for bot/app participants. */
  id: string | null
  /** "user" for human users, "bot" for app/bot identities. */
  kind: "user" | "bot"
}

export interface RecordingItem {
  /** Composite primary key: callId::filename. Stable across reloads. */
  id: string
  /** Teams call ID -- multiple recordings from the same call share this. */
  callId: string
  /** Recording filename (e.g. "Marc Catch-Up-20260526_120954-Meeting Recording.mp4"). */
  filename: string
  /** SharePoint sharing URL pointing to the .mp4 in the owner's OneDrive. */
  url: string
  /** ISO 8601 duration string (e.g. "PT15M1.826S"). */
  durationIso: string
  /** Meeting organizer oid, when set (scheduled meetings have this; 1:1s don't). */
  organizerOid: string | null
  /** Who started the recording. */
  initiatorOid: string | null
  /** When the recording event was posted (close to when recording finalized). */
  eventCreatedDateTime: string
  /** Chat the recording event was posted in. */
  chatId: string
  /** Chat topic (null for 1:1s and unnamed group chats). */
  chatTopic: string | null
  /** "oneOnOne" | "group" | "meeting". */
  chatType: ChatType
  /** Participants list from the matching callEndedEventMessageDetail event in
   * the same chat (joined by callId). Empty when no callEnded event was seen
   * (e.g. the chat was scanned beyond the message that carried it). */
  participants: Participant[]
}

/** A chat container grouping all in-window recordings from one chat. */
export interface RecordingContainer {
  /** Chat that owns these recordings. */
  chatId: string
  /** Chat topic (null for 1:1s and unnamed group chats). */
  chatTopic: string | null
  /** "oneOnOne" | "group" | "meeting". */
  chatType: ChatType
  /** All recordings from this chat that fall within the scan window. Sorted
   * by eventCreatedDateTime descending (newest first). */
  recordings: RecordingItem[]
  /** True when this chat's message scan hit the per-chat page cap before
   * reaching the window edge. Some in-window recordings may be missing.
   * Surfaced loudly in the UI -- never silent. */
  truncated: boolean
}

export interface RecordingsResult {
  /** One container per chat that had >=1 in-window recording. */
  containers: RecordingContainer[]
  /** Number of chats whose messages we actually scanned. */
  chatsScanned: number
  /** True if we hit the chat-paging cap and there may be more chats unread. */
  truncated: boolean
}

interface ChatListItem {
  id: string
  topic: string | null
  chatType: ChatType
  lastUpdatedDateTime: string
}

interface ChatMessage {
  id: string
  createdDateTime: string
  eventDetail?: {
    "@odata.type"?: string
    callId?: string
    callRecordingDisplayName?: string
    callRecordingUrl?: string
    callRecordingDuration?: string
    callRecordingStatus?: string
    meetingOrganizer?: { user?: { id?: string } }
    initiator?: { user?: { id?: string } }
    /** Present on callEndedEventMessageDetail events. */
    callParticipants?: {
      participant?: {
        user?: { displayName?: string; id?: string }
        application?: { displayName?: string; id?: string }
      }
    }[]
  }
}

function extractParticipants(
  cp: NonNullable<ChatMessage["eventDetail"]>["callParticipants"],
): Participant[] {
  if (!cp) return []
  const out: Participant[] = []
  const seen = new Set<string>()
  for (const entry of cp) {
    const p = entry.participant
    if (!p) continue
    let displayName = ""
    let id: string | null = null
    let kind: "user" | "bot" = "user"
    if (p.user) {
      displayName = (p.user.displayName ?? "").trim()
      id = p.user.id ?? null
      kind = "user"
    } else if (p.application) {
      displayName = (p.application.displayName ?? "").trim() || "(bot)"
      id = p.application.id ?? null
      kind = "bot"
    } else {
      continue
    }
    const key = id ?? `${kind}::${displayName}`
    if (seen.has(key)) continue
    seen.add(key)
    out.push({ displayName, id, kind })
  }
  return out
}

async function getGraphToken(
  msal: PublicClientApplication,
  scopes: string[],
): Promise<string> {
  const account = msal.getActiveAccount()
  if (!account) throw new Error("No active account")
  const r = await msal.acquireTokenSilent({ account, scopes })
  return r.accessToken
}

/** Graph GET with retry: honors Retry-After on 429, exponential backoff on
 * 502/503/504. Up to 5 retries before throwing. Refreshes token each attempt
 * so stale tokens don't block long-running paginated scans. */
async function graphJson<T>(
  msal: PublicClientApplication,
  path: string,
  scopes: string[],
): Promise<T> {
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`
  const MAX_RETRIES = 5
  let retryDelay = 1000
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const token = await getGraphToken(msal, scopes)
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    })
    if (resp.ok) return resp.json() as Promise<T>
    if (
      attempt < MAX_RETRIES &&
      (resp.status === 429 || resp.status === 502 || resp.status === 503 || resp.status === 504)
    ) {
      const ra = resp.headers.get("Retry-After")
      const wait = ra
        ? Math.min(parseInt(ra, 10) * 1000, 60_000)
        : retryDelay + Math.random() * retryDelay
      await new Promise<void>((resolve) => setTimeout(resolve, wait))
      retryDelay = Math.min(retryDelay * 2, 30_000)
      continue
    }
    const body = await resp.text()
    throw new Error(
      `Graph ${path}: ${resp.status} ${resp.statusText} \u2014 ${body.slice(0, 300)}`,
    )
  }
  // All retries exhausted -- this path is unreachable due to the throw above,
  // but TypeScript requires a return/throw after the loop.
  throw new Error(`Graph ${path}: max retries exceeded`)
}

export interface ListRecordingsOptions {
  /** Start of the date window (ms since epoch, inclusive).
   * Chats whose lastUpdatedDateTime is before this are skipped; recordings
   * whose event date is before this are excluded. */
  fromMs: number
  /** End of the date window (ms since epoch, inclusive). Defaults to now. */
  toMs?: number
  /** Cap chat pagination. Default 10 pages of 50 = up to 500 chats. */
  maxChatPages?: number
  /** Cap per-chat message page scans. Default 30 pages (~1500 msgs).
   * If a chat hits this cap before reaching the window edge, the container
   * gets truncated=true and the UI shows a loud warning. Never silent. */
  maxMsgPagesPerChat?: number
  /** Progress callback for UI status updates. */
  onProgress?: (note: string) => void
}

/**
 * Discover all call recordings the user has access to within [fromMs, toMs].
 *
 * Algorithm:
 *   1. Page through /me/chats. Stop early once the oldest chat in a page is
 *      older than fromMs (chats are returned by lastUpdatedDateTime desc).
 *   2. Filter to chats whose lastUpdatedDateTime >= fromMs.
 *   3. For each such chat, page through messages (newest-first) until
 *      createdDateTime < fromMs (window edge) or the per-chat page cap is hit.
 *      If the cap fires before the edge, container.truncated = true.
 *   4. Extract messages where eventDetail.@odata.type is
 *      "#microsoft.graph.callRecordingEventMessageDetail", status is "success",
 *      AND the recording's own createdDateTime is within [fromMs, toMs].
 *   5. Also collect callEndedEventMessageDetail events to get participant lists.
 *   6. Chats are scanned in a concurrency pool of ~4.
 *   7. Deduplicate recordings by (callId, filename) within each chat.
 *   8. Group recordings by chatId into RecordingContainer objects.
 *   9. Join participants by callId; fall back to chat membership for 1:1s
 *      and group chats with no callEnded event.
 *
 * Catches scheduled meetings, 1:1 calls, group-chat calls, multi-recording calls,
 * and recordings stored in others' OneDrives. No calendar lookup, no search.
 */
export async function listRecordings(
  msal: PublicClientApplication,
  options: ListRecordingsOptions,
): Promise<RecordingsResult> {
  const {
    fromMs,
    toMs,
    maxChatPages = 10,
    maxMsgPagesPerChat = 30,
    onProgress,
  } = options
  const untilMs = toMs ?? Date.now()

  // Phase 1: page through chats (unchanged from prior implementation)
  let chatsPath: string | null = "/me/chats?$top=50"
  const allChats: ChatListItem[] = []
  let chatPages = 0
  while (chatsPath && chatPages < maxChatPages) {
    onProgress?.(`Listing chats (page ${chatPages + 1})\u2026`)
    const page: { value: ChatListItem[]; "@odata.nextLink"?: string } =
      await graphJson(msal, chatsPath, SCOPES)
    allChats.push(...page.value)
    const oldest = page.value[page.value.length - 1]?.lastUpdatedDateTime
    if (oldest && new Date(oldest).getTime() < fromMs) break
    chatsPath = page["@odata.nextLink"] ?? null
    chatPages++
  }
  const truncated = chatsPath !== null && chatPages >= maxChatPages

  const recentChats = allChats.filter((c) => {
    if (!c.lastUpdatedDateTime) return false
    return new Date(c.lastUpdatedDateTime).getTime() >= fromMs
  })

  // Phase 2: deep message scan -- paginate each chat until window edge or cap.
  // Shared accumulators (single JS thread, no actual races at await points).
  const hits = new Map<string, RecordingItem>()
  const participantsByCallId = new Map<string, Participant[]>()
  const chatTruncatedSet = new Set<string>()
  let progressDone = 0

  async function scanChat(chat: ChatListItem): Promise<void> {
    let msgPath: string | null =
      `/me/chats/${encodeURIComponent(chat.id)}/messages?$top=50`
    let msgPages = 0
    let passedWindowEdge = false

    try {
      while (msgPath !== null && msgPages < maxMsgPagesPerChat) {
        const msgPage: { value: ChatMessage[]; "@odata.nextLink"?: string } =
          await graphJson(msal, msgPath, SCOPES)
        msgPages++

        for (const m of msgPage.value) {
          const msgMs = new Date(m.createdDateTime).getTime()
          // Messages are newest-first. Once we're past the window edge, stop.
          if (msgMs < fromMs) {
            passedWindowEdge = true
            break
          }
          // Skip messages above the upper bound (custom ranges with a past
          // end date -- the newest messages are too recent for the window).
          if (msgMs > untilMs) continue

          const ed = m.eventDetail
          if (!ed) continue
          const t = ed["@odata.type"]

          if (
            t === "#microsoft.graph.callRecordingEventMessageDetail" &&
            ed.callRecordingStatus === "success" &&
            ed.callRecordingUrl
          ) {
            const callId = ed.callId ?? "(no-callid)"
            const filename = ed.callRecordingDisplayName ?? "(unnamed)"
            const id = `${callId}::${filename}`
            if (!hits.has(id)) {
              hits.set(id, {
                id,
                callId,
                filename,
                url: ed.callRecordingUrl,
                durationIso: ed.callRecordingDuration ?? "",
                organizerOid: ed.meetingOrganizer?.user?.id ?? null,
                initiatorOid: ed.initiator?.user?.id ?? null,
                eventCreatedDateTime: m.createdDateTime,
                chatId: chat.id,
                chatTopic: chat.topic,
                chatType: chat.chatType,
                participants: [],
              })
            }
          } else if (
            t === "#microsoft.graph.callEndedEventMessageDetail" &&
            ed.callId
          ) {
            // Latest participants list wins. For the same call, all callEnded
            // events carry the same list, but we keep the last-seen defensively.
            const parts = extractParticipants(ed.callParticipants)
            if (parts.length > 0) {
              participantsByCallId.set(ed.callId, parts)
            }
          }
        }

        if (passedWindowEdge) break
        msgPath = msgPage["@odata.nextLink"] ?? null
      }

      // If we exited the loop without passing the window edge and there are
      // still more pages, we hit the per-chat cap. Flag it -- never silent.
      if (!passedWindowEdge && msgPath !== null) {
        chatTruncatedSet.add(chat.id)
      }
    } catch (err) {
      // Bad chat shouldn't kill the whole scan. Log and continue.
      console.warn(`[m365-pull] Failed to scan chat ${chat.id}:`, err)
    } finally {
      progressDone++
      onProgress?.(
        `Scanning chats for recordings\u2026 (${progressDone}/${recentChats.length}, ${hits.size} found)`,
      )
    }
  }

  // Concurrency pool: 4 chat scans in parallel
  const CONCURRENCY = 4
  if (recentChats.length > 0) {
    let chatIdx = 0
    const workers = Array.from(
      { length: Math.min(CONCURRENCY, recentChats.length) },
      async () => {
        while (chatIdx < recentChats.length) {
          const i = chatIdx++
          await scanChat(recentChats[i])
        }
      },
    )
    await Promise.all(workers)
  }

  // Phase 3: join participants into recordings by callId
  for (const rec of hits.values()) {
    const parts = participantsByCallId.get(rec.callId)
    if (parts) rec.participants = parts
  }

  // Phase 4: fallback for recordings without participants. 1:1 ad-hoc calls
  // don't emit callEndedEventMessageDetail, so the join above produces nothing.
  // Fall back to chat membership (exact for 1:1; superset for group chats).
  const chatsNeedingFallback = new Set<string>()
  for (const rec of hits.values()) {
    if (rec.participants.length === 0) chatsNeedingFallback.add(rec.chatId)
  }
  if (chatsNeedingFallback.size > 0) {
    onProgress?.(`Filling participants for ${chatsNeedingFallback.size} chats...`)
    const membersByChat = new Map<string, Participant[]>()
    for (const chatId of chatsNeedingFallback) {
      try {
        type MemberResponse = {
          value: {
            displayName?: string
            userId?: string
            email?: string
          }[]
        }
        const resp: MemberResponse = await graphJson(
          msal,
          `/me/chats/${encodeURIComponent(chatId)}/members?$top=50`,
          ["ChatMember.Read"],
        )
        const members: Participant[] = []
        for (const m of resp.value ?? []) {
          if (!m.displayName) continue
          members.push({
            displayName: m.displayName,
            id: m.userId ?? null,
            kind: "user",
          })
        }
        membersByChat.set(chatId, members)
      } catch (err) {
        console.warn(
          `[m365-pull] Failed to fetch members for chat ${chatId}:`,
          err,
        )
      }
    }
    for (const rec of hits.values()) {
      if (rec.participants.length === 0) {
        const members = membersByChat.get(rec.chatId)
        if (members && members.length > 0) rec.participants = members
      }
    }
  }

  // Phase 5: group recordings into containers by chatId
  const containerMap = new Map<string, RecordingContainer>()
  for (const rec of hits.values()) {
    if (!containerMap.has(rec.chatId)) {
      containerMap.set(rec.chatId, {
        chatId: rec.chatId,
        chatTopic: rec.chatTopic,
        chatType: rec.chatType,
        recordings: [],
        truncated: chatTruncatedSet.has(rec.chatId),
      })
    }
    containerMap.get(rec.chatId)!.recordings.push(rec)
  }

  // Sort recordings within each container newest-first
  for (const container of containerMap.values()) {
    container.recordings.sort((a, b) =>
      b.eventCreatedDateTime.localeCompare(a.eventCreatedDateTime),
    )
  }

  // Sort containers by most-recent recording date descending
  const containers = Array.from(containerMap.values()).sort((a, b) => {
    const aDate = a.recordings[0]?.eventCreatedDateTime ?? ""
    const bDate = b.recordings[0]?.eventCreatedDateTime ?? ""
    return bDate.localeCompare(aDate)
  })

  return {
    containers,
    chatsScanned: recentChats.length,
    truncated,
  }
}

/** Build a human-readable, conflict-resistant filename for a recording's
 * transcript file. Format: <YYYY-MM-DD>-<HHMM>-<subject>-<callIdShort>.<ext>
 *
 * The subject is derived from (in order):
 *   1. chat topic, when set (scheduled meetings, named group chats)
 *   2. "with <other person>" for 1:1 chats (skips self by oid)
 *   3. the raw recording filename, stripped of its Teams-injected
 *      "-YYYYMMDD_HHMMSS" suffix and ".mp4" extension
 *
 * Examples:
 *   2026-05-26-1207-Marc-Catch-Up-6674e488.md       (scheduled meeting)
 *   2026-05-26-1436-with-Diego-Colombo-91703f65.md  (1:1 ad-hoc call)
 *   2026-05-21-0856-Show-and-Tell-MADE-fbc0311f.md  (chat-originated call)
 *
 * The 8-char callId suffix keeps it unique across same-subject calls without
 * dominating the filename. Subject is capped at 60 chars after slugification.
 */
export function buildTranscriptFilename(
  r: RecordingItem,
  userOid: string | null,
  extension: string,
): string {
  const d = new Date(r.eventCreatedDateTime)
  const pad = (n: number) => String(n).padStart(2, "0")
  const dateStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
  const timeStr = `${pad(d.getHours())}${pad(d.getMinutes())}`

  let subject = ""
  if (r.chatTopic && r.chatTopic.trim()) {
    subject = r.chatTopic.trim()
  } else if (r.chatType === "oneOnOne") {
    const userOidLower = userOid ? userOid.toLowerCase() : null
    const other = r.participants.find(
      (p) =>
        p.kind === "user" &&
        p.displayName &&
        !(userOidLower && p.id && p.id.toLowerCase() === userOidLower),
    )
    if (other) subject = `with ${other.displayName}`
  }
  if (!subject) {
    subject = r.filename
      .replace(/\.mp4$/i, "")
      .replace(/[-_]Meeting[-_ ]Recording$/i, "")
      .replace(/-\d{8}_\d{6}(UTC)?$/i, "")
      .trim() ||
      "recording"
  }

  const slug = subject
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60)
    .replace(/-+$/g, "") || "recording"

  const shortId = r.callId.replace(/-/g, "").slice(0, 8) || "unknown"
  const ext = extension.startsWith(".") ? extension : `.${extension}`

  return `${dateStr}-${timeStr}-${slug}-${shortId}${ext}`
}

/** Parse ISO 8601 duration to seconds (best-effort, handles PT?H?M?S). */
function parseDurationSeconds(iso: string): number {
  if (!iso) return 0
  const m = /^PT(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/.exec(
    iso,
  )
  if (!m) return 0
  const h = parseFloat(m[1] ?? "0")
  const min = parseFloat(m[2] ?? "0")
  const s = parseFloat(m[3] ?? "0")
  return h * 3600 + min * 60 + s
}

/** Format seconds as "1h 23m" or "23m" or "45s". */
export function formatDurationShort(iso: string): string {
  const total = Math.round(parseDurationSeconds(iso))
  if (total === 0) return ""
  const h = Math.floor(total / 3600)
  const m = Math.floor((total % 3600) / 60)
  const s = total % 60
  if (h > 0) return `${h}h ${m}m`
  if (m > 0) return `${m}m`
  return `${s}s`
}
