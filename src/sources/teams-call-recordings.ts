// Call recordings source -- discovers recordings across all Teams chats by
// walking /me/chats and extracting callRecordingEventMessageDetail events.
//
// This replaces the old calendar-based teams-transcripts.ts source, which
// couldn't see recordings from 1:1 calls, chat-originated calls, or recordings
// in other users' OneDrives (anything where the user wasn't the meeting
// organizer). Chat events surface every recording the user has access to,
// with a direct SharePoint URL.
//
// Transcript fetching uses the same SharePoint REST path as teams-recordings.ts
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

export interface RecordingsResult {
  recordings: RecordingItem[]
  /** Number of chats whose messages we actually scanned. */
  chatsScanned: number
  /** Total chats we paged through (including ones outside the window). */
  chatsTotal: number
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

async function graphJson<T>(
  msal: PublicClientApplication,
  path: string,
  scopes: string[],
): Promise<T> {
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`
  const token = await getGraphToken(msal, scopes)
  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!resp.ok) {
    const body = await resp.text()
    throw new Error(
      `Graph ${path}: ${resp.status} ${resp.statusText} \u2014 ${body.slice(0, 300)}`,
    )
  }
  return resp.json() as Promise<T>
}

export interface ListRecordingsOptions {
  /** Date window: only chats with activity in the last N days are scanned. */
  daysBack: number
  /** Cap chat pagination. Default 10 pages of 50 = up to 500 chats. */
  maxChatPages?: number
  /** Progress callback for UI status updates. */
  onProgress?: (note: string) => void
}

/**
 * Discover all call recordings the user has access to within the given date window.
 *
 * Algorithm (proven via Probe F in the recording-source investigation spike):
 *   1. Page through /me/chats. Stop early once the oldest chat in a page is
 *      older than the window (chats are returned by lastUpdatedDateTime desc).
 *   2. Filter to chats with recent activity (lastUpdatedDateTime >= window start).
 *   3. For each such chat, fetch the most recent 50 messages.
 *   4. Extract messages where eventDetail.@odata.type is
 *      "#microsoft.graph.callRecordingEventMessageDetail" and status is "success".
 *   5. Deduplicate by (callId, filename).
 *   6. Sort by event date descending.
 *
 * Catches scheduled meetings, 1:1 calls, group-chat calls, multi-recording calls,
 * and recordings stored in others' OneDrives. No calendar lookup, no search.
 */
export async function listRecordings(
  msal: PublicClientApplication,
  options: ListRecordingsOptions,
): Promise<RecordingsResult> {
  const { daysBack, maxChatPages = 10, onProgress } = options
  const sinceMs = Date.now() - daysBack * 24 * 60 * 60 * 1000

  // Phase 1: page through chats
  let chatsPath: string | null = "/me/chats?$top=50"
  const allChats: ChatListItem[] = []
  let pages = 0
  while (chatsPath && pages < maxChatPages) {
    onProgress?.(`Listing chats (page ${pages + 1})\u2026`)
    const page: { value: ChatListItem[]; "@odata.nextLink"?: string } =
      await graphJson(msal, chatsPath, SCOPES)
    allChats.push(...page.value)
    const oldest = page.value[page.value.length - 1]?.lastUpdatedDateTime
    if (oldest && new Date(oldest).getTime() < sinceMs) break
    chatsPath = page["@odata.nextLink"] ?? null
    pages++
  }
  const truncated = chatsPath !== null && pages >= maxChatPages

  const recentChats = allChats.filter((c) => {
    if (!c.lastUpdatedDateTime) return false
    return new Date(c.lastUpdatedDateTime).getTime() >= sinceMs
  })

  // Phase 2: scan messages in each recent chat. Collect both:
  //   - callRecordingEventMessageDetail with status:"success" -> recordings
  //   - callEndedEventMessageDetail                            -> participants by callId
  // Join client-side -- a single chat carries both event types for a given call.
  const hits = new Map<string, RecordingItem>()
  const participantsByCallId = new Map<string, Participant[]>()
  for (let i = 0; i < recentChats.length; i++) {
    const chat = recentChats[i]
    onProgress?.(`Scanning chat ${i + 1}/${recentChats.length}\u2026`)
    try {
      const msgs: { value: ChatMessage[] } = await graphJson(
        msal,
        `/me/chats/${encodeURIComponent(chat.id)}/messages?$top=50`,
        SCOPES,
      )
      for (const m of msgs.value) {
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
          // Latest participants list wins (later messages overwrite earlier).
          // For the same call, all callEnded events should carry the same list,
          // but we keep the last-seen one to be defensive about Teams variations.
          const parts = extractParticipants(ed.callParticipants)
          if (parts.length > 0) {
            participantsByCallId.set(ed.callId, parts)
          }
        }
      }
    } catch (err) {
      // Bad chat shouldn't kill the whole scan. Log and continue.
      console.warn(`[m365-pull] Failed to scan chat ${chat.id}:`, err)
    }
  }

  // Phase 3: join participants into recordings by callId
  for (const rec of hits.values()) {
    const parts = participantsByCallId.get(rec.callId)
    if (parts) rec.participants = parts
  }

  // Phase 4: fallback for recordings without participants. 1:1 ad-hoc calls
  // (not scheduled meetings) don't emit callEndedEventMessageDetail in their
  // chat, so the join above produces nothing. Fall back to chat membership:
  //   - For 1:1 chats this is exact (2 members = the user + 1 other)
  //   - For group chats it's a superset of call attendance, but honest
  //   - Members fetched once per unique chatId, attached to all that chat's recordings
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

  const recordings = Array.from(hits.values()).sort((a, b) =>
    b.eventCreatedDateTime.localeCompare(a.eventCreatedDateTime),
  )

  return {
    recordings,
    chatsScanned: recentChats.length,
    chatsTotal: allChats.length,
    truncated,
  }
}

/** Parse ISO 8601 duration to seconds (best-effort, handles PT?H?M?S). */
export function parseDurationSeconds(iso: string): number {
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
