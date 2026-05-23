import type { PublicClientApplication } from "@azure/msal-browser"

// Teams meeting transcripts source.
//
// Graph endpoints used:
//   /me/calendarView                       -> list events in a window
//   /me/onlineMeetings?$filter=joinWebUrl  -> resolve onlineMeeting from event
//   /me/onlineMeetings/{id}/transcripts    -> list transcripts for a meeting
//   /me/onlineMeetings/{id}/transcripts/{id}/content?$format=text/vtt
//
// Scopes used: Calendars.Read, OnlineMeetings.Read,
// OnlineMeetingTranscript.Read.All. All three are already admin-consented.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const LIST_SCOPES = ["Calendars.Read"]
const MEETING_SCOPES = ["OnlineMeetings.Read"]
const TRANSCRIPT_SCOPES = ["OnlineMeetingTranscript.Read.All"]

interface CalendarEvent {
  id: string
  subject: string
  start: { dateTime: string; timeZone?: string }
  end: { dateTime: string; timeZone?: string }
  isOnlineMeeting: boolean
  onlineMeeting?: { joinUrl?: string } | null
  organizer?: {
    emailAddress?: { name?: string; address?: string }
  }
  /** "singleInstance" | "occurrence" | "exception" | "seriesMaster".
   * calendarView returns occurrence/exception for recurring meetings,
   * singleInstance for one-offs. */
  type?: string
  /** Present on occurrence/exception events; points at the series master. */
  seriesMasterId?: string
}

interface CalendarPage {
  value: CalendarEvent[]
  "@odata.nextLink"?: string
}

export interface MeetingItem {
  /** Calendar event id (primary key for the user-facing list). */
  id: string
  subject: string
  startDateTime: string
  endDateTime: string
  organizerName: string
  organizerEmail: string
  joinUrl: string
  /** Meeting organizer's user oid, parsed from the joinUrl's context param.
   * Null if it couldn't be extracted. */
  organizerOid: string | null
  /** True if this event is an occurrence of a recurring series. */
  isRecurring: boolean
  /** Series master event id, when known. Lets us group occurrences. */
  seriesMasterId: string | null
}

/** Extract the organizer's user oid from a Teams joinUrl.
 *
 * Format: `...?context=%7b%22Tid%22%3a%22<tenant>%22%2c%22Oid%22%3a%22<oid>%22%7d`
 * After URL-decoding by the URL constructor, the `context` query param is a
 * JSON string with `Tid` and `Oid` fields.
 */
export function parseOrganizerOid(joinUrl: string): string | null {
  try {
    const url = new URL(joinUrl)
    const contextRaw = url.searchParams.get("context")
    if (!contextRaw) return null
    const context = JSON.parse(contextRaw) as { Oid?: string }
    return context.Oid ?? null
  } catch {
    return null
  }
}

export interface ListMeetingsResult {
  meetings: MeetingItem[]
  nextCursor: string | null
}

interface TranscriptMeta {
  id: string
  createdDateTime: string
}

export interface MeetingTranscriptsPayload {
  source: "teams.transcripts"
  eventId: string
  meetingId: string
  subject: string
  startDateTime: string
  endDateTime: string
  organizerName: string
  organizerEmail: string
  joinUrl: string
  fetchedAt: string
  transcriptCount: number
  transcripts: Array<{
    id: string
    createdDateTime: string
    contentType: string
    content: string
  }>
}

async function getToken(
  msal: PublicClientApplication,
  scopes: string[],
): Promise<string> {
  const account = msal.getActiveAccount()
  if (!account) throw new Error("No active account — sign in first.")
  try {
    const r = await msal.acquireTokenSilent({ account, scopes })
    return r.accessToken
  } catch (err) {
    console.warn("Silent token acquisition failed; redirecting:", err)
    await msal.acquireTokenRedirect({ scopes })
    throw new Error("Redirecting for consent…")
  }
}

async function graphJson<T>(
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
      `Graph ${path}: ${response.status} ${response.statusText} — ${body.slice(0, 300)}`,
    )
  }
  return response.json() as Promise<T>
}

async function graphText(
  msal: PublicClientApplication,
  path: string,
  scopes: string[],
  accept = "text/vtt",
): Promise<string> {
  const token = await getToken(msal, scopes)
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: accept },
  })
  if (!response.ok) {
    const body = await response.text()
    throw new Error(
      `Graph ${path}: ${response.status} ${response.statusText} — ${body.slice(0, 300)}`,
    )
  }
  return response.text()
}

/** List the user's online meetings (calendar events flagged as Teams) in a date
 * window. The cursor is the opaque @odata.nextLink for paging.
 */
export async function listMeetingsPage(
  msal: PublicClientApplication,
  cursor: string | null = null,
  options: { daysBack?: number } = {},
): Promise<ListMeetingsResult> {
  const daysBack = options.daysBack ?? 90
  let path: string
  if (cursor) {
    path = cursor
  } else {
    const startTime = new Date(
      Date.now() - daysBack * 24 * 60 * 60 * 1000,
    ).toISOString()
    const endTime = new Date().toISOString()
    const select =
      "id,subject,start,end,isOnlineMeeting,onlineMeeting,organizer,type,seriesMasterId"
    const params = new URLSearchParams({
      startDateTime: startTime,
      endDateTime: endTime,
      $top: "50",
      $select: select,
      $orderby: "start/dateTime desc",
    })
    path = `/me/calendarView?${params.toString()}`
  }

  const data = await graphJson<CalendarPage>(msal, path, LIST_SCOPES)
  // Filter on joinUrl presence only -- isOnlineMeeting is unreliable
  // (false for some events that genuinely have Teams meetings attached).
  const meetings: MeetingItem[] = data.value
    .filter((e) => Boolean(e.onlineMeeting?.joinUrl))
    .map((e) => {
      const joinUrl = e.onlineMeeting!.joinUrl!
      const isRecurring =
        e.type === "occurrence" ||
        e.type === "exception" ||
        Boolean(e.seriesMasterId)
      return {
        id: e.id,
        subject: e.subject || "(no subject)",
        startDateTime: e.start.dateTime,
        endDateTime: e.end.dateTime,
        organizerName: e.organizer?.emailAddress?.name ?? "",
        organizerEmail: e.organizer?.emailAddress?.address ?? "",
        joinUrl,
        organizerOid: parseOrganizerOid(joinUrl),
        isRecurring,
        seriesMasterId: e.seriesMasterId ?? null,
      }
    })
  return {
    meetings,
    nextCursor: data["@odata.nextLink"] ?? null,
  }
}

export interface FetchTranscriptsProgress {
  count: number
  total: number
  detail: string
}

/** Fetch all transcripts for a meeting (selected by calendar event/MeetingItem).
 *
 * Multi-step Graph dance:
 *   1. Resolve onlineMeeting id from joinUrl (filter on joinWebUrl).
 *   2. List transcripts for that meeting.
 *   3. Fetch VTT content for each transcript.
 *
 * Returns a payload regardless of whether transcripts exist (with empty list
 * if none). Throws only on Graph errors that prevent any progress.
 */
export async function fetchMeetingTranscripts(
  msal: PublicClientApplication,
  meeting: MeetingItem,
  options: {
    onProgress?: (progress: FetchTranscriptsProgress) => void
  } = {},
): Promise<MeetingTranscriptsPayload> {
  // Pre-flight: Microsoft Graph delegated permissions do not allow looking up
  // someone else's online meeting via /me/onlineMeetings, and the cross-user
  // /users/{org}/onlineMeetings endpoint requires OnlineMeetings.Read.All
  // (application-only — no delegated form exists). So transcripts via this
  // SPA are constrained to meetings the signed-in user organized.
  const account = msal.getActiveAccount()
  const userOid = account?.localAccountId ?? null
  if (
    meeting.organizerOid &&
    userOid &&
    meeting.organizerOid.toLowerCase() !== userOid.toLowerCase()
  ) {
    throw new Error(
      "Transcripts can only be downloaded for meetings you organized. " +
        "Microsoft Graph delegated permissions don't expose other users' meeting metadata.",
    )
  }

  options.onProgress?.({ count: 0, total: 0, detail: "Locating online meeting…" })

  // 1. Resolve onlineMeeting from joinUrl
  const filterValue = `joinWebUrl eq '${meeting.joinUrl.replace(/'/g, "''")}'`
  const meetingResp = await graphJson<{ value: Array<{ id: string }> }>(
    msal,
    `/me/onlineMeetings?$filter=${encodeURIComponent(filterValue)}`,
    MEETING_SCOPES,
  )
  const onlineMeeting = meetingResp.value?.[0]
  if (!onlineMeeting) {
    throw new Error(
      "No online meeting found for this calendar event — it may have been deleted in Teams.",
    )
  }
  const meetingId = onlineMeeting.id

  // 2. List transcripts for the meeting
  options.onProgress?.({ count: 0, total: 0, detail: "Listing transcripts…" })
  const transcriptsResp = await graphJson<{ value: TranscriptMeta[] }>(
    msal,
    `/me/onlineMeetings/${meetingId}/transcripts`,
    [...MEETING_SCOPES, ...TRANSCRIPT_SCOPES],
  )
  const transcriptMetas = transcriptsResp.value ?? []

  const baseShape = {
    source: "teams.transcripts" as const,
    eventId: meeting.id,
    meetingId,
    subject: meeting.subject,
    startDateTime: meeting.startDateTime,
    endDateTime: meeting.endDateTime,
    organizerName: meeting.organizerName,
    organizerEmail: meeting.organizerEmail,
    joinUrl: meeting.joinUrl,
    fetchedAt: new Date().toISOString(),
  }

  if (transcriptMetas.length === 0) {
    return { ...baseShape, transcriptCount: 0, transcripts: [] }
  }

  // 3. Fetch each transcript's VTT content
  const transcripts: MeetingTranscriptsPayload["transcripts"] = []
  for (let i = 0; i < transcriptMetas.length; i++) {
    const t = transcriptMetas[i]
    options.onProgress?.({
      count: i,
      total: transcriptMetas.length,
      detail: `Fetching transcript ${i + 1} of ${transcriptMetas.length}`,
    })
    const content = await graphText(
      msal,
      `/me/onlineMeetings/${meetingId}/transcripts/${t.id}/content?$format=text/vtt`,
      [...MEETING_SCOPES, ...TRANSCRIPT_SCOPES],
      "text/vtt",
    )
    transcripts.push({
      id: t.id,
      createdDateTime: t.createdDateTime,
      contentType: "text/vtt",
      content,
    })
  }
  options.onProgress?.({
    count: transcripts.length,
    total: transcripts.length,
    detail: "Done.",
  })

  return { ...baseShape, transcriptCount: transcripts.length, transcripts }
}

export function meetingDisplayName(m: MeetingItem): string {
  return m.subject || "(no subject)"
}
