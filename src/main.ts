import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser"
import { config } from "./config"
import "./style.css"
import {
  listChatsPage,
  fetchChatById,
  fetchChatMessages,
  chatDisplayName,
  buildChatArchiveFilename,
  type TeamsChatItem,
  type TeamsChatMessage,
} from "./sources/teams-chats"
import { formatDateStamp } from "./sources/filename-format"
import {
  listRecordings,
  buildTranscriptFilename,
  type RecordingItem,
  type RecordingContainer,
} from "./sources/teams-call-recordings"
import {
  resolveRecordingFromUrl,
  fetchRecordingTranscripts,
} from "./sources/teams-recordings"
import { vttToMarkdown } from "./format/transcript-markdown"
import { renderChatMarkdown } from "./format/chat-markdown"
import { saveAsText } from "./destinations/browser"
import { saveTextToOneDrive, getOneDriveFolderWebUrl } from "./destinations/onedrive"
import {
  loadUserPrefs,
  saveUserPrefs,
  type UserPrefs,
  type Destination,
  type ChatRange,
} from "./cache/prefs"
import {
  loadCachedChats,
  saveCachedChats,
  clearCachedChats,
  ageMs,
  formatAge,
} from "./cache/chats-cache"
import { loadMarks, saveMarks } from "./cache/marks"
import { loadIgnored, saveIgnored } from "./cache/ignored"
import { loadChatPrefs, saveChatPrefs } from "./cache/chat-prefs"
import { loadRecordingPrefs, saveRecordingPrefs } from "./cache/recording-prefs"
import { loadUIState, saveUIState } from "./cache/ui-state"
import {
  loadOneDriveState,
  saveOneDriveState,
  mergeStates,
  deviceIdentifier,
  type AppState,
  type ChatPrefs,
  type RecordingPrefs,
} from "./state/onedrive-state"

const app = document.getElementById("app") as HTMLDivElement

if (!config.clientId || !config.tenantId) {
  app.innerHTML =
    '<p class="empty">Missing MSAL config \u2014 set <code>clientId</code> and <code>tenantId</code> in <code>src/config.ts</code>. See README.md.</p>'
  throw new Error("Missing MSAL config")
}

const msal = new PublicClientApplication({
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
    redirectUri: window.location.origin + "/auth-callback",
  },
})

await msal.initialize()
const redirectResp = await msal.handleRedirectPromise()
if (redirectResp?.account) msal.setActiveAccount(redirectResp.account)

// MSAL returns the auth code to the anonymous /auth-callback route, which is kept OUTSIDE the
// EasyAuth `authenticated` gate so the #code fragment survives the round-trip (the gate would
// otherwise 302 the return to /.auth/login and strip the fragment -> login loop). Once the
// account is established in the MSAL cache here, route into the real app at /.
if (window.location.pathname === "/auth-callback") {
  window.location.replace("/")
}

if (!msal.getActiveAccount()) {
  const cached = msal.getAllAccounts()
  if (cached.length > 0) {
    // Returning visit \u2014 reuse the account MSAL already cached (sync, instant).
    msal.setActiveAccount(cached[0])
  }
}

// ----- State -----

interface ChatsState {
  chats: TeamsChatItem[]
}

interface RecordingsState {
  containers: RecordingContainer[]
  chatsScanned: number
  truncated: boolean
}

type SortKey = "marked-first" | "recent" | "name"

interface FilterState {
  search: string
  enabledTypes: Set<string>
  sortKey: SortKey
  markedOnly: boolean
  showIgnored: boolean
}

type SyncStatus = "idle" | "syncing" | "synced" | "error" | "offline"

const KNOWN_TYPES: { id: string; label: string }[] = [
  { id: "oneOnOne", label: "1:1" },
  { id: "group", label: "Group" },
  { id: "meeting", label: "Meeting" },
]

const SIGNIN_SCOPES = [
  "User.Read",
  "Chat.Read",
  "Files.Read.All",
  "Files.ReadWrite",
  "Files.ReadWrite.AppFolder",
]

let chatsState: ChatsState = { chats: [] }
let recordingsState: RecordingsState = { containers: [], chatsScanned: 0, truncated: false }
/** Fast lookup: chatId \u2192 RecordingContainer (populated by background recordings scan). */
let recordingsMap: Map<string, RecordingContainer> = new Map()
/** True while the background recordings scan is running; rows show pending indicator. */
/** True after the background recordings scan has completed (success or error); distinguishes
 * "no entry yet — still checking" from "no entry — scan done, couldn't determine". */

let filterState: FilterState = {
  search: "",
  enabledTypes: new Set(KNOWN_TYPES.map((t) => t.id)),
  sortKey: "marked-first",
  markedOnly: false,
  showIgnored: false,
}
let markedIds: Set<string> = new Set()
let ignoredIds: Set<string> = new Set()
let chatPrefs: Record<string, ChatPrefs> = {}
let recordingPrefs: Record<string, RecordingPrefs> = {}
let userPrefs: UserPrefs = {
  destination: "browser",
  oneDriveFolder: "/m365-pull/teams-chats",
}
let syncStatus: SyncStatus = "idle"
let lastSyncedAt: Date | null = null
let lastSyncError: string | null = null

/** Ephemeral artifact selection — NOT synced, NOT in OneDrive AppState.
 * Key format: "msg:{chatId}" for Messages, "rec:{rec.id}" for Recordings. */
const selectedArtifacts: Set<string> = new Set()
/** Ephemeral expand state — NOT persisted across page reloads. */
const expandedChatIds: Set<string> = new Set()

// ----- Progressive-disclosure toolbar state (3-section toolbar, 2026-06-18) -----

/** True once a load has completed at least once. Gates "reload on setting change"
 * (replaces the old `!refreshchats.hidden` signal) and post-load section reveal. */
let hasLoadedOnce = false
/** Last recording-scan window, kept so the checklist's recordings Retry can re-run. */
let lastRecordingWindow: { fromMs: number; toMs: number } | null = null

// Build B: Include is a LOAD-SCOPE. The live §1 chips write userPrefs.includeMessages/
// includeRecordings, but §1 collapses to a receipt after load, so the scope can only
// change via the picker (pre-load, or Change→re-pick→Load). We CAPTURE the chip state
// at load time into these module vars; render/filter/action logic reads THESE (not the
// live chip state) so the listed experience always matches what the load actually pulled.
// Defaults both-on so any pre-load render matches today's behavior.
let loadIncludeMessages = true
let loadIncludeRecordings = true

type LoadStepState = "pending" | "running" | "done" | "error"
interface LoadStep {
  id: "signin" | "chats" | "recordings" | "done"
  label: string
  detail?: string
  state: LoadStepState
  error?: string
}
/** Live checklist model shown in §1 during a load. */
let loadSteps: LoadStep[] = []

// ----- Helpers -----

function el<T extends HTMLElement>(id: string): T {
  const e = document.getElementById(id)
  if (!e) throw new Error(`#${id} not in DOM`)
  return e as T
}

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"]/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!,
  )
}

function formatDate(iso: string): string {
  if (!iso) return ""
  return new Date(iso).toLocaleString()
}

function formatDateShort(d: Date): string {
  return d.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
  })
}

/** Format an ISO 8601 duration (e.g. "PT15M1.826S") into a readable string like "15m 1s". */
function formatIsoDuration(iso: string): string {
  const m = iso.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?/)
  if (!m) return iso
  const h = parseInt(m[1] || "0", 10)
  const min = parseInt(m[2] || "0", 10)
  const sec = Math.floor(parseFloat(m[3] || "0"))
  const parts: string[] = []
  if (h > 0) parts.push(`${h}h`)
  if (min > 0) parts.push(`${min}m`)
  if (sec > 0 || parts.length === 0) parts.push(`${sec}s`)
  return parts.join(" ")
}

/** Artifact IDs for a chat, SCOPED to the captured load-time Include scope:
 * "msg:{chatId}" (only when Messages is in scope) + "rec:{rec.id}" for each
 * confirmed recording (only when Recordings is in scope). Select-all, the global
 * select-all count, and per-chat selection all read this, so a suppressed type
 * is never selectable. */
function getArtifactIds(chatId: string, recContainer: RecordingContainer | undefined): string[] {
  const ids: string[] = []
  if (loadIncludeMessages) ids.push(`msg:${chatId}`)
  if (loadIncludeRecordings && recContainer && recContainer.recordings.length > 0) {
    recContainer.recordings.forEach((r) => ids.push(`rec:${r.id}`))
  }
  return ids
}

// ----- Phase 2 favorites: per-stream keys over the existing synced `marks` set -----
//
// A chat has up to TWO favorite targets, both stored in the SAME `markedIds`
// string set (so the additive union merge in onedrive-state.ts keeps working):
//   - messages stream   -> bare `chatId`            (matches the legacy v1 key)
//   - recordings stream  -> `${chatId}::rec`         (collision-safe suffix)
// A bare chatId never contains "::", so the two keyspaces never overlap, and a
// recordings-stream key can never be mistaken for the old per-recording
// `callId::filename` marks (those are dropped on migration).

const REC_STREAM_SUFFIX = "::rec"

/** localStorage key tracking whether the Phase 2 favorites migration has run. */
function favMigrationKey(userKey: string): string {
  return `m365-pull.favver.${userKey}`
}
function favSchemaVersion(userKey: string): number {
  try {
    return parseInt(localStorage.getItem(favMigrationKey(userKey)) || "1", 10) || 1
  } catch {
    return 1
  }
}
function setFavSchemaVersion(userKey: string, v: number): void {
  try {
    localStorage.setItem(favMigrationKey(userKey), String(v))
  } catch {
    /* ignore quota / disabled storage */
  }
}

function recStreamKey(chatId: string): string {
  return `${chatId}${REC_STREAM_SUFFIX}`
}
function isMessagesFavorited(chatId: string): boolean {
  return markedIds.has(chatId)
}
function isRecordingsFavorited(chatId: string): boolean {
  return markedIds.has(recStreamKey(chatId))
}
/** True when EITHER stream of this chat is favorited (drives collapsed ★ state). */
function isChatFavorited(chatId: string): boolean {
  return isMessagesFavorited(chatId) || isRecordingsFavorited(chatId)
}
/** Map any favorite stream key back to its owning chatId. */
function favoriteKeyToChatId(key: string): string {
  return key.endsWith(REC_STREAM_SUFFIX) ? key.slice(0, -REC_STREAM_SUFFIX.length) : key
}
/** Distinct chatIds that have at least one favorited stream. */
function favoritedChatIds(): Set<string> {
  return new Set([...markedIds].map(favoriteKeyToChatId))
}

/** Migrate a marks set in place from legacy (v1, whole-container bare-chatId) to
 * Phase 2 (v2, per-stream). For each legacy bare-chatId keep, ALSO favorite the
 * recordings stream (chatId::rec) so nothing the user previously kept is lost.
 * Drops the truly-legacy per-recording `callId::filename` and `rec:` marks.
 * Idempotent while marks are still in their pre-toggle (legacy) shape; gated by
 * the caller so it isn't re-run after the user starts toggling streams. */
function migrateMarksToStreams(marks: Set<string>): boolean {
  let changed = false
  for (const id of [...marks]) {
    if (id.endsWith(REC_STREAM_SUFFIX)) {
      continue // already a Phase 2 recordings-stream key
    }
    if (id.startsWith("rec:")) {
      // Pre-Phase-1 prefixed recording-container mark -> normalize to both streams.
      const bare = id.slice(4)
      marks.delete(id)
      marks.add(bare)
      marks.add(recStreamKey(bare))
      changed = true
      continue
    }
    if (id.includes("::")) {
      // Legacy per-recording mark (callId::filename) from a much older build -> drop.
      marks.delete(id)
      changed = true
      continue
    }
    // Bare chatId = legacy whole-container keep -> favorite the recordings stream too.
    if (!marks.has(recStreamKey(id))) {
      marks.add(recStreamKey(id))
      changed = true
    }
  }
  return changed
}

/** True iff a preview is a system event that represents REAL meeting/call
 * activity (callRecording / callEnded / callStarted / callTranscript), as
 * opposed to roster/membership churn (members*, *Renamed, *Pinned, etc.).
 *
 * Graph stamps the same `lastMessagePreview` slot for both kinds of system
 * event, but only the call/meeting ones reflect real activity. The eventDetail
 * `@odata.type` (e.g. `#microsoft.graph.callRecordingEventMessageDetail`)
 * carries the discriminator \u2014 we match the substring "call" case-insensitively.
 *
 * Defensive: returns false when `eventDetail`/`@odata.type` is absent, so a
 * chat with no usable discriminator degrades to the createdDateTime sentinel
 * (current behaviour \u2014 never crashes, never falsely counts roster churn). */
function isCallActivityEvent(preview: TeamsChatItem["lastMessagePreview"]): boolean {
  const odataType = preview?.eventDetail?.["@odata.type"]
  return typeof odataType === "string" && odataType.toLowerCase().includes("call")
}

/** Derive the "last real conversation activity" timestamp (ms) for a chat. */
function chatActivityDate(chat: TeamsChatItem): number {
  const preview = chat.lastMessagePreview
  const isRealMessage = preview?.messageType === "message"
  if ((isRealMessage || isCallActivityEvent(preview)) && preview?.createdDateTime) {
    return new Date(preview.createdDateTime).getTime()
  }
  return new Date(chat.createdDateTime).getTime()
}

// ----- Range helpers -----

const DAY_MS = 24 * 60 * 60 * 1000

/** Format a Date as "yyyy-mm-dd" in local time for date inputs. */
function toLocalDateString(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, "0")
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
}

/** Midnight (00:00:00.000) of the most recent Monday in local time. */
function startOfThisWeekMonday(): number {
  const now = new Date()
  const dayOfWeek = now.getDay() // 0=Sun, 1=Mon, \u2026 6=Sat
  const daysSinceMonday = (dayOfWeek + 6) % 7
  return new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysSinceMonday).getTime()
}

/** Midnight (00:00:00.000) of a yyyy-mm-dd date in local time. */
function startOfLocalDay(yyyyMmDd: string): number {
  const [y, m, d] = yyyyMmDd.split("-").map(Number)
  return new Date(y, m - 1, d, 0, 0, 0, 0).getTime()
}

/** 23:59:59.999 of a yyyy-mm-dd date in local time. */
function endOfLocalDay(yyyyMmDd: string): number {
  const [y, m, d] = yyyyMmDd.split("-").map(Number)
  return new Date(y, m - 1, d, 23, 59, 59, 999).getTime()
}


/** Most recent chatPrefs.lastSync timestamp across all synced chats, or null. */
function mostRecentChatSync(prefs: Record<string, { lastSync?: string }>): number | null {
  let most: number | null = null
  for (const v of Object.values(prefs)) {
    if (v.lastSync) {
      const t = Date.parse(v.lastSync)
      if (!isNaN(t) && (most === null || t > most)) most = t
    }
  }
  return most
}


/** Map a persisted ChatRange to a concrete [cutoffMs, untilMs] window.
 *  For since-last-download with no prior sync, falls back to last-7d. */
function computeChatWindow(
  range: ChatRange,
  prefs: Record<string, { lastSync?: string }>,
): { cutoffMs: number; untilMs: number } {
  const now = Date.now()
  switch (range.kind) {
    case "this-week":
      return { cutoffMs: startOfThisWeekMonday(), untilMs: now }
    case "last-7d":
      return { cutoffMs: now - 7 * DAY_MS, untilMs: now }
    case "last-30d":
      return { cutoffMs: now - 30 * DAY_MS, untilMs: now }
    case "since-last-download": {
      const t = mostRecentChatSync(prefs)
      return { cutoffMs: t !== null ? t : now - 7 * DAY_MS, untilMs: now }
    }
    case "custom": {
      const cutoffMs = range.customFrom ? startOfLocalDay(range.customFrom) : now - 7 * DAY_MS
      const untilMs = range.customTo ? endOfLocalDay(range.customTo) : now
      return { cutoffMs, untilMs }
    }
  }
}

/** Human-readable label for a ChatRange (used in status messages). */
function chatRangeLabel(range: ChatRange, prefs: Record<string, { lastSync?: string }>): string {
  switch (range.kind) {
    case "this-week": return "this week"
    case "last-7d": return "last 7 days"
    case "last-30d": return "last 30 days"
    case "since-last-download": {
      const t = mostRecentChatSync(prefs)
      if (t === null) return "since last download (none yet \u2014 showing 7 days)"
      return `since last download (${formatDateShort(new Date(t))})`
    }
    case "custom":
      if (range.customFrom && range.customTo) return `${range.customFrom} to ${range.customTo}`
      if (range.customFrom) return `from ${range.customFrom}`
      return "custom range"
  }
}

function setStatus(text: string, kind: "info" | "error" = "info"): void {
  const status = el<HTMLDivElement>("status")
  status.textContent = text
  status.className = kind === "error" ? "error" : ""
}

/** Update the quiet persistent scan-progress indicator (separate from the transient status line). */
function setScanStatus(text: string): void {
  const scanEl = document.getElementById("scan-status")
  if (scanEl) scanEl.textContent = text
}

function userCacheKey(): string {
  const account = msal.getActiveAccount()
  if (!account) return "anon"
  return account.localAccountId || account.homeAccountId || account.username
}

function syncUserPrefsToUI(): void {
  const destinationSel = document.getElementById("destination") as
    | HTMLSelectElement
    | null
  if (destinationSel) destinationSel.value = userPrefs.destination
  const folderEl = document.getElementById("onedrive-folder") as
    | HTMLSpanElement
    | null
  const editBtn = document.getElementById("edit-folder") as
    | HTMLButtonElement
    | null
  if (folderEl) folderEl.hidden = userPrefs.destination !== "onedrive"
  if (editBtn) editBtn.textContent = userPrefs.oneDriveFolder
  // Populate the "\u2197 Open" link to the destination folder in OneDrive on the web.
  void refreshOneDriveFolderLink()

  // Chat range
  const chatRange: ChatRange = userPrefs.chatRange ?? { kind: "last-7d" }
  const chatRangeEl = document.getElementById("chat-range") as HTMLSelectElement | null
  if (chatRangeEl) chatRangeEl.value = chatRange.kind
  // Custom date inputs are render-when-custom (renderCustomRangeField reads the
  // current selection + stored customFrom/To and (re)builds the inputs inline).
  renderCustomRangeField()

  // Chat marked-include toggle (default ON: undefined \u2192 true)
  const markedIncludeBtn = document.getElementById("marked-include") as HTMLButtonElement | null
  if (markedIncludeBtn)
    markedIncludeBtn.classList.toggle("active", userPrefs.markedInclude !== false)

  // Include-messages toggle (default ON)
  const includeMessagesBtn = document.getElementById("include-messages") as HTMLButtonElement | null
  if (includeMessagesBtn)
    includeMessagesBtn.classList.toggle("active", userPrefs.includeMessages !== false)

  // Include-recordings toggle (default ON)
  const includeRecordingsBtn = document.getElementById("include-recordings") as HTMLButtonElement | null
  if (includeRecordingsBtn)
    includeRecordingsBtn.classList.toggle("active", userPrefs.includeRecordings !== false)

  // Hide-downloaded toggle
  const hideBtn = document.getElementById("hide-downloaded") as HTMLButtonElement | null
  if (hideBtn) hideBtn.classList.toggle("active", !!userPrefs.hideDownloaded)

  // Change 3: Grouped/Flat view toggle (default "flat" when unset)
  const viewMode = userPrefs.viewMode === "grouped" ? "grouped" : "flat"
  const viewGroupedBtn = document.getElementById("view-grouped") as HTMLButtonElement | null
  if (viewGroupedBtn) viewGroupedBtn.classList.toggle("active", viewMode === "grouped")
  const viewFlatBtn = document.getElementById("view-flat") as HTMLButtonElement | null
  if (viewFlatBtn) viewFlatBtn.classList.toggle("active", viewMode === "flat")
}

/** Populate the "\u2197 Open" link next to the OneDrive folder path. */
async function refreshOneDriveFolderLink(): Promise<void> {
  const link = document.getElementById("open-folder") as HTMLAnchorElement | null
  if (!link) return
  if (userPrefs.destination !== "onedrive") {
    link.hidden = true
    return
  }
  const webUrl = await getOneDriveFolderWebUrl(msal, userPrefs.oneDriveFolder)
  if (webUrl) {
    link.href = webUrl
    link.hidden = false
  } else {
    link.hidden = true
  }
}

function typeLabel(type: string): string {
  return KNOWN_TYPES.find((t) => t.id === type)?.label ?? type
}

function buildLocalState(): AppState {
  return {
    version: 2,
    updatedAt: new Date().toISOString(),
    updatedBy: deviceIdentifier(),
    marks: [...markedIds].sort(),
    ignored: ignoredIds.size > 0 ? [...ignoredIds].sort() : undefined,
    chatPrefs: Object.keys(chatPrefs).length > 0 ? chatPrefs : undefined,
    recordingPrefs:
      Object.keys(recordingPrefs).length > 0 ? recordingPrefs : undefined,
    userPrefs,
  }
}

function updateSyncIndicator(): void {
  const badge = document.getElementById("sync-badge")
  if (!badge) return
  let text = ""
  let title = ""
  let cls = "sync-badge"
  switch (syncStatus) {
    case "idle":
      text = "not synced"
      cls += " idle"
      title = "Sign in to sync favorites across devices"
      break
    case "syncing":
      text = "syncing\u2026"
      cls += " syncing"
      title = "Saving state to OneDrive"
      break
    case "synced":
      text = lastSyncedAt
        ? `synced ${formatAge(Date.now() - lastSyncedAt.getTime())}`
        : "synced"
      cls += " synced"
      title = "Favorites are saved to OneDrive (/Apps/m365-pull/state.json)"
      break
    case "error":
      text = "sync error"
      cls += " error"
      title = lastSyncError ?? "Failed to sync state \u2014 using local only"
      break
    case "offline":
      text = "offline \u00b7 local only"
      cls += " offline"
      title = "Could not reach OneDrive; favorites stay on this device"
      break
  }
  badge.textContent = text
  badge.className = cls
  badge.title = title
}

// ----- OneDrive sync -----

let saveTimer: number | null = null
let savePromise: Promise<void> | null = null
const SAVE_DEBOUNCE_MS = 1500

function scheduleOneDriveSave(): void {
  if (saveTimer !== null) window.clearTimeout(saveTimer)
  saveTimer = window.setTimeout(() => {
    saveTimer = null
    void flushOneDriveSave()
  }, SAVE_DEBOUNCE_MS)
}

async function flushOneDriveSave(): Promise<void> {
  if (savePromise) {
    try {
      await savePromise
    } catch {
      /* fall through */
    }
  }
  syncStatus = "syncing"
  updateSyncIndicator()
  const state = buildLocalState()
  savePromise = saveOneDriveState(msal, state).then(
    () => {
      syncStatus = "synced"
      lastSyncedAt = new Date()
      lastSyncError = null
      updateSyncIndicator()
    },
    (err: Error) => {
      syncStatus = "error"
      lastSyncError = err.message
      console.warn("OneDrive save failed:", err)
      updateSyncIndicator()
    },
  )
  await savePromise
  savePromise = null
}

async function pullAndMergeOneDriveState(): Promise<void> {
  syncStatus = "syncing"
  updateSyncIndicator()
  try {
    const remote = await loadOneDriveState(msal)
    const local = buildLocalState()
    const merged = mergeStates(local, remote)
    const mergedMarks = new Set(merged.marks)
    // Phase 2 migration for marks arriving from a legacy (v1) remote state: a v1
    // remote holds only whole-container bare-chatId keeps, so promoting each to
    // favorite BOTH streams is safe and one-time (mergeStates already stamps the
    // written-back state as v2, so this path never re-fires once OneDrive is v2).
    if (remote?.version === 1) {
      migrateMarksToStreams(mergedMarks)
      // Persist the migrated marks back through `merged` so the OneDrive writeback
      // below stores v2 with BOTH streams — not the stale pre-migration set (which
      // would silently drop the recordings-stream favorites until the next save).
      merged.marks = [...mergedMarks].sort()
    }
    const marksChanged =
      mergedMarks.size !== markedIds.size ||
      [...mergedMarks].some((id) => !markedIds.has(id))
    markedIds = mergedMarks
    saveMarks(userCacheKey(), markedIds)
    const mergedIgnored = new Set(merged.ignored ?? [])
    const ignoredChanged =
      mergedIgnored.size !== ignoredIds.size ||
      [...mergedIgnored].some((id) => !ignoredIds.has(id))
    ignoredIds = mergedIgnored
    saveIgnored(userCacheKey(), ignoredIds)
    const mergedPrefs = merged.chatPrefs ?? {}
    const prefsChanged =
      JSON.stringify(mergedPrefs) !== JSON.stringify(chatPrefs)
    chatPrefs = mergedPrefs
    saveChatPrefs(userCacheKey(), chatPrefs)
    const mergedRecordingPrefs = merged.recordingPrefs ?? {}
    const recordingPrefsChanged =
      JSON.stringify(mergedRecordingPrefs) !== JSON.stringify(recordingPrefs)
    recordingPrefs = mergedRecordingPrefs
    saveRecordingPrefs(userCacheKey(), recordingPrefs)
    let userPrefsChanged = false
    if (merged.userPrefs) {
      userPrefsChanged =
        JSON.stringify(merged.userPrefs) !== JSON.stringify(userPrefs)
      userPrefs = merged.userPrefs
      saveUserPrefs(userCacheKey(), userPrefs)
      if (userPrefsChanged) syncUserPrefsToUI()
    }
    const changed =
      marksChanged || ignoredChanged || prefsChanged || recordingPrefsChanged || userPrefsChanged
    if (changed) {
      await saveOneDriveState(msal, {
        ...merged,
        updatedAt: new Date().toISOString(),
        updatedBy: deviceIdentifier(),
      })
      rerenderContainerList()
    } else if (!remote) {
      await saveOneDriveState(msal, local)
    }
    syncStatus = "synced"
    lastSyncedAt = new Date()
    lastSyncError = null
  } catch (err) {
    syncStatus = "error"
    lastSyncError = (err as Error).message
    console.warn("OneDrive sync failed:", err)
  } finally {
    updateSyncIndicator()
  }
}

// ----- Filter + sort for containers (pure) -----

function applyContainerFiltersAndSort(chats: TeamsChatItem[]): TeamsChatItem[] {
  const q = filterState.search.trim().toLowerCase()
  let result = chats.filter((c) => {
    const isIgnored = ignoredIds.has(c.id)
    if (filterState.showIgnored) {
      // "Show ignored" mode: show ONLY ignored chats
      if (!isIgnored) return false
    } else {
      // Normal mode: hide ignored chats
      if (isIgnored) return false
    }
    if (!filterState.enabledTypes.has(c.chatType)) return false
    if (filterState.markedOnly && !isChatFavorited(c.id)) return false
    // Build B: Messages OFF => recordings-only. A chat with no recording has
    // nothing to show in this scope, so filter it out. (Until the recording scan
    // lands, no chat has a recording yet — but §2/§3 aren't revealed until the
    // scan completes, so the user never sees the mid-scan empty list.)
    if (!loadIncludeMessages) {
      const rc = recordingsMap.get(c.id)
      if (!rc || rc.recordings.length === 0) return false
    }
    // hideDownloaded: hide containers where the chat archive has been downloaded
    if (userPrefs.hideDownloaded && chatPrefs[c.id]?.lastSync) return false
    if (q) {
      const name = chatDisplayName(c).toLowerCase()
      if (!name.includes(q)) return false
    }
    return true
  })
  const byRecent = (a: TeamsChatItem, b: TeamsChatItem) =>
    chatActivityDate(b) - chatActivityDate(a)
  const byName = (a: TeamsChatItem, b: TeamsChatItem) =>
    chatDisplayName(a).localeCompare(chatDisplayName(b))
  if (filterState.sortKey === "name") {
    result = [...result].sort(byName)
  } else if (filterState.sortKey === "recent") {
    result = [...result].sort(byRecent)
  } else {
    result = [...result].sort((a, b) => {
      const aM = isChatFavorited(a.id) ? 1 : 0
      const bM = isChatFavorited(b.id) ? 1 : 0
      if (aM !== bM) return bM - aM
      return byRecent(a, b)
    })
  }
  return result
}

function countByType(chats: TeamsChatItem[]): Map<string, number> {
  const counts = new Map<string, number>()
  for (const c of chats) counts.set(c.chatType, (counts.get(c.chatType) || 0) + 1)
  return counts
}

// ----- Render -----

function render(): void {
  const account = msal.getActiveAccount()
  if (!account) {
    app.innerHTML = `
      <div class="signin-card">
        <h1>m365-pull</h1>
        <p>Sign in with your Microsoft account to browse and download your Teams chats and call recordings.</p>
        <button id="signin" class="primary">Sign in with Microsoft</button>
      </div>
    `
    el<HTMLButtonElement>("signin").addEventListener("click", () => {
      void msal.loginRedirect({ scopes: SIGNIN_SCOPES })
    })
    return
  }

  markedIds = loadMarks(userCacheKey())
  ignoredIds = loadIgnored(userCacheKey())
  // Phase 2 favorites migration (gated, runs once per device): legacy v1 marks
  // are whole-container keeps stored as bare chatIds. Convert each to favorite
  // BOTH streams (keep the bare chatId for messages + add chatId::rec for
  // recordings) so nothing previously kept is lost. The local fav-schema flag
  // prevents re-running after the user starts toggling streams independently
  // (which would otherwise resurrect a deliberately un-favorited stream). Marks
  // arriving later via the OneDrive merge are handled in pullAndMergeOneDriveState
  // (gated on the remote state's own version).
  if (favSchemaVersion(userCacheKey()) < 2) {
    if (migrateMarksToStreams(markedIds)) {
      saveMarks(userCacheKey(), markedIds)
    }
    setFavSchemaVersion(userCacheKey(), 2)
  }
  chatPrefs = loadChatPrefs(userCacheKey())
  recordingPrefs = loadRecordingPrefs(userCacheKey())
  userPrefs = loadUserPrefs(userCacheKey())
  hydrateUIStateFromStorage()

  app.innerHTML = `
    <header>
      <h1>m365-pull</h1>
      <div class="user">
        <span class="user-name">${escapeHtml(account.username)}</span>
        <span id="sync-badge" class="sync-badge idle" title="">not synced</span>
        <button id="open-settings" class="icon-button" title="Settings" aria-label="Settings">\u2699</button>
        <button id="signout">Sign out</button>
      </div>
    </header>
    <div id="settings-modal" class="modal" hidden>
      <div class="modal-backdrop"></div>
      <div class="modal-card" role="dialog" aria-labelledby="settings-title" aria-modal="true">
        <header class="modal-header">
          <h2 id="settings-title">Settings</h2>
          <button id="settings-close" class="icon-button" aria-label="Close">\u2715</button>
        </header>
        <div class="modal-body">
          <label class="form-field">
            <span class="form-label">OneDrive folder path</span>
            <input id="settings-folder" type="text" class="form-input" />
            <span class="form-help">Where downloaded files land in your OneDrive when the "OneDrive folder" destination is selected. Files sync to wherever you have OneDrive mounted (Windows OneDrive client, Mac OneDrive, WSL via OneDrive mapping). Example: <code>/m365-pull/teams-chats</code></span>
          </label>
          <label class="form-field">
            <span class="form-label">Default destination</span>
            <select id="settings-destination" class="form-input">
              <option value="browser">Browser (save dialog)</option>
              <option value="onedrive">OneDrive folder</option>
            </select>
            <span class="form-help">Where to save downloads by default. You can override this from the main action row.</span>
          </label>
          <div class="form-field">
            <span class="form-label">Account</span>
            <span class="form-static">${escapeHtml(account.username)}</span>
            <span class="form-help">Signed-in user. Sign out from the header to switch accounts.</span>
          </div>
          <div class="form-field">
            <span class="form-label">State storage</span>
            <span class="form-static"><code>/Apps/m365-pull/state.json</code> in your OneDrive</span>
            <span class="form-help">Favorites, last-download timestamps, and these settings sync across all your devices via this single file.</span>
          </div>
        </div>
        <footer class="modal-footer">
          <button id="settings-cancel">Cancel</button>
          <button id="settings-save" class="primary">Save</button>
        </footer>
      </div>
    </div>
    <main>
      <!-- \u00a71 \u2014 LOAD (only section visible before content loads) -->
      <section class="toolbar-section" id="section-load">
        <div class="section-head">Load</div>
        <div class="section-body" id="load-picker">
          <span class="field">
            <span class="label">Show from:</span>
            <select id="chat-range" title="Date range for chat list">
              <option value="this-week">This week</option>
              <option value="last-7d" selected>Last 7 days</option>
              <option value="last-30d">Last 30 days</option>
              <option value="since-last-download">Since last pull</option>
              <option value="custom">Custom range\u2026</option>
            </select>
          </span>
          <span class="field" id="chat-custom-range-field"></span>
          <button class="chip" id="marked-include" title="Always show favorited chats regardless of range">\u2605 Always include favorites</button>
          <span class="field">
            <span class="label">Include:</span>
            <button class="chip active" id="include-messages" title="Include chat message archives">Messages</button>
            <button class="chip active" id="include-recordings" title="Include call recordings/transcripts">Recordings</button>
          </span>
          <button id="loadchats" class="primary">\u2b07 Load my Teams chats</button>
        </div>
        <div class="section-body load-checklist" id="load-checklist" hidden></div>
        <div class="section-body load-receipt" id="load-receipt" hidden></div>
      </section>

      <!-- \u00a72 \u2014 VIEW (revealed after load) -->
      <section class="toolbar-section" id="section-view" hidden>
        <div class="section-head">View</div>
        <div class="section-body">
          <input id="search" type="search" placeholder="Search chats by name\u2026" />
          <div class="chips" id="type-chips">
            ${KNOWN_TYPES.map(
              (t) =>
                `<button class="chip active" data-type="${t.id}">${escapeHtml(t.label)} <span class="chip-count">0</span></button>`,
            ).join("")}
          </div>
          <span class="field">
            <span class="label">Sort by:</span>
            <select id="sortby">
              <option value="marked-first">Favorites first \u00b7 then recent</option>
              <option value="recent">Most recent activity</option>
              <option value="name">Name (A\u2013Z)</option>
            </select>
          </span>
          <span class="field">
            <span class="label">View:</span>
            <span class="view-toggle" role="group" aria-label="View mode">
              <button class="seg" id="view-flat" data-view="flat" title="Flat \\u2014 every artifact (messages + each recording) as its own row">Flat</button>
              <button class="seg" id="view-grouped" data-view="grouped" title="Grouped \\u2014 chat rows that expand into their artifacts">Grouped</button>
            </span>
          </span>
          <button class="chip marked-only" id="markedonly">\u2605 Favorites only</button>
          <button class="chip" id="hide-downloaded">Hide downloaded</button>
          <button class="chip show-ignored" id="showignored" hidden>\u2298 Show ignored</button>
        </div>
      </section>

      <!-- \u00a73 \u2014 DOWNLOAD & BULK ACTIONS (revealed with \u00a72) -->
      <section class="toolbar-section" id="section-download" hidden>
        <div class="section-head">Download &amp; bulk actions</div>
        <div class="section-body">
          <span class="field">
            <span class="label">Destination:</span>
            <select id="destination" title="Where to save downloads">
              <option value="browser">Browser (save dialog)</option>
              <option value="onedrive">OneDrive folder</option>
            </select>
            <span class="onedrive-folder" id="onedrive-folder" hidden>
              <button class="link-button" id="edit-folder" title="Click to change">/m365-pull/teams-chats</button>
              <a class="link-button" id="open-folder" target="_blank" rel="noopener" hidden title="Open this folder in OneDrive on the web">\u2197 Open</a>
            </span>
          </span>
          <span class="field">
            <span class="label">Download history:</span>
            <select id="lookback" title="Download history per chat (messages)">
              <option value="7">Last 7 days</option>
              <option value="30" selected>Last 30 days</option>
              <option value="90">Last 90 days</option>
              <option value="all">All messages</option>
              <option value="since-last-download">Since last pull</option>
            </select>
          </span>
          <label class="field select-all-field">
            <input type="checkbox" id="select-all-global" />
            <span class="label">Select all</span>
          </label>
          <button id="bulk-containers" class="bulk-action" hidden></button>
          <button id="download-selected" class="bulk-action" hidden></button>
          <button class="chip clear-ignored" id="clearignored" hidden>\u2298 Clear ignored</button>
        </div>
      </section>

      <div id="scan-status" class="scan-status"></div>
      <div id="status"></div>
      <ul id="chats" class="chat-list"></ul>
    </main>
  `

  wireGlobalHandlers(account)
  syncUIControlsFromState()
  updateSyncIndicator()

  void pullAndMergeOneDriveState().then(() => rerenderContainerList())
}

function wireGlobalHandlers(account: AccountInfo): void {
  el<HTMLButtonElement>("signout").addEventListener("click", () => {
    void msal.logoutRedirect({ account })
  })

  // Container load (Reload now lives on the §1 receipt; Refresh button removed).
  el<HTMLButtonElement>("loadchats").addEventListener("click", () => {
    void initialLoadChats()
  })

  // §3 global "Select all" \u2014 select/clear every artifact across all loaded chats.
  el<HTMLInputElement>("select-all-global").addEventListener("change", () => {
    const cb = el<HTMLInputElement>("select-all-global")
    if (cb.checked) {
      for (const c of chatsState.chats) {
        for (const id of getArtifactIds(c.id, recordingsMap.get(c.id))) {
          selectedArtifacts.add(id)
        }
      }
    } else {
      selectedArtifacts.clear()
    }
    rerenderContainerList()
  })

  // Sync favorites button (pulls every favorited stream on click)
  el<HTMLButtonElement>("bulk-containers").addEventListener("click", () => {
    void syncFavorites()
  })

  // Phase 1: Download selected button
  el<HTMLButtonElement>("download-selected").addEventListener("click", () => {
    void downloadSelectedArtifacts()
  })

  // Chat range dropdown \u2014 custom date inputs are RENDERED inline only when
  // "Custom range\u2026" is selected (render-when-custom, not the hidden-attr trick).
  el<HTMLSelectElement>("chat-range").addEventListener("change", () => {
    const kind = el<HTMLSelectElement>("chat-range").value as ChatRange["kind"]
    if (kind === "custom") {
      // Seed defaults if the stored range isn't already custom.
      if (userPrefs.chatRange?.kind !== "custom") {
        userPrefs = { ...userPrefs, chatRange: {
          kind: "custom",
          customFrom: toLocalDateString(new Date(Date.now() - 7 * DAY_MS)),
          customTo: toLocalDateString(new Date()),
        } }
      }
    } else {
      userPrefs = { ...userPrefs, chatRange: { kind } }
    }
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    renderCustomRangeField()
    // Do NOT auto-load here. Changing the dropdown (including selecting "custom")
    // only persists the preference and reveals/hides the date inputs. The user
    // must press an explicit Load / Refresh button to trigger a data fetch.
    // (Bug fix: dropdown change was firing a reload before "custom" dates could
    // be set, and before any explicit user intent to reload.)
  })

  // chat-from / chat-to listeners are attached inside renderCustomRangeField,
  // since those inputs only exist in the DOM when "Custom range…" is selected.
  // Render them now if the stored range is already custom.
  renderCustomRangeField()

  // Marked-include toggle \u2014 persisted to OneDrive userPrefs; default ON.
  el<HTMLButtonElement>("marked-include").addEventListener("click", () => {
    const next = userPrefs.markedInclude === false ? true : false
    userPrefs = { ...userPrefs, markedInclude: next }
    el<HTMLButtonElement>("marked-include").classList.toggle("active", next)
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    // Reload so the marked-include enrichment runs (or stops running).
    reloadChatsWithCurrentSettings()
  })

  // Include is a LOAD-SCOPE (both-on default). Can't turn both off \u2014 block the
  // last-on chip from turning off (tooltip "Include at least one"). Build A keeps
  // behavior unchanged (both default on); Build B wires the actual scope.
  el<HTMLButtonElement>("include-messages").addEventListener("click", () => {
    const turningOff = userPrefs.includeMessages !== false
    if (turningOff && userPrefs.includeRecordings === false) {
      const btn = el<HTMLButtonElement>("include-messages")
      btn.title = "Include at least one"
      setStatus("Include at least one (Messages or Recordings).", "error")
      return
    }
    const next = userPrefs.includeMessages === false ? true : false
    userPrefs = { ...userPrefs, includeMessages: next }
    el<HTMLButtonElement>("include-messages").classList.toggle("active", next)
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
  })

  el<HTMLButtonElement>("include-recordings").addEventListener("click", () => {
    const turningOff = userPrefs.includeRecordings !== false
    if (turningOff && userPrefs.includeMessages === false) {
      const btn = el<HTMLButtonElement>("include-recordings")
      btn.title = "Include at least one"
      setStatus("Include at least one (Messages or Recordings).", "error")
      return
    }
    const next = userPrefs.includeRecordings === false ? true : false
    userPrefs = { ...userPrefs, includeRecordings: next }
    el<HTMLButtonElement>("include-recordings").classList.toggle("active", next)
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
  })

  // Change 3: Grouped/Flat view toggle \u2014 persisted in synced userPrefs.
  const setViewMode = (mode: "grouped" | "flat") => {
    if ((userPrefs.viewMode ?? "flat") === mode) return
    userPrefs = { ...userPrefs, viewMode: mode }
    el<HTMLButtonElement>("view-grouped").classList.toggle("active", mode === "grouped")
    el<HTMLButtonElement>("view-flat").classList.toggle("active", mode === "flat")
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    rerenderContainerList()
  }
  el<HTMLButtonElement>("view-grouped").addEventListener("click", () => setViewMode("grouped"))
  el<HTMLButtonElement>("view-flat").addEventListener("click", () => setViewMode("flat"))

  // Chat lookback: persists the download-depth selection only.
  el<HTMLSelectElement>("lookback").addEventListener("change", () => {
    saveUIPrefs()
  })

  // Hide-downloaded toggle
  el<HTMLButtonElement>("hide-downloaded").addEventListener("click", () => {
    const next = !userPrefs.hideDownloaded
    userPrefs = { ...userPrefs, hideDownloaded: next }
    el<HTMLButtonElement>("hide-downloaded").classList.toggle("active", next)
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    rerenderContainerList()
  })

  // Container filters
  const search = el<HTMLInputElement>("search")
  let searchTimer: number | null = null
  search.addEventListener("input", () => {
    if (searchTimer) window.clearTimeout(searchTimer)
    searchTimer = window.setTimeout(() => {
      filterState.search = search.value
      rerenderContainerList()
      saveUIPrefs()
    }, 150)
  })
  el<HTMLDivElement>("type-chips")
    .querySelectorAll<HTMLButtonElement>(".chip[data-type]")
    .forEach((chip) => {
      chip.addEventListener("click", () => {
        const t = chip.dataset.type!
        if (filterState.enabledTypes.has(t)) {
          filterState.enabledTypes.delete(t)
          chip.classList.remove("active")
        } else {
          filterState.enabledTypes.add(t)
          chip.classList.add("active")
        }
        rerenderContainerList()
        saveUIPrefs()
      })
    })
  el<HTMLSelectElement>("sortby").addEventListener("change", (e) => {
    filterState.sortKey = (e.target as HTMLSelectElement).value as SortKey
    rerenderContainerList()
    saveUIPrefs()
  })
  el<HTMLButtonElement>("markedonly").addEventListener("click", () => {
    filterState.markedOnly = !filterState.markedOnly
    el<HTMLButtonElement>("markedonly").classList.toggle(
      "active",
      filterState.markedOnly,
    )
    rerenderContainerList()
    saveUIPrefs()
  })

  // Show ignored toggle
  el<HTMLButtonElement>("showignored").addEventListener("click", () => {
    filterState.showIgnored = !filterState.showIgnored
    el<HTMLButtonElement>("showignored").classList.toggle(
      "active",
      filterState.showIgnored,
    )
    rerenderContainerList()
    saveUIPrefs()
  })

  // Clear all ignored
  el<HTMLButtonElement>("clearignored").addEventListener("click", () => {
    clearAllIgnored()
  })

  // Destination preference
  syncUserPrefsToUI()
  const destinationSel = el<HTMLSelectElement>("destination")
  destinationSel.addEventListener("change", () => {
    userPrefs = {
      ...userPrefs,
      destination: destinationSel.value as Destination,
    }
    saveUserPrefs(userCacheKey(), userPrefs)
    syncUserPrefsToUI()
    scheduleOneDriveSave()
  })
  // Inline folder label opens Settings
  el<HTMLButtonElement>("edit-folder").addEventListener("click", () => {
    openSettingsModal()
  })

  // Settings modal wiring
  el<HTMLButtonElement>("open-settings").addEventListener("click", () => {
    openSettingsModal()
  })
  el<HTMLButtonElement>("settings-close").addEventListener("click", () => {
    closeSettingsModal()
  })
  el<HTMLButtonElement>("settings-cancel").addEventListener("click", () => {
    closeSettingsModal()
  })
  el<HTMLButtonElement>("settings-save").addEventListener("click", () => {
    saveSettingsModal()
  })
  // Click backdrop to close
  el<HTMLDivElement>("settings-modal")
    .querySelector(".modal-backdrop")
    ?.addEventListener("click", () => closeSettingsModal())
  // Esc to close
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      const modal = document.getElementById("settings-modal")
      if (modal && !modal.hidden) closeSettingsModal()
    }
  })

  window.setInterval(updateSyncIndicator, 30_000)
}

/** Read module state and stash it in localStorage so reloads remember
 * the user's dropdown selections and chip filters. */
function saveUIPrefs(): void {
  const lookbackEl = document.getElementById("lookback") as
    | HTMLSelectElement
    | null
  saveUIState(userCacheKey(), {
    lookback: lookbackEl?.value,
    chatFilter: {
      search: filterState.search,
      enabledTypes: [...filterState.enabledTypes],
      sortKey: filterState.sortKey,
      markedOnly: filterState.markedOnly,
      showIgnored: filterState.showIgnored,
    },
  })
}

/** Restore filter state and dropdown values from localStorage.
 * Runs after the HTML is rendered (handlers wired later read these values). */
function hydrateUIStateFromStorage(): void {
  const saved = loadUIState(userCacheKey())
  if (saved.chatFilter) {
    filterState = {
      search: saved.chatFilter.search ?? "",
      enabledTypes: new Set(
        saved.chatFilter.enabledTypes ?? KNOWN_TYPES.map((t) => t.id),
      ),
      sortKey: saved.chatFilter.sortKey ?? "marked-first",
      markedOnly: saved.chatFilter.markedOnly ?? false,
      showIgnored: saved.chatFilter.showIgnored ?? false,
    }
  }
}

/** Push module state into the freshly-rendered DOM controls. Called once
 * after wireGlobalHandlers so handler-attached defaults don't fight us. */
function syncUIControlsFromState(): void {
  const saved = loadUIState(userCacheKey())
  // Chat lookback (download depth only \u2014 list range is synced via syncUserPrefsToUI)
  const lookbackEl = document.getElementById("lookback") as HTMLSelectElement | null
  if (lookbackEl && saved.lookback) lookbackEl.value = saved.lookback
  // Chat filter UI: search, sortby, type chips, markedonly
  const searchEl = document.getElementById("search") as HTMLInputElement | null
  if (searchEl) searchEl.value = filterState.search
  const sortbyEl = document.getElementById("sortby") as HTMLSelectElement | null
  if (sortbyEl) sortbyEl.value = filterState.sortKey
  const markedOnlyChat = document.getElementById("markedonly")
  if (markedOnlyChat) markedOnlyChat.classList.toggle("active", filterState.markedOnly)
  const showIgnoredBtn = document.getElementById("showignored")
  if (showIgnoredBtn) showIgnoredBtn.classList.toggle("active", filterState.showIgnored)
  document
    .querySelectorAll<HTMLButtonElement>("#type-chips .chip[data-type]")
    .forEach((chip) => {
      const t = chip.dataset.type
      const enabled = t ? filterState.enabledTypes.has(t) : true
      chip.classList.toggle("active", enabled)
    })
}

// ----- Container list rendering -----

function rerenderContainerList(): void {
  const list = el<HTMLUListElement>("chats")
  const filtered = applyContainerFiltersAndSort(chatsState.chats)
  if (chatsState.chats.length === 0) {
    list.innerHTML = ""
    return
  }
  if (filtered.length === 0) {
    list.innerHTML = `<li class="empty">No containers match these filters. ${
      filterState.markedOnly
        ? "Favorite a chat\u2019s Messages or Recordings stream (expand a row) to add it here."
        : "Loosen the filters."
    }</li>`
  } else {
    // Change 3: render grouped (chat rows that expand) or flat (every artifact
    // its own top-level row). Both reuse the same artifact-row helpers, so
    // Select / Download / Favorite behave identically across modes.
    const viewMode = userPrefs.viewMode === "grouped" ? "grouped" : "flat"
    if (viewMode === "flat") {
      list.innerHTML = renderFlatArtifactRows(filtered)
    } else {
      list.innerHTML = filtered
        .map((c) => renderContainerRow(c, recordingsMap.get(c.id)))
        .join("")
    }
    list.querySelectorAll<HTMLButtonElement>(".container-action").forEach((btn) => {
      btn.addEventListener("click", () => {
        const id = btn.dataset.chatId!
        const name = btn.dataset.chatName!
        void syncContainer(id, name, btn)
      })
    })
    // Change 1: per-artifact direct download (works in BOTH grouped & flat).
    // Reuses the exact primitives downloadSelectedArtifacts calls; the clicked
    // button shows its own progress/disabled state.
    list.querySelectorAll<HTMLButtonElement>(".artifact-download").forEach((btn) => {
      btn.addEventListener("click", () => {
        if (btn.dataset.artKind === "recording") {
          const recId = btn.dataset.recId!
          void downloadRecordingTranscript(recId, btn)
        } else {
          const chatId = btn.dataset.chatId!
          const chatName = btn.dataset.chatName || "chat"
          void downloadChat(chatId, chatName, btn)
        }
      })
    })
    // --- Phase 2: per-stream favorite toggles (live in the expanded view) ---
    list.querySelectorAll<HTMLButtonElement>(".fav-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const chatId = btn.dataset.chatId!
        if (btn.dataset.stream === "recordings") {
          toggleFavoriteRecordings(chatId)
        } else {
          toggleFavoriteMessages(chatId)
        }
      })
    })
    list.querySelectorAll<HTMLButtonElement>(".ignore-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        toggleIgnore(btn.dataset.chatId!)
      })
    })

    // --- Phase 1: expand/collapse ---
    list.querySelectorAll<HTMLButtonElement>(".expand-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const chatId = btn.dataset.chatId!
        const li = btn.closest<HTMLLIElement>("li.chat-row")
        const artifactRowsDiv = li?.querySelector<HTMLDivElement>(".artifact-rows")
        if (!li || !artifactRowsDiv) return
        if (expandedChatIds.has(chatId)) {
          expandedChatIds.delete(chatId)
          artifactRowsDiv.hidden = true
          btn.setAttribute("aria-expanded", "false")
          btn.textContent = "\u25b8"
          btn.title = "Expand artifacts"
          li.classList.remove("expanded")
        } else {
          expandedChatIds.add(chatId)
          artifactRowsDiv.hidden = false
          btn.setAttribute("aria-expanded", "true")
          btn.textContent = "\u25be"
          btn.title = "Collapse artifacts"
          li.classList.add("expanded")
        }
      })
    })

    // --- Phase 1: select-all checkboxes (group level) ---
    list.querySelectorAll<HTMLInputElement>(".select-all-check").forEach((cb) => {
      // Set indeterminate for "some but not all selected" — can't do via HTML attr
      const chatId = cb.dataset.chatId!
      const rc = recordingsMap.get(chatId)
      const aids = getArtifactIds(chatId, rc)
      const selCount = aids.filter((id) => selectedArtifacts.has(id)).length
      if (selCount > 0 && selCount < aids.length) cb.indeterminate = true

      cb.addEventListener("change", () => {
        const cid = cb.dataset.chatId!
        const container = recordingsMap.get(cid)
        const artifactIds = getArtifactIds(cid, container)
        if (cb.checked) {
          artifactIds.forEach((id) => selectedArtifacts.add(id))
        } else {
          artifactIds.forEach((id) => selectedArtifacts.delete(id))
        }
        cb.indeterminate = false
        // Sync individual artifact checkboxes
        list.querySelectorAll<HTMLInputElement>(".artifact-check").forEach((artCb) => {
          if (artifactIds.includes(artCb.dataset.artifactId ?? "")) {
            artCb.checked = cb.checked
          }
        })
        updateSelectedButton()
      })
    })

    // --- Phase 1: individual artifact checkboxes ---
    list.querySelectorAll<HTMLInputElement>(".artifact-check").forEach((cb) => {
      cb.addEventListener("change", () => {
        const artifactId = cb.dataset.artifactId!
        const chatId = cb.closest<HTMLDivElement>(".artifact-row")?.dataset.chatId ?? ""
        if (cb.checked) {
          selectedArtifacts.add(artifactId)
        } else {
          selectedArtifacts.delete(artifactId)
        }
        // Update the group select-all checkbox for this chat
        if (chatId) {
          const rc = recordingsMap.get(chatId)
          const aids = getArtifactIds(chatId, rc)
          const sc = aids.filter((id) => selectedArtifacts.has(id)).length
          list.querySelectorAll<HTMLInputElement>(".select-all-check").forEach((groupCb) => {
            if (groupCb.dataset.chatId === chatId) {
              if (sc === 0) { groupCb.checked = false; groupCb.indeterminate = false }
              else if (sc === aids.length) { groupCb.checked = true; groupCb.indeterminate = false }
              else { groupCb.checked = false; groupCb.indeterminate = true }
            }
          })
        }
        updateSelectedButton()
      })
    })
  }
  updateTypeCountChips()
  updateIgnoredControlsVisibility()
  updateContainerSummary(filtered.length)
  updateBulkButtons()
  updateSelectedButton()
}

/** Show/Clear-ignored controls render ONLY when there is something ignored.
 * Show-ignored lives in §2 (VIEW); Clear-ignored is an action in §3 (DOWNLOAD). */
function updateIgnoredControlsVisibility(): void {
  const ignored = ignoredIds.size
  const clearBtn = document.getElementById("clearignored") as HTMLButtonElement | null
  if (clearBtn) clearBtn.hidden = ignored === 0
  const showBtn = document.getElementById("showignored") as HTMLButtonElement | null
  if (showBtn) showBtn.hidden = ignored === 0
}

function updateContainerSummary(visible: number): void {
  const total = chatsState.chats.length
  const favorited = favoritedChatIds().size
  const ignored = ignoredIds.size
  const ignoredNote = ignored > 0 ? ` \u00b7 ${ignored} ignored` : ""
  // Recording scan progress lives in #scan-status (Item 4); omit from main status.
  const filtered = visible < total
  if (filtered) {
    setStatus(`Showing ${visible} of ${total} loaded \u00b7 ${favorited} favorited${ignoredNote}`)
  } else {
    setStatus(`${total} chats loaded \u00b7 ${favorited} favorited${ignoredNote}`)
  }
}

function renderContainerRow(
  chat: TeamsChatItem,
  recContainer: RecordingContainer | undefined,
): string {
  const name = chatDisplayName(chat)
  const isMarked = isChatFavorited(chat.id)
  const isIgnored = ignoredIds.has(chat.id)
  const lastSync = chatPrefs[chat.id]?.lastSync
  // Item 5: capitalise "Downloaded"; make no-download state explicit.
  const downloadedTag = lastSync
    ? ` \u00b7 Downloaded ${formatDateShort(new Date(lastSync))}`
    : " \u00b7 Not downloaded yet"
  const sub = `${typeLabel(chat.chatType)} \u00b7 last activity ${formatDate(new Date(chatActivityDate(chat)).toISOString())}${downloadedTag}`

  // Phase 1 recording indicator: only show badge when recordings are confirmed.
  // All other states (pending, none, unknown) = silence ("no-recording = silence" design).
  let recIndicatorHtml = ""
  if (recContainer !== undefined && recContainer.recordings.length > 0) {
    const n = recContainer.recordings.length
    const truncNote = recContainer.truncated ? " (may be incomplete)" : ""
    recIndicatorHtml = `<span class="rec-indicator" aria-label="${n} recording${n !== 1 ? "s" : ""}${truncNote}" title="${n} recording${n !== 1 ? "s" : ""} in range${truncNote}"><span aria-hidden="true">\uD83C\uDF99 ${n}</span></span>`
  }

  // Effective download scope = captured load-time Include \u2229 what this row actually has.
  const wantMessages = loadIncludeMessages
  const hasRecordings = recContainer !== undefined && recContainer.recordings.length > 0
  const wantRecordings = loadIncludeRecordings && hasRecordings
  let downloadBtnInner: string
  let downloadBtnDisabled = ""
  let downloadBtnTitle = ""
  let downloadBtnAriaLabel: string
  if (wantMessages && wantRecordings) {
    downloadBtnInner = `Download <span aria-hidden="true">\uD83D\uDCAC\uD83C\uDF99</span>`
    downloadBtnAriaLabel = "Download messages and recordings"
  } else if (wantMessages) {
    downloadBtnInner = `Download <span aria-hidden="true">\uD83D\uDCAC</span>`
    downloadBtnAriaLabel = "Download messages"
  } else if (wantRecordings) {
    downloadBtnInner = `Download <span aria-hidden="true">\uD83C\uDF99</span>`
    downloadBtnAriaLabel = "Download recordings"
  } else {
    downloadBtnInner = "Download"
    downloadBtnDisabled = " disabled"
    downloadBtnTitle = ` title="Nothing to download for current Include settings"`
    downloadBtnAriaLabel = "Download \u2014 nothing to download for current Include settings"
  }

  // Phase 1 expand/select-all state
  const isExpanded = expandedChatIds.has(chat.id)
  const artifactIds = getArtifactIds(chat.id, recContainer)
  const selectedCount = artifactIds.filter((id) => selectedArtifacts.has(id)).length
  const allSelected = selectedCount > 0 && selectedCount === artifactIds.length
  // indeterminate state (someSelected) must be set via JS after render \u2014 not expressible in HTML

  return `
    <li class="chat-row${isMarked ? " marked" : ""}${isIgnored ? " ignored" : ""}${isExpanded ? " expanded" : ""}">
      <div class="chat-row-header">
        <button class="expand-toggle" data-chat-id="${escapeHtml(chat.id)}" aria-expanded="${isExpanded ? "true" : "false"}" title="${isExpanded ? "Collapse artifacts" : "Expand artifacts"}">${isExpanded ? "\u25be" : "\u25b8"}</button>
        <input type="checkbox" class="select-all-check" data-chat-id="${escapeHtml(chat.id)}"${allSelected ? " checked" : ""} title="Select all artifacts in this chat" aria-label="Select all artifacts for ${escapeHtml(name)}">
        <span class="fav-state${isMarked ? " favorited" : ""}" title="${isMarked ? "Favorited \u2014 expand to change which streams" : "Not favorited \u2014 expand to favorite a stream"}" aria-label="${isMarked ? "Favorited" : "Not favorited"}">${isMarked ? "\u2605" : "\u2606"}</span>
        <div class="chat-info">
          <div class="chat-name">${escapeHtml(name)}</div>
          <div class="chat-sub">${escapeHtml(sub)}</div>
        </div>
        ${recIndicatorHtml}
        <button class="ignore-toggle${isIgnored ? " ignored" : ""}" data-chat-id="${escapeHtml(chat.id)}" title="${isIgnored ? "Un-ignore this container" : "Ignore this container"}" aria-label="${isIgnored ? "Un-ignore" : "Ignore"}" aria-pressed="${isIgnored ? "true" : "false"}">${isIgnored ? "\u2299" : "\u2298"}</button>
        <button class="container-action"${downloadBtnDisabled} data-chat-id="${escapeHtml(chat.id)}" data-chat-name="${escapeHtml(name)}" aria-label="${downloadBtnAriaLabel}"${downloadBtnTitle}>${downloadBtnInner}</button>
      </div>
      <div class="artifact-rows"${isExpanded ? "" : " hidden"}>
        ${renderArtifactRows(chat, recContainer)}
      </div>
    </li>
  `
}

/** Shared markup for the Messages artifact row.
 *
 * context "grouped" -> a <div> inside a chat's expanded .artifact-rows; the row
 *   label is just "Messages" (the chat name is on the parent chat row).
 * context "flat" -> a top-level <li> in the flat list; the chat name becomes the
 *   primary label so the row is identifiable on its own.
 *
 * Carries: the per-artifact Select checkbox, the messages-stream Favorite ★, and
 * the new per-artifact Download button (Change 1). */
function messagesArtifactRowHtml(
  chat: TeamsChatItem,
  context: "grouped" | "flat",
): string {
  const chatId = chat.id
  const chatName = chatDisplayName(chat)
  const flat = context === "flat"
  const msgArtId = `msg:${chatId}`
  const msgSelected = selectedArtifacts.has(msgArtId)
  const msgFav = isMessagesFavorited(chatId)
  const lastSync = chatPrefs[chatId]?.lastSync
  const dlTag = lastSync
    ? `Downloaded ${formatDateShort(new Date(lastSync))}`
    : "Not downloaded yet"
  const lastActivity = formatDate(new Date(chatActivityDate(chat)).toISOString())
  const nameRaw = flat ? chatName : "Messages"
  const subRaw = flat
    ? `Messages \u00b7 ${lastActivity} \u00b7 ${dlTag}`
    : `${lastActivity} \u00b7 ${dlTag}`
  const tag = flat ? "li" : "div"
  const flatClass = flat ? " artifact-flat" : ""
  const ctxSuffix = flat ? ` for ${chatName}` : ""
  return `
    <${tag} class="artifact-row${flatClass}" data-artifact-id="${escapeHtml(msgArtId)}" data-chat-id="${escapeHtml(chatId)}">
      <input type="checkbox" class="artifact-check" data-artifact-id="${escapeHtml(msgArtId)}"${msgSelected ? " checked" : ""} aria-label="Select Messages artifact${escapeHtml(ctxSuffix)}">
      <button class="fav-toggle${msgFav ? " favorited" : ""}" data-stream="messages" data-chat-id="${escapeHtml(chatId)}" title="${msgFav ? "Un-favorite the Messages stream" : "Favorite the Messages stream \u2014 synced on every Sync"}" aria-label="${msgFav ? "Un-favorite Messages stream" : "Favorite Messages stream"}" aria-pressed="${msgFav ? "true" : "false"}">${msgFav ? "\u2605" : "\u2606"}</button>
      <span class="artifact-type-icon" aria-hidden="true">\uD83D\uDCAC</span>
      <div class="artifact-info">
        <div class="artifact-name">${escapeHtml(nameRaw)}</div>
        <div class="artifact-sub">${escapeHtml(subRaw)}</div>
      </div>
      <button class="artifact-download" data-art-kind="messages" data-chat-id="${escapeHtml(chatId)}" data-chat-name="${escapeHtml(chatName)}" title="Download messages now" aria-label="Download messages${escapeHtml(ctxSuffix)}"><span aria-hidden="true">\u2b07</span></button>
    </${tag}>
  `
}

/** Shared markup for a single Recording artifact row.
 *
 * context "grouped" -> a <div> indented under the recordings group header (which
 *   carries the shared recordings-stream ★); the row itself shows no ★.
 * context "flat" -> a top-level <li>; since there's no group header, EACH row
 *   carries the shared recordings-stream ★ (toggling any toggles the one
 *   chatId::rec favorite), and the chat name leads the label.
 *
 * Carries: the per-artifact Select checkbox and the new per-artifact Download
 * button (Change 1). Recordings are never individually favoritable. */
function recordingArtifactRowHtml(
  chat: TeamsChatItem,
  rec: RecordingItem,
  context: "grouped" | "flat",
): string {
  const chatId = chat.id
  const chatName = chatDisplayName(chat)
  const flat = context === "flat"
  const recArtId = `rec:${rec.id}`
  const recSelected = selectedArtifacts.has(recArtId)
  const dateLabel = formatDate(rec.eventCreatedDateTime)
  const duration = formatIsoDuration(rec.durationIso)
  const attendees = rec.participants
    .filter((p) => p.kind === "user")
    .map((p) => p.displayName)
    .filter(Boolean)
    .join(", ")
  const recLastSync = recordingPrefs[rec.id]?.lastSync
  const recTag = recLastSync
    ? ` \u00b7 Downloaded ${formatDateShort(new Date(recLastSync))}`
    : ""
  const baseSub = [duration, attendees].filter(Boolean).join(" \u00b7 ")
  const tag = flat ? "li" : "div"
  const flatClass = flat ? " artifact-flat" : ""
  const recFav = isRecordingsFavorited(chatId)
  const favBtn = flat
    ? `<button class="fav-toggle${recFav ? " favorited" : ""}" data-stream="recordings" data-chat-id="${escapeHtml(chatId)}" title="${recFav ? "Un-favorite the Recordings stream" : "Favorite the Recordings stream \u2014 grabs all recordings on every Sync"}" aria-label="${recFav ? "Un-favorite Recordings stream" : "Favorite Recordings stream"}" aria-pressed="${recFav ? "true" : "false"}">${recFav ? "\u2605" : "\u2606"}</button>`
    : ""
  const nameRaw = flat ? chatName : `Recording \u2014 ${dateLabel}`
  const subRaw = flat
    ? `Recording ${dateLabel}${baseSub ? " \u00b7 " + baseSub : ""}${recTag}`
    : `${baseSub}${recTag}`
  return `
    <${tag} class="artifact-row${flatClass}" data-artifact-id="${escapeHtml(recArtId)}" data-chat-id="${escapeHtml(chatId)}">
      <input type="checkbox" class="artifact-check" data-artifact-id="${escapeHtml(recArtId)}"${recSelected ? " checked" : ""} aria-label="Select Recording artifact${flat ? escapeHtml(` for ${chatName}`) : ""}">
      ${favBtn}
      <span class="artifact-type-icon" aria-hidden="true">\uD83C\uDF99</span>
      <div class="artifact-info">
        <div class="artifact-name">${escapeHtml(nameRaw)}</div>
        <div class="artifact-sub">${escapeHtml(subRaw)}</div>
      </div>
      <button class="artifact-download" data-art-kind="recording" data-rec-id="${escapeHtml(rec.id)}" title="Download this recording now" aria-label="Download recording from ${escapeHtml(dateLabel)}"><span aria-hidden="true">\u2b07</span></button>
    </${tag}>
  `
}

/** Render the expanded artifact rows (Messages + per-recording) for a chat in
 * GROUPED mode. Called by renderContainerRow. */
function renderArtifactRows(
  chat: TeamsChatItem,
  recContainer: RecordingContainer | undefined,
): string {
  // Build B: Messages OFF => recordings-only; suppress the Messages artifact row.
  let html = loadIncludeMessages ? messagesArtifactRowHtml(chat, "grouped") : ""

  if (recContainer && recContainer.recordings.length > 0) {
    const n = recContainer.recordings.length
    const chatId = chat.id
    const recFav = isRecordingsFavorited(chatId)
    // Recordings-stream favorite lives on a group header (the stream is the unit;
    // individual recordings are immutable and not separately favoritable).
    html += `
      <div class="artifact-group-header" data-chat-id="${escapeHtml(chatId)}">
        <button class="fav-toggle${recFav ? " favorited" : ""}" data-stream="recordings" data-chat-id="${escapeHtml(chatId)}" title="${recFav ? "Un-favorite the Recordings stream" : "Favorite the Recordings stream \u2014 grabs all recordings on every Sync"}" aria-label="${recFav ? "Un-favorite Recordings stream" : "Favorite Recordings stream"}" aria-pressed="${recFav ? "true" : "false"}">${recFav ? "\u2605" : "\u2606"}</button>
        <span class="artifact-group-label">Recordings (${n})</span>
      </div>
    `
    for (const rec of recContainer.recordings) {
      html += recordingArtifactRowHtml(chat, rec, "grouped")
    }
  }

  return html
}

/** Render the FLAT view: every artifact as its own top-level <li> row, each
 * carrying its chat's name for identifiability. Reuses the same row helpers as
 * grouped mode so Select / Download / Favorite behave identically.
 *
 * When sort is "recent" (by date) we perform a GLOBAL sort across all artifacts
 * so that individual recordings from a recurring meeting interleave
 * chronologically with artifacts from other containers, rather than clustering
 * back-to-back under their shared container. For name / marked-first sorts the
 * chats array is already correctly ordered at the container level, so we emit
 * artifacts in container order (which is the natural expected grouping for
 * those sort keys). */
function renderFlatArtifactRows(chats: TeamsChatItem[]): string {
  if (filterState.sortKey === "recent") {
    // Build every artifact with its own per-artifact sort timestamp, then sort
    // globally so e.g. 4 recordings of "Weekly Standup" from different weeks
    // interleave with other items that happened between those weeks.
    const items: { ms: number; html: string }[] = []
    for (const chat of chats) {
      if (loadIncludeMessages) {
        items.push({
          ms: chatActivityDate(chat),
          html: messagesArtifactRowHtml(chat, "flat"),
        })
      }
      const rc = recordingsMap.get(chat.id)
      if (rc && rc.recordings.length > 0) {
        for (const rec of rc.recordings) {
          items.push({
            ms: new Date(rec.eventCreatedDateTime).getTime(),
            html: recordingArtifactRowHtml(chat, rec, "flat"),
          })
        }
      }
    }
    items.sort((a, b) => b.ms - a.ms)
    return items.map((i) => i.html).join("")
  }
  if (filterState.sortKey === "marked-first") {
    // Build a flat per-artifact list where each artifact carries its own sort
    // keys: primary = whether THIS artifact's specific stream is favorited
    // (per-stream, not whole-chat), secondary = date descending (same
    // per-artifact timestamps as the "recent" branch above).
    const items: { fav: boolean; ms: number; html: string }[] = []
    for (const chat of chats) {
      if (loadIncludeMessages) {
        items.push({
          fav: isMessagesFavorited(chat.id),     // messages-stream key = bare chatId
          ms: chatActivityDate(chat),
          html: messagesArtifactRowHtml(chat, "flat"),
        })
      }
      const rc = recordingsMap.get(chat.id)
      if (rc && rc.recordings.length > 0) {
        for (const rec of rc.recordings) {
          items.push({
            fav: isRecordingsFavorited(chat.id), // recordings-stream key (per recStreamKey)
            ms: new Date(rec.eventCreatedDateTime).getTime(),
            html: recordingArtifactRowHtml(chat, rec, "flat"),
          })
        }
      }
    }
    // Two-key sort: favorited artifacts first (true before false), then date desc.
    items.sort((a, b) => {
      if (a.fav !== b.fav) return a.fav ? -1 : 1
      return b.ms - a.ms
    })
    return items.map((i) => i.html).join("")
  }
  // name: the chats array is already sorted at the container level; emit
  // artifacts in container order so items from the same named container stay
  // together (the natural expectation for name sort).
  const parts: string[] = []
  for (const chat of chats) {
    // Build B: Messages OFF => recordings-only; suppress the Messages artifact row.
    if (loadIncludeMessages) parts.push(messagesArtifactRowHtml(chat, "flat"))
    const rc = recordingsMap.get(chat.id)
    if (rc && rc.recordings.length > 0) {
      for (const rec of rc.recordings) {
        parts.push(recordingArtifactRowHtml(chat, rec, "flat"))
      }
    }
  }
  return parts.join("")
}

/** Favoriting any stream clears the chat's ignored state (Favorite and Ignore
 * are mutually exclusive — you can't ignore something you're favoriting). */
function clearIgnoreOnFavorite(chatId: string): void {
  if (ignoredIds.has(chatId)) {
    ignoredIds.delete(chatId)
    saveIgnored(userCacheKey(), ignoredIds)
  }
}

/** Toggle the Favorite state of a chat's MESSAGES stream (key = bare chatId). */
function toggleFavoriteMessages(chatId: string): void {
  if (markedIds.has(chatId)) {
    markedIds.delete(chatId)
  } else {
    markedIds.add(chatId)
    clearIgnoreOnFavorite(chatId)
  }
  saveMarks(userCacheKey(), markedIds)
  rerenderContainerList()
  updateBulkButtons()
  scheduleOneDriveSave()
}

/** Toggle the Favorite state of a chat's RECORDINGS stream (key = chatId::rec). */
function toggleFavoriteRecordings(chatId: string): void {
  const key = recStreamKey(chatId)
  if (markedIds.has(key)) {
    markedIds.delete(key)
  } else {
    markedIds.add(key)
    clearIgnoreOnFavorite(chatId)
  }
  saveMarks(userCacheKey(), markedIds)
  rerenderContainerList()
  updateBulkButtons()
  scheduleOneDriveSave()
}

// Timer for the transient undo affordance shown after ignoring a container.
let undoIgnoreTimer: number | null = null
let undoIgnoreId: string | null = null

/** Toggle the ignored state for a container. Ignoring clears BOTH favorite
 * streams (Favorite and Ignore are mutually exclusive: Favorited / Inbox / Ignored). */
function toggleIgnore(id: string): void {
  const wasIgnored = ignoredIds.has(id)
  if (wasIgnored) {
    ignoredIds.delete(id)
  } else {
    ignoredIds.add(id)
    // Clear any favorite on either stream of this chat.
    const recKey = recStreamKey(id)
    if (markedIds.has(id) || markedIds.has(recKey)) {
      markedIds.delete(id)
      markedIds.delete(recKey)
      saveMarks(userCacheKey(), markedIds)
    }
  }
  saveIgnored(userCacheKey(), ignoredIds)
  rerenderContainerList()
  scheduleOneDriveSave()

  if (!wasIgnored) {
    if (undoIgnoreTimer !== null) window.clearTimeout(undoIgnoreTimer)
    undoIgnoreId = id
    const chatItem = chatsState.chats.find((c) => c.id === id)
    const label = chatItem ? `"${escapeHtml(chatDisplayName(chatItem))}"` : "Container"
    const statusEl = document.getElementById("status")
    if (statusEl) {
      statusEl.innerHTML = `${label} ignored. <button class="link-button" id="undo-ignore">Undo</button>`
      const undoBtn = document.getElementById("undo-ignore")
      if (undoBtn) {
        undoBtn.addEventListener("click", () => {
          if (undoIgnoreId) {
            ignoredIds.delete(undoIgnoreId)
            undoIgnoreId = null
            if (undoIgnoreTimer !== null) {
              window.clearTimeout(undoIgnoreTimer)
              undoIgnoreTimer = null
            }
            saveIgnored(userCacheKey(), ignoredIds)
            rerenderContainerList()
            scheduleOneDriveSave()
            setStatus("")
          }
        })
      }
    }
    undoIgnoreTimer = window.setTimeout(() => {
      undoIgnoreTimer = null
      undoIgnoreId = null
      setStatus("")
    }, 5000)
  }
}

/** Remove ALL ignored ids at once. */
function clearAllIgnored(): void {
  if (ignoredIds.size === 0) return
  const count = ignoredIds.size
  if (undoIgnoreTimer !== null) {
    window.clearTimeout(undoIgnoreTimer)
    undoIgnoreTimer = null
  }
  undoIgnoreId = null
  ignoredIds = new Set()
  saveIgnored(userCacheKey(), ignoredIds)
  rerenderContainerList()
  scheduleOneDriveSave()
  setStatus(`Cleared ${count} ignored container${count === 1 ? "" : "s"}.`)
}

function updateTypeCountChips(): void {
  const counts = countByType(chatsState.chats)
  el<HTMLDivElement>("type-chips")
    .querySelectorAll<HTMLButtonElement>(".chip[data-type]")
    .forEach((chip) => {
      const t = chip.dataset.type!
      const n = counts.get(t) ?? 0
      const span = chip.querySelector<HTMLSpanElement>(".chip-count")
      if (span) span.textContent = String(n)
    })
}

// ----- §1 progressive disclosure: picker / checklist / receipt -----

/** §1 shows the LOAD picker (the date/include/Load controls). */
function showLoadPicker(): void {
  const picker = document.getElementById("load-picker")
  const checklist = document.getElementById("load-checklist")
  const receipt = document.getElementById("load-receipt")
  if (picker) picker.hidden = false
  if (checklist) checklist.hidden = true
  if (receipt) receipt.hidden = true
}

/** §1 shows the step-by-step loading CHECKLIST. */
function showLoadChecklistView(): void {
  const picker = document.getElementById("load-picker")
  const checklist = document.getElementById("load-checklist")
  const receipt = document.getElementById("load-receipt")
  if (picker) picker.hidden = true
  if (checklist) checklist.hidden = false
  if (receipt) receipt.hidden = true
}

/** Reveal §2 VIEW and §3 DOWNLOAD (progressive disclosure: hidden until a load completes). */
function revealPostLoadSections(): void {
  const view = document.getElementById("section-view")
  const dl = document.getElementById("section-download")
  if (view) view.hidden = false
  if (dl) dl.hidden = false
}

/** §1 collapses to a one-line RECEIPT, and §2/§3 are revealed. */
function showLoadReceipt(): void {
  const picker = document.getElementById("load-picker")
  const checklist = document.getElementById("load-checklist")
  const receipt = document.getElementById("load-receipt")
  if (picker) picker.hidden = true
  if (checklist) checklist.hidden = true
  if (receipt) receipt.hidden = false
  renderLoadReceipt()
  revealPostLoadSections()
}

/** Render the receipt line: ✓ N chats · <range> · ★ favorites included [Change] [⟳ Reload]. */
function renderLoadReceipt(): void {
  const host = document.getElementById("load-receipt")
  if (!host) return
  const n = chatsState.chats.length
  const rangeStr = chatRangeLabel(userPrefs.chatRange ?? { kind: "last-7d" }, chatPrefs)
  const favNote = userPrefs.markedInclude !== false ? " \u00b7 \u2605 favorites included" : ""
  host.innerHTML = `
    <span class="receipt-text">\u2713 ${n} chat${n === 1 ? "" : "s"} \u00b7 ${escapeHtml(rangeStr)}${favNote}</span>
    <button class="link-button" id="load-change">Change</button>
    <button class="link-button" id="load-reload">\u27f3 Reload</button>
  `
  host.querySelector<HTMLButtonElement>("#load-change")
    ?.addEventListener("click", () => showLoadPicker())
  host.querySelector<HTMLButtonElement>("#load-reload")
    ?.addEventListener("click", () => void refreshChats())
}

// ----- §1 loading checklist (real phase completion, not a timed bar) -----

function initLoadSteps(): void {
  loadSteps = [
    { id: "signin", label: "Signed in", state: "done" },
    { id: "chats", label: "Finding chats", state: "pending" },
  ]
  // Build B: with Recordings OFF the scan is skipped entirely, so omit its
  // checklist step (any stray updateLoadStep("recordings", ...) safely no-ops).
  if (loadIncludeRecordings) {
    loadSteps.push({ id: "recordings", label: "Scanning recordings", state: "pending" })
  }
  loadSteps.push({ id: "done", label: "Done", state: "pending" })
  renderLoadChecklist()
}

function updateLoadStep(id: LoadStep["id"], patch: Partial<LoadStep>): void {
  const step = loadSteps.find((s) => s.id === id)
  if (!step) return
  Object.assign(step, patch)
  renderLoadChecklist()
}

function renderLoadChecklist(): void {
  const host = document.getElementById("load-checklist")
  if (!host) return
  host.innerHTML = loadSteps
    .map((s) => {
      const icon =
        s.state === "done" ? "\u2713"
        : s.state === "running" ? "\u27f3"
        : s.state === "error" ? "\u2715"
        : "\u00b7"
      const text = s.state === "error" ? (s.error ?? "Failed") : (s.detail ?? s.label)
      const retry =
        s.state === "error"
          ? ` <button class="link-button load-retry" data-step="${s.id}">Retry</button>`
          : ""
      return `<div class="load-step state-${s.state}"><span class="load-ico" aria-hidden="true">${icon}</span><span class="load-text">${escapeHtml(text)}</span>${retry}</div>`
    })
    .join("")
  host.querySelectorAll<HTMLButtonElement>(".load-retry").forEach((btn) => {
    btn.addEventListener("click", () => retryLoadStep(btn.dataset.step as LoadStep["id"]))
  })
}

function retryLoadStep(step: LoadStep["id"]): void {
  if (step === "chats") {
    void initialLoadChats()
  } else if (step === "recordings" && lastRecordingWindow) {
    updateLoadStep("recordings", { state: "running", detail: "Scanning recordings\u2026", error: undefined })
    void backgroundLoadRecordings(lastRecordingWindow.fromMs, lastRecordingWindow.toMs)
  }
}

/** Mark the load complete: stamp hasLoadedOnce, collapse §1 to receipt, reveal §2/§3. */
function finalizeLoad(): void {
  hasLoadedOnce = true
  updateLoadStep("done", { state: "done", detail: "Done" })
  showLoadReceipt()
}

/** Reload the chat list with current settings, but only once a load has happened
 * (replaces the old `!refreshchats.hidden` gate). Used by range/favorite-include changes. */
function reloadChatsWithCurrentSettings(): void {
  if (!hasLoadedOnce) return
  clearCachedChats(userCacheKey())
  chatsState = { chats: [] }
  recordingsState = { containers: [], chatsScanned: 0, truncated: false }
  recordingsMap = new Map()
  el<HTMLUListElement>("chats").innerHTML = ""
  void initialLoadChats()
}

/** Render the custom-range From/To inputs inline ONLY when "Custom range…" is the
 * selected chat-range (render-when-custom; empties the field otherwise). */
function renderCustomRangeField(): void {
  const field = document.getElementById("chat-custom-range-field") as HTMLSpanElement | null
  if (!field) return
  const rangeSel = document.getElementById("chat-range") as HTMLSelectElement | null
  const kind = rangeSel?.value ?? userPrefs.chatRange?.kind ?? "last-7d"
  if (kind !== "custom") {
    field.innerHTML = ""
    return
  }
  const range = userPrefs.chatRange
  const today = toLocalDateString(new Date())
  const from =
    (range?.kind === "custom" && range.customFrom) ||
    toLocalDateString(new Date(Date.now() - 7 * DAY_MS))
  const to =
    (range?.kind === "custom" && range.customTo) || today
  field.innerHTML = `
    <span class="label">From</span>
    <input type="date" id="chat-from" title="From (inclusive)" value="${from}" max="${today}" />
    <span class="label">to</span>
    <input type="date" id="chat-to" title="To (inclusive)" value="${to}" max="${today}" />
  `
  field.querySelector<HTMLInputElement>("#chat-from")?.addEventListener("change", applyCustomChatRange)
  field.querySelector<HTMLInputElement>("#chat-to")?.addEventListener("change", applyCustomChatRange)
}

/** Persist a custom From/To selection without auto-reloading.
 * The user must press Load / Refresh to actually fetch data with the new range.
 * Values are clamped to today so future dates can't slip through even if the
 * browser's max attribute is bypassed. */
function applyCustomChatRange(): void {
  const fromEl = document.getElementById("chat-from") as HTMLInputElement | null
  const toEl = document.getElementById("chat-to") as HTMLInputElement | null
  const today = toLocalDateString(new Date())
  const clampToToday = (v: string | undefined) =>
    v && v > today ? today : v
  userPrefs = { ...userPrefs, chatRange: {
    kind: "custom",
    customFrom: clampToToday(fromEl?.value || undefined),
    customTo: clampToToday(toEl?.value || undefined),
  } }
  saveUserPrefs(userCacheKey(), userPrefs)
  scheduleOneDriveSave()
  // No auto-reload — the user explicitly presses Load / Refresh.
}

// ----- Chats loading -----

async function initialLoadChats(): Promise<void> {
  const userKey = userCacheKey()
  // Build B: CAPTURE the §1 Include scope at load time. This is the single entry
  // point for first-Load AND every Reload (refreshChats / reloadChatsWithCurrentSettings
  // both funnel here), so capturing here covers all cases. Render/filter/action
  // logic reads these captured values, not the live chips.
  loadIncludeMessages = userPrefs.includeMessages !== false // default ON
  loadIncludeRecordings = userPrefs.includeRecordings !== false // default ON
  // Selection is per-load + scope-dependent; clear it so no stale (now-out-of-scope)
  // artifact IDs linger across a reload.
  selectedArtifacts.clear()

  // §1 switches to the step-by-step loading checklist for the duration of the load.
  // (initLoadSteps reads loadIncludeRecordings, so it must run AFTER the capture.)
  initLoadSteps()
  showLoadChecklistView()
  updateLoadStep("chats", { state: "running", detail: "Finding chats\u2026" })

  const cached = loadCachedChats(userKey)
  if (cached && cached.chats.length > 0) {
    chatsState = { chats: cached.chats }
    rerenderContainerList()
    setStatus(
      `Showing ${cached.chats.length} cached chats from ${formatAge(ageMs(cached))}. Refreshing\u2026`,
    )
  } else {
    setStatus("Loading chats\u2026")
    el<HTMLUListElement>("chats").innerHTML = ""
  }
  const loadBtn = el<HTMLButtonElement>("loadchats")
  loadBtn.disabled = true

  const range: ChatRange = userPrefs.chatRange ?? { kind: "last-7d" }
  const { cutoffMs, untilMs } = computeChatWindow(range, chatPrefs)
  const rangeStr = chatRangeLabel(range, chatPrefs)

  const kept: TeamsChatItem[] = []
  let cursor: string | null = null
  let consecutiveOutOfWindowPages = 0
  const PAGE_HARD_CAP = 30

  try {
    let pageCount = 0
    do {
      const page = await listChatsPage(msal, cursor)
      cursor = page.nextCursor
      pageCount++

      const inWindow = page.chats.filter((c) => {
        const t = chatActivityDate(c)
        return t >= cutoffMs && t <= untilMs
      })

      const allBelowCutoff = page.chats.every(
        (c) => chatActivityDate(c) < cutoffMs,
      )

      if (inWindow.length > 0) {
        kept.push(...inWindow)
        consecutiveOutOfWindowPages = 0
      } else if (allBelowCutoff) {
        consecutiveOutOfWindowPages++
      } else {
        consecutiveOutOfWindowPages = 0
      }

      setStatus(`Loading recent containers\u2026 (${kept.length} in window so far)`)

      if (consecutiveOutOfWindowPages >= 2) {
        cursor = null
      }
    } while (cursor !== null && pageCount < PAGE_HARD_CAP)

    // Marked-include enrichment: fetch any marked chats that fell outside the
    // window so they always appear, regardless of timestamp staleness.
    const markedIncludeOn = userPrefs.markedInclude !== false // default ON
    if (markedIncludeOn && markedIds.size > 0) {
      const keptIds = new Set(kept.map((c) => c.id))
      // Favorites are per-stream keys; collapse to distinct owning chatIds before
      // resolving any that fell outside the window (a chatId::rec recordings-stream
      // favorite must still pull its chat into view).
      const missingMarked = [...favoritedChatIds()].filter((id) => !keptIds.has(id))
      if (missingMarked.length > 0) {
        setStatus(`Fetching ${missingMarked.length} favorited container(s) outside window\u2026`)
        for (const id of missingMarked) {
          try {
            const chat = await fetchChatById(msal, id)
            if (chat) kept.push(chat)
          } catch (err) {
            console.warn(
              "[m365-pull] Skipping favorited container (fetch failed):",
              id,
              (err as Error).message,
            )
          }
        }
      }
    }

    kept.sort((a, b) => chatActivityDate(b) - chatActivityDate(a))

    chatsState = { chats: kept }
    rerenderContainerList()
    saveCachedChats(userKey, kept)
    updateLoadStep("chats", {
      state: "done",
      detail: `Found ${kept.length} chat${kept.length === 1 ? "" : "s"} \u00b7 \u2605 ${favoritedChatIds().size} favorites`,
    })
    setStatus(`${kept.length} chats (${rangeStr}).`)
  } catch (err) {
    if (kept.length > 0) {
      chatsState = { chats: kept }
      rerenderContainerList()
      saveCachedChats(userKey, kept)
      updateLoadStep("chats", {
        state: "done",
        detail: `Found ${kept.length} chats (partial) \u00b7 \u2605 ${favoritedChatIds().size} favorites`,
      })
    } else {
      // Hard failure with nothing loaded \u2014 surface the error + Retry on the chats
      // line and STOP (don't reveal §2/§3 or collapse to a receipt).
      updateLoadStep("chats", { state: "error", error: (err as Error).message })
    }
    setStatus(
      `Load failed: ${(err as Error).message}${
        cached
          ? " \u2014 showing cached."
          : kept.length > 0
            ? ` \u2014 showing ${kept.length} partially loaded chats.`
            : ""
      }`,
      "error",
    )
  } finally {
    loadBtn.disabled = false
  }

  // If nothing loaded, stay in the checklist (chats step shows error + Retry).
  if (chatsState.chats.length === 0) {
    updateLoadStep("recordings", { state: "pending", detail: undefined })
    return
  }

  // Build B: Recordings OFF => SKIP the recording scan entirely (real fetch savings).
  // The recordings checklist step was omitted in initLoadSteps, so just finalize the
  // load directly (backgroundLoadRecordings would normally call finalizeLoad).
  if (!loadIncludeRecordings) {
    rerenderContainerList() // messages-only list (no recording artifacts)
    finalizeLoad()
    return
  }

  // After chats load, start background recordings scan using the same window.
  const range2 = userPrefs.chatRange ?? { kind: "last-7d" }
  const { cutoffMs: fromMs, untilMs: toMs } = computeChatWindow(range2, chatPrefs)
  void backgroundLoadRecordings(fromMs, toMs)
}

async function refreshChats(): Promise<void> {
  selectedArtifacts.clear() // ephemeral — artifact IDs may change after reload
  clearCachedChats(userCacheKey())
  chatsState = { chats: [] }
  el<HTMLUListElement>("chats").innerHTML = ""
  // Reset recordings state; initialLoadChats will re-kick the scan.
  recordingsState = { containers: [], chatsScanned: 0, truncated: false }
  recordingsMap = new Map()
  await initialLoadChats()
}

// ----- Background recordings scan -----

/** Scan recordings using the unified chat window. Runs in the background after
 * chats load; does NOT block the container list. Renders rows with a pending
 * indicator while scanning, then updates counts when the scan lands. */
async function backgroundLoadRecordings(fromMs: number, toMs: number): Promise<void> {
  lastRecordingWindow = { fromMs, toMs }
  recordingsMap = new Map()
  rerenderContainerList()
  updateLoadStep("recordings", { state: "running", detail: "Scanning recordings\u2026", error: undefined })
  try {
    const result = await listRecordings(msal, {
      fromMs,
      toMs,
      onProgress: (note) => {
        setScanStatus(`\uD83C\uDF99 ${note}`)
        // Surface real scan progress (e.g. "38/142") on the checklist line.
        updateLoadStep("recordings", { state: "running", detail: `Scanning recordings\u2026 ${note}` })
      },
    })
    recordingsState.containers = result.containers
    recordingsState.chatsScanned = result.chatsScanned
    recordingsState.truncated = result.truncated
    recordingsMap = new Map(result.containers.map((c) => [c.chatId, c]))
    const recTotal = result.containers.reduce((s, c) => s + c.recordings.length, 0)
    const withRecs = result.containers.filter((c) => c.recordings.length > 0).length
    const rangeStr = chatRangeLabel(userPrefs.chatRange ?? { kind: "last-7d" }, chatPrefs)
    updateLoadStep("recordings", {
      state: "done",
      detail: `Scanned recordings \u00b7 ${recTotal} found across ${withRecs} chat${withRecs === 1 ? "" : "s"}`,
    })
    setStatus(
      `${chatsState.chats.length} chats (${rangeStr}) \u00b7 ${recTotal} recording(s) across ${withRecs} chat(s)${result.truncated ? " \u00b7 (chat list truncated \u2014 narrow window)" : ""}.`,
    )
  } catch (err) {
    console.warn("[m365-pull] Recordings scan failed:", err)
    // Per-step error + Retry on the recordings line only; chats are still loaded,
    // so we still finalize (reveal §2/§3 + receipt) in the finally below.
    updateLoadStep("recordings", { state: "error", error: `Recordings scan failed: ${(err as Error).message}` })
    setStatus(`Recordings scan failed: ${(err as Error).message}`, "error")
  } finally {
    rerenderContainerList()
    setScanStatus("") // clear progress; row indicators show per-row state
    finalizeLoad()
  }
}

// ----- Bulk buttons -----

/** Count favorited STREAMS among loaded chats (messages + recordings count
 * separately) and update the Sync-favorites button. */
function countFavoritedStreams(): number {
  let n = 0
  for (const c of chatsState.chats) {
    // Build B: only count streams that are in the captured load scope, so the
    // "Sync favorites (N)" count matches what syncFavorites will actually pull.
    if (loadIncludeMessages && isMessagesFavorited(c.id)) n++
    if (loadIncludeRecordings && isRecordingsFavorited(c.id)) n++
  }
  return n
}

/** Update the Sync-favorites button. Pulls every favorited stream when clicked. */
function updateBulkButtons(): void {
  const streams = countFavoritedStreams()
  const containerBtn = document.getElementById("bulk-containers") as HTMLButtonElement | null
  if (containerBtn) {
    if (streams > 0) {
      containerBtn.hidden = false
      containerBtn.textContent = `Sync favorites (${streams})`
      // Honest copy: there is no backend/background. Sync runs on click.
      containerBtn.title = "Syncs your favorited chats & recordings now (runs when you click)."
    } else {
      containerBtn.hidden = true
    }
  }
}

/** Compute selected artifact count and update the #download-selected button label/visibility. */
function updateSelectedButton(): void {
  const btn = document.getElementById("download-selected") as HTMLButtonElement | null
  const count = selectedArtifacts.size
  if (btn) {
    if (count > 0) {
      btn.hidden = false
      btn.textContent = `Download selected (${count})`
    } else {
      btn.hidden = true
    }
  }
  // §3 global "Select all" checkbox reflects whole-list selection state.
  const globalCb = document.getElementById("select-all-global") as HTMLInputElement | null
  if (globalCb) {
    let total = 0
    for (const c of chatsState.chats) {
      total += getArtifactIds(c.id, recordingsMap.get(c.id)).length
    }
    globalCb.checked = count > 0 && count === total
    globalCb.indeterminate = count > 0 && count < total
  }
}

// ----- Downloading a chat -----

/** Compute the real YYYY-MM-DD span of the messages INCLUDED in a download.
 * Prefers the actual min/max message timestamps (so "all" and "since last
 * download" resolve to concrete dates, not the literal words). Falls back to
 * the resolved query window (since -> today) when no messages were included
 * (e.g. an empty incremental pull), so the filename still carries a real span. */
function computeMessageRange(
  messages: TeamsChatMessage[],
  since: Date | null,
): { rangeStart: string; rangeEnd: string } {
  const times = messages
    .map((m) => (m.createdDateTime ? new Date(m.createdDateTime).getTime() : NaN))
    .filter((t) => !Number.isNaN(t))
  if (times.length > 0) {
    return {
      rangeStart: formatDateStamp(new Date(Math.min(...times))),
      rangeEnd: formatDateStamp(new Date(Math.max(...times))),
    }
  }
  const end = new Date()
  const start = since ?? end
  return {
    rangeStart: formatDateStamp(start),
    rangeEnd: formatDateStamp(end),
  }
}

async function downloadChat(
  chatId: string,
  chatName: string,
  button: HTMLButtonElement,
): Promise<boolean> {
  const lookbackSel = el<HTMLSelectElement>("lookback")
  const lookback = lookbackSel.value

  let since: Date | null = null
  let sinceLabel: string
  let usingIncremental = false

  if (lookback === "all") {
    sinceLabel = "all-time"
  } else if (lookback === "since-last-download") {
    const priorIso = chatPrefs[chatId]?.lastSync
    if (priorIso) {
      since = new Date(priorIso)
      sinceLabel = `since last download (${formatDateShort(since)})`
      usingIncremental = true
    } else {
      since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000)
      sinceLabel = `last 30 days (no prior download for this chat)`
    }
  } else {
    since = new Date(Date.now() - parseInt(lookback, 10) * 24 * 60 * 60 * 1000)
    sinceLabel = `since ${formatDateShort(since)}`
  }

  setStatus(`Fetching "${chatName}" ${sinceLabel}\u2026`)
  button.disabled = true
  const originalLabel = button.textContent
  button.textContent = "Fetching\u2026"
  try {
    const opts: Parameters<typeof fetchChatMessages>[2] = {
      onProgress: ({ count, oldestSeen }) => {
        const back = oldestSeen ? `back to ${formatDateShort(oldestSeen)}` : "scanning\u2026"
        setStatus(`Fetching "${chatName}" ${sinceLabel} \u00b7 ${count} messages, ${back}`)
        button.textContent = `Fetching\u2026 (${count})`
      },
    }
    if (since) opts.since = since
    const messages = await fetchChatMessages(msal, chatId, opts)
    const destination = userPrefs.destination
    const pulledAt = new Date()
    // Phase 3: versioned name — the real YYYY-MM-DD span of the INCLUDED messages
    // plus the pulled-at stamp. Same name for BOTH destinations so every pull is
    // its own dated file (sort-by-name reveals the version history).
    const { rangeStart, rangeEnd } = computeMessageRange(messages, since)
    const versionedName = buildChatArchiveFilename(chatName, {
      pulledAt,
      rangeStart,
      rangeEnd,
      extension: ".md",
    })
    const nowIso = pulledAt.toISOString()
    const chatType = chatsState.chats.find((c) => c.id === chatId)?.chatType ?? "chat"
    const markdownBody = renderChatMarkdown(messages, {
      title: chatName,
      chatId,
      chatType,
      fetchedAt: nowIso,
      lookback,
      sinceIso: since?.toISOString() ?? null,
    })

    let result: { saved: boolean; reason?: string; path?: string; webUrl?: string }
    if (destination === "onedrive") {
      const fullPath = `${userPrefs.oneDriveFolder.replace(/\/$/, "")}/${versionedName}`
      setStatus(`${messages.length} messages fetched. Saving to OneDrive (${userPrefs.oneDriveFolder})\u2026`)
      result = await saveTextToOneDrive(msal, fullPath, markdownBody, "text/markdown")
    } else {
      setStatus(`${messages.length} messages fetched. Saving\u2026`)
      result = await saveAsText(versionedName, markdownBody, {
        extension: ".md",
        description: "Markdown",
        mimeType: "text/markdown",
      })
    }

    if (result.saved) {
      chatPrefs = {
        ...chatPrefs,
        [chatId]: {
          ...(chatPrefs[chatId] ?? {}),
          lastSync: new Date().toISOString(),
        },
      }
      saveChatPrefs(userCacheKey(), chatPrefs)
      scheduleOneDriveSave()
      rerenderContainerList()
      const incrementalNote = usingIncremental ? " (incremental)" : ""
      if (destination === "onedrive") {
        void refreshOneDriveFolderLink()
        const where = result.path ? `OneDrive (${result.path})` : "OneDrive"
        setStatus(
          `\u2713 Saved ${messages.length} messages from "${chatName}" to ${where}${incrementalNote}.`,
        )
      } else {
        setStatus(
          `\u2713 Saved ${messages.length} messages from "${chatName}"${incrementalNote}.`,
        )
      }
      return true
    } else if (result.reason === "cancelled") {
      setStatus(`Save cancelled. (${messages.length} messages fetched but not written.)`)
      return false
    } else if (result.reason === "unsupported") {
      setStatus(
        "Browser save not supported here \u2014 use Microsoft Edge or another Chromium-based browser.",
        "error",
      )
      return false
    } else {
      setStatus(`Save failed: ${result.reason}`, "error")
      return false
    }
  } catch (err) {
    setStatus(`Error: ${(err as Error).message}`, "error")
    return false
  } finally {
    button.disabled = false
    button.textContent = originalLabel || "Download"
  }
}

// ----- Downloading recording transcripts -----

/** Outcome of a single transcript download. "cross-tenant" is NOT a failure \u2014
 * the recording lives in another org's SharePoint and isn't accessible via this
 * account; callers count it separately from real failures. */
type TranscriptOutcome = "ok" | "fail" | "cross-tenant"

async function downloadRecordingTranscript(
  recordingId: string,
  button: HTMLButtonElement,
): Promise<TranscriptOutcome> {
  // Search all containers for the recording
  let recording: RecordingItem | undefined
  for (const container of recordingsState.containers) {
    recording = container.recordings.find((r) => r.id === recordingId)
    if (recording) break
  }
  if (!recording) {
    setStatus("Recording not found in current list. Refresh and try again.", "error")
    return "fail"
  }
  const originalLabel = button.textContent
  button.disabled = true
  button.textContent = "Resolving\u2026"
  const subject = recording.chatTopic?.trim() || recording.filename
  try {
    const resolved = await resolveRecordingFromUrl(msal, recording.url)
    button.textContent = "Fetching transcripts\u2026"
    setStatus(`Fetching transcripts for "${subject}"\u2026`)
    const payload = await fetchRecordingTranscripts(msal, resolved, {
      onProgress: ({ stage, count, total }) => {
        const counter =
          typeof count === "number" && typeof total === "number"
            ? ` (${count}/${total})`
            : ""
        setStatus(`${subject}: ${stage}${counter}`)
      },
    })
    if (payload.transcriptCount === 0) {
      setStatus(
        `No transcripts attached to "${subject}". (Transcription may not have been enabled.)`,
      )
      return "fail"
    }
    button.textContent = "Saving\u2026"
    const account = msal.getActiveAccount()
    const userOid = account?.localAccountId ?? null
    const filename = buildTranscriptFilename(recording, userOid, ".md")
    const destination = userPrefs.destination
    let result: { saved: boolean; reason?: string; path?: string; webUrl?: string }

    const combinedVtt = payload.transcripts
      .map((t) => t.content)
      .join("\n\n")
    const attendees =
      recording.participants.length > 0
        ? recording.participants
            .filter((p) => p.kind === "user")
            .map((p) => p.displayName)
            .filter(Boolean)
            .join(", ")
        : ""
    const dateLabel = formatDate(recording.eventCreatedDateTime)
    const kindLabel =
      recording.chatType === "oneOnOne"
        ? "1:1"
        : recording.chatType === "group"
          ? "Group"
          : "Meeting"
    const metadata = [
      { label: "Date", value: dateLabel },
      { label: "Chat type", value: kindLabel },
    ]
    if (recording.callId) {
      metadata.push({ label: "Call ID", value: recording.callId })
    }
    if (attendees) {
      metadata.push({ label: "Attendees", value: attendees })
    }
    const markdownBody = vttToMarkdown(combinedVtt, {
      title: subject,
      sourceUrl: recording.url,
      metadata,
    })

    if (destination === "onedrive") {
      const fullPath = `${userPrefs.oneDriveFolder.replace(/\/$/, "")}/${filename}`
      result = await saveTextToOneDrive(
        msal,
        fullPath,
        markdownBody,
        "text/markdown",
      )
    } else {
      result = await saveAsText(filename, markdownBody, {
        extension: ".md",
        description: "Markdown",
        mimeType: "text/markdown",
      })
    }
    if (result.saved) {
      recordingPrefs = {
        ...recordingPrefs,
        [recording.id]: {
          ...(recordingPrefs[recording.id] ?? {}),
          lastSync: new Date().toISOString(),
        },
      }
      saveRecordingPrefs(userCacheKey(), recordingPrefs)
      scheduleOneDriveSave()
      rerenderContainerList()
      if (destination === "onedrive") {
        const where = result.path ? `OneDrive (${result.path})` : "OneDrive"
        setStatus(`\u2713 Saved ${payload.transcriptCount} transcript(s) for "${subject}" to ${where}.`)
      } else {
        setStatus(`\u2713 Saved ${payload.transcriptCount} transcript(s) for "${subject}".`)
      }
      return "ok"
    } else if (result.reason === "cancelled") {
      setStatus(`Save cancelled. (${payload.transcriptCount} transcripts fetched but not written.)`)
      return "fail"
    } else {
      setStatus(`Save failed: ${result.reason}`, "error")
      return "fail"
    }
  } catch (err) {
    // Cross-tenant recording: not a real failure \u2014 the .mp4 lives in another
    // org's SharePoint and isn't reachable via this account.
    if ((err as { crossTenant?: boolean }).crossTenant) {
      setStatus(
        `\u2297 "${subject}" \u2014 from another organization \u2014 Microsoft won't let you download it.`,
      )
      return "cross-tenant"
    }
    setStatus(`Error: ${(err as Error).message}`, "error")
    console.error("[m365-pull] downloadRecordingTranscript failed:", err)
    return "fail"
  } finally {
    button.disabled = false
    button.textContent = originalLabel || "Download transcript"
  }
}

/** Sync all recordings in a container row. */
async function downloadContainerTranscripts(
  chatId: string,
  button: HTMLButtonElement,
): Promise<void> {
  const container = recordingsMap.get(chatId)
    ?? recordingsState.containers.find((c) => c.chatId === chatId)
  if (!container) {
    setStatus("Container not found in current list. Refresh and try again.", "error")
    return
  }
  if (container.recordings.length === 0) {
    setStatus("No recordings in range for this container.", "error")
    return
  }
  const originalLabel = button.textContent
  button.disabled = true
  let ok = 0
  let fail = 0
  let crossTenant = 0
  for (let i = 0; i < container.recordings.length; i++) {
    const rec = container.recordings[i]
    button.textContent = `Downloading ${i + 1}/${container.recordings.length}\u2026`
    const tempBtn = document.createElement("button")
    const outcome = await downloadRecordingTranscript(rec.id, tempBtn)
    if (outcome === "ok") ok++
    else if (outcome === "cross-tenant") crossTenant++
    else fail++
  }
  button.disabled = false
  button.textContent = originalLabel || "Download"
  setStatus(
    `Download complete \u2014 saved ${ok} transcript${ok !== 1 ? "s" : ""}${crossTenant > 0 ? `, ${crossTenant} from another organization (unavailable)` : ""}${fail > 0 ? `, ${fail} didn\u2019t come through` : ""}.`,
  )
}

// ----- Unified container sync -----

/** Sync a single container: messages and/or transcripts depending on the
 * global include-messages / include-recordings toggles. Files stay separate. */
async function syncContainer(
  chatId: string,
  chatName: string,
  button: HTMLButtonElement,
): Promise<boolean> {
  // Build B: use the CAPTURED load-time scope so the per-chat download matches the
  // listed experience (not the live chips, which can't change post-load anyway).
  const includeMessages = loadIncludeMessages
  const includeRecordings = loadIncludeRecordings

  if (!includeMessages && !includeRecordings) {
    setStatus("No artifact type selected \u2014 enable Include: Messages and/or Include: Recordings.", "error")
    return false
  }

  let anyOk = false

  if (includeMessages) {
    const ok = await downloadChat(chatId, chatName, button)
    if (ok) anyOk = true
  }

  if (includeRecordings) {
    const container = recordingsMap.get(chatId)
      ?? recordingsState.containers.find((c) => c.chatId === chatId)
    if (container && container.recordings.length > 0) {
      await downloadContainerTranscripts(chatId, button)
      anyOk = true
    }
  }

  return anyOk
}

// ----- Sync favorites (every favorited stream, pulled when the user clicks) -----

/** Pull every favorited STREAM among the loaded chats:
 *   favorited messages stream  -> downloadChat(chatId)
 *   favorited recordings stream -> all of that chat's recordings
 * There is NO backend/background — this runs only on click. */
async function syncFavorites(): Promise<void> {
  type StreamTask =
    | { kind: "messages"; chat: TeamsChatItem }
    | { kind: "recordings"; chat: TeamsChatItem }
  const tasks: StreamTask[] = []
  for (const chat of chatsState.chats) {
    // Build B: only sync streams that are in the captured load scope, matching the
    // listed experience. A suppressed type is simply not acted on; the stored
    // favorite is untouched (not deleted).
    if (loadIncludeMessages && isMessagesFavorited(chat.id)) tasks.push({ kind: "messages", chat })
    if (loadIncludeRecordings && isRecordingsFavorited(chat.id)) tasks.push({ kind: "recordings", chat })
  }
  if (tasks.length === 0) return

  const bulkBtn = el<HTMLButtonElement>("bulk-containers")
  const originalLabel = bulkBtn.textContent
  bulkBtn.disabled = true
  let ok = 0
  let fail = 0
  for (let i = 0; i < tasks.length; i++) {
    const task = tasks[i]
    bulkBtn.textContent = `Syncing ${i + 1}/${tasks.length}\u2026`
    const rowBtn =
      (document.querySelector(
        `.container-action[data-chat-id="${CSS.escape(task.chat.id)}"]`,
      ) as HTMLButtonElement | null) ?? document.createElement("button")
    if (task.kind === "messages") {
      const success = await downloadChat(task.chat.id, chatDisplayName(task.chat), rowBtn)
      if (success) ok++
      else fail++
    } else {
      // Recordings stream: pull every recording in this chat. downloadContainerTranscripts
      // no-ops with a status if the chat has no (scanned) recordings.
      const container =
        recordingsMap.get(task.chat.id) ??
        recordingsState.containers.find((c) => c.chatId === task.chat.id)
      if (container && container.recordings.length > 0) {
        await downloadContainerTranscripts(task.chat.id, rowBtn)
        ok++
      } else {
        fail++
      }
    }
  }
  bulkBtn.disabled = false
  bulkBtn.textContent = originalLabel || ""
  setStatus(
    `Sync complete \u2014 ${ok} favorited stream${ok !== 1 ? "s" : ""} pulled${fail > 0 ? `, ${fail} skipped or didn\u2019t come through` : ""}.`,
  )
  updateBulkButtons()
}

// ----- Download selected artifacts (Phase 1 ephemeral selection) -----

async function downloadSelectedArtifacts(): Promise<void> {
  if (selectedArtifacts.size === 0) return
  const btn = document.getElementById("download-selected") as HTMLButtonElement | null
  if (!btn) return
  const artifacts = [...selectedArtifacts]
  const originalLabel = btn.textContent
  btn.disabled = true
  let done = 0
  let ok = 0

  for (const artifactId of artifacts) {
    done++
    btn.textContent = `Downloading ${done}/${artifacts.length}\u2026`
    const tempBtn = document.createElement("button")

    if (artifactId.startsWith("msg:")) {
      const chatId = artifactId.slice(4)
      const chat = chatsState.chats.find((c) => c.id === chatId)
      if (chat) {
        const success = await downloadChat(chatId, chatDisplayName(chat), tempBtn)
        if (success) ok++
      }
    } else if (artifactId.startsWith("rec:")) {
      const recId = artifactId.slice(4)
      const outcome = await downloadRecordingTranscript(recId, tempBtn)
      if (outcome === "ok") ok++
    }
  }

  // Clear selection
  selectedArtifacts.clear()
  btn.disabled = false
  btn.textContent = originalLabel || ""
  updateSelectedButton()
  // Reset all artifact and group checkboxes in the DOM
  document
    .querySelectorAll<HTMLInputElement>(".artifact-check, .select-all-check")
    .forEach((cb) => {
      cb.checked = false
      cb.indeterminate = false
    })
  setStatus(
    `Download complete \u2014 ${ok} of ${artifacts.length} artifact${artifacts.length !== 1 ? "s" : ""} saved.`,
  )
}

// ----- Settings modal -----

function openSettingsModal(): void {
  const modal = el<HTMLDivElement>("settings-modal")
  el<HTMLInputElement>("settings-folder").value = userPrefs.oneDriveFolder
  el<HTMLSelectElement>("settings-destination").value = userPrefs.destination
  modal.hidden = false
  setTimeout(() => el<HTMLInputElement>("settings-folder").focus(), 50)
}

function closeSettingsModal(): void {
  el<HTMLDivElement>("settings-modal").hidden = true
}

function saveSettingsModal(): void {
  const rawFolder = el<HTMLInputElement>("settings-folder").value.trim()
  const destination = el<HTMLSelectElement>("settings-destination")
    .value as Destination

  if (!rawFolder) {
    setStatus("Settings: folder path cannot be empty.", "error")
    return
  }
  const normalized =
    (rawFolder.startsWith("/") ? rawFolder : "/" + rawFolder)
      .replace(/\\+/g, "/")
      .replace(/\/+/g, "/")
      .replace(/\/$/, "") || "/"

  const changed =
    normalized !== userPrefs.oneDriveFolder ||
    destination !== userPrefs.destination

  userPrefs = {
    ...userPrefs,
    oneDriveFolder: normalized,
    destination,
  }
  saveUserPrefs(userCacheKey(), userPrefs)
  syncUserPrefsToUI()
  if (changed) scheduleOneDriveSave()
  closeSettingsModal()
  setStatus("Settings saved.")
}

render()
