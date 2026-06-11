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
} from "./sources/teams-chats"
import {
  listRecordings,
  buildTranscriptFilename,
  type ChatType,
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
import { saveTextToOneDrive } from "./destinations/onedrive"
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
  type RecordingRange,
} from "./state/onedrive-state"

const app = document.getElementById("app") as HTMLDivElement

if (!config.clientId || !config.tenantId) {
  app.innerHTML =
    '<p class="empty">Missing MSAL config — set <code>clientId</code> and <code>tenantId</code> in <code>src/config.ts</code>. See README.md.</p>'
  throw new Error("Missing MSAL config")
}

const msal = new PublicClientApplication({
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
    redirectUri: window.location.origin,
  },
})

await msal.initialize()
const redirectResp = await msal.handleRedirectPromise()
if (redirectResp?.account) msal.setActiveAccount(redirectResp.account)

// ----- State -----

type SourceId = "teams.chats" | "teams.recordings"

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
}

interface RecordingFilterState {
  search: string
  sortKey: SortKey
  markedOnly: boolean
  enabledKinds: Set<ChatType>
}

type SyncStatus = "idle" | "syncing" | "synced" | "error" | "offline"

const KNOWN_TYPES: { id: string; label: string }[] = [
  { id: "oneOnOne", label: "1:1" },
  { id: "group", label: "Group" },
  { id: "meeting", label: "Meeting" },
]

const CHAT_TYPE_KINDS: { id: ChatType; label: string }[] = [
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

let currentSource: SourceId = "teams.chats"
let chatsState: ChatsState = { chats: [] }
let recordingsState: RecordingsState = { containers: [], chatsScanned: 0, truncated: false }
let filterState: FilterState = {
  search: "",
  enabledTypes: new Set(KNOWN_TYPES.map((t) => t.id)),
  sortKey: "marked-first",
  markedOnly: false,
}
let recordingFilterState: RecordingFilterState = {
  search: "",
  sortKey: "marked-first",
  markedOnly: false,
  enabledKinds: new Set(CHAT_TYPE_KINDS.map((k) => k.id)),
}
let markedIds: Set<string> = new Set()
let chatPrefs: Record<string, ChatPrefs> = {}
let recordingPrefs: Record<string, RecordingPrefs> = {}
let userPrefs: UserPrefs = {
  destination: "browser",
  oneDriveFolder: "/m365-pull/teams-chats",
}
let syncStatus: SyncStatus = "idle"
let lastSyncedAt: Date | null = null
let lastSyncError: string | null = null

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

// ----- Recording range helpers -----

const DAY_MS = 24 * 60 * 60 * 1000

/** Format a Date as "yyyy-mm-dd" in local time for date inputs. */
function toLocalDateString(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, "0")
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
}

/** Midnight (00:00:00.000) of the most recent Monday in local time. */
function startOfThisWeekMonday(): number {
  const now = new Date()
  const dayOfWeek = now.getDay() // 0=Sun, 1=Mon, … 6=Sat
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

/** Map a persisted RecordingRange to a concrete [fromMs, toMs] window.
 * For since-last-download with no prior sync, falls back to last-7d. */
function computeRecordingWindow(range: RecordingRange): { fromMs: number; toMs: number } {
  const now = Date.now()
  switch (range.kind) {
    case "this-week":
      return { fromMs: startOfThisWeekMonday(), toMs: now }
    case "last-7d":
      return { fromMs: now - 7 * DAY_MS, toMs: now }
    case "last-30d":
      return { fromMs: now - 30 * DAY_MS, toMs: now }
    case "since-last-download": {
      const t = mostRecentRecordingSync()
      return { fromMs: t !== null ? t : now - 7 * DAY_MS, toMs: now }
    }
    case "custom": {
      const fromMs = range.customFrom ? startOfLocalDay(range.customFrom) : now - 7 * DAY_MS
      const toMs = range.customTo ? endOfLocalDay(range.customTo) : now
      return { fromMs, toMs }
    }
  }
}

/** Human-readable label for a RecordingRange (used in status messages). */
function rangeLabel(range: RecordingRange): string {
  switch (range.kind) {
    case "this-week": return "this week"
    case "last-7d": return "last 7 days"
    case "last-30d": return "last 30 days"
    case "since-last-download": {
      const t = mostRecentRecordingSync()
      if (t === null) return "since last download (none yet \u2014 showing 7 days)"
      return `since last download (${formatDateShort(new Date(t))})`
    }
    case "custom":
      if (range.customFrom && range.customTo) return `${range.customFrom} to ${range.customTo}`
      if (range.customFrom) return `from ${range.customFrom}`
      return "custom range"
  }
}

// ----- Chat range helpers -----

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

/** Most recent recordingPrefs.lastSync timestamp across all synced recordings, or null. */
function mostRecentRecordingSync(): number | null {
  let most: number | null = null
  for (const v of Object.values(recordingPrefs)) {
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

  // Recording range
  const range: RecordingRange = userPrefs.recordingRange ?? { kind: "last-7d" }
  const rangeEl = document.getElementById("recording-range") as HTMLSelectElement | null
  if (rangeEl) rangeEl.value = range.kind
  const customEl = document.getElementById("custom-range-inputs") as HTMLSpanElement | null
  if (customEl) customEl.hidden = range.kind !== "custom"
  if (range.kind === "custom") {
    const fromEl = document.getElementById("recording-from") as HTMLInputElement | null
    const toEl = document.getElementById("recording-to") as HTMLInputElement | null
    if (fromEl && range.customFrom) fromEl.value = range.customFrom
    if (toEl && range.customTo) toEl.value = range.customTo
  }

  // Chat range
  const chatRange: ChatRange = userPrefs.chatRange ?? { kind: "last-7d" }
  const chatRangeEl = document.getElementById("chat-range") as HTMLSelectElement | null
  if (chatRangeEl) chatRangeEl.value = chatRange.kind
  const chatCustomEl = document.getElementById("chat-custom-range-inputs") as HTMLSpanElement | null
  if (chatCustomEl) chatCustomEl.hidden = chatRange.kind !== "custom"
  if (chatRange.kind === "custom") {
    const chatFromEl = document.getElementById("chat-from") as HTMLInputElement | null
    const chatToEl = document.getElementById("chat-to") as HTMLInputElement | null
    if (chatFromEl && chatRange.customFrom) chatFromEl.value = chatRange.customFrom
    if (chatToEl && chatRange.customTo) chatToEl.value = chatRange.customTo
  }

  // Chat marked-include toggle (default ON: undefined → true)
  const markedIncludeBtn = document.getElementById("marked-include") as HTMLButtonElement | null
  if (markedIncludeBtn)
    markedIncludeBtn.classList.toggle("active", userPrefs.markedInclude !== false)

  // Recording marked-include toggle (default ON: undefined → true)
  const recMarkedIncludeBtn = document.getElementById("rec-marked-include") as HTMLButtonElement | null
  if (recMarkedIncludeBtn)
    recMarkedIncludeBtn.classList.toggle("active", userPrefs.recordingMarkedInclude !== false)

  // Hide-downloaded toggle
  const hideBtn = document.getElementById("hide-downloaded") as HTMLButtonElement | null
  if (hideBtn) hideBtn.classList.toggle("active", !!userPrefs.hideDownloaded)
}

function typeLabel(type: string): string {
  return KNOWN_TYPES.find((t) => t.id === type)?.label ?? type
}

function buildLocalState(): AppState {
  return {
    version: 1,
    updatedAt: new Date().toISOString(),
    updatedBy: deviceIdentifier(),
    marks: [...markedIds].sort(),
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
      title = "Sign in to sync marks across devices"
      break
    case "syncing":
      text = "syncing…"
      cls += " syncing"
      title = "Saving state to OneDrive"
      break
    case "synced":
      text = lastSyncedAt
        ? `synced ${formatAge(Date.now() - lastSyncedAt.getTime())}`
        : "synced"
      cls += " synced"
      title = "Marks are saved to OneDrive (/Apps/m365-pull/state.json)"
      break
    case "error":
      text = "sync error"
      cls += " error"
      title = lastSyncError ?? "Failed to sync state — using local only"
      break
    case "offline":
      text = "offline · local only"
      cls += " offline"
      title = "Could not reach OneDrive; marks stay on this device"
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
    const marksChanged =
      mergedMarks.size !== markedIds.size ||
      [...mergedMarks].some((id) => !markedIds.has(id))
    markedIds = mergedMarks
    saveMarks(userCacheKey(), markedIds)
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
      marksChanged || prefsChanged || recordingPrefsChanged || userPrefsChanged
    if (changed) {
      await saveOneDriveState(msal, {
        ...merged,
        updatedAt: new Date().toISOString(),
        updatedBy: deviceIdentifier(),
      })
      rerenderChatList()
      rerenderRecordingsList()
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

// ----- Filter + sort for chats (pure) -----

function applyFiltersAndSort(chats: TeamsChatItem[]): TeamsChatItem[] {
  const q = filterState.search.trim().toLowerCase()
  let result = chats.filter((c) => {
    if (!filterState.enabledTypes.has(c.chatType)) return false
    if (filterState.markedOnly && !markedIds.has(c.id)) return false
    if (q) {
      const name = chatDisplayName(c).toLowerCase()
      if (!name.includes(q)) return false
    }
    return true
  })
  const byRecent = (a: TeamsChatItem, b: TeamsChatItem) =>
    a.lastUpdatedDateTime < b.lastUpdatedDateTime ? 1 : -1
  const byName = (a: TeamsChatItem, b: TeamsChatItem) =>
    chatDisplayName(a).localeCompare(chatDisplayName(b))
  if (filterState.sortKey === "name") {
    result = [...result].sort(byName)
  } else if (filterState.sortKey === "recent") {
    result = [...result].sort(byRecent)
  } else {
    result = [...result].sort((a, b) => {
      const aM = markedIds.has(a.id) ? 1 : 0
      const bM = markedIds.has(b.id) ? 1 : 0
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

function applyRecordingFiltersAndSort(
  containers: RecordingContainer[],
): RecordingContainer[] {
  const q = recordingFilterState.search.trim().toLowerCase()
  let result = containers.filter((c) => {
    const markId = `rec:${c.chatId}`
    if (recordingFilterState.markedOnly && !markedIds.has(markId)) return false
    if (!recordingFilterState.enabledKinds.has(c.chatType)) return false
    // Hide-downloaded: exclude containers where every recording has been synced
    if (userPrefs.hideDownloaded && c.recordings.length > 0) {
      const allDownloaded = c.recordings.every((r) => !!recordingPrefs[r.id]?.lastSync)
      if (allDownloaded) return false
    }
    if (q) {
      const haystack =
        `${c.chatTopic ?? ""} ${c.recordings.map((r) => r.filename).join(" ")}`.toLowerCase()
      if (!haystack.includes(q)) return false
    }
    return true
  })
  const mostRecentDate = (c: RecordingContainer): string =>
    c.recordings[0]?.eventCreatedDateTime ?? ""
  const byRecent = (a: RecordingContainer, b: RecordingContainer) =>
    mostRecentDate(b).localeCompare(mostRecentDate(a))
  const byName = (a: RecordingContainer, b: RecordingContainer) =>
    (a.chatTopic ?? "").localeCompare(b.chatTopic ?? "")
  if (recordingFilterState.sortKey === "name") {
    result = [...result].sort(byName)
  } else if (recordingFilterState.sortKey === "recent") {
    result = [...result].sort(byRecent)
  } else {
    result = [...result].sort((a, b) => {
      const aM = markedIds.has(`rec:${a.chatId}`) ? 1 : 0
      const bM = markedIds.has(`rec:${b.chatId}`) ? 1 : 0
      if (aM !== bM) return bM - aM
      return byRecent(a, b)
    })
  }
  return result
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
  // Migrate: drop old per-recording marks that used the callId::filename composite key.
  // They don't map to containers; re-mark at container level with the rec: prefix.
  let migratedMarks = false
  for (const id of [...markedIds]) {
    if (id.includes("::")) {
      markedIds.delete(id)
      migratedMarks = true
    }
  }
  if (migratedMarks) saveMarks(userCacheKey(), markedIds)
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
        <button id="open-settings" class="icon-button" title="Settings" aria-label="Settings">⚙</button>
        <button id="signout">Sign out</button>
      </div>
    </header>
    <div id="settings-modal" class="modal" hidden>
      <div class="modal-backdrop"></div>
      <div class="modal-card" role="dialog" aria-labelledby="settings-title" aria-modal="true">
        <header class="modal-header">
          <h2 id="settings-title">Settings</h2>
          <button id="settings-close" class="icon-button" aria-label="Close">✕</button>
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
            <span class="form-help">Marks, last-download timestamps, and these settings sync across all your devices via this single file.</span>
          </div>
        </div>
        <footer class="modal-footer">
          <button id="settings-cancel">Cancel</button>
          <button id="settings-save" class="primary">Save</button>
        </footer>
      </div>
    </div>
    <main>
      <div class="actions actions-global">
        <label class="sort-label">
          Source
          <select id="source-select">
            <option value="teams.chats">Teams chats</option>
            <option value="teams.recordings">Call recordings</option>
          </select>
        </label>
        <span class="label">to</span>
        <select id="destination" title="Where to save downloads">
          <option value="browser">Browser (save dialog)</option>
          <option value="onedrive">OneDrive folder</option>
        </select>
        <span class="onedrive-folder" id="onedrive-folder" hidden>
          <button class="link-button" id="edit-folder" title="Click to change">/m365-pull/teams-chats</button>
        </span>
      </div>
      <div class="actions" id="actions-chats">
        <button id="loadchats" class="primary">Load my Teams chats</button>
        <button id="refreshchats" hidden>Refresh</button>
        <span class="label">Show chats from:</span>
        <select id="chat-range" title="Date range for chat list">
          <option value="this-week">This week</option>
          <option value="last-7d" selected>Last 7 days</option>
          <option value="last-30d">Last 30 days</option>
          <option value="since-last-download">Since last download</option>
          <option value="custom">Custom range\u2026</option>
        </select>
        <span id="chat-custom-range-inputs" class="custom-range" hidden>
          <input type="date" id="chat-from" title="From (inclusive)" />
          <span class="label">to</span>
          <input type="date" id="chat-to" title="To (inclusive)" />
        </span>
        <button class="chip" id="marked-include" title="Always show marked chats regardless of range">\u2605 Always include marked</button>
        <span class="label" aria-hidden="true" style="opacity:0.35;padding:0 0.25rem;">\u2502</span>
        <span class="label">Download history per chat:</span>
        <select id="lookback" title="Download history per chat">
          <option value="7">Last 7 days</option>
          <option value="30" selected>Last 30 days</option>
          <option value="90">Last 90 days</option>
          <option value="all">All messages</option>
          <option value="since-last-download">Since last download</option>
        </select>
        <button id="bulk-chats" class="bulk-action" hidden></button>
      </div>
      <div class="actions" id="actions-transcripts" hidden>
        <button id="load-recordings" class="primary">Load my recordings</button>
        <button id="refresh-recordings" hidden>Refresh</button>
        <span class="label">Range:</span>
        <select id="recording-range" title="Date range for recordings">
          <option value="this-week">This week</option>
          <option value="last-7d" selected>Last 7 days</option>
          <option value="last-30d">Last 30 days</option>
          <option value="since-last-download">Since last download</option>
          <option value="custom">Custom range\u2026</option>
        </select>
        <span id="custom-range-inputs" class="custom-range" hidden>
          <input type="date" id="recording-from" title="From (inclusive)" />
          <span class="label">to</span>
          <input type="date" id="recording-to" title="To (inclusive)" />
        </span>
        <button class="chip" id="rec-marked-include" title="Always show marked containers regardless of range">\u2605 Always include marked</button>
        <button id="bulk-recordings" class="bulk-action" hidden></button>
      </div>

      <div class="filters" id="filters" hidden>
        <input id="search" type="search" placeholder="Search chats by name…" />
        <div class="chips" id="type-chips">
          ${KNOWN_TYPES.map(
            (t) =>
              `<button class="chip active" data-type="${t.id}">${escapeHtml(t.label)} <span class="chip-count">0</span></button>`,
          ).join("")}
        </div>
        <label class="sort-label">
          Sort
          <select id="sortby">
            <option value="marked-first">Marked first · then recent</option>
            <option value="recent">Most recent activity</option>
            <option value="name">Name (A–Z)</option>
          </select>
        </label>
        <button class="chip marked-only" id="markedonly">★ Marked only</button>
      </div>
      <div class="filters" id="filters-recordings" hidden>
        <input id="search-recordings" type="search" placeholder="Search recordings by topic or filename…" />
        <div class="chips" id="kind-chips">
          ${CHAT_TYPE_KINDS.map(
            (k) =>
              `<button class="chip active" data-kind="${k.id}">${escapeHtml(k.label)} <span class="chip-count">0</span></button>`,
          ).join("")}
        </div>
        <label class="sort-label">
          Sort
          <select id="sortby-recordings">
            <option value="marked-first">Marked first · then recent</option>
            <option value="recent">Most recent</option>
            <option value="name">Subject (A–Z)</option>
          </select>
        </label>
        <button class="chip marked-only" id="markedonly-recordings">★ Marked only</button>
        <button class="chip" id="hide-downloaded">Hide downloaded</button>
      </div>
      <div id="status"></div>
      <ul id="chats" class="chat-list"></ul>
      <ul id="recordings" class="chat-list" hidden></ul>
    </main>
  `

  wireGlobalHandlers(account)
  syncUIControlsFromState()
  updateSyncIndicator()
  applySourceVisibility()

  void pullAndMergeOneDriveState().then(() => rerenderChatList())
}

function wireGlobalHandlers(account: AccountInfo): void {
  el<HTMLButtonElement>("signout").addEventListener("click", () => {
    void msal.logoutRedirect({ account })
  })

  // Source switcher
  const sourceSel = el<HTMLSelectElement>("source-select")
  sourceSel.value = currentSource
  sourceSel.addEventListener("change", () => {
    currentSource = sourceSel.value as SourceId
    applySourceVisibility()
    saveUIPrefs()
  })

  // Chats actions -- ensure source state matches before running
  el<HTMLButtonElement>("loadchats").addEventListener("click", () => {
    switchSourceIfNeeded("teams.chats")
    void initialLoadChats()
  })
  el<HTMLButtonElement>("refreshchats").addEventListener("click", () => {
    switchSourceIfNeeded("teams.chats")
    void refreshChats()
  })

  // Transcripts actions
  el<HTMLButtonElement>("load-recordings").addEventListener("click", () => {
    switchSourceIfNeeded("teams.recordings")
    void initialLoadRecordings()
  })
  el<HTMLButtonElement>("refresh-recordings").addEventListener("click", () => {
    switchSourceIfNeeded("teams.recordings")
    void refreshRecordings()
  })

  // Bulk download buttons
  el<HTMLButtonElement>("bulk-chats").addEventListener("click", () => {
    void bulkDownloadChats()
  })
  el<HTMLButtonElement>("bulk-recordings").addEventListener("click", () => {
    void bulkDownloadRecordings()
  })
  // Recording range dropdown + custom date inputs — persisted to OneDrive userPrefs
  el<HTMLSelectElement>("recording-range").addEventListener("change", () => {
    const kind = el<HTMLSelectElement>("recording-range").value as RecordingRange["kind"]
    const customEl = el<HTMLSpanElement>("custom-range-inputs")
    customEl.hidden = kind !== "custom"
    if (kind === "custom") {
      // Default date inputs when first revealed
      const fromEl = el<HTMLInputElement>("recording-from")
      const toEl = el<HTMLInputElement>("recording-to")
      if (!fromEl.value) fromEl.value = toLocalDateString(new Date(Date.now() - 7 * DAY_MS))
      if (!toEl.value) toEl.value = toLocalDateString(new Date())
      userPrefs = { ...userPrefs, recordingRange: {
        kind,
        customFrom: fromEl.value || undefined,
        customTo: toEl.value || undefined,
      } }
    } else {
      userPrefs = { ...userPrefs, recordingRange: { kind } }
    }
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    // If recordings have already been loaded, reload with the new range
    if (!el<HTMLButtonElement>("refresh-recordings").hidden) {
      void refreshRecordings()
    }
  })

  const applyCustomRange = () => {
    const fromEl = document.getElementById("recording-from") as HTMLInputElement | null
    const toEl = document.getElementById("recording-to") as HTMLInputElement | null
    userPrefs = { ...userPrefs, recordingRange: {
      kind: "custom",
      customFrom: fromEl?.value || undefined,
      customTo: toEl?.value || undefined,
    } }
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    // If recordings have already been loaded, reload with the new range
    if (!el<HTMLButtonElement>("refresh-recordings").hidden) {
      void refreshRecordings()
    }
  }
  el<HTMLInputElement>("recording-from").addEventListener("change", applyCustomRange)
  el<HTMLInputElement>("recording-to").addEventListener("change", applyCustomRange)

  el<HTMLButtonElement>("hide-downloaded").addEventListener("click", () => {
    const next = !userPrefs.hideDownloaded
    userPrefs = { ...userPrefs, hideDownloaded: next }
    el<HTMLButtonElement>("hide-downloaded").classList.toggle("active", next)
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    rerenderRecordingsList()
  })

  // Chat range dropdown + custom date inputs — controls the list window; persisted
  // to OneDrive userPrefs (cross-device), mirroring the recording range pattern.
  el<HTMLSelectElement>("chat-range").addEventListener("change", () => {
    const kind = el<HTMLSelectElement>("chat-range").value as ChatRange["kind"]
    const chatCustomEl = el<HTMLSpanElement>("chat-custom-range-inputs")
    chatCustomEl.hidden = kind !== "custom"
    if (kind === "custom") {
      const fromEl = el<HTMLInputElement>("chat-from")
      const toEl = el<HTMLInputElement>("chat-to")
      if (!fromEl.value) fromEl.value = toLocalDateString(new Date(Date.now() - 7 * DAY_MS))
      if (!toEl.value) toEl.value = toLocalDateString(new Date())
      userPrefs = { ...userPrefs, chatRange: {
        kind,
        customFrom: fromEl.value || undefined,
        customTo: toEl.value || undefined,
      } }
    } else {
      userPrefs = { ...userPrefs, chatRange: { kind } }
    }
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    // If chats are already loaded, clear and reload with the new range.
    if (!el<HTMLButtonElement>("refreshchats").hidden) {
      clearCachedChats(userCacheKey())
      chatsState = { chats: [] }
      el<HTMLUListElement>("chats").innerHTML = ""
      void initialLoadChats()
    }
  })

  const applyCustomChatRange = () => {
    const fromEl = document.getElementById("chat-from") as HTMLInputElement | null
    const toEl = document.getElementById("chat-to") as HTMLInputElement | null
    userPrefs = { ...userPrefs, chatRange: {
      kind: "custom",
      customFrom: fromEl?.value || undefined,
      customTo: toEl?.value || undefined,
    } }
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    if (!el<HTMLButtonElement>("refreshchats").hidden) {
      clearCachedChats(userCacheKey())
      chatsState = { chats: [] }
      el<HTMLUListElement>("chats").innerHTML = ""
      void initialLoadChats()
    }
  }
  el<HTMLInputElement>("chat-from").addEventListener("change", applyCustomChatRange)
  el<HTMLInputElement>("chat-to").addEventListener("change", applyCustomChatRange)

  // Marked-include toggle — persisted to OneDrive userPrefs; default ON.
  el<HTMLButtonElement>("marked-include").addEventListener("click", () => {
    const next = userPrefs.markedInclude === false ? true : false
    userPrefs = { ...userPrefs, markedInclude: next }
    el<HTMLButtonElement>("marked-include").classList.toggle("active", next)
    saveUserPrefs(userCacheKey(), userPrefs)
    scheduleOneDriveSave()
    // Reload so the marked-include enrichment runs (or stops running).
    if (!el<HTMLButtonElement>("refreshchats").hidden) {
      clearCachedChats(userCacheKey())
      chatsState = { chats: [] }
      el<HTMLUListElement>("chats").innerHTML = ""
      void initialLoadChats()
    }
  })

  // Recording marked-include toggle — default ON; triggers reload when recordings are loaded.
  const recMarkedIncludeBtn = document.getElementById("rec-marked-include") as HTMLButtonElement | null
  if (recMarkedIncludeBtn) {
    recMarkedIncludeBtn.addEventListener("click", () => {
      const next = userPrefs.recordingMarkedInclude === false ? true : false
      userPrefs = { ...userPrefs, recordingMarkedInclude: next }
      recMarkedIncludeBtn.classList.toggle("active", next)
      saveUserPrefs(userCacheKey(), userPrefs)
      scheduleOneDriveSave()
      if (!el<HTMLButtonElement>("refresh-recordings").hidden) {
        void refreshRecordings()
      }
    })
  }

  // Chat lookback: persists the download-depth selection only. The list window
  // is now owned by the chat-range selector above.
  el<HTMLSelectElement>("lookback").addEventListener("change", () => {
    saveUIPrefs()
  })

  // Meeting filters
  const searchMeetings = el<HTMLInputElement>("search-recordings")
  let searchMeetingsTimer: number | null = null
  searchMeetings.addEventListener("input", () => {
    if (searchMeetingsTimer) window.clearTimeout(searchMeetingsTimer)
    searchMeetingsTimer = window.setTimeout(() => {
      recordingFilterState.search = searchMeetings.value
      rerenderRecordingsList()
      saveUIPrefs()
    }, 150)
  })
  el<HTMLSelectElement>("sortby-recordings").addEventListener("change", (e) => {
    recordingFilterState.sortKey = (e.target as HTMLSelectElement).value as SortKey
    rerenderRecordingsList()
    saveUIPrefs()
  })
  el<HTMLButtonElement>("markedonly-recordings").addEventListener("click", () => {
    recordingFilterState.markedOnly = !recordingFilterState.markedOnly
    el<HTMLButtonElement>("markedonly-recordings").classList.toggle(
      "active",
      recordingFilterState.markedOnly,
    )
    rerenderRecordingsList()
    saveUIPrefs()
  })
  el<HTMLDivElement>("kind-chips")
    .querySelectorAll<HTMLButtonElement>(".chip[data-kind]")
    .forEach((chip) => {
      chip.addEventListener("click", () => {
        const k = chip.dataset.kind as ChatType
        if (recordingFilterState.enabledKinds.has(k)) {
          recordingFilterState.enabledKinds.delete(k)
          chip.classList.remove("active")
        } else {
          recordingFilterState.enabledKinds.add(k)
          chip.classList.add("active")
        }
        rerenderRecordingsList()
        saveUIPrefs()
      })
    })

  // Chats filters
  const search = el<HTMLInputElement>("search")
  let searchTimer: number | null = null
  search.addEventListener("input", () => {
    if (searchTimer) window.clearTimeout(searchTimer)
    searchTimer = window.setTimeout(() => {
      filterState.search = search.value
      rerenderChatList()
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
        rerenderChatList()
        saveUIPrefs()
      })
    })
  el<HTMLSelectElement>("sortby").addEventListener("change", (e) => {
    filterState.sortKey = (e.target as HTMLSelectElement).value as SortKey
    rerenderChatList()
    saveUIPrefs()
  })
  el<HTMLButtonElement>("markedonly").addEventListener("click", () => {
    filterState.markedOnly = !filterState.markedOnly
    el<HTMLButtonElement>("markedonly").classList.toggle(
      "active",
      filterState.markedOnly,
    )
    rerenderChatList()
    saveUIPrefs()
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
  // Inline folder label opens Settings (replaced the old prompt() editor)
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

/** Coerce source state to match the action the user just took. If they clicked
 * a button belonging to a different source (e.g. Load my meetings while still
 * on chats), switch the source for them and update the UI. */
function switchSourceIfNeeded(wanted: SourceId): void {
  if (currentSource === wanted) return
  currentSource = wanted
  const sourceSel = document.getElementById("source-select") as
    | HTMLSelectElement
    | null
  if (sourceSel) sourceSel.value = wanted
  applySourceVisibility()
  saveUIPrefs()
}

/** Read module state and stash it in localStorage so reloads remember
 * the user's dropdown selections, chip filters, and source choice.
 * Note: lookback persists the message-download-depth selection; the list
 * window range is owned by userPrefs.chatRange (synced to OneDrive). */
function saveUIPrefs(): void {
  const lookbackEl = document.getElementById("lookback") as
    | HTMLSelectElement
    | null
  saveUIState(userCacheKey(), {
    currentSource,
    lookback: lookbackEl?.value,
    chatFilter: {
      search: filterState.search,
      enabledTypes: [...filterState.enabledTypes],
      sortKey: filterState.sortKey,
      markedOnly: filterState.markedOnly,
    },
    recordingFilter: {
      search: recordingFilterState.search,
      enabledKinds: [...recordingFilterState.enabledKinds],
      sortKey: recordingFilterState.sortKey,
      markedOnly: recordingFilterState.markedOnly,
    },
  })
}

/** Restore filter state, source choice, and dropdown values from localStorage.
 * Runs after the HTML is rendered (handlers wired later read these values). */
function hydrateUIStateFromStorage(): void {
  const saved = loadUIState(userCacheKey())
  if (saved.currentSource === "teams.chats" || saved.currentSource === "teams.recordings") {
    currentSource = saved.currentSource
  }
  if (saved.chatFilter) {
    filterState = {
      search: saved.chatFilter.search ?? "",
      enabledTypes: new Set(
        saved.chatFilter.enabledTypes ?? KNOWN_TYPES.map((t) => t.id),
      ),
      sortKey: saved.chatFilter.sortKey ?? "marked-first",
      markedOnly: saved.chatFilter.markedOnly ?? false,
    }
  }
  if (saved.recordingFilter) {
    recordingFilterState = {
      search: saved.recordingFilter.search ?? "",
      enabledKinds: new Set(
        (saved.recordingFilter.enabledKinds ?? CHAT_TYPE_KINDS.map((k) => k.id)).filter(
          (k): k is ChatType => k === "oneOnOne" || k === "group" || k === "meeting",
        ),
      ),
      sortKey: saved.recordingFilter.sortKey ?? "marked-first",
      markedOnly: saved.recordingFilter.markedOnly ?? false,
    }
  }
  // Note: lookback dropdown value is restored in syncUIControlsFromState
  // (called after handlers are wired and the DOM is interactive).
}

/** Push module state into the freshly-rendered DOM controls. Called once
 * after wireGlobalHandlers so handler-attached defaults don't fight us. */
function syncUIControlsFromState(): void {
  const saved = loadUIState(userCacheKey())
  // Source dropdown
  const sourceSel = document.getElementById("source-select") as HTMLSelectElement | null
  if (sourceSel) sourceSel.value = currentSource
  // Chat lookback (download depth only — list range is synced via syncUserPrefsToUI)
  const lookbackEl = document.getElementById("lookback") as HTMLSelectElement | null
  if (lookbackEl && saved.lookback) lookbackEl.value = saved.lookback
  // Recording range + chat range are synced via syncUserPrefsToUI (called from wireGlobalHandlers)
  // Chat filter UI: search, sortby, type chips, markedonly
  const searchEl = document.getElementById("search") as HTMLInputElement | null
  if (searchEl) searchEl.value = filterState.search
  const sortbyEl = document.getElementById("sortby") as HTMLSelectElement | null
  if (sortbyEl) sortbyEl.value = filterState.sortKey
  const markedOnlyChat = document.getElementById("markedonly")
  if (markedOnlyChat) markedOnlyChat.classList.toggle("active", filterState.markedOnly)
  document
    .querySelectorAll<HTMLButtonElement>("#type-chips .chip[data-type]")
    .forEach((chip) => {
      const t = chip.dataset.type
      const enabled = t ? filterState.enabledTypes.has(t) : true
      chip.classList.toggle("active", enabled)
    })
  // Meeting filter UI: search, sortby, kind chips, markedonly
  const searchMeetingsEl = document.getElementById("search-recordings") as HTMLInputElement | null
  if (searchMeetingsEl) searchMeetingsEl.value = recordingFilterState.search
  const sortbyMeetingsEl = document.getElementById("sortby-recordings") as HTMLSelectElement | null
  if (sortbyMeetingsEl) sortbyMeetingsEl.value = recordingFilterState.sortKey
  const markedOnlyMeetings = document.getElementById("markedonly-recordings")
  if (markedOnlyMeetings)
    markedOnlyMeetings.classList.toggle("active", recordingFilterState.markedOnly)
  document
    .querySelectorAll<HTMLButtonElement>("#kind-chips .chip[data-kind]")
    .forEach((chip) => {
      const k = chip.dataset.kind as ChatType | undefined
      const enabled = k ? recordingFilterState.enabledKinds.has(k) : true
      chip.classList.toggle("active", enabled)
    })
}

function applySourceVisibility(): void {
  const isChats = currentSource === "teams.chats"
  el<HTMLDivElement>("actions-chats").hidden = !isChats
  el<HTMLDivElement>("actions-transcripts").hidden = isChats
  el<HTMLUListElement>("chats").hidden = !isChats
  el<HTMLUListElement>("recordings").hidden = isChats
  updateFiltersVisibility()
  if (isChats) {
    updateMatchSummary(applyFiltersAndSort(chatsState.chats).length)
  } else {
    updateRecordingsSummary()
  }
}

// ----- Chat list rendering -----

function rerenderChatList(): void {
  const list = el<HTMLUListElement>("chats")
  const filtered = applyFiltersAndSort(chatsState.chats)
  if (chatsState.chats.length === 0) {
    list.innerHTML = ""
    return
  }
  if (filtered.length === 0) {
    list.innerHTML = `<li class="empty">No chats match these filters. ${
      filterState.markedOnly
        ? "Star some chats from the list to add them here."
        : "Loosen the filters."
    }</li>`
  } else {
    list.innerHTML = filtered.map(renderChatRow).join("")
    list.querySelectorAll<HTMLButtonElement>(".chat-action").forEach((btn) => {
      btn.addEventListener("click", () => {
        const id = btn.dataset.chatId!
        const name = btn.dataset.chatName!
        void downloadChat(id, name, btn)
      })
    })
    list.querySelectorAll<HTMLButtonElement>(".mark-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        toggleMark(btn.dataset.chatId!)
      })
    })
  }
  updateTypeCountChips()
  updateFiltersVisibility()
  if (currentSource === "teams.chats") updateMatchSummary(filtered.length)
  updateBulkButtons()
}

function updateMatchSummary(visible: number): void {
  const total = chatsState.chats.length
  const marked = markedIds.size
  const filtered = visible < total
  if (filtered) {
    setStatus(`Showing ${visible} of ${total} loaded · ${marked} marked`)
  } else {
    setStatus(`${total} chats loaded · ${marked} marked`)
  }
}

function renderChatRow(chat: TeamsChatItem): string {
  const name = chatDisplayName(chat)
  const isMarked = markedIds.has(chat.id)
  const lastSync = chatPrefs[chat.id]?.lastSync
  const downloadedTag = lastSync
    ? ` · downloaded ${formatDateShort(new Date(lastSync))}`
    : ""
  const sub = `${typeLabel(chat.chatType)} · last activity ${formatDate(chat.lastUpdatedDateTime)}${downloadedTag}`
  return `
    <li class="chat-row${isMarked ? " marked" : ""}">
      <button class="mark-toggle${isMarked ? " marked" : ""}" data-chat-id="${escapeHtml(chat.id)}" title="${isMarked ? "Unmark" : "Mark"} this chat" aria-label="${isMarked ? "Unmark" : "Mark"}">${isMarked ? "★" : "☆"}</button>
      <div class="chat-info">
        <div class="chat-name">${escapeHtml(name)}</div>
        <div class="chat-sub">${escapeHtml(sub)}</div>
      </div>
      <button class="chat-action" data-chat-id="${escapeHtml(chat.id)}" data-chat-name="${escapeHtml(name)}">Download</button>
    </li>
  `
}

/** Toggle a mark on any item (chat or meeting). The id format is distinct
 * between sources, so a single Set works. We rerender both lists because
 * the cost is negligible at typical loaded sizes and we don't need to know
 * which kind of id this is. */
function toggleMark(id: string): void {
  if (markedIds.has(id)) markedIds.delete(id)
  else markedIds.add(id)
  saveMarks(userCacheKey(), markedIds)
  rerenderChatList()
  rerenderRecordingsList()
  updateBulkButtons()
  scheduleOneDriveSave()
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

function updateFiltersVisibility(): void {
  el<HTMLDivElement>("filters").hidden =
    currentSource !== "teams.chats" || chatsState.chats.length === 0
  el<HTMLDivElement>("filters-recordings").hidden =
    currentSource !== "teams.recordings" || recordingsState.containers.length === 0
}

function showChatsRefreshButton(): void {
  el<HTMLButtonElement>("loadchats").hidden = true
  el<HTMLButtonElement>("refreshchats").hidden = false
}

// ----- Chats loading -----

async function initialLoadChats(): Promise<void> {
  const userKey = userCacheKey()
  const cached = loadCachedChats(userKey)
  if (cached && cached.chats.length > 0) {
    chatsState = { chats: cached.chats }
    rerenderChatList()
    showChatsRefreshButton()
    setStatus(
      `Showing ${cached.chats.length} cached chats from ${formatAge(ageMs(cached))}. Refreshing…`,
    )
  } else {
    setStatus("Loading chats…")
    el<HTMLUListElement>("chats").innerHTML = ""
  }
  const loadBtn = el<HTMLButtonElement>("loadchats")
  const refreshBtn = el<HTMLButtonElement>("refreshchats")
  loadBtn.disabled = true
  refreshBtn.disabled = true

  // Determine recency window from the dedicated chat-range selector (stored in
  // userPrefs.chatRange, synced cross-device via OneDrive). The old
  // lookback-as-window stopgap is gone; lookback now owns only download depth.
  const range: ChatRange = userPrefs.chatRange ?? { kind: "last-7d" }
  const { cutoffMs, untilMs } = computeChatWindow(range, chatPrefs)
  const rangeStr = chatRangeLabel(range, chatPrefs)

  // PROVEN METHOD — validated against a working reference implementation and a
  // live probe (80 chats for a 7-day window in ~5 pages, ZERO per-chat calls):
  //
  // (a) No per-chat enrichment: lastUpdatedDateTime from /me/chats is trusted
  //     directly. The probe confirmed it returns the correct recent set.
  //     Accepted gap: rare deeply-stale 1:1s (months-silent, one recent msg)
  //     may sit below the stop point and won't appear — explicit, accepted cost.
  //
  // (b) Jitter-tolerant stop: Graph's chat ordering is non-monotonic and
  //     unstable between requests (87 ordering violations measured in the probe).
  //     Stopping on a SINGLE out-of-window page is fragile — one stale chat at
  //     a page boundary caused a prior run to stop at page 1 with only 13
  //     results. We stop after 2 CONSECUTIVE pages where ALL items are BELOW
  //     cutoffMs. Items above untilMs (custom past-range upper bound) don't
  //     trigger the stop — we may not have reached the window yet.
  //     30-page hard cap as backstop.
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
        const t = new Date(c.lastUpdatedDateTime).getTime()
        return t >= cutoffMs && t <= untilMs
      })

      // Only count a page as "out of window" if ALL items are BELOW the lower
      // bound (cutoffMs). Items that are too recent (> untilMs, custom range)
      // don't mean we've passed the window — we just haven't gotten there yet.
      const allBelowCutoff = page.chats.every(
        (c) => new Date(c.lastUpdatedDateTime).getTime() < cutoffMs,
      )

      if (inWindow.length > 0) {
        kept.push(...inWindow)
        consecutiveOutOfWindowPages = 0
      } else if (allBelowCutoff) {
        consecutiveOutOfWindowPages++
      } else {
        // Items at/above cutoffMs exist but none in [cutoffMs, untilMs] —
        // keep paging (we haven't reached the window's lower edge yet).
        consecutiveOutOfWindowPages = 0
      }

      setStatus(`Loading recent chats… (${kept.length} in window so far)`)

      // Jitter-tolerant stop: 2 consecutive fully-below-cutoff pages.
      if (consecutiveOutOfWindowPages >= 2) {
        cursor = null
      }
    } while (cursor !== null && pageCount < PAGE_HARD_CAP)

    // Marked-include enrichment: fetch any marked chats that fell outside the
    // window so they always appear, regardless of timestamp staleness.
    const markedIncludeOn = userPrefs.markedInclude !== false // default ON
    if (markedIncludeOn && markedIds.size > 0) {
      const keptIds = new Set(kept.map((c) => c.id))
      const missingMarked = [...markedIds].filter((id) => !keptIds.has(id))
      if (missingMarked.length > 0) {
        setStatus(`Fetching ${missingMarked.length} marked chat(s) outside window…`)
        for (const id of missingMarked) {
          try {
            const chat = await fetchChatById(msal, id)
            if (chat) kept.push(chat)
          } catch (err) {
            // Fail soft — a single marked chat failing should not abort the load.
            console.warn(
              "[m365-pull] Skipping marked chat (fetch failed):",
              id,
              (err as Error).message,
            )
          }
        }
      }
    }

    kept.sort(
      (a, b) =>
        new Date(b.lastUpdatedDateTime).getTime() -
        new Date(a.lastUpdatedDateTime).getTime(),
    )

    chatsState = { chats: kept }
    rerenderChatList()
    saveCachedChats(userKey, kept)
    showChatsRefreshButton()
    setStatus(`${kept.length} chats (${rangeStr}).`)
  } catch (err) {
    // Partial recovery: save and show however many chats were kept before
    // the failure so progress isn't lost.
    if (kept.length > 0) {
      chatsState = { chats: kept }
      rerenderChatList()
      saveCachedChats(userKey, kept)
    }
    setStatus(
      `Load failed: ${(err as Error).message}${
        cached
          ? " — showing cached."
          : kept.length > 0
            ? ` — showing ${kept.length} partially loaded chats.`
            : ""
      }`,
      "error",
    )
  } finally {
    loadBtn.disabled = false
    refreshBtn.disabled = false
  }
}

async function refreshChats(): Promise<void> {
  clearCachedChats(userCacheKey())
  chatsState = { chats: [] }
  el<HTMLUListElement>("chats").innerHTML = ""
  await initialLoadChats()
}

// ----- Meetings rendering -----

function rerenderRecordingsList(): void {
  const list = el<HTMLUListElement>("recordings")
  const filtered = applyRecordingFiltersAndSort(recordingsState.containers)
  if (recordingsState.containers.length === 0) {
    list.innerHTML = `<li class="empty">No recordings found in this window. Try widening the window.</li>`
  } else if (filtered.length === 0) {
    list.innerHTML = `<li class="empty">No recording containers match these filters. ${
      recordingFilterState.markedOnly
        ? "Star some containers to add them here."
        : "Loosen the filters."
    }</li>`
  } else {
    list.innerHTML = filtered.map(renderRecordingContainerRow).join("")
    list.querySelectorAll<HTMLButtonElement>(".chat-action").forEach((btn) => {
      btn.addEventListener("click", () => {
        const chatId = btn.dataset.recContainerId!
        void downloadContainerTranscripts(chatId, btn)
      })
    })
    list.querySelectorAll<HTMLButtonElement>(".mark-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const chatId = btn.dataset.recContainerId
        if (chatId) toggleMark(`rec:${chatId}`)
      })
    })
  }
  updateFiltersVisibility()
  updateChatTypeChips()
  if (currentSource === "teams.recordings") {
    updateRecordingsSummary()
  }
  updateBulkButtons()
}

/** Mirror of updateTypeCountChips for the recording kind chips (1:1 / Group / Meeting). */
function updateChatTypeChips(): void {
  const counts = new Map<string, number>()
  for (const c of recordingsState.containers) {
    counts.set(c.chatType, (counts.get(c.chatType) || 0) + 1)
  }
  const chips = document.getElementById("kind-chips")
  if (!chips) return
  chips
    .querySelectorAll<HTMLButtonElement>(".chip[data-kind]")
    .forEach((chip) => {
      const k = chip.dataset.kind!
      const n = counts.get(k) ?? 0
      const span = chip.querySelector<HTMLSpanElement>(".chip-count")
      if (span) span.textContent = String(n)
    })
}

/** Compute marked subsets for each loaded source and update bulk-action buttons. */
function updateBulkButtons(): void {
  const markedChats = chatsState.chats.filter((c) => markedIds.has(c.id))
  const markedContainers = recordingsState.containers.filter((c) =>
    markedIds.has(`rec:${c.chatId}`),
  )
  const chatBtn = document.getElementById("bulk-chats") as HTMLButtonElement | null
  const meetingBtn = document.getElementById("bulk-recordings") as HTMLButtonElement | null
  if (chatBtn) {
    if (markedChats.length > 0) {
      chatBtn.hidden = false
      chatBtn.textContent = `Download ${markedChats.length} marked`
    } else {
      chatBtn.hidden = true
    }
  }
  if (meetingBtn) {
    if (markedContainers.length > 0) {
      meetingBtn.hidden = false
      meetingBtn.textContent = `Sync ${markedContainers.length} marked`
    } else {
      meetingBtn.hidden = true
    }
  }
}

/** Render a container row. Mark key is "rec:{chatId}" -- distinct from chat marks
 * (bare chatId) so the two keyspaces never collide in the shared markedIds set. */
function renderRecordingContainerRow(c: RecordingContainer): string {
  const markId = `rec:${c.chatId}`
  const isMarked = markedIds.has(markId)
  const title = c.chatTopic?.trim() || "(unnamed chat)"
  const kindLabel =
    c.chatType === "oneOnOne" ? "1:1" : c.chatType === "group" ? "Group" : "Meeting"

  // Last downloaded: most recent recording sync across the container
  let lastSyncMs: number | null = null
  for (const r of c.recordings) {
    const t = recordingPrefs[r.id]?.lastSync
      ? Date.parse(recordingPrefs[r.id].lastSync!)
      : null
    if (t !== null && (lastSyncMs === null || t > lastSyncMs)) lastSyncMs = t
  }
  const downloadedTag =
    lastSyncMs !== null
      ? ` · downloaded ${formatDateShort(new Date(lastSyncMs))}`
      : ""

  const n = c.recordings.length
  let dateTag = ""
  if (n === 1) {
    dateTag = ` · ${formatDateShort(new Date(c.recordings[0].eventCreatedDateTime))}`
  } else if (n > 1) {
    const latestStr = formatDateShort(new Date(c.recordings[0].eventCreatedDateTime))
    const earliestStr = formatDateShort(new Date(c.recordings[n - 1].eventCreatedDateTime))
    dateTag = earliestStr === latestStr ? ` · ${latestStr}` : ` · ${earliestStr} – ${latestStr}`
  }
  const sub = `${kindLabel} · ${n} recording${n !== 1 ? "s" : ""} in range${dateTag}${downloadedTag}`

  // Truncation badge: visible warning when the message scan hit the per-chat cap
  const truncBadge = c.truncated
    ? ` <span class="truncated-badge" title="Message scan hit the per-chat cap — narrow the window for complete results">⚠ scan limit</span>`
    : ""

  return `
    <li class="chat-row${isMarked ? " marked" : ""}">
      <button class="mark-toggle${isMarked ? " marked" : ""}" data-rec-container-id="${escapeHtml(c.chatId)}" title="${isMarked ? "Unmark" : "Mark"} this container" aria-label="${isMarked ? "Unmark" : "Mark"}">${isMarked ? "★" : "☆"}</button>
      <div class="chat-info">
        <div class="chat-name">${escapeHtml(title)}${truncBadge}</div>
        <div class="chat-sub">${escapeHtml(sub)}</div>
      </div>
      <button class="chat-action" data-rec-container-id="${escapeHtml(c.chatId)}">Sync transcripts</button>
    </li>
  `
}
function updateRecordingsSummary(): void {
  const n = recordingsState.containers.length
  const filtered = applyRecordingFiltersAndSort(recordingsState.containers).length
  const marked = recordingsState.containers.filter((c) =>
    markedIds.has(`rec:${c.chatId}`),
  ).length
  const totalRecs = recordingsState.containers.reduce(
    (sum, c) => sum + c.recordings.length,
    0,
  )
  const { chatsScanned, truncated } = recordingsState
  const label = rangeLabel(userPrefs.recordingRange ?? { kind: "last-7d" })
  if (n === 0) {
    setStatus(`No recordings for ${label}. Widen the window.`)
  } else if (filtered < n) {
    setStatus(
      `Showing ${filtered} of ${n} container(s) · ${totalRecs} recording(s) · ${marked} marked · ${chatsScanned} chat(s) scanned${truncated ? " · (chat list truncated — narrow window)" : ""}`,
    )
  } else {
    setStatus(
      `${n} container(s) · ${totalRecs} recording(s) (${label}) · ${marked} marked · ${chatsScanned} chat(s) scanned${truncated ? " · (chat list truncated — narrow window)" : ""}`,
    )
  }
}
function showRecordingsRefreshButton(): void {
  el<HTMLButtonElement>("load-recordings").hidden = true
  el<HTMLButtonElement>("refresh-recordings").hidden = false
}

// ----- Recordings loading -----

async function initialLoadRecordings(): Promise<void> {
  const loadBtn = el<HTMLButtonElement>("load-recordings")
  loadBtn.disabled = true
  loadBtn.textContent = "Loading\u2026"
  try {
    const range: RecordingRange = userPrefs.recordingRange ?? { kind: "last-7d" }
    const { fromMs, toMs } = computeRecordingWindow(range)
    const result = await listRecordings(msal, {
      fromMs,
      toMs,
      onProgress: (note) => setStatus(note),
    })

    let containers = result.containers

    // Marked-include enrichment: always show rec:-marked containers even when
    // their chat fell outside the scan window (0 in-window recordings is fine).
    const recMarkedIncludeOn = userPrefs.recordingMarkedInclude !== false // default ON
    if (recMarkedIncludeOn) {
      const scannedIds = new Set(containers.map((c) => c.chatId))
      const missingMarked = [...markedIds]
        .filter((id) => id.startsWith("rec:"))
        .map((id) => id.slice(4)) // strip "rec:" prefix
        .filter((chatId) => !scannedIds.has(chatId))

      if (missingMarked.length > 0) {
        setStatus(`Fetching ${missingMarked.length} marked container(s) outside window\u2026`)
        for (const chatId of missingMarked) {
          try {
            const chat = await fetchChatById(msal, chatId)
            if (chat) {
              containers = [
                ...containers,
                {
                  chatId: chat.id,
                  chatTopic: chat.topic ?? null,
                  chatType: chat.chatType as ChatType,
                  recordings: [],
                  truncated: false,
                },
              ]
            }
          } catch (err) {
            console.warn(
              "[m365-pull] Skipping marked container (fetch failed):",
              chatId,
              (err as Error).message,
            )
          }
        }
      }
    }

    recordingsState.containers = containers
    recordingsState.chatsScanned = result.chatsScanned
    recordingsState.truncated = result.truncated
    rerenderRecordingsList()
    showRecordingsRefreshButton()
  } catch (err) {
    setStatus(`Failed to load recordings: ${(err as Error).message}`, "error")
  } finally {
    loadBtn.disabled = false
    loadBtn.textContent = "Load my recordings"
  }
}

async function refreshRecordings(): Promise<void> {
  recordingsState = { containers: [], chatsScanned: 0, truncated: false }
  el<HTMLUListElement>("recordings").innerHTML = ""
  await initialLoadRecordings()
}
// ----- Downloading a chat -----

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

  setStatus(`Fetching "${chatName}" ${sinceLabel}…`)
  button.disabled = true
  const originalLabel = button.textContent
  button.textContent = "Fetching…"
  try {
    const opts: Parameters<typeof fetchChatMessages>[2] = {
      onProgress: ({ count, oldestSeen }) => {
        const back = oldestSeen ? `back to ${formatDateShort(oldestSeen)}` : "scanning…"
        setStatus(`Fetching "${chatName}" ${sinceLabel} · ${count} messages, ${back}`)
        button.textContent = `Fetching… (${count})`
      },
    }
    if (since) opts.since = since
    const messages = await fetchChatMessages(msal, chatId, opts)
    const destination = userPrefs.destination
    const browserName = buildChatArchiveFilename(chatId, chatName, {
      withTimestamp: true,
      extension: ".md",
    })
    const onedriveName = buildChatArchiveFilename(chatId, chatName, {
      withTimestamp: false,
      extension: ".md",
    })
    const nowIso = new Date().toISOString()
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
      const fullPath = `${userPrefs.oneDriveFolder.replace(/\/$/, "")}/${onedriveName}`
      setStatus(`${messages.length} messages fetched. Saving to OneDrive (${userPrefs.oneDriveFolder})…`)
      result = await saveTextToOneDrive(msal, fullPath, markdownBody, "text/markdown")
    } else {
      setStatus(`${messages.length} messages fetched. Saving…`)
      result = await saveAsText(browserName, markdownBody, {
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
      rerenderChatList()
      const incrementalNote = usingIncremental ? " (incremental)" : ""
      if (destination === "onedrive") {
        const where = result.path ? `OneDrive (${result.path})` : "OneDrive"
        setStatus(
          `✓ Saved ${messages.length} messages from "${chatName}" to ${where}${incrementalNote}.`,
        )
      } else {
        setStatus(
          `✓ Saved ${messages.length} messages from "${chatName}"${incrementalNote}.`,
        )
      }
      return true
    } else if (result.reason === "cancelled") {
      setStatus(`Save cancelled. (${messages.length} messages fetched but not written.)`)
      return false
    } else if (result.reason === "unsupported") {
      setStatus(
        "Browser save not supported here — use Microsoft Edge or another Chromium-based browser.",
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

async function downloadRecordingTranscript(
  recordingId: string,
  button: HTMLButtonElement,
): Promise<boolean> {
  // Search all containers for the recording
  let recording: RecordingItem | undefined
  for (const container of recordingsState.containers) {
    recording = container.recordings.find((r) => r.id === recordingId)
    if (recording) break
  }
  if (!recording) {
    setStatus("Recording not found in current list. Refresh and try again.", "error")
    return false
  }
  const originalLabel = button.textContent
  button.disabled = true
  button.textContent = "Resolving…"
  const subject = recording.chatTopic?.trim() || recording.filename
  try {
    const resolved = await resolveRecordingFromUrl(msal, recording.url)
    button.textContent = "Fetching transcripts…"
    setStatus(`Fetching transcripts for "${subject}"…`)
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
      return false
    }
    button.textContent = "Saving…"
    const account = msal.getActiveAccount()
    const userOid = account?.localAccountId ?? null
    const filename = buildTranscriptFilename(recording, userOid, ".md")
    const destination = userPrefs.destination
    let result: { saved: boolean; reason?: string; path?: string; webUrl?: string }

    // Concat all VTT bodies (some recordings carry multiple transcripts) into
    // one combined markdown document. Single-transcript case is the common
    // path; the concat produces the same output as a direct vttToMarkdown call.
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
      rerenderRecordingsList()
      if (destination === "onedrive") {
        const where = result.path ? `OneDrive (${result.path})` : "OneDrive"
        setStatus(`✓ Saved ${payload.transcriptCount} transcript(s) for "${subject}" to ${where}.`)
      } else {
        setStatus(`✓ Saved ${payload.transcriptCount} transcript(s) for "${subject}".`)
      }
      return true
    } else if (result.reason === "cancelled") {
      setStatus(`Save cancelled. (${payload.transcriptCount} transcripts fetched but not written.)`)
      return false
    } else {
      setStatus(`Save failed: ${result.reason}`, "error")
      return false
    }
  } catch (err) {
    setStatus(`Error: ${(err as Error).message}`, "error")
    console.error("[m365-pull] downloadRecordingTranscript failed:", err)
    return false
  } finally {
    button.disabled = false
    button.textContent = originalLabel || "Download transcript"
  }
}

// ----- Bulk download (all marked items in current source) -----

async function bulkDownloadChats(): Promise<void> {
  const marked = chatsState.chats.filter((c) => markedIds.has(c.id))
  if (marked.length === 0) return
  const bulkBtn = el<HTMLButtonElement>("bulk-chats")
  const originalLabel = bulkBtn.textContent
  bulkBtn.disabled = true
  let ok = 0
  let fail = 0
  for (let i = 0; i < marked.length; i++) {
    const chat = marked[i]
    bulkBtn.textContent = `Downloading ${i + 1}/${marked.length}\u2026`
    const rowBtn =
      (document.querySelector(
        `.chat-action[data-chat-id="${CSS.escape(chat.id)}"]`,
      ) as HTMLButtonElement | null) ?? document.createElement("button")
    const success = await downloadChat(chat.id, chatDisplayName(chat), rowBtn)
    if (success) ok++
    else fail++
  }
  bulkBtn.disabled = false
  bulkBtn.textContent = originalLabel || ""
  setStatus(
    `Bulk chats complete: ${ok} succeeded${fail > 0 ? `, ${fail} failed` : ""}.`,
  )
  updateBulkButtons()
}

/** Sync all recordings in a container row. Loops downloadRecordingTranscript
 * over each recording in the container; reuses the existing pipeline untouched. */
async function downloadContainerTranscripts(
  chatId: string,
  button: HTMLButtonElement,
): Promise<void> {
  const container = recordingsState.containers.find((c) => c.chatId === chatId)
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
  for (let i = 0; i < container.recordings.length; i++) {
    const rec = container.recordings[i]
    button.textContent = `Syncing ${i + 1}/${container.recordings.length}\u2026`
    const tempBtn = document.createElement("button")
    const success = await downloadRecordingTranscript(rec.id, tempBtn)
    if (success) ok++
    else fail++
  }
  button.disabled = false
  button.textContent = originalLabel || "Sync transcripts"
  setStatus(
    `Container sync complete: ${ok} transcript(s) saved${fail > 0 ? `, ${fail} failed` : ""}.`,
  )
}

async function bulkDownloadRecordings(): Promise<void> {
  const markedContainers = recordingsState.containers.filter((c) =>
    markedIds.has(`rec:${c.chatId}`),
  )
  if (markedContainers.length === 0) return
  const bulkBtn = el<HTMLButtonElement>("bulk-recordings")
  const originalLabel = bulkBtn.textContent
  bulkBtn.disabled = true
  let ok = 0
  let fail = 0
  for (let i = 0; i < markedContainers.length; i++) {
    const container = markedContainers[i]
    bulkBtn.textContent = `Syncing container ${i + 1}/${markedContainers.length}\u2026`
    for (const rec of container.recordings) {
      const tempBtn = document.createElement("button")
      const success = await downloadRecordingTranscript(rec.id, tempBtn)
      if (success) ok++
      else fail++
    }
  }
  bulkBtn.disabled = false
  bulkBtn.textContent = originalLabel || ""
  setStatus(
    `Bulk sync complete: ${ok} recording(s) succeeded${fail > 0 ? `, ${fail} failed` : ""}.`,
  )
  updateBulkButtons()
}
// ----- Settings modal -----

function openSettingsModal(): void {
  const modal = el<HTMLDivElement>("settings-modal")
  el<HTMLInputElement>("settings-folder").value = userPrefs.oneDriveFolder
  el<HTMLSelectElement>("settings-destination").value = userPrefs.destination
  modal.hidden = false
  // Focus the folder field for instant typing
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
  // Normalize: convert backslashes to forward slashes, ensure leading slash,
  // strip trailing slash
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
