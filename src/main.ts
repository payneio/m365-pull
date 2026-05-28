import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser"
import { config } from "./config"
import "./style.css"
import {
  listChatsPage,
  fetchChatMessages,
  chatDisplayName,
  type TeamsChatItem,
  type TeamsChatMessage,
} from "./sources/teams-chats"
import {
  listRecordings,
  formatDurationShort,
  buildTranscriptFilename,
  type ChatType,
  type RecordingItem,
} from "./sources/teams-call-recordings"
import {
  resolveRecordingFromUrl,
  fetchRecordingTranscripts,
} from "./sources/teams-recordings"
import { vttToMarkdown } from "./format/transcript-markdown"
import { saveAsJson, saveAsText } from "./destinations/browser"
import { saveToOneDrive, saveTextToOneDrive, loadFromOneDrive } from "./destinations/onedrive"
import {
  loadUserPrefs,
  saveUserPrefs,
  type UserPrefs,
  type Destination,
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
  nextCursor: string | null
}

interface RecordingsState {
  recordings: RecordingItem[]
  daysBack: number
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
let chatsState: ChatsState = { chats: [], nextCursor: null }
let recordingsState: RecordingsState = { recordings: [], daysBack: 90, chatsScanned: 0, truncated: false }
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

function applyRecordingFiltersAndSort(recordings: RecordingItem[]): RecordingItem[] {
  const q = recordingFilterState.search.trim().toLowerCase()
  let result = recordings.filter((r) => {
    if (recordingFilterState.markedOnly && !markedIds.has(r.id)) return false
    if (!recordingFilterState.enabledKinds.has(r.chatType)) return false
    if (q) {
      const haystack = `${r.filename} ${r.chatTopic ?? ""}`.toLowerCase()
      if (!haystack.includes(q)) return false
    }
    return true
  })
  const byRecent = (a: RecordingItem, b: RecordingItem) =>
    a.eventCreatedDateTime < b.eventCreatedDateTime ? 1 : -1
  const byName = (a: RecordingItem, b: RecordingItem) =>
    (a.chatTopic || a.filename).localeCompare(b.chatTopic || b.filename)
  if (recordingFilterState.sortKey === "name") {
    result = [...result].sort(byName)
  } else if (recordingFilterState.sortKey === "recent") {
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

function mergeMessagesById(
  oldMessages: TeamsChatMessage[],
  newMessages: TeamsChatMessage[],
): TeamsChatMessage[] {
  const byId = new Map<string, TeamsChatMessage>()
  for (const m of oldMessages) if (m.id) byId.set(m.id, m)
  for (const m of newMessages) if (m.id) byId.set(m.id, m)
  const merged = [...byId.values()]
  merged.sort((a, b) => {
    const at = a.createdDateTime ? Date.parse(a.createdDateTime) : 0
    const bt = b.createdDateTime ? Date.parse(b.createdDateTime) : 0
    return at - bt
  })
  return merged
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
        <span class="label">Lookback:</span>
        <select id="lookback" title="Lookback window for downloads">
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
        <span class="label">Window:</span>
        <select id="daysback" title="How far back to search the calendar">
          <option value="7">Last 7 days</option>
          <option value="30">Last 30 days</option>
          <option value="90" selected>Last 90 days</option>
          <option value="365">Last year</option>
        </select>
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
      </div>
      <div id="status"></div>
      <ul id="chats" class="chat-list"></ul>
      <ul id="recordings" class="chat-list" hidden></ul>
      <div id="loadmore-row" class="loadmore-row" hidden>
        <button id="loadmore">Load more</button>
      </div>
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
  el<HTMLSelectElement>("daysback").addEventListener("change", (e) => {
    recordingsState.daysBack = parseInt((e.target as HTMLSelectElement).value, 10) || 90
    saveUIPrefs()
  })

  // Chat lookback dropdown persists too
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

  // Shared load-more button (context-dependent)
  el<HTMLButtonElement>("loadmore").addEventListener("click", () => {
    if (currentSource === "teams.chats") void loadMoreChats()
    // recordings source has no pagination -- loadmore-row stays hidden
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
 * the user's dropdown selections, chip filters, and source choice. */
function saveUIPrefs(): void {
  const lookbackEl = document.getElementById("lookback") as
    | HTMLSelectElement
    | null
  const daysbackEl = document.getElementById("daysback") as
    | HTMLSelectElement
    | null
  saveUIState(userCacheKey(), {
    currentSource,
    lookback: lookbackEl?.value,
    daysBack: daysbackEl?.value,
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
  if (saved.daysBack) {
    const n = parseInt(saved.daysBack, 10)
    if (Number.isFinite(n) && n > 0) recordingsState.daysBack = n
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
  // Chat lookback
  const lookbackEl = document.getElementById("lookback") as HTMLSelectElement | null
  if (lookbackEl && saved.lookback) lookbackEl.value = saved.lookback
  // Meeting daysBack
  const daysbackEl = document.getElementById("daysback") as HTMLSelectElement | null
  if (daysbackEl) daysbackEl.value = String(recordingsState.daysBack)
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
    updateLoadMore()
    updateMatchSummary(applyFiltersAndSort(chatsState.chats).length)
  } else {
    el<HTMLDivElement>("loadmore-row").hidden = true
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
        : "Loosen the filters or load more chats."
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
  const more = chatsState.nextCursor ? " · more available" : ""
  if (filtered) {
    setStatus(`Showing ${visible} of ${total} loaded · ${marked} marked${more}`)
  } else {
    setStatus(`${total} chats loaded · ${marked} marked${more}`)
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
    currentSource !== "teams.recordings" || recordingsState.recordings.length === 0
}

function updateLoadMore(): void {
  const row = el<HTMLDivElement>("loadmore-row")
  const btn = el<HTMLButtonElement>("loadmore")
  if (currentSource !== "teams.chats") return
  if (chatsState.nextCursor) {
    row.hidden = false
    btn.textContent = `Load more chats (${chatsState.chats.length} loaded so far)`
    btn.disabled = false
  } else {
    row.hidden = true
  }
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
    chatsState = { chats: cached.chats, nextCursor: cached.nextCursor }
    rerenderChatList()
    updateLoadMore()
    showChatsRefreshButton()
    const moreNote = cached.nextCursor ? " (partial · Load more available)" : " (full list)"
    setStatus(
      `Showing ${cached.chats.length} cached chats from ${formatAge(ageMs(cached))}${moreNote}. Refreshing first page…`,
    )
  } else {
    setStatus("Loading chats…")
    el<HTMLUListElement>("chats").innerHTML = ""
  }
  const loadBtn = el<HTMLButtonElement>("loadchats")
  const refreshBtn = el<HTMLButtonElement>("refreshchats")
  loadBtn.disabled = true
  refreshBtn.disabled = true
  try {
    const page = await listChatsPage(msal, null)
    chatsState = { chats: page.chats, nextCursor: page.nextCursor }
    rerenderChatList()
    updateLoadMore()
    saveCachedChats(userKey, chatsState.chats, chatsState.nextCursor)
    showChatsRefreshButton()
  } catch (err) {
    setStatus(
      `Refresh failed: ${(err as Error).message}${cached ? " — showing cached." : ""}`,
      "error",
    )
  } finally {
    loadBtn.disabled = false
    refreshBtn.disabled = false
  }
}

async function loadMoreChats(): Promise<void> {
  if (!chatsState.nextCursor) return
  const btn = el<HTMLButtonElement>("loadmore")
  btn.disabled = true
  btn.textContent = "Loading more…"
  try {
    const page = await listChatsPage(msal, chatsState.nextCursor)
    chatsState.chats = [...chatsState.chats, ...page.chats]
    chatsState.nextCursor = page.nextCursor
    rerenderChatList()
    updateLoadMore()
    saveCachedChats(userCacheKey(), chatsState.chats, chatsState.nextCursor)
  } catch (err) {
    setStatus(`Load more failed: ${(err as Error).message}`, "error")
    btn.disabled = false
    btn.textContent = "Load more chats"
  }
}

async function refreshChats(): Promise<void> {
  clearCachedChats(userCacheKey())
  chatsState = { chats: [], nextCursor: null }
  el<HTMLUListElement>("chats").innerHTML = ""
  el<HTMLDivElement>("loadmore-row").hidden = true
  await initialLoadChats()
}

// ----- Meetings rendering -----

function rerenderRecordingsList(): void {
  const list = el<HTMLUListElement>("recordings")
  const filtered = applyRecordingFiltersAndSort(recordingsState.recordings)
  if (recordingsState.recordings.length === 0) {
    list.innerHTML = `<li class="empty">No recordings found in this window. Try widening the window.</li>`
  } else if (filtered.length === 0) {
    list.innerHTML = `<li class="empty">No recordings match these filters. ${
      recordingFilterState.markedOnly
        ? "Star some recordings to add them here."
        : "Loosen the filters."
    }</li>`
  } else {
    list.innerHTML = filtered.map(renderRecordingRow).join("")
    list.querySelectorAll<HTMLButtonElement>(".chat-action").forEach((btn) => {
      btn.addEventListener("click", () => {
        const eventId = btn.dataset.recordingId!
        void downloadRecordingTranscript(eventId, btn)
      })
    })
    list.querySelectorAll<HTMLButtonElement>(".mark-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const id = btn.dataset.recordingId
        if (id) toggleMark(id)
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
  for (const r of recordingsState.recordings) {
    counts.set(r.chatType, (counts.get(r.chatType) || 0) + 1)
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
  const markedMeetings = recordingsState.recordings.filter((m) => markedIds.has(m.id))
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
    if (markedMeetings.length > 0) {
      meetingBtn.hidden = false
      meetingBtn.textContent = `Download ${markedMeetings.length} marked`
    } else {
      meetingBtn.hidden = true
    }
  }
}

function renderRecordingRow(r: RecordingItem): string {
  const account = msal.getActiveAccount()
  const userOid = account?.localAccountId?.toLowerCase() ?? null
  const youInitiated =
    r.initiatorOid !== null && userOid !== null &&
    r.initiatorOid.toLowerCase() === userOid
  const title = r.chatTopic?.trim() || r.filename
  const start = formatDate(r.eventCreatedDateTime)
  const dur = formatDurationShort(r.durationIso)
  const isMarked = markedIds.has(r.id)
  const lastSync = recordingPrefs[r.id]?.lastSync
  const downloadedTag = lastSync
    ? ` · downloaded ${formatDateShort(new Date(lastSync))}`
    : ""
  const kindLabel =
    r.chatType === "oneOnOne" ? "1:1" : r.chatType === "group" ? "Group" : "Meeting"
  const initiatorTag = youInitiated ? " · you started" : ""
  const sub = `${start} · ${kindLabel}${dur ? ` · ${dur}` : ""}${initiatorTag}${downloadedTag}`

  // Participants line. Show up to 4 names; "+N more" beyond that. Skip self by
  // oid match. Bots labeled with parens to disambiguate from human names.
  let participantsLine = ""
  if (r.participants.length > 0) {
    const others = r.participants.filter(
      (p) => !(userOid && p.id && p.id.toLowerCase() === userOid),
    )
    const labels = others.map((p) =>
      p.kind === "bot" ? `(${p.displayName})` : p.displayName,
    )
    const visible = labels.slice(0, 4)
    const extra = labels.length - visible.length
    const tail = extra > 0 ? `, +${extra} more` : ""
    participantsLine = `With ${visible.join(", ")}${tail}`
  }

  return `
    <li class="chat-row${isMarked ? " marked" : ""}">
      <button class="mark-toggle${isMarked ? " marked" : ""}" data-recording-id="${escapeHtml(r.id)}" title="${isMarked ? "Unmark" : "Mark"} this recording" aria-label="${isMarked ? "Unmark" : "Mark"}">${isMarked ? "★" : "☆"}</button>
      <div class="chat-info">
        <div class="chat-name">${escapeHtml(title)}</div>
        <div class="chat-sub">${escapeHtml(sub)}</div>
        ${participantsLine ? `<div class="chat-participants">${escapeHtml(participantsLine)}</div>` : ""}
      </div>
      <button class="chat-action" data-recording-id="${escapeHtml(r.id)}">Download transcript</button>
    </li>
  `
}

function updateRecordingsSummary(): void {
  const n = recordingsState.recordings.length
  const filtered = applyRecordingFiltersAndSort(recordingsState.recordings).length
  const marked = recordingsState.recordings.filter((r) => markedIds.has(r.id)).length
  const { chatsScanned, truncated } = recordingsState
  if (n === 0) {
    setStatus(`No recordings in last ${recordingsState.daysBack} days. Widen the window.`)
  } else if (filtered < n) {
    setStatus(`Showing ${filtered} of ${n} · ${marked} marked · from ${chatsScanned} chat(s)${truncated ? " · (chat list truncated — narrow window or add pagination)" : ""}`)
  } else {
    setStatus(`${n} recording(s) from ${chatsScanned} chat(s) (last ${recordingsState.daysBack} days) · ${marked} marked${truncated ? " · (chat list truncated — narrow window or add pagination)" : ""}`)
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
  loadBtn.textContent = "Loading…"
  try {
    const result = await listRecordings(msal, {
      daysBack: recordingsState.daysBack,
      onProgress: (note) => setStatus(note),
    })
    recordingsState.recordings = result.recordings
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
  recordingsState = { recordings: [], daysBack: recordingsState.daysBack, chatsScanned: 0, truncated: false }
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
    const stableId = chatId.replace(/[^a-zA-Z0-9]/g, "")
    const browserName = `teams-chat-${stableId}-${Date.now()}.json`
    const onedriveName = `teams-chat-${stableId}.json`
    const nowIso = new Date().toISOString()
    const baseSnapshot = {
      source: "teams.chats",
      chatId,
      chatName,
      lookback,
      since: since?.toISOString() ?? null,
      fetchedAt: nowIso,
      messageCount: messages.length,
      messages,
    }

    let result: { saved: boolean; reason?: string; path?: string; webUrl?: string }
    let mergedTotal = messages.length
    let newCount = messages.length
    if (destination === "onedrive") {
      const fullPath = `${userPrefs.oneDriveFolder.replace(/\/$/, "")}/${onedriveName}`
      setStatus(
        `${messages.length} messages fetched. Merging with existing archive at OneDrive (${userPrefs.oneDriveFolder})…`,
      )
      const existing = await loadFromOneDrive<{
        messages?: TeamsChatMessage[]
        firstDownloadAt?: string
      }>(msal, fullPath)
      const oldMessages = existing.found ? (existing.data?.messages ?? []) : []
      const firstDownloadAt = existing.data?.firstDownloadAt ?? nowIso
      const merged = mergeMessagesById(oldMessages, messages)
      mergedTotal = merged.length
      newCount = mergedTotal - oldMessages.length
      const archive = {
        source: "teams.chats",
        chatId,
        chatName,
        schemaVersion: 2,
        firstDownloadAt,
        lastDownloadAt: nowIso,
        lastLookback: lookback,
        lastSince: since?.toISOString() ?? null,
        messageCount: merged.length,
        messages: merged,
      }
      result = await saveToOneDrive(msal, fullPath, archive)
    } else {
      setStatus(`${messages.length} messages fetched. Saving…`)
      result = await saveAsJson(browserName, baseSnapshot)
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
          `✓ "${chatName}" archive updated in ${where}: +${newCount} new · ${mergedTotal} total${incrementalNote}.`,
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
  const recording = recordingsState.recordings.find((r) => r.id === recordingId)
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
    bulkBtn.textContent = `Downloading ${i + 1}/${marked.length}…`
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

async function bulkDownloadRecordings(): Promise<void> {
  const marked = recordingsState.recordings.filter((r) => markedIds.has(r.id))
  if (marked.length === 0) return
  const bulkBtn = el<HTMLButtonElement>("bulk-recordings")
  const originalLabel = bulkBtn.textContent
  bulkBtn.disabled = true
  let ok = 0
  let fail = 0
  for (let i = 0; i < marked.length; i++) {
    const recording = marked[i]
    bulkBtn.textContent = `Downloading ${i + 1}/${marked.length}…`
    const rowBtn =
      (document.querySelector(
        `.chat-action[data-recording-id="${CSS.escape(recording.id)}"]`,
      ) as HTMLButtonElement | null) ?? document.createElement("button")
    const success = await downloadRecordingTranscript(recording.id, rowBtn)
    if (success) ok++
    else fail++
  }
  bulkBtn.disabled = false
  bulkBtn.textContent = originalLabel || ""
  setStatus(
    `Bulk recordings complete: ${ok} succeeded${fail > 0 ? `, ${fail} failed` : ""}.`,
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
