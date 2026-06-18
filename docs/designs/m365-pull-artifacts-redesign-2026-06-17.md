# m365-pull — Redesign: Artifacts, not Containers (2026-06-17)

**Status: LOCKED — building.** Supersedes the unified-containers direction (2026-06-12).
Grounded against the code by zen-architect: the artifact granularity already exists in the
data and the synced ledger — this is a presentation + selection change, not a schema rewrite.

## The reframe
The chat is a **folder you open**; the artifacts inside are the atoms.
- **Chat** = an expandable group ("folder"). Never downloaded itself.
- **Messages** = exactly **1** artifact per chat, always present (a chat *is* its messages).
- **Recordings** = **0..N** artifacts per chat — each separate and immutable. A chat may have
  **many** recordings, or **none**.

## Two orthogonal verbs (this kills the "Download kept" confusion)
- **Select** (`☐` checkbox) — *ephemeral*, per artifact (or whole group via select-all).
  Drives **"Download selected (N)"**. One-off; ignores favorite state. **NOT synced.**
- **Favorite** (`★`) — *persistent + synced*, per **STREAM**: a chat's **Messages** stream
  and/or its **Recordings** stream. Drives **"Sync favorites (N)"**. (Renames the old
  "Keep/marked".) Individual recordings are **not** individually favoritable (they're
  immutable) — favoriting the *recordings stream* means "grab this chat's recordings on
  every sync."

**Naming:** "Favorite", not "Keep" — *Keep* implies the un-kept get discarded. Use the `★` glyph.

## Honest "sync" — no backend, no background
There is no server process. **Favorite = membership in the sync set.** **Sync pulls every
favorited stream WHEN THE USER CLICKS IT.** Never label it "auto/always" — it's a manual
Sync over a remembered set. The empty/first-run copy must say so plainly.

## No-recording = silence
A chat with no recording shows **nothing** for recordings — no `?`, no empty row. Recording
info appears **only once it's truth** (after the recordings scan). The old `?` came from
showing recording state before scanning; it's gone.

## Phased build
- **Phase 1** (proves the model; no schema, no migration): expand chat → a **Messages** row
  + each **recording** as separate selectable rows → checkbox **Select** (ephemeral) →
  **"Download selected (N)"** over the existing per-artifact download primitives
  (`downloadChat(chatId)`, `downloadRecordingTranscript(rec.id)`). No-recording renders clean.
  Existing mark/bulk path left working & untouched.
- **Phase 2** (Favorite): rename mark→Favorite; move to **per-stream** (messages / recordings
  stream); migrate existing bare-chatId marks → favorite **both** streams (no loss);
  **"Sync favorites (N)"** replaces the bulk action; bump synced-state `v1→v2` **AND** add the
  new fields to `mergeStates` (else they're silently dropped on cross-device merge — fail-quiet
  data loss). Collapsed row shows favorite **state** (read-only `★`); toggle lives in expanded.
- **Phase 3** (versioned filenames): dated + ranged names so **sort-by-name reveals download
  history**.
  - chat: `<Name>__chat__pulled-<YYYY-MM-DD-HHMM>__<rangeStart>_to_<rangeEnd>.txt`
  - recording: `<Name>__rec-<callDate>__pulled-<YYYY-MM-DD-HHMM>.transcript.txt`
  - sanitize Windows-prohibited chars in `<Name>`; pulled-date is the primary version key.

## Parked (with reason)
- **Recurring-meeting SERIES auto-grab** (cross-occurrence): each weekly occurrence is a *new
  chat id*, so clustering occurrences is a heuristic that can **fail-quiet** (pull the wrong
  meeting's recording). Park until it can be made fail-*visible* (grouping that never changes
  what's pulled). You can still favorite any single chat's recordings stream today.
- Message-count pre-flight and OneDrive-vs-browser destination already exist.

## Taste decisions taken
- Favorite toggle in the **expanded** view; collapsed shows a read-only `★` state (promotable later).
- Filenames keep human-readable spaces in the chat name (sanitized); `pulled-<date-time>` for
  collision-safe versioning.
