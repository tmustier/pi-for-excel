# Persistence and restore

## Summary

Persistence is what Pi for Excel remembers when the taskpane goes away, and restore is what it brings back when the taskpane opens again. Everything Pi remembers lives in the add-in's own browser storage on this computer (IndexedDB, plus a little `localStorage` and a private file area), never in the workbook file and never on a server. The one exception is a small token Pi writes into an unsaved workbook so that its [automatic backups](../workbook/backups-and-revert.md) can find it again. There is no save button and no "Saving…" indicator: sessions, tab layouts and settings are saved quietly as things happen, and a save that fails says nothing. What comes back is the set of tabs that were open in this workbook, in order, with the same active tab and each tab's conversation, model and thinking level. What does not come back: anything that was only in memory, such as composer drafts, queued messages, the response that was streaming, chosen tab titles, the recently closed list and the error banner. This document owns *what is saved when, and where*; every other document's **Persistence.** paragraph links here.

## The simple case

The user opens Pi in a saved workbook, asks a question in "Chat 1", and Pi answers. The moment Pi's answer completes, the session is saved, linked to this workbook, and recorded in the workbook's tab layout. The user opens a second tab, asks something else, gets an answer, and leaves that tab active.

The user closes the taskpane (or Excel, or reloads the add-in). Nothing happens at that moment: everything that matters was already saved.

The next time the user opens Pi in the same workbook, the taskpane shows "Chat 1" and "Chat 2" in the same order, with the second tab active and both conversations exactly as they were after their last completed message. The composer is empty. The tab labels are "Chat 1" and "Chat 2" even if the user had renamed them.

## The interaction, event by event

Persistence has two moments the user can notice: a **save**, which happens without any visible sign, and a **restore**, which happens while the taskpane is starting.

```mermaid
stateDiagram-v2
    [*] --> starting : taskpane opens
    starting --> restoring : workbook identity resolved (layout key chosen)
    restoring --> ready : layout found (tabs rebuilt, active tab selected)
    restoring --> ready : no layout (latest linked session, or a blank tab)
    ready --> ready : message completes (session saved, after the first answer)
    ready --> ready : tab created, closed, moved or switched (layout saved)
    ready --> ready : setting changed (setting saved at once)
    ready --> [*] : taskpane goes away (nothing saved; memory-only state lost)
```

### What is saved, and where

| What | Where | When it is written | Scope |
| --- | --- | --- | --- |
| Sessions: messages, model, thinking level, saved title, usage | IndexedDB database `pi-for-excel`, stores `sessions` and `sessions-metadata` | See [When a session is saved](#when-a-session-is-saved) | All workbooks |
| Workbook link and "latest session" pointer | Settings store | With every session save, when the workbook has an identity | Per [workbook identity](#workbook-identity) |
| Tab layout (open session ids, order, active tab) | Settings store, one entry per workbook identity | Every time a tab is created, closed, moved or switched, and when the taskpane regains focus | Per workbook identity; one shared entry for workbooks without one |
| Recently closed | Memory only (up to 10 items) | When a tab closes | This taskpane, until it goes away |
| Settings (language, proxy, execution mode, model-switch behavior, auto-compaction flag, rules, conventions, skills, extensions, connections, web search, MCP servers) | Settings store | The moment the user changes them | All workbooks, except workbook rules (per workbook identity) |
| API keys and OAuth access tokens | IndexedDB store `provider-keys` | On login, and when an expired OAuth token is refreshed at startup | All workbooks |
| OAuth grants (refresh tokens) | Settings store, one entry per provider | On login and refresh | All workbooks |
| Custom gateways and fetched model lists | IndexedDB stores `custom-providers` and `model-catalogs` | When added or refreshed | All workbooks |
| [Automatic backups](../workbook/backups-and-revert.md) | Independent compressed backups and indexed summaries; 512 MiB shared ceiling, no count/age limit | State captured before a supported mutation, saved after it succeeds | Tagged per workbook identity; observed identity changes move history |
| Workbook change audit log | One settings entry, at most 500 entries | After each mutation | All workbooks |
| [Files workspace](../files/the-files-workspace.md) | The browser's private file area ("Sandboxed workspace"), or a folder on disk the user connected ("Local folder"); file tags and the Files audit trail in the settings store | When files are imported, written or deleted | All workbooks; files are tagged with the workbook they came from |
| [Manual backups](../workbook/manual-backups.md) | Files workspace, under `manual-backups/full-workbook/v1/`, one folder per workbook identity | On `/backup` | Per workbook identity |
| Extension storage | One settings entry, at most 1 MB per extension | When an extension writes | All workbooks |
| Disclosure bar dismissal, experimental flags | `localStorage` | When dismissed or toggled | All workbooks |
| Document token (`pi.workbookInstanceId`) | Inside the workbook file, as an add-in setting | Before the first automatic backup in a workbook that has never been saved | That file |

> Technical note: "settings store" is a key-value store inside the same IndexedDB database. Most of the rows above are therefore one database, owned by the add-in's web origin (`https://pi-for-excel.vercel.app` for the hosted add-in). Anything that clears that origin's site data clears all of them at once.

### When a session is saved

- **The first save** happens when Pi's first model response in the session ends. That includes a model response that ended because the user pressed Stop or because the provider failed: the transcript then holds an assistant message (marked "Request aborted" or "Error:"), and that counts. Until then the session exists only in memory.
- **After that**, every completed message saves the whole session: the user's message as it joins the transcript, each model response when it finishes, each tool result, and each steer or follow-up when it is delivered. A model response that is still streaming is never saved part-way.
- **Forced saves** write the session even if Pi has never answered: closing a tab, duplicating a tab (the copy), changing a tab's model (in place, or the forked copy), and restoring a blank tab from the tab layout.
- **Not saved on their own:** renaming a tab before Pi has answered (the title waits for the first save), changing the thinking level (saved with the next completed message), and compaction (the compacted conversation is saved with the next completed message).
- **Nothing is saved when the taskpane goes away.** There is no save on unload.

Each save rewrites the whole session record. It also links the session to the current workbook if it has no link yet (an existing link is never changed), and records it as the workbook's latest session.

### Restoring when the taskpane opens

1. Pi works out the [workbook identity](#workbook-identity) and reads the tab layout stored for it.
2. **If a layout exists**, Pi opens one tab per listed session, in order. A session that is in storage comes back with its messages, model, thinking level and saved title. A session that is not in storage (it was never saved, or it was deleted) comes back as a blank tab, which is force-saved at once under a new id. The tab that was active becomes active again; if it cannot be found, the first tab does.
3. **If no layout exists** (the first time Pi opens in this workbook), Pi opens one tab. If the workbook has an identity, that tab is filled with the latest session linked to this workbook, if it has any messages; otherwise it is blank. If the workbook has no identity, the tab is filled with the most recently changed session in storage, from any workbook.
4. If the layout could not be read at all (a storage error), Pi starts with one blank tab and leaves the stored layout untouched, so that a later open can still use it. A layout that was read but is malformed is replaced.

Every restored tab gets a fresh default title: "Chat 1", "Chat 2", … in tab order, whatever it was called before. Confirmed in Excel on the web after a language-change reload: three tabs (one of them "Chat 4" before the reload) came back as tabs 1 to 3, with the same active tab and each tab's model and thinking level.

### What a reload brings back and what it loses

| Comes back | Lost |
| --- | --- |
| The tabs open in this workbook, their order, the active tab | Composer drafts in every tab (confirmed in Excel on the web) |
| Each tab's completed messages, tool cards and thinking blocks | The model response that was streaming, and any tool call that had not finished |
| Each tab's model and thinking level | Steers, follow-ups, queued prompts and queued commands |
| Saved titles (in the [resume picker](../sessions/resuming-a-session.md) only) | Chosen tab titles on the tabs themselves; tabs show "Chat N" again (confirmed in Excel on the web) |
| All settings, logins, rules and conventions | Recently closed: the undo toast, Ctrl/Cmd+Shift+T and `/reopen` find nothing ("No recently closed tab", confirmed in Excel on the web) |
| Automatic backups, including across saves and reloads (subject to deletion, eviction or storage loss) | The error banner, the working indicator, open overlays and menus, toasts |
| Files in the Files workspace | Permission to a connected local folder, if the browser no longer grants it (Pi falls back to the sandboxed workspace) |
| | A session that never received an answer from Pi: the tab comes back blank |

## Workbook identity

Pi recognises "the same workbook" by its file location. It takes the workbook's URL or path as Excel reports it, drops anything after `?` or `#`, and stores only a one-way hash of it (`url_sha256:…`); the raw path is never stored. The workbook name Pi shows (for example on backups) is the file-name part of that location.

- **A saved workbook on disk or in OneDrive/SharePoint** has an identity as soon as the taskpane opens. Pi asks Excel for the current location every time it needs it, so a **Save As** or a **rename** is picked up without a reload, the next time Pi looks (at the latest when the taskpane next regains focus).
- **A new workbook that has never been saved** ("Book1") has no identity. Its tab layout goes to one shared slot used by every workbook without an identity, its sessions are not linked to any workbook, workbook-level [rules](../guidance/rules.md) cannot be saved ("Can't identify this workbook right now — try saving the file first."), [manual backups](../workbook/manual-backups.md) are refused, and its automatic backups are kept under the document token instead.
- **Excel on the web** reported no location at all in the workbook used for verification, which was opened through an anonymous sharing link; every workbook there therefore behaves like an unsaved one. The only tab layout in storage was the shared slot, `workbook.tabLayout.v1.__global__` (confirmed). Confirmed in Excel on the web (guest access); signed-in workbooks have not been checked. See [host differences](host-differences.md#workbook-identity-and-saving).

What happens when the identity changes while Pi is open:

| Change | Sessions | Tab layout | Automatic backups | Workbook rules |
| --- | --- | --- | --- | --- |
| First save of a new workbook | The next save of each open session links it to the new identity; sessions not saved again stay unlinked | Written under the new identity from the next layout save; the shared slot keeps its old copy | Token-scoped history moves to the saved identity | Become available, starting empty |
| Save As / rename of a saved workbook | Existing sessions stay linked to the old identity; the resume picker in the new file lists them under "Show sessions from all workbooks" | Written under the new identity; the old file keeps its last layout, so both files reopen the same sessions | When observed on the same open document, history moves without copying, retaining any destination history | Stay with the old identity; the renamed file starts with none |
| Save As a copy of a never-saved workbook | As for a first save | As for a first save | The open document's history follows its new identity; no duplicate history is kept for the old identity | As for a first save |

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | Saved with the session; on restore the model is looked up again so its limits are current, and kept as saved if it no longer exists. | Changing the model is refused while busy; when idle it forces a save. |
| Thinking level | Saved with the session and restored. | Saved with the next completed message. |
| Execution mode | A global setting, saved the moment it changes. Not part of a session. | Saved at once. |
| Files in the workspace, or an overlay open | Files persist independently of sessions. Overlays are never restored. | Importing during a turn persists the files at once. |
| Tab state (new, restored, another tab busy) | A new tab is not saved until Pi's first answer (or a forced save). A restored tab saves under its original id. Another tab's activity does not affect this tab's saves. | Each tab saves its own messages as they complete. |
| Host (desktop or web) | Same mechanism; on Excel on the web workbooks may have no identity, and storage belongs to the browser profile. | No effect. |

Changing any of these mid-turn changes what the next save writes; nothing already saved is rewritten except by the next save of the same session.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | No effect on storage. | The partial model response is saved, marked "Request aborted", and counts as Pi's first answer in a new session. |
| The user sends something else | The draft is never saved; sending makes it a message, saved once the session has been saved before. | Steers and follow-ups are saved when delivered, not when queued. |
| Tab switched, created, closed or reopened | Each change rewrites the tab layout. Closing force-saves the closed tab's session and puts it in recently closed (memory only). | Same. Closing a responding tab after "Stop and close" saves the stopped session. |
| An overlay or menu opens | No effect. | No effect. |
| Model or thinking level changes | A model change forces a save; a thinking-level change waits for the next message. | Model change refused; thinking level saved with the next message. |
| The provider fails | Not applicable. | The failed model response is saved with its "Error:" text and counts as Pi's first answer. |
| The taskpane goes away | The draft is lost. | Completed messages are already saved; the streaming response, the queue and the draft are lost. In a session Pi never answered, nothing is saved and the tab comes back blank. |
| The workbook changes underneath | Excel owns the workbook. Saving it does not clear automatic backups. | A save during a turn does not expire earlier backups. |
| Focus moves to the grid or another window | No effect. Regaining focus re-reads the workbook identity and saves the tab layout under it. | Same. |
| The context window fills | Not applicable. | Compaction's result is saved with the next completed message, not when compaction ends. |

After any interrupt, storage holds the last completed message of every tab that had been saved at least once; nothing partial is ever written.

## Interactions with other systems

**Persistence.** This document.

**Workbook changes and recovery.** Automatic backups persist through saves, AutoSave and reloads. Their 512 MiB ceiling is shared across all workbooks; exceeding it evicts oldest first, without a protected newest floor. Lists use pages of summaries, and restoring loads only the chosen backup. See [backups and revert](../workbook/backups-and-revert.md).

**Other tabs and sessions.** Each tab saves its own session. The tab layout is one record per workbook, rewritten in full on every change. For two taskpanes writing to the same storage, see [several taskpanes](several-taskpanes.md).

**Execution mode and approvals.** The execution mode is saved globally as "yolo" (Auto) or "safe" (Confirm) under `execution.mode.v1` and read once when the taskpane opens (confirmed in Excel on the web: toggling to Confirm in one taskpane stored "safe" while a second taskpane kept running in Auto). Approvals are never saved.

**Models and providers.** Credentials are stored per provider for all workbooks. OAuth access tokens that have expired are refreshed only when the taskpane opens; see [provider errors and offline](provider-errors-and-offline.md).

**Offline and provider errors.** Storage is local, so saving and restoring work offline.

**Notifications.** None. A successful save is silent, and so is a failed one.

**Settings and preferences.** Every setting is saved the moment it changes. Most are read once at startup, so a change made in another taskpane appears here after a reload.

**Host differences.** Excel desktop on macOS and Windows keeps the add-in's storage in the Office webview's data folder; Excel on the web keeps it in the browser profile. They never share it. See [host differences](host-differences.md).

**Language.** The UI language is a saved setting, read at startup. Restored conversations are not translated.

**Accessibility and keyboard.** No interaction.

## Storage failures, quota and clearing

- **Backup failure is visible.** A workbook change already applied stays applied if its backup cannot be stored. Pi reports that the change has no available backup rather than implying it can be reverted. Backup storage failures do not prevent ordinary session/settings writes. Other save paths have their own notification behavior.
- **Quota.** Storage is best effort: the host can refuse writes or evict the whole origin under pressure. Automatic history uses a fixed ceiling rather than trusting free-space estimates; backup-write failures are visible. Neither that ceiling nor being under quota guarantees durability for other Pi data.
- **Storage unavailable at startup.** If the database cannot be opened, startup fails with "Initialization failed: {msg}", or after 60 seconds with "Initialization failed: taskpane timeout".
- **Clearing storage.** Clearing the Office webview's cache (on macOS, deleting `~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/`), the browser's site data on Excel on the web, or running in a private window removes every session, tab layout, setting, login, backup, Files workspace file (sandboxed workspace) and manual backup at once. The next open shows the welcome screen. The document token inside a workbook survives, but nothing refers to it any more.

**Automatic-history upgrade.** Existing backups are copied once, preserving their IDs and contents. The source is deleted only in the same successful commit as the copy and completion marker. Failed migration leaves the old data intact and readable; sessions/settings still open, but new backups may be unavailable until a retry succeeds. The warning says "Backups could not be upgraded. Existing backups are unchanged; new backups may be unavailable until storage recovers."

> Technical note: [#809](https://github.com/tmustier/pi-for-excel/pull/809) advances the shared IndexedDB schema from v2 to v3. The schema upgrade only creates empty stores/indexes; the payload copy runs afterwards so a failed copy cannot block the rest of Pi. Older builds cannot open any Pi data on an upgraded origin: on Mac the old bundle displays the raw lower-database-version initialization error. Test branches on isolated origins; clearing storage to remove this lock also deletes chats and sign-ins. This is a forward-only change, not a reversible data downgrade.

The automatic-history ceiling counts compressed backups and metadata, not all Pi data or database overhead. It does not reserve space or prevent the host evicting the whole origin. Mac storage-usage estimates were stale and persistence was denied; no quota estimate controls retention. Deleting Blob-backed history freed its files immediately in the Mac test. Real disk-full failure and long-term eviction remain unmeasured.

## Edge cases

- A tab where the user sent a message and reloaded before Pi answered comes back blank: the message was never saved.
- Closing a blank tab force-saves an empty session; it is listed in the resume picker as "Untitled". Observed in storage in Excel on the web (a closed blank tab left a one-message record with no title; the one message is Pi's hidden system message).
- Renaming a tab Pi has not answered, then reloading, loses the title; renaming an answered tab saves the title, but the tab still shows "Chat N" after reload.
- `/compact` followed immediately by a reload brings back the conversation as it was before compaction.
- Two unsaved workbooks, or two workbooks on Excel on the web without an identity, share one tab layout: whichever taskpane changed tabs last decides what the next unsaved workbook opens with.
- Opening Pi for the first time in an unsaved workbook restores the most recently changed session from any workbook, if the shared layout slot is empty.
- Save and reload preserve history, including the first backup in an anonymous-link web workbook that can store its document token.
- Automatic history has no count limit. Crossing the shared 512 MiB ceiling removes the oldest backups, potentially all history for another workbook.
- Recently closed holds 10 items per taskpane and is gone on reload, even though the closed sessions themselves are saved and can be found in the resume picker.
- A workbook whose file is moved or renamed outside Excel gets a new identity the next time it opens: its tab layout, links, rules and backups stay behind under the old one.

## Open questions and verification

Backup-specific updates above follow main's merged fixes and [#809](https://github.com/tmustier/pi-for-excel/pull/809); current Mac/web observations and limitations are recorded in [backups and revert](../workbook/backups-and-revert.md#open-questions-and-verification). The older pass below is historical. In particular, its save-clearing, silent-backup-failure and ignored-count-setting observations do not describe the current backup contract. Non-backup persistence observations remain outside this storage update.

- Pass of 2026-10-26 in Excel on the web: restore after a reload (tabs, order, active tab, model and thinking level), the draft lost, the shared layout slot, the stored execution mode, `navigator.storage.persisted()` false, and backups cleared when a taskpane opens were confirmed. Later in the pass, a failed save (write made to fail), compaction before a reload, and a reload with an unanswered message were also confirmed (below). Real quota exhaustion and Save As were not exercised; Save As needs desktop.
- Two tabs both labelled "GPT-6.1 Sol" showed different context windows after a reload ("/ 500k" and "/ 272k"). They were on different providers (`openai-codex` and `openai`); restore keeps the saved provider (`src/taskpane/sessions.ts:54`). See [the status bar](../ui/the-status-bar.md#open-questions-and-verification).
- Composer drafts, recently closed and chosen titles are lost on reload: confirmed in Excel on the web (draft "pd draft probe" gone, Ctrl+Shift+T toasting "No recently closed tab" after a reload in which a tab had been closed before, tab labelled "Chat 1" for a session saved as "Budget review").
- Workbook identity on Excel on the web was null (`document.url` empty and `getFilePropertiesAsync` returning no URL) for a workbook opened through an anonymous edit link; the only tab layout in storage was the shared slot. Confirmed in Excel on the web. Signed-in OneDrive and SharePoint workbooks were not checked.
- A failed save is never reported to the user (`src/taskpane/sessions.ts:264`, `src/taskpane/tab-layout-persistence.ts:62`, `src/workbook/recovery/log-store.ts:43`). Confirmed in Excel on the web with the session write made to fail: only a console warning "[pi] Session save failed: QuotaExceededError", no toast, and after a reload the answered turn was gone. Data loss then only shows after a reload; may be worth treating as a bug.
- Storage persistence is never requested (`requestPersistence` and `getQuotaInfo` in `src/storage/local/app-storage.ts:39-45` have no callers). Read from code; eviction behavior in WKWebView and WebView2 not tested.
- The backups retention setting ("Retention set to {count} backups") is saved but never applied: new backups are trimmed to the fixed 120 (`src/workbook/recovery-log.ts:420`). Read from code. Looks like a bug.
- Compaction is not saved until the next completed message (`src/commands/builtins/export.ts` replaces messages without a save). Confirmed in Excel on the web: after "/compact" at 75% ("Summarized 97,729 tokens") and a reload, the tab came back with the earlier summary ("Summarized 68,496 tokens") and at 75% again.
- A tab whose message got no answer comes back blank: confirmed in Excel on the web with the model request held over the DevTools protocol and the taskpane reloaded; the tab returned as an empty state.
- Restoring the global latest session into an unsaved workbook (`src/taskpane/sessions.ts:111-113`) can show a conversation from an unrelated workbook. Read from code; probably intended as a legacy fallback, but surprising.
- Save As behavior (layout moves, links stay) was read from code and the `office-document-url` tests. Confirmed in Excel for Mac that the new identity is picked up without a reload and that sessions linked to the old name are listed only under "Show sessions from all workbooks"; where the layout is written was not checked.

Verified against pi-for-excel commit `ec453df`.
