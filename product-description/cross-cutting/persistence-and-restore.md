# Persistence and restore

## Summary

Persistence is what Pi remembers when the taskpane goes away. Sessions, layouts, drafts, Not sent rows, chosen titles, stable tab numbers, recently closed state and settings live in this origin's browser storage, not on a Pi server or inside the workbook. A small document token in the workbook lets location-less layouts and backups find the same document again. Saves happen as state changes, with feedback on failures. Open menus, transient toasts and streaming execution do not resume. This document owns what is saved when and where.

## The simple case

The user opens Pi in a saved workbook and asks a question in Chat 1. The first user message saves before its answer; completed responses/tool results update the session. The user opens a second tab and leaves it active.

The user closes the taskpane (or Excel, or reloads the add-in). Nothing happens at that moment: everything that matters was already saved.

Reopening restores tab order, active tab, saved conversations, chosen titles, stable numbers and each tab's draft/Not sent rows. Held messages are never sent automatically. A response interrupted by closing the pane does not continue streaming.

## The interaction, event by event

Persistence has two moments the user can notice: a **save**, which happens without any visible sign, and a **restore**, which happens while the taskpane is starting.

```mermaid
stateDiagram-v2
    [*] --> starting : taskpane opens
    starting --> restoring : workbook identity resolved (layout key chosen)
    restoring --> ready : layout found (tabs rebuilt, active tab selected)
    restoring --> ready : no layout (latest linked session, or a blank tab)
    ready --> ready : first user message or later message completes (session saved)
    ready --> ready : tab created, closed, moved or switched (layout saved)
    ready --> ready : setting changed (setting saved at once)
    ready --> [*] : taskpane goes away (nothing saved; memory-only state lost)
```

### What is saved, and where

| What | Where | When it is written | Scope |
| --- | --- | --- | --- |
| Sessions: messages, model, thinking level, saved title, usage | IndexedDB database `pi-for-excel`, stores `sessions` and `sessions-metadata` | See [When a session is saved](#when-a-session-is-saved) | All workbooks |
| Workbook link and "latest session" pointer | Settings store | With every session save, when the workbook has an identity | Per [workbook identity](#workbook-identity) |
| Tab layout (session ids, order, active tab, titles/numbers, recently closed) | Settings store | Tab/layout changes and scope refresh | Per workbook identity or location-less document scope |
| Recently closed | Saved layout state (up to 10 items) | When a tab closes | Workbook layout scope; survives reload |
| Settings (language, Theme, proxy, execution mode, auto-compaction, rules, conventions, skills, extensions, connections, web search, MCP servers) | Settings store | The moment the user commits them | All workbooks, except workbook rules (per workbook identity) |
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

- **The first save** happens with the first user message, before Pi answers. Draft/held state also persists through the composer-state path.
- **After that**, every completed message saves the whole session: the user's message as it joins the transcript, each model response when it finishes, each tool result, and each steer or follow-up when it is delivered. A model response that is still streaming is never saved part-way.
- **Forced saves** also preserve session changes without waiting for an answer, including closing/duplicating a tab and explicit model/resume actions. Model continuity choices belong to [the picker](../setup/the-model-picker.md), not the former global fork preference.
- **Explicit saves** also preserve chosen titles and completed compaction immediately. Model/resume choices follow [the model picker](../setup/the-model-picker.md); thinking changes for a busy tab apply next turn.
- **Nothing is saved when the taskpane goes away.** There is no save on unload.

Each save rewrites the whole session record. It also links the session to the current workbook if it has no link yet (an existing link is never changed), and records it as the workbook's latest session.

### Restoring when the taskpane opens

1. Pi works out the [workbook identity](#workbook-identity) and reads the tab layout stored for it.
2. **If a layout exists**, Pi opens one tab per listed session, in order. A session that is in storage comes back with its messages, model, thinking level and saved title. A session that is not in storage (it was never saved, or it was deleted) comes back as a blank tab, which is force-saved at once under a new id. The tab that was active becomes active again; if it cannot be found, the first tab does.
3. **If no layout exists** (the first time Pi opens in this workbook), Pi opens one tab. If the workbook has an identity, that tab is filled with the latest session linked to this workbook, if it has any messages; otherwise it is blank. If the workbook has no identity, the tab is filled with the most recently changed session in storage, from any workbook.
4. If the layout could not be read at all (a storage error), Pi starts with one blank tab and leaves the stored layout untouched, so that a later open can still use it. A layout that was read but is malformed is replaced.

Restored tabs retain chosen titles and stable numbers. The earlier renumbering/title-loss observation is historical. Resume/reopen focuses an already-open session instead of duplicating it.

### What a reload brings back and what it loses

| Comes back | Lost |
| --- | --- |
| The tabs open in this workbook, their order, active tab and drafts | Unsaved changes since the latest successful persistence write |
| Each tab's completed messages, tool cards and thinking blocks | The model response that was streaming, and any tool call that had not finished |
| Each tab's saved model/thinking choice and held prompt rows | Live execution of queued commands; held prompts never auto-send |
| Saved titles, stable tab numbers and recently closed state | Transient close toast |
| All settings, logins, rules and conventions | Open settings pages and unsaved editor drafts |
| Automatic backups, including across saves and reloads (subject to deletion, eviction or storage loss) | The error banner, the working indicator, open overlays and menus, toasts |
| Files in the Files workspace | Permission to a connected local folder, if the browser no longer grants it (Pi falls back to the sandboxed workspace) |
| First user message even without an answer | Any in-flight response not yet persisted |

## Workbook identity

Pi recognises "the same workbook" by its file location. It takes the workbook's URL or path as Excel reports it, drops anything after `?` or `#`, and stores only a one-way hash of it (`url_sha256:…`); the raw path is never stored. The workbook name Pi shows (for example on backups) is the file-name part of that location.

- **A saved workbook on disk or in OneDrive/SharePoint** has an identity as soon as the taskpane opens. Pi asks Excel for the current location every time it needs it, so a **Save As** or a **rename** is picked up without a reload, the next time Pi looks (at the latest when the taskpane next regains focus).
- **A workbook without a saved location** uses a document-token layout scope rather than sharing a global layout. File-specific [rules](../guidance/rules.md) and [manual backups](../workbook/manual-backups.md) still need a usable workbook identity. Automatic history can use the document token.
- **Excel on the web** can report no location through an anonymous link. The older global-layout observation is historical; document-token scoping now distinguishes those workbooks. Signed-in location reporting was not separately checked. See [host differences](host-differences.md#workbook-identity-and-saving).

What happens when the identity changes while Pi is open:

| Change | Sessions | Tab layout | Automatic backups | Workbook rules |
| --- | --- | --- | --- | --- |
| First save of a new workbook | The next save of each open session links it to the new identity; sessions not saved again stay unlinked | Written under the new identity from the next layout save; the shared slot keeps its old copy | Token-scoped history moves to the saved identity | Become available, starting empty |
| Save As / rename of a saved workbook | Existing sessions stay linked to the old identity; the resume picker in the new file lists them under "Show sessions from all workbooks" | Written under the new identity; the old file keeps its last layout, so both files reopen the same sessions | When observed on the same open document, history moves without copying, retaining any destination history | Stay with the old identity; the renamed file starts with none |
| Save As a copy of a never-saved workbook | As for a first save | As for a first save | The open document's history follows its new identity; no duplicate history is kept for the old identity | As for a first save |

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | Saved with the session; explicit resume/switch choices follow the model-picker contract. | Busy changes apply to the next turn, not the running one. |
| Thinking level | Saved with the session and restored. | Saved with the next completed message. |
| Execution mode | A global setting, saved the moment it changes. Not part of a session. | Saved at once. |
| Files in the workspace, or an overlay open | Files persist independently of sessions. Overlays are never restored. | Importing during a turn persists the files at once. |
| Tab state (new, restored, another tab busy) | First user message saves; tabs own their draft/held state. | Each tab saves its own messages as they complete. |
| Host (desktop or web) | Same mechanism; on Excel on the web workbooks may have no identity, and storage belongs to the browser profile. | No effect. |

Changing any of these mid-turn changes what the next save writes; nothing already saved is rewritten except by the next save of the same session.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | No effect on storage. | The partial model response is saved, marked "Request aborted", and counts as Pi's first answer in a new session. |
| The user sends something else | Draft and first message persist. | Undelivered prompts become saved held rows; they do not auto-send on restore. |
| Tab switched, created, closed or reopened | Layout and recently closed state persist; closing force-saves the session. | Stop and close saves the stopped session. |
| An overlay or menu opens | No effect. | No effect. |
| Model or thinking level changes | Explicit model choice follows the picker; thinking is tab-owned. | Busy choices apply next turn. |
| The provider fails | Not applicable. | The failed model response is saved with its "Error:" text and counts as Pi's first answer. |
| The taskpane goes away | Persisted drafts restore. | Completed messages, first user message and held rows restore; live streaming/commands do not restart. |
| The workbook changes underneath | Excel owns the workbook. Saving it does not clear automatic backups. | A save during a turn does not expire earlier backups. |
| Focus moves to the grid or another window | No effect. Regaining focus re-reads the workbook identity and saves the tab layout under it. | Same. |
| The context window fills | Not applicable. | Completed compaction saves immediately in its originating tab. |

After any interrupt, storage holds the last completed message of every tab that had been saved at least once; nothing partial is ever written.

## Interactions with other systems

**Persistence.** This document.

**Workbook changes and recovery.** Automatic backups persist through saves, AutoSave and reloads. Their 512 MiB ceiling is shared across all workbooks; exceeding it evicts oldest first, without a protected newest floor. Lists use pages of summaries, and restoring loads only the chosen backup. See [backups and revert](../workbook/backups-and-revert.md).

**Other tabs and sessions.** Each tab saves its own session. The tab layout is one record per workbook, rewritten in full on every change. For two taskpanes writing to the same storage, see [several taskpanes](several-taskpanes.md).

**Execution mode and approvals.** Auto/Confirm is saved globally and synchronized across panes through settings updates/focus refresh. Approvals themselves are not restored.

**Models and providers.** Credentials are stored per provider for all workbooks. OAuth access tokens that have expired are refreshed only when the taskpane opens; see [provider errors and offline](provider-errors-and-offline.md).

**Offline and provider errors.** Storage is local, so saving and restoring work offline.

**Notifications.** Successful saves are normally silent. Failed session/layout persistence reports the problem; a conflicting session save preserves the local conversation as a copy instead of replacing a newer stored revision.

**Settings and preferences.** Changes save as committed. Shared settings/guidance refresh through update events and focus/visibility; this does not merge live conversations or subscribe every open page to all storage changes.

**Host differences.** Excel desktop on macOS and Windows keeps the add-in's storage in the Office webview's data folder; Excel on the web keeps it in the browser profile. They never share it. See [host differences](host-differences.md).

**Language.** The UI language is a saved setting, read at startup. Restored conversations are not translated.

**Accessibility and keyboard.** No interaction.

## Storage failures, quota and clearing

- **Backup failure is visible.** A workbook change already applied stays applied if its backup cannot be stored. Pi reports that the change has no available backup rather than implying it can be reverted. Backup storage failures do not prevent ordinary session/settings writes. Other save paths have their own notification behavior.
- **Quota.** Storage is best effort: the host can refuse writes or evict the whole origin under pressure. Automatic history uses a fixed ceiling rather than trusting free-space estimates; backup-write failures are visible. Neither that ceiling nor being under quota guarantees durability for other Pi data.
- **Storage unavailable at startup.** An actual database error can fail initialization. A slow startup alone is not fatal at 60 seconds: delay feedback appears at 12/60 seconds. Blocked upgrades ask for older panes to close; version changes close the older connection and ask to reopen.
- **Clearing storage.** Clearing the Office webview's cache (on macOS, deleting `~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/`), the browser's site data on Excel on the web, or running in a private window removes every session, tab layout, setting, login, backup, Files workspace file (sandboxed workspace) and manual backup at once. The next open shows the welcome screen. The document token inside a workbook survives, but nothing refers to it any more.

**Automatic-history upgrade.** Existing backups are copied once, preserving their IDs and contents. The source is deleted only in the same successful commit as the copy and completion marker. Failed migration leaves the old data intact and readable; sessions/settings still open, but new backups may be unavailable until a retry succeeds. The warning says "Backups could not be upgraded. Existing backups are unchanged; new backups may be unavailable until storage recovers."

> Technical note: [#809](https://github.com/tmustier/pi-for-excel/pull/809) advances the shared IndexedDB schema from v2 to v3. The schema upgrade only creates empty stores/indexes; the payload copy runs afterwards so a failed copy cannot block the rest of Pi. Older builds cannot open any Pi data on an upgraded origin: on Mac the old bundle displays the raw lower-database-version initialization error. Test branches on isolated origins; clearing storage to remove this lock also deletes chats and sign-ins. This is a forward-only change, not a reversible data downgrade.

The automatic-history ceiling counts compressed backups and metadata, not all Pi data or database overhead. It does not reserve space or prevent the host evicting the whole origin. Mac storage-usage estimates were stale and persistence was denied; no quota estimate controls retention. Deleting Blob-backed history freed its files immediately in the Mac test. Real disk-full failure and long-term eviction remain unmeasured.

## Edge cases

- A first user message survives reload before its answer; live execution does not resume.
- Closing a blank tab force-saves an empty session; it is listed in the resume picker as "Untitled". Observed in storage in Excel on the web (a closed blank tab left a one-message record with no title; the one message is Pi's hidden system message).
- Chosen titles and completed compaction survive reload; drafts/Not sent rows are restored, not sent automatically.
- Location-less layouts use document scope instead of one global slot. Legacy state can migrate once; it is not a promise of live pane synchronization.
- Save and reload preserve history, including the first backup in an anonymous-link web workbook that can store its document token.
- Automatic history has no count limit. Crossing the shared 512 MiB ceiling removes the oldest backups, potentially all history for another workbook.
- Recently closed holds up to 10 persisted items; resume/reopen deduplicates sessions already open.
- A workbook whose file is moved or renamed outside Excel gets a new identity the next time it opens: its tab layout, links, rules and backups stay behind under the old one.

## Open questions and verification

The pass below is historical. Its draft/title/recently-closed loss, global layout, late first save, delayed compaction save, silent failures and missing persistence request are superseded by #789/#795/#796/#798. Backup changes follow [#809](https://github.com/tmustier/pi-for-excel/pull/809); its migration/ceiling evidence and limitations above are retained. These source corrections do not claim a new reload/storage-failure acceptance run.

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
