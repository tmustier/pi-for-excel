# The workbook

## Summary

The workbook is the Excel file open in the window that owns the taskpane, and the only thing in Excel that Pi touches. Pi reads and changes it through its tools, one tool call at a time, inside a [turn](turns.md). This document owns the facts every other document relies on: what Pi can read and what it can change, which tool calls count as a *mutation*, the *workbook lock* that keeps mutations from running at the same time, the automatic *backups* taken before mutations and when they disappear, the *execution mode* (Auto or Confirm), how Pi recognizes "the same workbook" from one session to the next, what Excel's own Undo does with Pi's changes, and what is different for a workbook that has never been saved. Nothing here has its own button: the user meets these rules through tool cards in the transcript, the "Locking…" label on a tab, the "Auto"/"Confirm" control in the status bar, approval dialogs, and the Backups dialog. All of it works with or without a provider connected (only a turn needs one), and on Excel desktop and Excel on the web, with the host differences listed below.

## The simple case

The user asks "Put a total under the Revenue column." Pi reads the sheet first: a tool card "Read Sheet1!A1:D20" appears, runs, and completes. Reading changes nothing and needs no permission. Pi then writes: a card "Edit Sheet1!D21 — 1 changed" appears. Before the value lands in the grid, Pi has recorded what D21 held, so the change can be reverted. The grid shows the new total as soon as the card completes.

Pi did not ask before writing, because the execution mode is Auto by default. If the user had switched the status-bar control to "Confirm", a dialog "Allow workbook mutation in Safe mode?" would have appeared before the write and the turn would have waited for the answer.

If the total is wrong, the user types `/revert` or opens Backups from the [utilities menu](../ui/the-utilities-menu.md) and clicks "Restore": D21 goes back to what it held, and a toast says "Reverted Sheet1!D21". When the user saves the workbook in Excel, the backups for it are cleared, because they are meant to cover the stretch between saves.

## What Pi can read and change

Pi's built-in workbook tools, as the user sees them on tool cards. Each is described in full in [reading](../workbook/reading.md) or [writing](../workbook/writing.md).

| Tool | What it does for the user | Read or mutation | Automatic backup |
| --- | --- | --- | --- |
| `get_workbook_overview` | Lists sheets, their size, header rows, tables, charts and named ranges, or details one sheet. | Read | Not applicable |
| `read_range` | Reads values (as a table or CSV), or values with formulas and number formats. | Read | Not applicable |
| `search_workbook` | Finds cells by value or formula text across visible sheets. | Read | Not applicable |
| `trace_dependencies` | Shows what a cell depends on, or what depends on it, as a tree. | Read | Not applicable |
| `explain_formula` | Explains one formula cell in plain words with its direct references. | Read | Not applicable |
| `write_cells` | Writes values and formulas from a start cell. | Mutation | Values and formulas |
| `fill_formula` | Puts one formula in the top-left cell and fills it across a range. | Mutation | Values and formulas |
| `modify_structure` | Inserts or deletes rows and columns; adds, deletes, renames, duplicates, hides or unhides sheets. | Mutation | The structure change, and deleted values and formulas within the size limit |
| `format_cells` | Fonts, fills, number formats, borders, alignment, wrap, column width, row height, merge. | Mutation | Those format properties |
| `conditional_format` | Adds or clears conditional formatting rules on a range. | Mutation | The range's rules |
| `comments` | Reads comments, or adds, edits, replies to, deletes, resolves or reopens one. | Read for "read", mutation otherwise | The comment thread |
| `charts` | Lists charts and captures a chart image, or creates, updates or deletes a chart. | Read for list and image, mutation otherwise | Create and update only; delete has none |
| `view_settings` | Reads, or changes gridlines, headings, freeze panes, tab color, sheet visibility, standard width, the active sheet. | Read for "get", mutation otherwise | None |
| `workbook_history` | Lists, restores, deletes or clears backups. | Mutation for restore only | Restoring creates a backup of what it overwrote |

Three more built-in tools, `instructions`, `conventions` and `skills`, change Pi's own saved settings (rules, formatting conventions, skills), not the workbook; they count as reads. A few tools outside the built-in set are also present in the default configuration and can change the workbook: direct Office.js execution (`execute_office_js`) and a Python transform that writes results back (`python_transform_range`). Both count as mutations; direct Office.js has no automatic backup. What they do is described in [experimental](../extensions/experimental.md).

Pi has no tool for named ranges, tables, pivot tables, data validation, page layout or printing. It sees named ranges, tables and object counts through the overview, and can only change those through direct Office.js.

> Technical note: classification is by tool name and, for `view_settings`, `comments`, `charts` and `workbook_history`, by the `action` input. A tool the classifier does not know is treated as a mutation.

## The interaction, event by event

Every tool call goes through the same gate. The user sees the states below on the tool card and the tab.

```mermaid
stateDiagram-v2
    [*] --> called : model response asks for a tool
    called --> reading : read (no lock, no approval)
    reading --> done : result (nothing recorded)
    called --> asking : mutation in Confirm mode
    asking --> failed : Cancel, Escape or backdrop (nothing changed)
    asking --> queued : "Allow once"
    called --> queued : mutation in Auto mode
    queued --> writing : lock free (tab may show "Locking…" until then)
    writing --> done : change made (backup stored, audit entry written)
    writing --> failed : Excel error (nothing or part changed)
    queued --> failed : turn stopped before its turn (nothing changed)
    done --> [*]
    failed --> [*]
```

### The tool call arrives

A tool call is decided by the model, not the user. The moment it arrives, Pi classifies it as a read or a mutation from its name and inputs, and reads the current execution mode. Reads go straight to Excel. Mutations go to the approval step (Confirm mode only) and then to the workbook lock. Addresses without a sheet name ("A1:D10") are resolved against whichever sheet is active in Excel at the moment the tool runs, not when the turn started.

### Ends at once

A mutation ends without touching the workbook when:

- **The user declines an approval.** The tool call fails with "Mutation cancelled by user (Confirm mode)." The model sees the failure and the turn continues. See [confirm mode and approvals](../workbook/confirm-mode-and-approvals.md).
- **The turn was stopped while the mutation was waiting.** When its place in the lock comes up, it fails with "Aborted" and does nothing.
- **The tool's own check blocks it.** `write_cells` and `fill_formula` refuse to overwrite cells that hold values or formulas unless the model explicitly asked to overwrite ("⛔ Write blocked — … contains N non-empty cell(s)."). This is not a failure: the card completes with "— blocked" and the model is told to ask the user. A blocked write still went through approval and the lock.

Reads never end this way: they cannot be declined, do not wait, and record nothing.

### The change is made

A mutation commits when it reaches Excel. From then on:

- The change is in the workbook. It stays there whatever happens to the turn (Stop, a provider error, closing the taskpane).
- For tools that take backups, the state of the target before the change has been captured, and the backup is stored once the change succeeds. See [Backups](#backups).
- An entry is added to the local change audit log (exportable with `/export audit`; see [export and copy](../sessions/export-and-copy.md)).
- If the change altered structure (rows, columns, sheets, sheet visibility, chart creation or deletion), Pi refreshes its picture of the workbook before its next request. See [auto-context](../turn/auto-context.md).

### While a mutation runs

The tab is responding, as for any tool call; see [turns](turns.md). While this tab's mutation holds the workbook lock, the tab cannot be closed: its close button is disabled with the tooltip "Wait for workbook changes to complete", and closing it another way shows "Wait for workbook changes to finish before closing this tab". The status bar shows a small "lock" badge while the active tab holds the lock and "lock…" while it waits.

### Finishes

The tool card shows the result: for writes, a "Changes" table of before and after values; for every built-in mutation tool, an "Explain these changes" drawer; and " — no backup" in the card header when no backup could be taken. See [writing](../workbook/writing.md). If Excel rejected the change, the card shows the error and the model sees it. A change that Excel applied partly before failing is not rolled back and has no backup.

## The workbook lock

Only one mutation runs at a time in a taskpane, across all its tabs. Reads are never locked: a tab can read while another tab writes.

When a tab's mutation must wait because another mutation is running or already waiting, the tab shows "Locking…" next to its title until its mutation starts. The wait has no time limit. Mutations run in the order they asked for the lock. Restores started by the user (from the Backups dialog or `/revert`) take the same lock and are attributed to the active tab.

The lock lives in the taskpane's memory. Two taskpanes, for example the same workbook open in two Excel windows, do not share it. See [several taskpanes](../cross-cutting/several-taskpanes.md).

## Backups

A backup is a snapshot of what a mutation is about to change, saved by Pi so the user can [revert](../workbook/backups-and-revert.md) it. The UI calls them "Backups" ("Snapshots saved before Pi modifies data").

- **What is captured.** Depends on the tool (see the table above): for cell writes, the values and formulas of the target range; for formatting, the format properties Pi can set; for conditional formats, the range's rules; for comments, the thread; for charts, the chart's properties; for structure, enough to undo the change, including deleted values and formulas. A backup never captures cell formatting when the mutation was a write, so reverting a write does not undo formatting that Excel's fill copied.
- **When it is not taken.** No backup for `view_settings`, chart deletes, direct Office.js, or a mutation larger than 20,000 cells. No backup when the workbook has no identity and the host cannot store a document token (see [Unsaved and new workbooks](#unsaved-and-new-workbooks)). The tool card says " — no backup" and the result tells the model "ℹ️ Backup not created …".
- **Where it lives.** In the add-in's browser storage on this computer, never in the workbook file. One list holds the backups of every workbook, each tagged with its workbook's identity; the user only ever sees the current workbook's.
- **Retention.** The list keeps the 120 most recent backups across all workbooks; older ones are dropped. The Backups dialog shows at most the 40 most recent for the current workbook. The "Keep at most" setting in the Backups dialog is saved but, at this commit, does not change the cap (see [Open questions](#open-questions-and-verification)).
- **Cleared on save.** Every 4 seconds Pi asks Excel whether the workbook has unsaved changes. When it sees the workbook go from unsaved to saved, it deletes every backup for that workbook. It also deletes them the first time it looks, if the workbook is already saved at that moment; in practice, opening or reloading the taskpane on a saved workbook clears its backups. Hence the warning in the Backups dialog: "Backups will be cleared when this workbook is saved in Excel." If the host cannot report the saved state, backups are never cleared this way.

## The execution mode

The execution mode decides whether Pi applies mutations immediately (**Auto**, the default) or asks before each one (**Confirm**). It is one setting for the whole add-in: all tabs, all workbooks, kept across reloads. It is changed with the status-bar control, the "Auto mode" switch in Settings, or `/yolo`; a change applies from the next mutation and Pi is told about it in its next request. Reads never ask. Some risky actions ask in either mode. All of this is owned by [confirm mode and approvals](../workbook/confirm-mode-and-approvals.md).

## Workbook identity

Pi recognizes a workbook by an identity derived from the file's location (a hash of its path or URL), never by storing the path itself. The identity links sessions, the tab layout, workbook rules, backups and manual backups to the workbook. Consequences the user can see:

- Renaming, moving or "Save As" gives the workbook a new identity. Backups, manual backups and workbook rules stay with the old identity and are no longer shown; sessions linked to the old identity are no longer preferred on restore (see [sessions and tabs](sessions-and-tabs.md)).
- Two copies of the same file in different folders are different workbooks.
- A workbook with no location (never saved) has no identity; see below.

## Unsaved and new workbooks

A workbook that has never been saved ("Book1") has no path-derived identity.

- **Automatic backups** still work on Excel: the first time Pi needs one, it stores a random token inside the document's add-in settings and files the backup under that token. Reading the list never creates the token. When the workbook is saved for the first time it gets a real identity, and the token's backups are cleared at that moment, just as any save clears backups.
- **Manual backups** refuse: "Current workbook identity is unavailable; cannot create or restore manual full-workbook backups safely." See [manual backups](../workbook/manual-backups.md).
- **Workbook rules** cannot be set ("Can't identify this workbook right now — try saving the file first."). See [rules](../guidance/rules.md).

## Excel's Undo

Pi's changes are made through Excel's add-in API, not through the grid. Excel's own Undo (Ctrl/Cmd+Z in the grid, or the Undo button) does not step back through them; the way to take back a Pi change is a [revert](../workbook/backups-and-revert.md). Ctrl/Cmd+Z inside the taskpane is not Excel's Undo either: there it reopens the last closed tab (see [reopening closed tabs](../sessions/reopening-closed-tabs.md)).

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | No effect on what tools can do. The model decides which tools to call. Tool output is cut at a smaller size for models with a context window under 128k tokens (see [reading](../workbook/reading.md)). | Model changes are refused while the tab is busy. |
| Thinking level | No effect. | No effect. |
| Execution mode | Auto: mutations run as called. Confirm: each mutation waits for an approval. | Applies from the next mutation; an approval dialog already open is not affected. |
| Files in the workspace, or an overlay open | Files are not part of the workbook. An open overlay does not pause tools, but an approval dialog replaces any other approval dialog. | No effect on tools already running. |
| Tab state (new, restored, another tab busy) | No difference between new and restored tabs. Another tab's mutation makes this tab's mutation wait ("Locking…"). | Reads are never delayed by another tab. |
| Host (desktop or web) | Excel on the web keeps files saved continuously, which may clear backups within seconds (see [Open questions](#open-questions-and-verification)). Never-saved workbooks only occur on desktop. | No effect. |

A modifier changed while a tool runs takes effect at the next tool call, never in the middle of one.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Not applicable: no tool call exists yet. | A read in progress finishes or is discarded. A mutation waiting for the lock or for approval does nothing once its turn comes. A mutation already in Excel completes and keeps its backup. Escape with an approval dialog open closes the dialog (a decline) instead of stopping. |
| The user sends something else | Not applicable. | A steer is delivered after the running tools finish; it never interrupts a mutation. |
| Tab switched, created, closed or reopened | Not applicable. | Tools keep running in their tab. A tab holding the lock cannot be closed. |
| An overlay or menu opens | Not applicable. | Tools keep running. A second approval dialog closes the first one, which counts as a decline. |
| Model or thinking level changes | Not applicable. | Model change refused while busy; thinking level does not affect tools. |
| The provider fails | Not applicable. | Tools already running finish; their changes stay. |
| The taskpane goes away | Not applicable. | A mutation already sent to Excel may complete without its backup being stored. Waiting mutations are lost. The lock disappears with the taskpane. |
| The workbook changes underneath | Auto-context tells Pi about the user's edits at send time. | Pi's tools see the grid as it is when they run. Addresses without a sheet name follow the sheet the user has active. A revert later writes back the captured state over the user's newer edits in that range. Saving clears all backups. |
| Focus moves to the grid or another window | Not applicable. | Tools keep running. If the user is editing a cell, Excel may hold Pi's calls until the edit ends. |
| The context window fills | Not applicable. | No effect on the workbook; large tool output is cut before it reaches the conversation. |

Whatever interrupts a turn, workbook changes already made stay made and can be reverted from their backups until the next save.

## Interactions with other systems

**Persistence.** Backups and the change audit log are stored in the add-in's browser storage on this computer, separate from sessions. The execution mode is stored there too. Nothing Pi records is written into the workbook file, except the document token for never-saved workbooks. See [persistence and restore](../cross-cutting/persistence-and-restore.md).

**Workbook changes and recovery.** This document owns them. The user-facing recovery flows are in [backups and revert](../workbook/backups-and-revert.md) and [manual backups](../workbook/manual-backups.md).

**Other tabs and sessions.** All tabs share the workbook lock, the backup list and the execution mode. A backup is not tied to the tab or session that made it; any tab can revert it.

**Execution mode and approvals.** See [The execution mode](#the-execution-mode).

**Models and providers.** None, beyond the output-size scaling for small context windows.

**Offline and provider errors.** Tools talk only to Excel, so they keep working while the provider is unreachable; the turn that called them may fail afterwards.

**Notifications.** Tool cards, "Locking…" on tabs, the "lock" badge in the status bar, approval dialogs, and the "Reverted {address}" toast. There is no notification when backups are cleared on save.

**Settings and preferences.** The "Auto mode" switch in Settings, and the retention field in the Backups dialog.

**Host differences.** Desktop Excel on macOS and Windows and Excel on the web run the same tools. Differences come from the host: whether the saved state can be read, continuous saving on the web, and whether a never-saved workbook can carry a document token. See [host differences](../cross-cutting/host-differences.md).

**Language.** Tool card headers are mostly English in every UI language (only "Overview" is translated); tool results are English because they are written for the model.

**Accessibility and keyboard.** Nothing here has its own shortcut. "Locking…" is part of the tab's text; the status-bar lock badges carry tooltips only.

## Edge cases

- A tab whose own model response asks for two mutations at once shows "Locking…" on itself while the first runs, and its close button becomes enabled during that time (see [Open questions](#open-questions-and-verification)).
- A no-op write (the same values written again) still takes a backup, with "0 changes".
- Reading a very hidden sheet works if Pi names it; search skips hidden sheets unless Pi names the sheet.
- Backups for workbook A count against the same 120-entry cap as workbook B; heavy work in one workbook can push out another workbook's backups.
- After a save, `/revert` says "No backups for this workbook yet" even though Pi changed the workbook earlier in the session.

## Open questions and verification

- The "Keep at most" retention value is saved (`src/workbook/recovery/log-store.ts:64`) but never read when backups are stored: `src/workbook/recovery-log.ts:420` always caps the list at `MAX_RECOVERY_ENTRIES` (120), across all workbooks. Looks like a bug.
- Opening or reloading the taskpane on a saved workbook clears its backups on the first poll (`src/workbook/save-boundary-monitor.ts:74`, covered by `tests/recovery-save-boundary.test.ts`). Intentional per the test, but it means backups never survive a taskpane reload unless the workbook has unsaved changes at that moment.
- Excel on the web saves continuously. If `workbook.isDirty` is reported there, backups would be cleared within about 4 seconds of each save, making revert nearly unusable on the web. Not confirmed by hand; the most important item to verify.
- Two taskpanes in the same browser storage (two workbooks open at once) each load the backup list once and then rewrite the whole list on every change (`src/workbook/recovery-log.ts:375-383` and `:420-421`). One taskpane can silently erase backups the other stored after it loaded. Suspected bug; read from code.
- Lock state per tab is set from the most recent coordinator event (`src/taskpane/init.ts:1501-1512`). When one tab has two mutations in flight, the second one's "queued" event marks the tab "waiting" while it actually holds the lock, so it shows "Locking…" and can be closed mid-write; the first one's "completed" marks it idle before the second starts. Suspected bug.
- `workbook_history` delete and clear are classified as reads (`src/tools/execution-policy.ts:73`), so Pi can delete every backup without an approval in Confirm mode and without the lock. Possibly intentional (they do not touch the workbook) but it removes the user's safety net silently.
- The README's scope says direct Office.js execution is off by default, but `createAllTools()` registers `execute_office_js` in every configuration (`src/tools/index.ts:50`) and the system prompt advertises it. This document follows the code.
- Whether Excel's Undo stack is cleared, kept or extended by Pi's changes on each host has not been checked; the glossary's "Excel's own Undo does not undo Pi's changes" is assumed.
- Whether Excel holds add-in calls while a cell is in edit mode, and for how long, has not been checked.
- A write captured but not yet stored as a backup when the taskpane closes is lost. Read from code.

Verified against pi-for-excel commit `ec453df`.
