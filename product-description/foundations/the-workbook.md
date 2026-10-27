# The workbook

## Summary

The workbook is the Excel file open in the window that owns the taskpane, and the only thing in Excel that Pi touches. Pi reads and changes it through its tools, one tool call at a time, inside a [turn](turns.md). This document owns the facts every other document relies on: what Pi can read and what it can change, which tool calls count as a *mutation*, the *workbook lock* that keeps mutations from running at the same time, the automatic *backups* taken before mutations and when they disappear, the *execution mode* (Auto or Confirm), how Pi recognizes "the same workbook" from one session to the next, what Excel's own Undo does with Pi's changes, and what is different for a workbook that has never been saved. Nothing here has its own button: the user meets these rules through tool cards in the transcript, the "Locking…" label on a tab, the "Auto"/"Confirm" control in the status bar, approval dialogs, and the Backups dialog. All of it works with or without a provider connected (only a turn needs one), and on Excel desktop and Excel on the web, with the host differences listed below.

## The simple case

The user asks "Put a total under the Revenue column." Pi reads the sheet first: a tool card "Read Sheet1!A1:D20" appears, runs, and completes. Reading changes nothing and needs no permission. Pi then writes: a card "Edit Sheet1!D21 — 1 changed" appears. Before the value lands in the grid, Pi has recorded what D21 held, so the change can be reverted. The grid shows the new total as soon as the card completes. (Confirmed in Excel on the web.)

Pi did not ask before writing, because the execution mode is Auto by default. In Confirm, the queued approval names the tab and intended action before the write can acquire the lock.

If the total is wrong, the user can use `/revert` or Backups from the [utilities menu](../ui/the-utilities-menu.md). Restore previews/confirmation make the overwritten state explicit. Saving and reloading do not clear automatic backups.

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
- **The tool's own check blocks it.** Checked writes/fills refuse occupied targets without overwrite permission, unsafe merge/spill/protection targets and occupied targets too large to back up. Refusals are errors; preflight can refuse before approval, then inspection rechecks under the lock.

Reads cannot be declined and record no mutation. They wait for earlier registered writes in their own runtime and use the bounded Excel host policy.

### The change is made

A mutation commits when it reaches Excel. From then on:

- The change is in the workbook. It stays there whatever happens to the turn (Stop, a provider error, closing the taskpane).
- For tools that take backups, the state of the target before the change has been captured, and the backup is stored once the change succeeds. See [Backups](#backups).
- An entry is added to the local change audit log (exportable with `/export audit`; see [export and copy](../sessions/export-and-copy.md)).
- If the change altered structure (rows, columns, sheets, sheet visibility, chart creation or deletion), Pi refreshes its picture of the workbook before its next request. See [auto-context](../turn/auto-context.md).

### While a mutation runs

The tab is responding, as for any tool call; see [turns](turns.md). While this tab's mutation holds the workbook lock, the tab cannot be closed: its close button is disabled with the tooltip "Wait for workbook changes to complete", and closing it another way shows "Wait for workbook changes to finish before closing this tab". The status bar shows a small "lock" badge while the active tab holds the lock and "lock…" while it waits.

### Finishes

The tool card is the receipt: changed-cell samples, verification scope and backup status. Partial checked writes name and back up what landed. A failed read-back is not called an unapplied write; an unanswered submission may have an unknown outcome. Other partial multi-step changes remain without a guaranteed inverse. See [writing](../workbook/writing.md).

## The workbook lock

Only one mutation runs at a time in a taskpane, across all its tabs. Reads do not take the lock, but wait for writes registered earlier in the same runtime. Reads in another tab can observe a write in progress.

When a tab's mutation must wait because another mutation is running or already waiting, the tab shows "Locking…" next to its title until its mutation starts. The wait has no time limit. Mutations run in the order they asked for the lock. Restores started by the user (from the Backups dialog or `/revert`) take the same lock and are attributed to the active tab.

The lock lives in the taskpane's memory. Two taskpanes, for example the same workbook open in two Excel windows, do not share it. See [several taskpanes](../cross-cutting/several-taskpanes.md).

## Backups

A backup is a snapshot of what a mutation is about to change, saved by Pi so the user can [revert](../workbook/backups-and-revert.md) it. The UI calls them "Backups" ("Snapshots saved before Pi modifies data").

- **What is captured.** Depends on the tool (see the table above): for cell writes, the values and formulas of the target range; for formatting, the format properties Pi can set; for conditional formats, the range's rules; for comments, the thread; for charts, the chart's properties; for structure, enough to undo the change, including deleted values and formulas. A backup never captures cell formatting when the mutation was a write, so reverting a write does not undo formatting that Excel's fill copied.
- **When it is not taken.** No backup for `view_settings`, chart deletes, direct Office.js, or a mutation larger than 20,000 cells. No backup when the workbook has no identity and the host cannot store a document token (see [Unsaved and new workbooks](#unsaved-and-new-workbooks)). The tool card says " — no backup" and the result tells the model "ℹ️ Backup not created …".
- **Where it lives.** In the add-in's browser storage on this computer, never in the workbook file. One list holds the backups of every workbook, each tagged with its workbook's identity; the user only ever sees the current workbook's.
- **Retention.** "Keep at most" applies per workbook, from 5 to 120 (default 120), with a global cap of 240. The dialog lists the full retained set so counts and Clear all agree.
- **Save and reload.** Neither deletes backups, on desktop or web. Only explicit deletion and retention remove them; the old save-boundary warning and clearing behavior are gone.
- **Other taskpanes.** Storage mutations reload current entries under a shared storage lock, avoiding stale-list replacement. This is not a cross-taskpane workbook write lock.

## The execution mode

The execution mode decides whether Pi applies mutations immediately (**Auto**, the default) or asks before each one (**Confirm**). It is one setting for the whole add-in: all tabs, all workbooks, kept across reloads. It is changed with the status-bar control, the "Auto mode" switch in Settings, or `/yolo`; a change applies from the next mutation and Pi is told about it in its next request. Reads never ask. Some risky actions ask in either mode. All of this is owned by [confirm mode and approvals](../workbook/confirm-mode-and-approvals.md).

## Workbook identity

Pi recognizes a workbook by an identity derived from the file's location (a hash of its path or URL), never by storing the path itself. The identity links sessions, the tab layout, workbook rules, backups and manual backups to the workbook. Consequences the user can see:

- Renaming, moving or "Save As" gives the workbook a new identity. Backups, manual backups and workbook rules stay with the old identity and are no longer shown; sessions linked to the old identity are no longer preferred on restore (see [sessions and tabs](sessions-and-tabs.md)).
- Two copies of the same file in different folders are different workbooks.
- A workbook with no location (never saved) has no identity; see below.

## Unsaved and new workbooks

A workbook that has never been saved ("Book1") has no path-derived identity. On Excel on the web, a workbook opened through an "anyone with the link" sharing link without signing in has none either: Excel reports no location for it, so it behaves exactly like a never-saved workbook (confirmed in Excel on the web).

- **Automatic backups** use a durable document token when no path identity is available. Saving does not delete them; a different identity can change which list is shown. Reading the list does not create a token.
- **Manual backups** refuse: "Current workbook identity is unavailable; cannot create or restore manual full-workbook backups safely." (confirmed in Excel on the web). See [manual backups](../workbook/manual-backups.md).
- **Workbook rules** cannot be set ("Can't identify this workbook right now — try saving the file first."). See [rules](../guidance/rules.md).

## Excel's Undo

Pi's changes are made through Excel's add-in API, not through the grid. On Excel on the web they still join Excel's own Undo stack: Ctrl/Cmd+Z in the grid (or the Undo button) takes back Pi's writes and Pi's restores one step at a time, and Ctrl/Cmd+Y puts them back (confirmed in Excel on the web). Pi is not told: its backups still describe the change, and a later revert writes the captured state over whatever Undo left. In one run the stack stopped short of Pi's first write of the session, the one that created the document token, which suggests that saving the token clears Excel's Undo history; this is a hypothesis. Whether Excel desktop behaves the same has not been checked. A [revert](../workbook/backups-and-revert.md) remains the way to take back a specific Pi change regardless of what happened since. Ctrl/Cmd+Z inside the taskpane is not Excel's Undo: there it reopens the last closed tab (see [reopening closed tabs](../sessions/reopening-closed-tabs.md)).

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | No effect on what tools can do. The model decides which tools to call. Tool output is cut at a smaller size for models with a context window under 128k tokens (see [reading](../workbook/reading.md)). | Model changes are refused while the tab is busy. |
| Thinking level | No effect. | No effect. |
| Execution mode | Auto: mutations run as called. Confirm: each mutation waits for an approval. | Applies from the next mutation; an approval dialog already open is not affected. |
| Files in the workspace, or an overlay open | Files are not part of the workbook. An open overlay does not pause tools; approvals queue. | No effect on tools already running. |
| Tab state (new, restored, another tab busy) | No difference between new and restored tabs. Another tab's mutation makes this tab's mutation wait ("Locking…"). | Reads are never delayed by another tab. |
| Host (desktop or web) | Saving and taskpane loads never clear backups. Workbooks without a location can use a document token. Excel's Undo is not Pi's cross-host recovery guarantee. | No effect. |

A modifier changed while a tool runs takes effect at the next tool call, never in the middle of one.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Not applicable: no tool call exists yet. | Pending approvals and lock waits cancel. Checked writes settle submitted chunks and report applied work and backup status. Other submitted host calls have bounded grace and may report unknown outcome. Escape dismisses the top modal before acting on the turn. |
| The user sends something else | Not applicable. | A steer is delivered after the running tools finish; it never interrupts a mutation. |
| Tab switched, created, closed or reopened | Not applicable. | Tools keep running in their tab. A tab holding the lock cannot be closed. |
| An overlay or menu opens | Not applicable. | Tools keep running. Further approvals queue behind the current request. |
| Model or thinking level changes | Not applicable. | Model change refused while busy; thinking level does not affect tools. |
| The provider fails | Not applicable. | Tools already running finish; their changes stay. |
| The taskpane goes away | Not applicable. | A mutation already sent to Excel may complete without its backup being stored. Waiting mutations are lost. The lock disappears with the taskpane. |
| The workbook changes underneath | Auto-context reports new external edits before requests, mid-turn included. | Reads/writes see current grid state; the lock does not exclude co-authors. User restore previews later edits and asks before replacing them. Saving keeps backups. |
| Focus moves to the grid or another window | Not applicable. | Tools continue. Edit-mode rejections can wait/retry for up to 60 seconds before the first successful sync; a submitted mutation is not replayed after sync. |
| The context window fills | Not applicable. | No effect on the workbook; large tool output is cut before it reaches the conversation. |

Whatever interrupts a turn, workbook changes already made stay made and can be reverted while their backups are retained; saving does not remove them.

## Interactions with other systems

**Persistence.** Backups and the change audit log are stored in the add-in's browser storage on this computer, separate from sessions. The execution mode is stored there too. Nothing Pi records is written into the workbook file, except the document token for never-saved workbooks. See [persistence and restore](../cross-cutting/persistence-and-restore.md).

**Workbook changes and recovery.** This document owns them. The user-facing recovery flows are in [backups and revert](../workbook/backups-and-revert.md) and [manual backups](../workbook/manual-backups.md).

**Other tabs and sessions.** All tabs share the workbook lock, the backup list and the execution mode. A backup is not tied to the tab or session that made it; any tab can revert it.

**Execution mode and approvals.** See [The execution mode](#the-execution-mode).

**Models and providers.** None, beyond the output-size scaling for small context windows.

**Offline and provider errors.** Tools talk only to Excel, so they keep working while the provider is unreachable; the turn that called them may fail afterwards.

**Notifications.** Tool cards, "Locking…" on tabs, the "lock" badge, queued approval dialogs and revert outcomes. Saving does not clear backups.

**Settings and preferences.** The "Auto mode" switch in Settings, and the retention field in the Backups dialog.

**Host differences.** Tools share paging, checked-write and bounded wait rules. Anonymous-link workbooks may lack a location; optional state APIs differ. Neither desktop nor web clears backups on save. See [host differences](../cross-cutting/host-differences.md).

**Language.** Tool card headers are mostly English in every UI language (only "Overview" is translated); tool results are English because they are written for the model.

**Accessibility and keyboard.** Nothing here has its own shortcut. "Locking…" is part of the tab's text; the status-bar lock badges carry tooltips only.

## Edge cases

- Parallel mutations in one tab can wait on that tab's own earlier write. Counted lock ownership keeps the tab non-closeable while any write holds the lock.
- A no-op write (the same values written again) still takes a backup, with "0 changes".
- Reading a very hidden sheet works if Pi names it; search skips hidden sheets unless Pi names the sheet.
- Backups are retained per workbook, but the global 240-entry cap can still evict old entries across workbooks.

## Open questions and verification

- Merged coordinator/recovery source inspected: per-runtime read barrier, counted lock ownership, durable backups, applied retention, atomic storage updates and always-approved deletion. Old WB/REVERT/APPROVE findings remain in the recorded runs, not as present-tense defects here.
- A real GPT-6 Luna turn read and wrote a formula; independent Excel read-back returned 8 and the receipt included a backup ID. A merge-child write was refused. See [the workbook checklist](../verification/workbook.md) for the wider acceptance work still required.
- Optional host APIs, desktop acceptance, concurrent-pane recovery and unknown-outcome timing were not all rerun. A pane closed between write and backup storage can still lose the backup.

Verified against pi-for-excel commit `1f51cf8` (merged source; old dated findings above are historical, superseded by the current contracts).
