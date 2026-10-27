# Confirm mode and approvals

## Summary

As of [PR #808](https://github.com/tmustier/pi-for-excel/pull/808) (WT-01,
WT-05), **Auto** is the default: ordinary workbook changes run without a
dialog. **Confirm** asks before each change unless the user has allowed
reversible workbook changes for the rest of this turn. Irreversible actions
still ask in either mode. The dialog is the confirmation; Pi presents a plan
for multi-step work but is not instructed to ask the same question in chat.

### Switching the mode

The mode is one saved setting per browser profile, shared by tabs and
workbooks. The status-bar Auto/Confirm menu, Settings and `/yolo` change it.
The next tool gate reads the current mode; changing mode does not answer a
dialog already open.

## The simple case

### Approval asked

In Confirm, the user asks Pi to fill a column. Reading does not ask. Before
writing, a dialog titled "Allow this change to the workbook?" names the target,
shows humanized inputs and, where available, what existing cells it replaces.

### Answered

The user can choose:

- **Allow once:** run this call only.
- **Allow for the rest of this turn:** run this and later reversible workbook
  calls in this run without another dialog, including requests already queued.
- **Decline:** refuse this call; Pi sees the refusal and may continue.
- **Stop Pi:** withdraw pending requests and stop the turns that asked.

The turn grant is never saved and does not cover the next turn, another tab,
guidance changes, MCP calls, backup deletion or irreversible workbook calls.
Stopping aborts the run and its pending approvals; a stopped grant cannot
approve anything afterwards.

## Risky actions ask in any mode

| Call | Auto | Confirm |
| --- | --- | --- |
| Workbook reads | No | No |
| Reversible workbook mutations | No | Yes, unless allowed for this turn |
| Structural deletes exceeding automatic backup capacity | Always | Always |
| Chart deletion | Always | Always |
| Destructive direct Office.js (delete, clear or remove) | Always | Always |
| Direct code reaching browser authority | Always | Always |
| Other direct Office.js | No | Yes, unless allowed for this turn |
| `workbook_history` delete or clear | Always | Always |
| User-started restore | Only a later-edit overwrite warning when needed | Same |

Structural deletion checks the used data before asking. If more than 20,000
cells would be lost without an automatic backup, the dialog warns "This can't
be undone from Backups. Pi asks for this in every mode." Chart deletes have
the same warning. Direct Office.js uses conservative code scanning; a false
positive costs a dialog, and this heuristic is not a sandbox.

Where the workbook has a canonical identity, destructive dialogs offer
**Take a full backup first**. It captures/downloads the full workbook using
the same mechanism as `/backup`, then allows the call. Failure does not allow
the change. Hosts with only a document-instance identity cannot use this
full-file fallback. See [manual backups](manual-backups.md).

The Python/LibreOffice bridge has its own first-use consent, remembered for
that bridge address. Guidance, Files, extensions and MCP have separate
policies; workbook turn consent never overrides them.

### While the dialog is open

Approvals share one queue across tabs. Parallel requests are shown in arrival
order, with a waiting count and the requesting tab named where needed. A new
request does not decline the one already shown. Approval happens before a
call waits for the workbook lock, so a pending dialog does not hold that lock.
Calls which preflight would refuse are refused without asking.

The rest of the taskpane is inert while the dialog is open. Initial focus is
on the dialog, not Allow, so Enter intended for the composer cannot approve a
change. Tab reaches native buttons; Escape, close and backdrop decline. Focus
returns after closing. Buttons wrap into readable rows in a narrow pane.

## Cancel and interrupt

Stop withdraws both shown and queued requests through the run's abort signal.
It cannot recall an Excel batch already submitted. A secondary preparation
action, such as taking a full backup, is awaited before approval and cannot
allow a withdrawn request afterwards.

## Interactions with other systems

The mode survives reloads. Per-turn permission does not. Declining changes no
cells and creates no backup. Approved changes remain after Stop or failure;
use the [turn receipt or Backups](backups-and-revert.md) to undo them.

The mode is reflected in Pi's system prompt. Switching Auto/Confirm changes
the prompt prefix once; a turn grant does not. Dialog labels and notes use
translated UI keys; tool contracts and agent-facing guidance remain English.
Dialogs work in desktop Excel and Excel on the web without browser-native
prompts. Provider failure cannot answer a local approval.

## `/yolo`

`/yolo status` reports the mode. `on` selects Auto, `off` selects Confirm,
and `toggle` switches. An invalid argument changes nothing and shows usage.
The command does not bypass or answer an open approval.

## Modifiers

Model and thinking choices do not change the consent boundary. Workbook
identity controls whether full-file fallback is available; provider failure
does not approve a request. Tabs share mode, but not per-turn grants.

### Ends at once

Reads, user-started restores without later edits, and rejected preflight
checks do not open a workbook approval. An invalid mode argument changes nothing.

## Edge cases

Decline is not Stop: Pi may try another operation. A turn grant never approves
an irreversible action. Saving the mode is separate from answering the dialog.

## Open questions and verification

PR #808: real GPT-6 Luna turns in Excel on the web wrote two cells in separate
calls after one turn grant; the next turn asked again. Auto asked before a
20,001-cell sheet deletion, and Stop withdrew it without deleting the sheet.
Public-interface tests cover queued/future turn grants, new-run isolation,
irreversible consent and abort during secondary actions. Office.js detection
is tested including bracket and Unicode-escaped forms. Native desktop
acceptance for these additions remains separate from the web evidence.
