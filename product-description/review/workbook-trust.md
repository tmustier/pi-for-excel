# Review: workbook trust and safety

This review asks whether the documented behavior is the best behavior for the people who use it. It does not check whether the documents describe that behavior correctly. Users are analysts and finance professionals working in Excel desktop or on the web, sometimes with co-authors. The review also looks from the model's side.

The target is `main` as documented (`ec453df`) plus the fixes in flight, especially [#787](https://github.com/tmustier/pi-for-excel/pull/787), [#790](https://github.com/tmustier/pi-for-excel/pull/790) (the "Consent outside the workbook" rule), [#791](https://github.com/tmustier/pi-for-excel/pull/791), [#792](https://github.com/tmustier/pi-for-excel/pull/792), [#793](https://github.com/tmustier/pi-for-excel/pull/793) and [#794](https://github.com/tmustier/pi-for-excel/pull/794). Defects already in [bug-triage.md](../bug-triage.md) or [agent-description/bug-triage.md](../../agent-description/bug-triage.md) are not raised again. Code references are to the PR branches.

## Summary

The overall model is right: Auto by default, a backup before every change it can cover, Confirm for users who want to review each change, and reads that never ask. The fix PRs close most of the trust gaps found so far: the approval queue, backups that survive a save, writes that report what actually happened, and auto-context that is stored and accurate.

Three structural problems remain.

1. **Auto mode lets through workbook changes that cannot be undone.** Under #790, Auto asks before deleting a *file* Pi didn't create. It does not ask before deleting a 50,000-cell *sheet*, deleting a chart, or running direct Office.js, and none of those can be restored from Backups. #791 already refuses overwrites too large to back up, but deletes of the same size go ahead.
2. **Backups now last longer, but restores are as blunt as before.** #794 rightly stops deleting backups on save. But a restore still writes old values over whatever is in the cells now, without checking or saying so. So the longer a backup lives, the more likely `/revert` is to destroy later work.
3. **The model isn't told about changes it didn't make in its own tab.** #792 drops every change made through the add-in. That includes the user's own reverts and other tabs' writes, so the agent keeps reasoning from cells that are no longer there.

## Proposals, ranked

| ID | Proposal | Kind | Touches |
| --- | --- | --- | --- |
| WT-01 | Auto asks before workbook changes that cannot be undone | product call | #790, #791, #794 |
| WT-02 | Restore checks for later edits and says how old the backup is | fix | #794 |
| WT-03 | Backups expire by age as well as by count | product call | #794 |
| WT-04 | The approval dialog gets a way out (Stop) and no Enter-to-allow default | fix | #794, #793 |
| WT-05 | Confirm asks once per change, not twice, and offers "Allow for this turn" | product call | #794 (prompt) |
| WT-06 | Tell the agent about reverts and other tabs' writes | fix | #792, #794 |
| WT-07 | Put the backup handle in the result text | fix | #791, #794 |
| WT-08 | Report formulas replaced by constants | fix | #791 |
| WT-09 | Approvals say what will be overwritten | fix | #794, #791 |
| WT-10 | Mutations require a sheet-qualified address (B-17) | product call | #787, #791 |

### WT-01: Auto asks before workbook changes that cannot be undone

- **Planned behavior.** [The workbook § Backups](../foundations/the-workbook.md#backups): no backup for `view_settings`, chart deletes, direct Office.js, or anything above 20,000 cells. Auto asks for none of these. In `modify_structure` on #794 (`src/tools/modify-structure.ts:147`, `:217`, `:273`), `delete_rows`, `delete_columns` and `delete_sheet` with more than 20,000 cells of data log "Checkpoint capture was skipped" and delete anyway. `execute_office_js` runs unasked in Auto when its code uses only the Excel API ([confirm mode § Risky actions](../workbook/confirm-mode-and-approvals.md#risky-actions-ask-in-any-mode)), and it has no backup. Even a restored sheet comes back without its formatting.
- **Proposed.** Apply #790's Auto-mode rule ("asks … before irreversible loss") to the workbook. In Auto, a change that will have no backup and destroys content asks first, with a dialog such as "This can't be undone from Backups. Delete sheet 'Data' (52,000 cells)?". That covers structure deletes over the cap, chart deletes, and Office.js code that deletes or clears. Non-destructive changes without a backup (`view_settings`, freeze panes) still run. Where the workbook has an identity, the dialog offers "Take a full backup first" (`/backup`).
- **Why.** A finance user leaves Auto on because they were told every change can be reverted. A deleted assumptions sheet is the one change they can't get back, and on the web Excel's Undo is the only fallback. #791 already reasons this way for overwrites ("Fill refused: … a backup can't hold 200,000 cells"), and #790 for files. Without this, the most destructive workbook action is the least protected. Claude Code draws the same line: checkpoints cover file edits, and "actions that affect remote systems … can't be checkpointed. You control those with your permission mode" ([How Claude Code works](https://code.claude.com/docs/en/how-claude-code-works)).
- **Cost and risk.** Low to medium. The tool already knows a backup will be skipped before it deletes. For Office.js, telling "destructive" from "not" is a heuristic, like the existing risk lint. A false positive costs one dialog.
- **Product call:** should Auto ever destroy unbackable content without asking? Recommended: no.

### WT-02: Restore checks for later edits and says how old the backup is

- **Planned behavior.** [Backups and revert § Cancel and interrupt](../workbook/backups-and-revert.md#cancel-and-interrupt): "A restore writes the captured state over the user's newer edits in the same range without warning (confirmed …)". `/revert` restores the latest backup "whatever made it". On #794, backups survive saves and reloads, and are capped at 120 per workbook.
- **Proposed.** Before restoring cell values, compare the cells with what Pi left there (the after-state is already computed for the Changes diff; `src/workbook/recovery-log.ts:299-323`). If they differ, ask: "3 cells in Sheet1!B2:D20 have changed since Pi's edit. Restore anyway / Cancel." The `/revert` toast and the dialog say whose change and how old it is: "Reverted Pi's write to Sheet1!B2:D20 from yesterday 16:40". Pi's own `workbook_history restore` gets the same check as an error the model can relay.
- **Why.** On `main`, backups mostly lasted only until a save, so a stale restore was rare. #794 makes them last days. A user who types `/revert` on Monday to undo "the last thing Pi did" can wipe Friday afternoon's manual work, or a co-author's. Restoring an old write after rows were inserted writes the values to the old address, which now holds other data ([edge cases](../workbook/backups-and-revert.md#edge-cases)).
- **Cost and risk.** Medium: the after-state has to be stored with the backup, which is about twice the size per backup. Structure and format restores can start by checking only the range shape (they already fail on a shape change).
- **Fix.** It is a precondition for #794's longer retention.

### WT-03: Backups expire by age as well as by count

- **Planned behavior.** #794: "Saving never deletes backups; only 'Keep at most N' does", with N up to 120 per workbook and a global cap of 240 (`src/workbook/recovery/constants.ts`).
- **Proposed.** Keep "saving never deletes". Also expire backups after a fixed age (for example 30 days, or when the workbook has not been opened in Pi for that long), shown in the dialog as "Backups are kept on this computer for 30 days". Replace the old warning "Backups will be cleared when this workbook is saved in Excel." with that sentence.
- **Why.** Each backup is a copy of cell contents in the browser's storage, outside the file's permissions and retention. For finance users, confidential numbers then outlive the file and stay readable on a shared or managed machine. A count cap alone also keeps a six-month-old backup at the top of `/revert` for a workbook touched once (see WT-02). And the global cap of 240 lets heavy work in one workbook evict another's backups.
- **Cost and risk.** Low: an age filter where retention already runs.
- **Product call:** how long backups are kept, and whether to show it.

### WT-04: The approval dialog gets a way out and no Enter-to-allow default

- **Planned behavior.** #794's `src/ui/approval-dialog.ts`: two buttons, Decline and the confirm label. Initial focus is on "Allow once" (`confirmButton.focus()`, `data-autofocus`), and everything behind the dialog is `inert`, including the Stop button. Requests queue with "N more waiting"; only Stop withdraws them.
- **Proposed.** Add "Stop Pi" to the dialog: it declines this request and every queued request from the same turn, then stops the turn. Put initial focus on the dialog itself (or on Decline), not on Allow.
- **Why.** The dialog opens on its own, mid-turn, often while the user is typing a steer. With focus on Allow, the Enter that was meant to send the message approves a change the user hasn't read. With five changes waiting and Stop unreachable, the only way out is to decline five times, and after each decline the model is told to carry on.
- **Cost and risk.** Low. The queue already withdraws a stopped turn's requests.
- **Fix.**

### WT-05: Confirm asks once per change, not twice, and offers "Allow for this turn"

- **Planned behavior.** [Confirm mode § Edge cases](../workbook/confirm-mode-and-approvals.md#edge-cases): "Pi often asks in chat before calling the tool … and the dialog then asks again (confirmed)". The Confirm prompt says "Ask for explicit user confirmation before mutating workbook tools" and "reconfirm" destructive operations (`src/prompt/system-prompt.ts:80-87`). Workflow rule 7 adds "present a plan and get approval first". Every call asks separately, so building a model is a dozen dialogs.
- **Proposed.** In Confirm, the dialog is the confirmation. The prompt keeps "present a plan before multi-step changes" but drops "ask for explicit confirmation before mutating" and "reconfirm". The dialog adds a third choice, "Allow for the rest of this turn" (never persisted, withdrawn by Stop).
- **Why.** Asking twice teaches users to click through, which defeats Confirm. Comparable tools offer a scoped grant: Claude Code's approval prompt can allow a tool for the rest of the session ([permissions](https://code.claude.com/docs/en/permissions)), and VS Code agent approvals have session-scoped allow levels ([approvals](https://code.visualstudio.com/docs/agents/run/approvals)). Copilot in Excel separates "plan" from "edit" as modes rather than asking per edit ([FAQ](https://support.microsoft.com/en-us/office/frequently-asked-questions-about-agent-mode-in-excel-1cfd906d-40b4-46be-8e2d-65b893e28a02)).
- **Cost and risk.** Low for the prompt (a one-off prompt-cache change). Medium for the turn-scoped grant. Risk: a grant covers changes the user didn't see. Scoping it to one turn and keeping destructive changes (WT-01) always asking limits that.
- **Product call:** is Confirm "review every change" or "review the plan"?

### WT-06: Tell the agent about reverts and other tabs' writes

- **Planned behavior.** #792's change tracker skips every event with `triggerSource: "ThisLocalAddin"` (`src/context/change-tracker.ts:176`) so that Pi's own writes aren't reported as the user's. User-started restores (Backups dialog, `/revert`), another tab's Pi, and `execute_office_js` all write through the same add-in, so they are skipped too. [Backups and revert](../workbook/backups-and-revert.md#interactions-with-other-systems) already says "Pi in other tabs is not told that a revert happened".
- **Proposed.** The coordinator already attributes every mutation to a tab. Feed the tracker from it: changes made by another tab, or by a user restore, become a stored note for this tab, for example "The user reverted Sheet1!B2:D20 (your write_cells from 10:42)" or "Another Pi tab changed Calc!A1:F40".
- **Why.** After the user reverts, the agent still believes its write is in place. Its next "verified" summary, or its next formula that depends on those cells, is wrong, and it has no reason to re-read. This is the honesty gap #792 fixed for user edits, reopened for the user's most deliberate edit.
- **Cost and risk.** Low to medium. These are cache-safe stored notes, the same mechanism #792 introduced.
- **Fix.**

### WT-07: Put the backup handle in the result text

- **Planned behavior.** Backup IDs are only in `details`, which never reaches the model ([agent tools/backups.md](../../agent-description/tools/backups.md#the-simple-case): "mutation results never told it the IDs"). The model learns that a backup exists only when none was made.
- **Proposed.** Every backed-up mutation ends with one line: "Backup `5d2c0a1e` saved; `/revert` or workbook_history restore undoes this."
- **Why.** The model can then say truthfully "I can undo that", restore exactly its own change instead of guessing from a time-sorted list, and tell the user how to undo it themselves. A few tokens in the result tail, so no effect on the prompt cache.
- **Cost and risk.** Trivial.
- **Fix.**

### WT-08: Report formulas replaced by constants

- **Planned behavior.** #791 reports cells not written, text Excel converted, validation breaks and formula errors. It does not report that a write replaced formulas with typed values. With #791's rule ("If the user asked to change those cells, set allow_overwrite=true"), the model overwrites without asking more often.
- **Proposed.** When a write replaces formula cells with constants, the result says so: "⚠️ 12 formulas replaced by values in Model!D5:D16 (D5 was =C5*(1+$B$2))". The tool card shows the same line.
- **Why.** Hard-coding over formulas is the classic silent error in finance models. Results still look right today and are wrong once an input changes. The prompt already says "Prefer formulas", but nothing detects when that rule is broken. #791 already loads the before-formulas (`src/tools/mutation/range-write.ts:278`).
- **Cost and risk.** Low.
- **Fix.**

### WT-09: Approvals say what will be overwritten

- **Planned behavior.** #794's dialog shows the card summary and the input ("Write Sheet1!D1", with start, size and data). #794's `checkBeforeApproval` already inspects the target, and #791 asks it to reuse `inspectWriteTarget`.
- **Proposed.** When the target holds data, the dialog adds a line: "Replaces 40 existing cells, including 12 formulas". When the change will have no backup, it adds "Can't be undone from Backups".
- **Why.** Confirm users approve risk, not addresses. "Write Sheet1!D1:D40" is safe on an empty column and harmful on the forecast. The information is already read before the dialog opens.
- **Cost and risk.** Low once #791 and #794 share the inspection.
- **Fix.**

### WT-10: Mutations require a sheet-qualified address

- **Planned behavior.** [B-17](../../agent-description/bug-triage.md#b-17-unqualified-addresses-follow-the-users-active-sheet) is open as a product call. Unqualified addresses resolve against the sheet the user has active when the batch runs (confirmed: a write landed on `'My Sheet'!B2` after the user switched sheets). #787 makes reads say which sheet they used. Writes still follow the user.
- **Proposed.** Option A: mutating tools reject unqualified addresses with "Name the sheet, e.g. Sheet1!B2". Reads keep the fallback and the #787 note.
- **Why.** Users keep working while Pi runs, and switching sheets is the most ordinary thing they do. A write on the wrong sheet passes overwrite protection when the cells there are empty, and nothing flags it. The cost to the model is one retry, once.
- **Cost and risk.** Low: a schema-description change plus a check (one-off prompt-cache change).
- **Product call** (already pending as B-17). Recommended: A.

## Right as planned; do not change

- **Auto by default, with backups as the safety net.** Asking on every edit trains users to click through. Copilot in Excel's edit mode also applies changes directly and relies on undo and version history ([FAQ](https://support.microsoft.com/en-us/office/frequently-asked-questions-about-agent-mode-in-excel-1cfd906d-40b4-46be-8e2d-65b893e28a02)), and Claude Code pairs automatic edits with checkpoints ([checkpointing](https://code.claude.com/docs/en/checkpointing)). WT-01 tightens Auto's edges; it does not change the default.
- **Reads never ask, never lock, never back up.** Analysis must stay frictionless, and asking for reads would make Confirm unusable.
- **#794's approval queue.** Nothing is declined without a click, each call is asked once and before the lock, and a call the tool would refuse is refused without asking. Each rule removes a way for the model to be told "the user refused" when they didn't.
- **#794's "saving never deletes backups".** The save boundary meant nothing on the web or with AutoSave, and Excel's Undo is no substitute on desktop. Keep it, with WT-02 and WT-03 alongside.
- **`workbook_history` delete and clear ask in every mode (#794).** Removing the undo path is the one local-only action that is irreversible for the workbook.
- **#791's split between blocking and reporting.** Protection, merges and spills are refused with nothing written. Conversions, validation and errors are written and reported. Overwrites that can't be backed up are refused. The result then always matches the workbook.
- **#791's allow_overwrite rule.** An explicit "change A2" is consent. With backups up to 20,000 cells and refusal above, the extra round trip bought nothing. WT-08 and WT-09 cover the remaining risk.
- **#790's Auto rules for code and files.** Asking before new code runs and before Pi destroys a user's file, and failing closed without a dialog, is the right line. WT-01 extends the same line to the workbook.
- **Restores make their own backup, and user-started restores never ask.** The user already chose them, and a mistaken revert can be reverted.
- **#792's stored, change-only notes.** They fixed both honesty (user edits, structure) and cost (cache-stable prefix). WT-06 reuses the mechanism.

## Needs Thomas's call

WT-01 (Auto and unbackable destructive changes), WT-03 (backup age limit), WT-05 (what Confirm means; turn-scoped grant), WT-10 (B-17, option A).

Reviewed against `ec453df` and the heads of #787 and #790–#794 on 27 October 2026. Evidence comes from the documents, the PR descriptions and code on the PR branches. No new model turns were run for this review.
