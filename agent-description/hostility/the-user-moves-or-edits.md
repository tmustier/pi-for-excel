# The user moves or edits

## Summary

The user keeps using Excel while Pi works. They select other cells, switch sheets, and edit the cells the agent just read. This document covers what each of these does to a tool call and what the agent is told about it.

## The active sheet moves

Unqualified addresses are resolved against the active sheet when each batch runs. After the user activated `My Sheet`, `write_cells {start_cell: "B2"}` wrote to `'My Sheet'!B2` (confirmed). The result heading names the sheet; nothing else warns. A multi-batch tool can resolve different batches against different sheets if the user switches between them (read from code). See [B-17](../bug-triage.md#b-17-unqualified-addresses-follow-the-users-active-sheet).

## The selection moves

The selection is read once per model call, into the auto-context. A selection change during a tool call has no effect on the call. The next model call shows the new selection with no note that it changed. The agent can confuse "the user's selection" in the request with the selection now shown.

## The user edits cells

- **Between model calls.** Edits on watched sheets go into the change list and appear in the next model call only ([auto-context](../foundations/auto-context.md#the-change-list)). Edits on sheets created after the taskpane opened never appear. Pi's own writes appear as if the user made them ([B-08](../bug-triage.md#b-08-the-change-list-reports-pis-own-writes-misses-new-sheets-and-lasts-one-model-call)).
- **Between the agent's read and its write.** The agent writes from what it read. If the user changed those cells in between, `write_cells` with `allow_overwrite` overwrites the change. Without `allow_overwrite`, the write is blocked only if the cells were empty when the agent read them.
- **During a write.** Between `write_cells`' pre-read and its write, a user edit is overwritten without a block (read from code).
- **While typing.** See [cell-edit mode](cell-edit-mode.md).

## Structure changes by the user

Inserted or deleted rows shift every address the agent learned. `modify_structure delete_rows` takes a position, so a stale position deletes the wrong row. A sheet the user adds is missing from the blueprint ([B-12](../bug-triage.md#b-12-the-blueprint-goes-stale-and-is-re-sent-on-every-model-call)).

## What the agent is left knowing

Its picture of the workbook is a set of snapshots taken at different moments. Nothing timestamps them, and no tool checks that the state the agent read is still the state it writes into.

## Proposed fixes

[B-08](../bug-triage.md#b-08-the-change-list-reports-pis-own-writes-misses-new-sheets-and-lasts-one-model-call), [B-12](../bug-triage.md#b-12-the-blueprint-goes-stale-and-is-re-sent-on-every-model-call), [B-17](../bug-triage.md#b-17-unqualified-addresses-follow-the-users-active-sheet). Optionally, an `expected` parameter for `write_cells` (the values the agent read) so it can refuse when they have changed.

## Open questions and verification

- Confirmed: MOVE-01 to MOVE-03, AUTO-10 to AUTO-12.
- Read from code: the race windows.

Verified against pi-for-excel commit `ec453df`.
