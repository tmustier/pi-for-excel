# The user moves or edits

## Summary

The user keeps using Excel while Pi works. They select other cells, switch sheets, and edit the cells the agent just read. This document covers what each of these does to a tool call and what the agent is told about it.

## The active sheet moves

Reads can resolve unqualified addresses against the active sheet and disclose that choice. Guarded mutations now refuse an omitted sheet instead of following a selection change. The older observation that `write_cells {start_cell: "B2"}` followed the user to `My Sheet` is historical [B-17](../bug-triage.md#b-17-unqualified-addresses-follow-the-users-active-sheet), not the merged write contract. Explicit targets prevent this ambiguity, not stale-intent or co-author races.

## The selection moves

The selection is read once per model call, into the auto-context. A selection change during a tool call has no effect on the call. The next model call shows the new selection with no note that it changed. The agent can confuse "the user's selection" in the request with the selection now shown.

## The user edits cells

- **Between model calls.** Changed edit context is stored as a conversation note, not drained into one request only ([auto-context](../foundations/auto-context.md#the-change-list)). Added sheets acquire listeners. Identifiable own writes are suppressed, with range-scoped suppression when Excel reports Unknown; Remote changes are labelled as other people's edits. These notes do not lock cells or guarantee every remote race is observed.
- **Between the agent's read and its write.** The agent writes from what it read. If the user changed those cells in between, `write_cells` with `allow_overwrite` overwrites the change. Without `allow_overwrite`, the write is blocked only if the cells were empty when the agent read them.
- **During a write.** Between `write_cells`' pre-read and its write, a user edit is overwritten without a block (read from code).
- **While typing.** See [cell-edit mode](cell-edit-mode.md).

## Structure changes by the user

Inserted/deleted rows shift earlier addresses; a positional deletion can still use stale intent. Sheet/table/name lifecycle events now invalidate the blueprint, and changed structure is a durable note rather than the old permanently stale or repeatedly injected snapshot.

## What the agent is left knowing

Its picture of the workbook is a set of snapshots taken at different moments. Nothing timestamps them, and no tool checks that the state the agent read is still the state it writes into.

## Proposed fixes

The change-tracker, blueprint and explicit-mutation-target fixes are merged. Read/write races remain: re-read when necessary. There is no expected-old-value compare-and-swap parameter. Backup restore checks for later edits, which is separate from ordinary write preconditions.

## Open questions and verification

- Confirmed: MOVE-01 to MOVE-03, AUTO-10 to AUTO-12.
- Read from code: the race windows.

Verified against pi-for-excel commit `1f51cf8` (source inspection; MOVE/AUTO observations are historical).
