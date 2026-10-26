# The lock and parallel calls

## Summary

All tool calls from one model call start together. Mutating calls are serialized by the per-taskpane workbook lock, across every tab. Read calls are not serialized at all. This document covers what that means for the agent: reads that race writes, writes that wait behind other tabs, and Confirm-mode approvals that cancel each other.

## Parallel calls in one batch

The agent loop runs a batch's calls in parallel, and the request sets `parallel_tool_calls: true`. Results come back in call order. The agent is not told that order of *execution* is not guaranteed.

- **Read beside write.** `write_cells Calc!Z1 = "new"` and `read_range Calc!Z1` issued together: the read returned `old` (confirmed). The read ran before the write took the lock. See [B-14](../bug-triage.md#b-14-a-read-in-the-same-batch-as-a-write-can-see-the-old-value).
- **Write beside write.** Serialized in the order they reach the lock, which is call order in practice.

## Other tabs

Another tab's mutation holds the lock. This tab's mutations wait, with no time limit and no note in the result about the wait. Reads do not wait, so they can observe another tab's half-finished multi-batch change. The agent is never told another tab exists.

## Racing approvals

In Confirm mode each mutating call asks for approval *before* taking the lock, so all mutating calls in a batch ask at once. All approval dialogs share one overlay id. Opening the second closes the first, which resolves as declined. Calling the dialog twice: the first resolved `false` without a click (confirmed). The agent receives `Mutation cancelled by user (Confirm mode).` for every call but the last. By its prompt, it should treat that as the user's refusal. See [B-09](../bug-triage.md#b-09-parallel-confirm-mode-approvals-decline-each-other).

## Classification quirks

- `execute_office_js` always takes the lock, even for pure reads, and so waits behind other tabs.
- `workbook_history` `delete` and `clear` never take it.

## Open questions and verification

- Confirmed: LOCK-02, LOCK-05.
- Read from code: cross-tab waiting; no lock timeout.

Verified against pi-for-excel commit `ec453df`.
