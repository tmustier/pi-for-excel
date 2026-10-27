# Stop

## Summary

Stop cancels the turn, compaction and retry waits. It cannot undo a submitted Excel batch. Reads have a short stop grace; ordinary mutations have a longer one. Checked writes and fills report cancellation themselves, preserving a truthful account of applied chunks.

## By phase

| Phase | Effect of Stop |
| --- | --- |
| Forms the call | Partial response stays aborted; no completed call starts. |
| Rejected at once | Aborted before workbook work. |
| Waiting for approval or lock | Outstanding approval withdrawn; queued work cancelled before applying. |
| Before a batch is submitted | Cancellation checked; no new chunk sent. |
| Batch submitted | Cannot recall it. Checked writes await the current batch and report applied rows, or unknown outcome on timeout. |
| Excel never answers | Each Excel run is bounded at 60 seconds. Generic reads stop waiting after 500 ms, generic writes after 10 seconds; checked writes that report Stop themselves use their bounded run instead of being cut short by that grace. |
| Result | Completed results stay; the deliberate abort does not become a raw provider error. |

## What the agent knows next turn

History retains completed tool results and the aborted response. A stopped checked write states which chunks landed and whether remaining rows were not attempted or have an unknown outcome. Unknown means read back before retrying, not assume failure. Undelivered steers and follow-ups become saved Not sent rows, never unasked input next turn. See [stopping](../../product-description/turn/stopping.md).

## Open questions and verification

Read from merged coordinator, action queue and checked write paths at `1f51cf8`. Older delayed-host observations in verification are pre-fix evidence. Real desktop late completion remains untested here.

Verified against pi-for-excel commit `1f51cf8`.
