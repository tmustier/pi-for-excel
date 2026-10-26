# Stop

## Summary

The user's Stop aborts the turn ([product-description/turn/stopping.md](../../product-description/turn/stopping.md)). For a tool call, what happens depends on the phase it is in. Before Excel is touched, the call is refused with a thrown `Aborted`. Once the tool is running, it runs to completion, because no workbook tool checks the abort signal and an Excel batch cannot be cancelled. The turn then ends and no model call sees the result.

## By phase

| Phase | Effect of Stop |
| --- | --- |
| Forms the call (model streaming) | The partial response is kept as "Request aborted"; no tools run. |
| Rejected at once | The call returns `Aborted` / `Operation aborted` (thrown). |
| Waiting for approval or the lock | Approval stays open; after it resolves, the wrapper checks the signal and returns `Aborted` before running (read from code). |
| Touches Excel / while running | Runs to completion. Writes land. |
| The result | Stored in the session, but the agent sees it only if the user sends another message. |

## What the agent knows next turn

The history shows the tool call and its result (if it finished before the turn ended), followed by an aborted assistant message. The agent is not told the user pressed Stop beyond that. Steers stranded by the Stop button arrive unasked in the next turn (product-description, confirmed there).

## Open questions and verification

- Read from code; Stop during a running write was not timed in this pass.

Verified against pi-for-excel commit `ec453df`.
