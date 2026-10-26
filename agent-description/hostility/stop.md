# Stop

## Summary

The user's Stop aborts the turn ([product-description/turn/stopping.md](../../product-description/turn/stopping.md)). For a tool call, what happens depends on the phase it is in. Before Excel is touched, the call is refused with a thrown `Aborted`. Once the tool is running, the turn ends but the call's result depends on whether the write batch had reached Excel. If Excel never answers, Stop does not end the turn at all ([B-20](../bug-triage.md#b-20-excel-calls-have-no-timeout-and-a-hung-host-hangs-the-turn-past-stop)). The second pass timed all three with real turns (GPT-6 Luna), delaying `Excel.run` from the taskpane console to hold a write in flight.

## By phase

| Phase | Effect of Stop |
| --- | --- |
| Forms the call (model streaming) | The partial response is kept as "Request aborted"; no tools run. |
| Rejected at once | The call returns `Aborted` / `Operation aborted` (thrown). |
| Waiting for approval or the lock | Approval stays open; after it resolves, the wrapper checks the signal and returns `Aborted` before running (read from code). |
| Touches Excel, write not yet sent | Every `Excel.run` delayed 8 s before its batch ran; Stop pressed 5 s into `write_cells` of `Co!J1:J10`. The card showed `(no output)` and "Request aborted". The write never landed, although queued auto-context reads ran later (confirmed). |
| Touches Excel, write sent | Each batch ran at once but its return was delayed 6 s; Stop pressed while waiting. `Co!K1:K10` was written, the card showed the full result, then `Error: signal is aborted without reason` (confirmed). |
| Excel never answers | Stop pressed twice, 25 s and 100 s into a hung `read_range`; the turn kept running until the page was reloaded (confirmed). |
| The result | Stored in the session if the call finished, and the agent sees it only if the user sends another message. |

## What the agent knows next turn

The history shows the tool call and its result (if it finished before the turn ended), followed by an aborted assistant message. The agent is not told the user pressed Stop beyond that. Asked next turn to quote what it had received, the agent said for J: "I received no tool result for a J-column write, so I can't confirm that write happened", and for K quoted the full success result. For a call cut off by a hung host it said the result was "No result provided" (confirmed). In each case the agent could not tell whether the workbook had changed and did not offer to check. Steers stranded by the Stop button arrive unasked in the next turn (product-description, confirmed there).

## Open questions and verification

- Confirmed: STOP-01 to STOP-03 (real turns, delayed `Excel.run`).
- The user-visible `Error: signal is aborted without reason` after a completed write is a raw `AbortSignal` message ([B-21](../bug-triage.md#b-21-small-result-format-slips)).

Verified against pi-for-excel commit `ec453df`.
