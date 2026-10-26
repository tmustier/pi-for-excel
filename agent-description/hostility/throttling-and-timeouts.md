# Throttling and timeouts

## Summary

Excel batches are sent with no timeout and cannot be cancelled. If Excel is slow, throttled or never answers, the tool call waits as long as Excel does. A mutation holds the workbook lock for that whole time, so every other mutation in every tab waits too. The second pass loaded Excel on the web with bursts from the taskpane. It never returned a throttling error: overload showed up first as delay, then, after a burst of heavy formula writes, as a host that stopped answering altogether. Every later call hung, a real turn hung with it, and Stop did not end the turn ([B-20](../bug-triage.md#b-20-excel-calls-have-no-timeout-and-a-hung-host-hangs-the-turn-past-stop)).

## What the agent receives

Nothing until Excel answers. If Excel answers with a throttling error, it arrives as an error-shaped result with Office's message (read from code; never seen). The one timeout in the auto-context pipeline is the per-part limit (1.2–2.5 s), and when it hits, that part is silently left out.

Python tools have their own `timeout_ms`; Excel tools have none.

## Observed durations (web)

| Call | Time |
| --- | --- |
| Small `read_range` | 150–350 ms |
| `read_range` 10,000 × 5 | 650 ms |
| `get_workbook_overview`, 9 sheets | 1.6 s |
| `charts get_image` | 3.2 s |
| `fill_formula` 200,000 rows | 3.7 s |
| `write_cells` 60,000 × 4 | 10.2 s |

## Bursts (web, second pass)

| Burst | Outcome |
| --- | --- |
| 60 parallel `read_range` of 100 × 4 | All succeeded in 8.7 s. |
| 40 parallel `write_cells` of one cell | All succeeded; serialized by the lock, 34.3 s in total (about 0.86 s each). |
| 300 parallel raw `Excel.run` reads of 1,000 × 5 | All succeeded in 45 s. |
| 100 parallel raw writes of 2,000 × 5 values | All succeeded in 96 s. A `read_range` issued right after waited 15.3 s. |
| 150 parallel raw writes of 2,000 × 5 volatile formulas | Never answered. 10,800 of 1.5 million formulas landed. Every later `Excel.run` hung (a 2 × 2 read still waiting at 50 s); a real turn's `read_range` spun for 90 s; Stop did nothing. Excel showed only "Saving...". A page reload recovered Excel in 22 s; Pi restored its chats. |

## Proposed fix

[B-20](../bug-triage.md#b-20-excel-calls-have-no-timeout-and-a-hung-host-hangs-the-turn-past-stop).

## Open questions and verification

- Confirmed: TIME-01, TIME-02 (bursts above), STOP-03.
- No throttling error code was ever returned on the web; desktop behaviour was not observed.

Verified against pi-for-excel commit `ec453df`.
