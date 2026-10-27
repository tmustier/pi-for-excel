# Throttling and timeouts

## Summary

Each Excel run has a 60-second timeout. On timeout the tool ends and releases its lock, but the host batch cannot be cancelled and may finish later. The outcome is unknown: read back before retrying a mutation. Repeated timeouts explicitly tell the agent to stop calling workbook tools and ask the user to reload Excel. Stop has separate bounded grace rules; see [Stop](stop.md).

## What the agent receives

The tool result on timeout says Excel did not answer, the request may or may not have been applied, and the affected range must be read back. Explicit throttling errors retain Office's reason. Auto-context uses shorter per-source deadlines and does not wait through cell-edit mode.

Python tools have their own `timeout_ms`; Excel runs use the shared 60-second bound. An operation containing several runs can take longer overall.

## Observed durations (web)

The following measurements and bursts are historical observations at `ec453df`, before paging, chunking and timeouts merged; they are not current limits.

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

Merged: bounded Excel runs, outcome-unknown errors and bounded Stop. Host batches can still finish late; this is not cancellation at the Excel server.

## Open questions and verification

- Confirmed: TIME-01, TIME-02 (bursts above), STOP-03.
- No throttling error code was ever returned on the web; desktop behaviour was not observed.

Verified against pi-for-excel commit `1f51cf8` (code review; historical measurements labelled above).
