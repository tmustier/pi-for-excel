# Throttling and timeouts

## Summary

Excel batches are sent with no timeout and cannot be cancelled. If Excel is slow, throttled or never answers, the tool call waits as long as Excel does. A mutation holds the workbook lock for that whole time, so every other mutation in every tab waits too. Stop ends the turn only after the call returns. None of this was observed in this pass: Excel on the web answered every probe, in at most 10 s (a 60,000 × 4 write).

## What the agent receives

Nothing until Excel answers. If Excel answers with a throttling error, it arrives as an error-shaped result with Office's message (read from code). The one timeout in the auto-context pipeline is the per-part limit (1.2–2.5 s), and when it hits, that part is silently left out.

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

## Proposed fix

[B-20](../bug-triage.md#b-20-excel-calls-have-no-timeout-and-cannot-be-stopped).

## Open questions and verification

- Throttling was not reproduced.

Verified against pi-for-excel commit `ec453df`.
