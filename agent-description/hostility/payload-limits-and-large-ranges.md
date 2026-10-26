# Payload limits and large ranges

## Summary

Excel caps the size of each request to it and each response from it (about 5 MB on the web). Pi's tools load whole ranges, often with three grids per cell, so large sheets hit the cap. The agent is not told the cap exists, gets Office's raw message when it hits it, and is sometimes told a write failed when it succeeded. The numbers are in [limits](../foundations/limits.md#excel).

## What the agent receives

```
The response payload size has exceeded the limit. Please refer to the documentation: "https://docs.microsoft.com/office/dev/add-ins/concepts/resource-limits-and-performance-optimization#excel-add-ins".
```

wrapped by the tool. It gives no size, no safe chunk and no hint whether anything was changed.

## Per tool (all confirmed on a 200,001 × 5 sheet)

| Tool | Outcome |
| --- | --- |
| `read_range` 10,000 × 5 | Succeeds; truncated to 1,117 rows. |
| `read_range` 50,000 × 5, whole sheet, any `A:A` | Payload error, nothing returned. |
| `search_workbook` without `sheet` | Payload error for every query. |
| `get_workbook_overview {sheet: "Big"}` | Payload error. |
| `get_workbook_overview` | Succeeds (loads sizes only). |
| `fill_formula` 50,000 rows | Succeeds, no backup. |
| `fill_formula` 200,000 rows | Payload error **after filling** (false failure). |
| `write_cells` 30,000 × 3 | Succeeds, no backup. |
| `write_cells` 60,000 × 4 | Payload error **after writing** (false failure). |
| Auto-context selection, whole column | Silently omitted. |
| Auto-context selection, 5,000 rows | Sent: 311,761 characters on every model call. |

## How the agent copes

In a real turn ("how many rows on Big have Region = North, total Qty?"), the agent read a 6-row sample and then wrote its own `execute_office_js` aggregation over `C2:D200001`. It returned the right answer (66,667 rows, total 33,300,338). Agents can work around the cap, but only by using the escape hatch, and only when they guess the cap exists.

## Proposed fixes

[B-02](../bug-triage.md#b-02-large-writes-and-fills-report-failure-after-the-change-is-applied), [B-04](../bug-triage.md#b-04-the-selection-context-has-no-size-limit-and-disappears-for-whole-column-selections), [B-05](../bug-triage.md#b-05-large-sheets-break-reads-search-and-sheet-overviews-with-a-raw-payload-error). Add one prompt sentence: "On large sheets read at most ~10,000 rows × 5 columns per call and never whole columns."

## Open questions and verification

- Desktop limits not measured.

Verified against pi-for-excel commit `ec453df`.
