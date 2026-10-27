# Limits

## Summary

This document owns every number that bounds what the agent can send, receive or keep. Other documents link here. Values are for the default configuration. "Measured" values come from Excel on the web on 2026-10-26.

## Excel

| Limit | Value | Source |
| --- | --- | --- |
| Request and response payload per batch | about 5 MB on the web | Office documentation; measured below |
| `read_range` page | 500 rows / 2,500 cells; reduced on host response overflow | code and live web page |
| Whole-column/row reads | clipped to used data, then paged | code |
| Checked writes/fills | chunked; large verification sampled with whole-target error/blank scans | code |
| `search_workbook` | bounded block scan | code |
| Batch timeout | 60 seconds; outcome unknown, host may finish later | code |
| Cell-edit wait | up to 60 seconds; 250 ms initial delay, maximum 2 seconds | code |

## Tool-specific caps

| Cap | Value |
| --- | --- |
| `write_cells` verification preview | 8 rows × 6 columns |
| Change sample in details | 12 cells (UI only) |
| `execute_office_js` code | 20,000 characters |
| `execute_office_js` explanation | 50 characters |
| `execute_office_js` result | 8,000 characters of pretty-printed JSON, then `…` and "ℹ️ Result truncated to 8,000 characters." |
| `explain_formula` references | 8 by default, 20 at most |
| Overview headers per sheet | 8, then "… (+N more)" |
| Sheet preview | 5 rows |
| Automatic backup | 20,000 cells maximum; occupied checked-write targets above the cap refused |
| Recovery retention | 512 MiB compressed payload plus metadata across all workbooks; oldest first, no count/age limit or protected minimum; pages of 50 |

## Per-result truncation

Every tool result's text is cut when the call finishes:

- 50 KB (51,200 bytes of UTF-8) or 2,000 lines, whichever comes first;
- for models whose context window is under 128k, both scale with the window (floors 8 KB and 200 lines);
- head kept for workbook tools; tail kept for `python_run`, `tmux`, `mcp`, `execute_office_js`, `execute_wps_js`;
- notice: `[Output truncated: showing first 1,121 of 10,004 lines (50.0KB limit); limits: 2,000 lines / 50.0KB]`, with `; full output saved to Files workspace: {path}` when the full text was at most 512 KB and could be saved.

Paging now usually bounds read results before this generic truncation layer. A single large cell can still exhaust the output budget.

## History shaping

Before every model call:

- the 6 most recent tool results are kept whole (scaled below 128k windows, floor 2: 3 at 65k);
- older results over 1,200 characters, or holding an image, become a 500-character preview.

Compaction additionally marks answered tool results over 8,000 characters in the kept tail for compacted previews. Full card output stays stored. Results for the still-unanswered request remain whole.

## Auto-context

| Part | Limit |
| --- | --- |
| Workbook identity lookup | 1.2 s |
| Blueprint | 2.5 s |
| Files summary | 1.2 s |
| Selection | 2 seconds; 240 cells, at most 20 columns, 4,000 table characters; tall windows 15 head + 5 tail rows |
| Change list | bounded shared log with a separate cursor per runtime; 10 addresses per sheet displayed |

## Compaction

| Threshold | Value (200k window) |
| --- | --- |
| Reserve for the answer | min(16,384, window/2) |
| Hard trigger | min(window − reserve, 85% of window at ≥200k; 88% at ≥128k) |
| Kept recent tokens | 20,000 |
| Unknown window | treated as 200,000 |

The trigger depends on the selected model. Meter colours follow the soft-warning budget; yellow begins above 75% of that budget and red above it.

## Request size observed

Historical `ec453df` measurements below predate bounded selection and durable notes; they are not current request limits.

| What | Size |
| --- | --- |
| System prompt | 13,675 characters |
| Tool schemas | 36,557 characters (26 tools) |
| First request of a fresh tab | about 53 KB |
| Same, with a 5,000-row selection | 405,530 characters |

## Open questions and verification

- The desktop payload ceiling and any batch-count throttling were not measured.
- The 5 MB figure is from Office documentation; the measured boundaries above are what matters to the agent.

Verified against pi-for-excel commit `1f51cf8` (code review and live read/selection probes).
