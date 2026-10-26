# Limits

## Summary

This document owns every number that bounds what the agent can send, receive or keep. Other documents link here. Values are for the default configuration. "Measured" values come from Excel on the web on 2026-10-26.

## Excel

| Limit | Value | Source |
| --- | --- | --- |
| Request and response payload per batch | about 5 MB on the web | Office documentation; measured below |
| `read_range` that succeeds | 10,000 × 5 (values, formulas and formats loaded) | measured |
| `read_range` that fails | 50,000 × 5; any whole column | measured |
| `write_cells` that succeeds | 30,000 × 3 (4.2 s) | measured |
| `write_cells` that fails after writing | 60,000 × 4 (10 s) | measured |
| `fill_formula` that succeeds | 50,000 × 1 (1.6 s) | measured |
| `fill_formula` that fails after filling | 200,000 × 1 (3.7 s) | measured |
| `search_workbook` | fails for the whole workbook once any sheet's used range is too big to load | measured (200,001 × 5) |
| Batch timeout | none | code |
| Cell-edit mode | every batch fails in 3–12 ms on the web | measured |

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
| Automatic backup | skipped above 20,000 cells |

## Per-result truncation

Every tool result's text is cut when the call finishes:

- 50 KB (51,200 bytes of UTF-8) or 2,000 lines, whichever comes first;
- for models whose context window is under 128k, both scale with the window (floors 8 KB and 200 lines);
- head kept for workbook tools; tail kept for `python_run`, `tmux`, `mcp`, `execute_office_js`, `execute_wps_js`;
- notice: `[Output truncated: showing first 1,121 of 10,004 lines (50.0KB limit); limits: 2,000 lines / 50.0KB]`, with `; full output saved to Files workspace: {path}` when the full text was at most 512 KB and could be saved.

A 10,000 × 5 compact read was cut at row 1,117. The same read as CSV was cut at line 1,796.

## History shaping

Before every model call:

- the 6 most recent tool results are kept whole (scaled below 128k windows, floor 2: 3 at 65k);
- older results over 1,200 characters, or holding an image, become a 500-character preview.

## Auto-context

| Part | Limit |
| --- | --- |
| Workbook identity lookup | 1.2 s |
| Blueprint | 2.5 s |
| Files summary | 1.2 s |
| Selection | 1.5 s; window ±5 rows, 20 columns, **no row limit** |
| Change list | 50 changes kept; 10 addresses per sheet shown |

## Compaction

| Threshold | Value (200k window) |
| --- | --- |
| Reserve for the answer | min(16,384, window/2) |
| Hard trigger | min(window − reserve, 85% of window at ≥200k; 88% at ≥128k) |
| Kept recent tokens | 20,000 |
| Unknown window | treated as 200,000 |

The model used here has a 500k window, so auto-compaction triggers near 425k tokens.

## Request size observed

| What | Size |
| --- | --- |
| System prompt | 13,675 characters |
| Tool schemas | 36,557 characters (26 tools) |
| First request of a fresh tab | about 53 KB |
| Same, with a 5,000-row selection | 405,530 characters |

## Open questions and verification

- The desktop payload ceiling and any batch-count throttling were not measured.
- The 5 MB figure is from Office documentation; the measured boundaries above are what matters to the agent.

Verified against pi-for-excel commit `ec453df`.
