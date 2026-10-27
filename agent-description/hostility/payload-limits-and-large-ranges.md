# Payload limits and large ranges

## Summary

Excel's per-batch payload ceiling still exists. Pi avoids loading whole requested grids for common reads and checked writes: reads page, searches and sheet previews use bounded scans, and writes/fills chunk both submission and verification. The exact limits belong to [limits](../foundations/limits.md).

## What the agent receives

A read page names the actual address read, the requested extent and the next page. A wide request may also need column paging. Response overflow reduces page rows rather than immediately returning the host's raw error. A write result separates applied, unattempted and unknown-outcome rows, and states the verification scope.

## Per tool (all confirmed on a 200,001 × 5 sheet)

The heading is retained for historical links; current behavior below is code-reviewed unless a live observation is explicitly stated.

| Tool | Current behavior |
| --- | --- |
| `read_range` | At most 500 rows / 2,500 cells per page; whole-line addresses clipped to used data. |
| `search_workbook` | Scans bounded blocks; no need to return a whole used grid. |
| Sheet overview | Uses a bounded preview rather than loading the full sheet. |
| `write_cells`, `fill_formula` | Chunked writes and bounded read-back; no single huge verification response. Occupied targets exceeding the 20,000-cell backup cap are refused rather than overwritten without protection. |
| Large empty targets | May be written without a full backup; report the size reason and sampled versus whole-range checks. |
| Selection | Exact dimensions with bounded head/tail sample, even for a whole column. |

## How the agent copes

Follow the returned continuation address for detail. Prefer search, workbook functions or Python for whole-sheet aggregation; do not add a long page of numbers mentally. Python bridge input limits are checked before sending and include a row estimate.

## Proposed fixes

Paging, bounded sampling, chunked writes and truthful unknown outcomes are merged, not planned. Host limits can still affect a single unusually large cell or unsupported operation.

## Open questions and verification

In Excel on the web at `1f51cf8`, a request for J1:J800 returned J1:J500, next address J501:J800 and page sum 125,250. Selecting J:J produced a 1,274-character context sample with exact 1,048,576-row dimensions, first 15 and last 5 used rows. Desktop ceilings were not measured.

Verified against pi-for-excel commit `1f51cf8`.
