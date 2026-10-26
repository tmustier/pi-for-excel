# Verification: writing and other mutating tools

How to run this file: the fixture workbook from [README.md](README.md#how-to-run-a-pass), Auto mode unless the row says Confirm. Read the target back with `Excel.run` after each row, not with the tool under test.

## tools/writing-cells.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| WRITE-01 | P2 | probe | Simple write result text. | Empty `Calc!C1`. | Write a string. | `Written to …`, Verified values, `Changed cell(s): 1.` | pass |
| WRITE-02 | P2 | probe | Overwrite protection blocks and shows data. | `Spill!A3`. | Write without overwrite. | `⛔ **Write blocked**`. | pass |
| WRITE-03 | P2 | probe | Invalid formula syntax blocked before Excel. | — | `=SUM(A1`. | Blocked, unbalanced parentheses. | — |
| WRITE-04 | P2 | probe | Blocked spill reported as formula error. | `Spill!C1` with C3 occupied. | `=SEQUENCE(4)`. | `C1: #SPILL!`. | pass |
| WRITE-05 | P2 | turn | Agent asks before overwriting on an explicit request. | `Protected!A2` holds text. | "change A2 to say Output". | Agent asks first. | pass (B-18) |
| WRITE-06 | P2 | probe | Unlocked cell on a protected sheet is writable. | `Protected!B2`. | Write with overwrite. | Success. | pass |
| WRITE-07 | P2 | probe | Verification preview capped 8 × 6. | — | Write 1 × 8. | Preview 1×6 of 1×8. | pass |
| WRITE-08 | P1 | probe | Write inside a merge reports success, nothing written (suspected bug). | `Merged` A1:D1. | Write `Merged!B1`. | Success text, empty preview. | fail (B-01) |
| WRITE-09 | P1 | probe | 2×2 over a merge lands 1 of 4 (suspected bug). | `Merged` B4:C6. | Write 2×2 at B4. | `Changed cell(s): 1.` | fail (B-01) |
| WRITE-10 | P1 | probe | Overwriting a spill child breaks the anchor silently (suspected bug). | `Spill`. | Overwrite A3. | Clean result; A1 `#SPILL!`. | fail (B-03) |
| WRITE-11 | P1 | probe | Locked cell: generic message, no protection hint. | `Protected!C1`. | Write. | `You cannot perform the requested operation.` | fail (B-10) |
| WRITE-12 | P1 | probe | Data validation not enforced (suspected bug). | `Valid!A2`. | Write `Maybe`. | Success. | fail (B-03) |
| WRITE-13 | P2 | probe | Unqualified start cell lands on active sheet. | Activate `My Sheet`. | Write `B2`. | `'My Sheet'!B2`. | pass |
| WRITE-14 | P1 | probe | 60,000 × 4 write: payload error after writing (suspected bug). | `Calc`. | Write. | Error; row 60,000 present. | fail (B-02) |
| WRITE-15 | P1 | probe | Text coerced on write (suspected bug). | `Valid!C1`. | Write `00123`, `1/2`, `3-4`, `50%`. | 123, 46024, 46085, 0.5. | fail (B-07) |
| WRITE-16 | P1 | Confirm, probe | Write in Confirm mode with approval declined. | Confirm. | Decline. | Thrown `Mutation cancelled by user (Confirm mode).` | — |

## tools/filling-a-formula.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| FILL-01 | P2 | probe | Simple fill text. | `Tbl!D2:D7`. | `=B2*2`. | Pattern and example formulas. | pass |
| FILL-02 | P2 | probe | Fill next to a table extends it, unreported. | Same. | Read table range. | `Tbl!A1:D7`, header `Column1`. | fail (B-21) |
| FILL-03 | P2 | probe | Comma in range rejected. | — | `A1:A2,B1:B2`. | Single contiguous range error. | — |
| FILL-04 | P2 | probe | 50,000-row fill succeeds. | `Big!G2:G50001`. | Fill. | Success. | pass |
| FILL-05 | P1 | probe | 200,000-row fill: payload error after filling (suspected bug). | `Big!F2:F200001`. | Fill; read F200001. | Error; formula present. | fail (B-02) |
| FILL-06 | P1 | probe | Skipped backup gives the identity reason (suspected bug). | FILL-04. | Inspect details. | "workbook identity unavailable…". | fail (B-19) |

## tools/formatting.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| FMT-01 | P2 | probe | Multi-range formatting works. | `Dates`. | `A1,C1` bold. | `Formatted **Dates!A1,Dates!C1**: bold.` | pass |
| FMT-02 | P2 | probe | Conditional format with a value. | `Tbl`. | `cell_value` > 25. | Added. | pass |
| FMT-03 | P1 | probe | Protected sheet: generic error. | `Protected!A1`. | Bold. | `Error formatting: You cannot perform…` | fail (B-10) |
| FMT-04 | P2 | probe | Formatting a merged cell. | `Merged!B1`. | Bold. | `Formatted **Merged!B1**: bold.` | pass |

## tools/structure-and-view.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| STRUCT-01 | P2 | probe | Delete rows in a filtered table by position. | `Tbl`. | Delete row 3. | Deleted, no mention of table. | pass |
| STRUCT-02 | P2 | probe | Rename a protected sheet succeeds. | `Protected`. | Rename. | Renamed. | pass |
| STRUCT-03 | P1 | probe | Insert rows on protected sheet: generic error. | `Protected`. | Insert. | `Error: You cannot perform…` | fail (B-10) |
| STRUCT-04 | P2 | probe | View change notes no backup. | `Dates`. | Freeze 1 row. | `ℹ️ Backup not created.` | pass |

## tools/charts-and-comments.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| CHART-01 | P2 | probe | Create names the chart `Chart N`. | `Tbl`. | Create with title. | `Chart 'Chart 1' created`. | pass |
| CHART-02 | P2 | probe | `get_image` needs `name`. | — | Use `chart_name`. | `name is required`. | pass |
| CHART-03 | P3 | probe | Capture takes seconds. | — | `get_image`. | ~3 s, 600×360. | pass |
| COMMENT-01 | P2 | probe | Add and read a comment. | `Dates!A2`. | Add; read A1:D4. | One comment by Guest User. | pass |
| COMMENT-02 | P2 | probe | Wrong parameter name gives an internal JS error. | — | `cell` instead of `range`. | `Cannot read properties of undefined`. | fail (B-21) |

## tools/office-js.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| OJS-01 | P2 | probe | Result fenced JSON. | — | Return an object. | `Executed Office.js: …` and JSON. | pass |
| OJS-02 | P2 | probe | Missing load error. | — | Read `values` without load. | `The property 'values' is not available…` | pass |
| OJS-03 | P1 | probe | Protected write: generic error. | `Protected`. | Write. | `You cannot perform…` | fail (B-10) |
| OJS-04 | P3 | probe | 8,000-char cut holds ~130 rows of 5 columns. | `Big`. | Return 2,000 rows. | Truncation note. | pass |

## tools/backups.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| HIST-01 | P2 | probe | List shows backups with IDs. | After writes. | `list`. | Table. | pass |
| HIST-02 | P1 | probe, Confirm | `delete` skips approval in Confirm mode. | Confirm. | `delete`. | No dialog. | — |
