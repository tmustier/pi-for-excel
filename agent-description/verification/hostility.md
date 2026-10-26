# Verification: hostility

How to run this file: the fixture workbook from [README.md](README.md#how-to-run-a-pass). `user` rows need real input in the grid (CDP mouse and key events on the OneDrive page target).

## hostility/cell-edit-mode.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| EDIT-01 | P1 | user, probe | Reads fail in ms during edit mode. | Click a cell, type, no Enter. | `read_range`. | Edit-mode error, <20 ms. | fail (B-06) |
| EDIT-02 | P1 | user, probe | Writes fail; nothing written. | Same. | `write_cells`. | Edit-mode error. | fail (B-06) |
| EDIT-03 | P1 | user, turn | A real turn stops and asks the user to press Enter. | Three-step prompt; start typing after step 1. | Send. | Step 2 fails; turn ends. | fail (B-06) |
| EDIT-04 | P2 | user, probe | Registering a change handler fails during edit mode. | Same. | `onChanged.add`. | Exception. | pass |
| EDIT-05 | P1 | desktop | Desktop queues instead of failing. | Desktop Excel. | Same as EDIT-01. | Record. | fail. desktop (macOS): with a cell in edit mode (typing in the formula bar, no Enter), `range.load` failed after 0.55 s with "Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.", as on the web. Nothing is queued. |

## hostility/the-user-moves-or-edits.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| MOVE-01 | P2 | user, probe | User edits are reported with triggerSource Unknown; Pi's with ThisLocalAddin. | `onChanged` listener. | Edit; write. | Two sources. | pass |
| MOVE-02 | P1 | probe | Unqualified write follows the active sheet. | Activate `My Sheet`. | Write `B2`. | `'My Sheet'!B2`. | fail (B-17) |
| MOVE-03 | P2 | user, turn | The agent cannot answer "what did I change?" after a tool call. | Edit; ask with a tool call first. | Send. | "I wasn't given an edit log". | fail (B-08) |

## hostility/the-lock-and-parallel-calls.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| LOCK-01 | P2 | 2 tabs | A mutation waits for another tab's lock. | Two tabs. | Concurrent writes. | Serialized. | — |
| LOCK-02 | P1 | probe | Parallel read sees the pre-write value. | `Calc!Z1 = old`. | Write `new` and read together. | `old`. | fail (B-14) |
| LOCK-05 | P1 | probe | Second approval dialog declines the first. | — | Two `requestConfirmationDialog` calls. | First `false` without a click. | fail (B-09) |
| LOCK-06 | P1 | turn, Confirm | Parallel mutations in Confirm mode: all but the last declined unseen. | Confirm mode, GPT-6 Luna. | "make A1:B1 bold, give C1 a yellow fill, and write the word checked in H2". | One dialog (write); both `format_cells` `Mutation cancelled by user (Confirm mode).` | fail (B-09) |

## hostility/payload-limits-and-large-ranges.md

See READ-10, READ-11, OVER-05, SEARCH-02, FILL-05, WRITE-14, AUTO-06, AUTO-07.

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| SIZE-01 | P2 | turn | The agent works around the limit with `execute_office_js`. | `Big`. | Ask a count/sum question. | Escape hatch used; right answer. | pass |

## hostility/recalculation.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| RECALC-01 | P1 | probe | Manual mode: new formula verified as 0. | Manual. | Write `=A1+A2`. | `0.0000e+0`. | fail (B-03) |
| RECALC-02 | P2 | turn | The agent discovers Manual via the escape hatch. | Manual. | Ask for a formula's value. | `calculate()`, 102. | pass |

## hostility/context-pressure.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| CTX-01 | P2 | probe | Truncation notice text. | `Big`. | Read 10,000 × 5. | `[Output truncated: …]`. | pass |
| CTX-02 | P1 | turn | Shaped preview after 6 results. | Long turn. | Inspect request. | `[Compacted tool result]`. | — |
| CTX-03 | P2 | turn | `/compact` below 20k tokens does nothing. | 14k-token session. | `/compact`. | "Nothing to compact". | pass |
| CTX-04 | P1 | turn | Compaction summary lists answered work as in progress; frees little. | 63k-token session with large reads. | `/compact`; inspect `compactionSummary`. | Next Steps repeat done reads; 62,934 → 58,752 tokens. | fail (B-24) |

## hostility/co-authoring-and-autosave.md, throttling-and-timeouts.md, stop.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| COAUTH-01 | P1 | 2nd user | Remote edits arrive with source Remote and are not distinguished. | Second guest profile (CDP 9224). | Co-author types in `Calc!Z3`. | `{source: "Remote", triggerSource: "Unknown"}`. | fail (B-08) |
| COAUTH-02 | P1 | 2nd user, probe | Pi's writes keep `ThisLocalAddin` while co-authored. | As above. | Raw and tool writes after the co-author joins. | `triggerSource: "Unknown"`, also after the co-author leaves. | fail (B-08) |
| COAUTH-03 | P1 | 2nd user, probe | A co-author's Enter overwrites Pi's verified write. | Co-author types in `Co!B4`, no Enter. | Pi writes 555; co-author presses Enter. | Result "Verified 555"; cell 9. | fail (B-23) |
| COAUTH-04 | P1 | 2nd user, probe | A co-author's row insert moves Pi's write. | Co-author inserts a row above 2. | Pi writes `Total`, `=SUM(B2:B5)` at `Co!A6`. | Result A6:B6; cells in row 7, `=SUM(B3:B6)`. | fail (B-23) |
| COAUTH-05 | P2 | 2nd user, probe | A co-author's edit mode does not block Pi. | Co-author mid-edit. | `write_cells`, `read_range`. | Both succeed. | pass |
| SAVE-01 | P1 | probe | Backups survive on the web with unknown identity. | Writes over 30 min. | `list`. | Still listed. | pass |
| STOP-01 | P2 | turn | Stop before the write batch runs: nothing written, no result. | `Excel.run` delayed 8 s before the batch. | Stop at 5 s into `write_cells` J1:J10. | `(no output)`, J empty; next turn "no tool result". | pass |
| STOP-02 | P2 | turn | Stop after the batch ran: write lands, result kept. | Return delayed 6 s after the batch. | Stop while waiting (K1:K10). | K written; result shown; `Error: signal is aborted without reason`. | pass |
| STOP-03 | P1 | turn | Stop ends a turn whose Excel call hangs. | Hung host (TIME-02). | Stop twice. | Turn kept running until reload. | fail (B-20) |
| TIME-01 | P2 | probe | Bursts throttle with an error. | — | 60 reads, 40 writes, 300 raw reads, 100 raw writes in parallel. | All succeed, slower; no error. | pass |
| TIME-02 | P1 | probe | A hung batch never returns. | — | 150 parallel 10,000-cell volatile-formula writes. | No answer; later calls hang ≥ 50 s; recovered by reload. | fail (B-20) |
