# Verification: integration-fixes against the agent experience

**Build under test.** `integration-fixes` @ `a2b3f2d` (PRs #786–#799 merged onto main `fe7013f`), then the force-pushed rebuild @ `e48d83c`. After the rebuild, only the changes listed in INTEGRATION.md were retested: the #794 fill, the format_cells backup, the compaction toast and the selection note. Everything else was observed on `a2b3f2d`. AUTO-09, STRUCT-01, TIME-01 and TIME-02 were then run on `e48d83c`, and EDIT-05 was routed to the Mac verification thread. The two builds do not differ in the code those rows exercise. The force-pushed head **`8099cd6`** was then checked for the rows its fixes touch: OJS-03, N-1, N-2, N-4, N-5, B-16, SIZE-01, analyst task 5 and STRUCT-01, plus N-3 and the new #794, #789 and #793 commits (see "Rechecked on 8099cd6"). Finally **`cea7588`** was checked for the timeout handling only (#791 e1c3f29 `runMutation`, #797 47e8799): see "Rechecked on cea7588".

**Setup.** The tests ran on Excel on the web in an orb (`node scripts/orb/excel-web-session.mjs --check`), with the dev server serving the branch checkout. The workbooks were anonymous guest workbooks: `orb-afa50678b930-20261027013600.xlsx`, then `…20261027022640.xlsx`. The fixture follows [README.md](README.md#how-to-run-a-pass): Dates, Merged, Protected, Tbl (table Sales, filtered to North), Spill, Valid, My Sheet, Calc, Co, and Big (200,001 × 5).

Inputs came from three places:
- **Probes:** tools wrapped the way the runtime wraps them, called over CDP. The Python tools also went through `applyExperimentalToolGates`.
- **Real turns:** GPT-6 Luna (`openai-codex`, 272k window, thinking High), sent through the composer.
- **Grid input:** CDP mouse and key events, for the user on port 9223 and for a co-author in a second guest profile on port 9224.

Every model request was captured from CDP `Network.requestWillBeSent` on the taskpane target. Bodies are in `/tmp/vt/out/requests.jsonl` in the orb. The Python bridge ran in real mode with `PYTHON_BRIDGE_MODE=real`, without pandas, plus LibreOffice. Pyodide was tested with the bridge stopped. About 300 probe calls and 60 real turns.

**Result meaning.** `pass`: the build does what the fixing PRs say (#787, #791, #792, #794, #797, #798), or the checklist's expectation holds. A row that recorded a defect on main passes when the defect is gone. `fail`: it does not. `blocked`: not run in this pass, with the reason.

## Counts

| Set | Pass | Fail | Blocked |
| --- | ---: | ---: | ---: |
| Bug triage B-01 to B-25 | 25 | 0 | 0 |
| Checklist rows (foundations-and-reading, writing, hostility) | 125 | 0 | 1 |
| Analyst tasks | 4 succeeded | 1 partial | — |

Regressions found on `a2b3f2d`:
- The `fill_formula` timeout above about 4,500 cells.
- The refused format_cells backup.

Both are fixed and verified on `e48d83c`. Of the six low-severity bugs found on `a2b3f2d`, five are fixed on `8099cd6`; N-2 was only partly fixed there and is fixed on `cea7588` (see the last section).

Counts are as of `8099cd6`: B-16, OJS-03 and SIZE-01 moved from fail to pass, and task 5 from partial to succeeded.

## Bug triage (B-IDs)

| ID | Result | Note |
| --- | --- | --- |
| B-01 | pass | write_cells to `Merged!B1` and a 2×2 at `B4` are blocked with `isError`, the merge is named, and nothing is written. python_transform_range and fill_formula into merges are blocked the same way. |
| B-02 | pass | 60,000×4 write_cells: success in 3.5 s, partial-verification note, row 60,000 present, no-backup reason. Fill of 200,000 rows: **failed on a2b3f2d** (Excel `Timeout` after about 40 s, 4,500 formulas left behind while the result said "nothing was written"). **Passes on e48d83c**: 0.75 s, partial verification, first and last formulas correct. |
| B-03 | pass | Spill child write blocked with the anchor named. Manual calculation: `=A1+A2` verified as 3, with a Manual note. Validation violation reported. Gap: the Manual note appears only when formulas are written (new bug N-1). |
| B-04 | pass | 5,000-row selection: a 1,201-char note. `Dates!B:B`: "1,048,576 rows × 1 column" plus a sample with row numbers and column letters. |
| B-05 | pass | read_range on 50k rows and whole columns pages at 500 rows and names the next range. `search_workbook` across the whole workbook found `Big!B200000` in 2.4 s. The Big sheet overview works. |
| B-06 | pass | read_range waited 10.2 s and showed "Waiting for you to finish editing a cell". A write held for 58 s returned an `excel_cell_edit_mode` error saying nothing was written. A real turn waited about 17 s and then completed. The selection note says "Unavailable: Excel is in cell-edit mode". |
| B-07 | pass | Reads: `01/01/2024 [45292]`, `5.0%`, `0`, `a\|b`, `line1\nline2`, `1e-7`; a text `00123` stays `00123`. Writes: four conversions reported with stored and shown values and the apostrophe hint; `'00123` stored as text. |
| B-08 | pass | Pi's own writes are left out. An edit on a sheet added after start (`UserAdded2!F11`) is reported. Notes are stored and appear in every later call. Another tab's Pi writes are labelled. Co-author edits and row inserts are labelled, with a "may have moved" warning. |
| B-09 | pass | Confirm mode with three parallel mutations: one dialog at a time ("2 more waiting", "1 more waiting"), all three applied. |
| B-10 | pass | Protected sheet: write_cells, format_cells, modify_structure insert and fill_formula return `isError` and name the protection. Missing sheet: the sheet list is returned. Not covered: `execute_office_js` (OJS-03, in no PR's scope). **8099cd6:** execute_office_js is covered too (see OJS-03). |
| B-11 | pass | Overview has "Current state" (protected sheets) and the sheet view has "Sheet state" (spills). read_range marks merged, hidden or filtered rows, the table filter, spills, protection and Manual calculation. |
| B-12 | pass | A sheet added by the user produced a "Workbook structure (structure changed)" note, sent once and stored. Each request's input was the previous input plus new items in 70 of 70 consecutive call pairs, except where B-16 shaping and compaction rewrote history. |
| B-13 | pass | `=A1+A2` and `=SUM(A1:A2)` both trace to A1 and A2. |
| B-14 | pass | A write of `new` and a read of `Calc!Z1` in the same batch: the read returned `new`. |
| B-15 | pass | Pages of at most 500 rows, summaries placed before the grid, CSV pages, and the next range named. |
| B-16 | pass | Fixed on 8099cd6 (#792, commit 8cc8bf1). The wording is fixed ("no longer in your context; run the tool again"). The shaping boundary now moves in steps of 6. In one turn with 20 sequential ~1.5k-char `read_range` calls (21 requests in one session), earlier tool output was rewritten only twice: when the result count reached 12 and 18 (requests 13 and 19). Every other consecutive pair was an exact prefix extension. cacheRead per call: 0, 9,344, 9,344, 0, 12,800, 13,568, 14,464, 15,360, 11,904, 16,256, 17,152, 18,944, **9,344**, 17,408, 18,176, 19,072, 19,968, 20,864, **9,344**, 19,328, 20,096. The two bold drops are the step rewrites. The drops at calls 4 and 9 had an unchanged prefix, so they are on the provider side. `prefixChangeReasons` was `[]` on all 23 calls; the tracker still watches only model, system prompt and tools, so it does not show a shaping rewrite. On a2b3f2d the boundary moved on every call from the 7th result on. The agent answered correctly (799). |
| B-17 | pass | Product call made: an unqualified read says "the active sheet; the address named no sheet". An unqualified write follows the active sheet and names it in the result. |
| B-18 | pass | The prompt rule is reworded. In real Auto-mode turns the agent set `allow_overwrite` directly when asked. |
| B-19 | pass | "Too large to back up: 240,000 cells exceed the 20,000-cell backup limit … `/backup` first". |
| B-20 | pass | With `Excel.run` hung: read and write time out at 60 s with `excel_timeout`, and the second message says to ask the user to reload. In a real turn with every `Excel.run` hung, Stop ended the turn 2.6 s after it was pressed, and later turns worked. |
| B-21 | pass | Selection grid has row and column labels. `#1 priority` is not reported as an error. execute_office_js returns compact JSON. The table "Column1" note appears. comments with `cell` gives a clear error. The dialog title says "Allow this change to the workbook?". The Stop-after-write wording is now "outcome unknown" (see N-4). **8099cd6:** "the change stands" now shows after a slow write (N-4 fixed). |
| B-22 | pass | Bridge: `ModuleNotFoundError … local Python (bridge at https://localhost:3340)` in 30 ms. A timeout returns 504 after 2.0 s with no second run. Pyodide: backend named and stderr shown. Leftover: the description still promises pandas (N-3). |
| B-23 | pass | The co-author's Enter overwrote Pi's verified `Co!B4 = 555` with 9, and the next request said "By a co-author: Co: B4 (now 9)". A co-author row insert was reported with the may-have-moved warning. The agent re-read before writing its Total. |
| B-24 | pass | `/compact` on 4×2 read turns (~31k tokens): "Summarized 4 messages", every request under Done, "Next Steps: No pending steps", 7 kept results shrunk to previews, context meter 17% → 2%. The agent then answered from memory correctly. |
| B-25 | pass | stdout and stderr appear in the transform result. Ragged output padded with a warning. Merged destination blocked. Text conversion reported. A 20,000-row input refused before running with "About 6,084 rows fit" (the 6,084-row block then succeeded). Output limit reported with a row estimate. Pyodide has no limit (20,000 rows OK). |

## Checklist rows

### foundations-and-reading.md

| ID | Result | Note |
| --- | --- | --- |
| TOLD-01 | pass | Sections in the documented order. |
| TOLD-02 | pass | 26 tools; tool schemas 35,510 chars; prompt 13,976 chars (larger after #787, #791 and #794 rewrote descriptions, as the PRs say). |
| TOLD-03 | pass | The same 26 tools on every call (one `additional_tools` item, hash unchanged). |
| AUTO-01 | pass | Stored notes go before the user's prompt; the selection note is last. |
| AUTO-02 | pass | Exactly one selection note per request; it is not stored. |
| AUTO-03 | pass | No "context missing" re-sends; the blueprint is sent once. |
| AUTO-04 | pass | A user-added sheet appears in the next request's note. |
| AUTO-05 | pass | Selection window with row numbers and column letters. On e48d83c it ends with "(Background from Excel, not a message from the user …)". |
| AUTO-06 | pass | `Big!A1:A5000`: a 1,201-char note. |
| AUTO-07 | pass | Whole column described by address, size and sample. |
| AUTO-08 | pass | Cell-edit mode gives "Unavailable: Excel is in cell-edit mode". |
| AUTO-09 | pass | Rerun on e48d83c. With nothing relevant in the workspace (only `assistant-docs/`), `getContextSummary` returns `hasRelevantFiles: false`, and no Files note appeared in any of 60 captured requests. After `imports/probe-auto09.csv` was added, the next request carried `Files workspace (initial): imports/: 1 file (imports/probe-auto09.csv)`, and the agent named the file. |
| AUTO-10 | pass | Pi's own write is not listed. |
| AUTO-11 | pass | Edit on a sheet added after start is listed. |
| AUTO-12 | pass | The edit note is kept in calls 2+ of the turn and in later turns. |
| AUTO-13 | pass | Two tabs: Chat 1 received the user's edits that Chat 2 had already seen, plus "By Pi in another chat tab: Calc H1, H2". |
| CONTRACT-01 | pass | Excel failures in mutation tools carry `isError` and a hint. Read failures are still `isError: false` (not in scope). |
| CONTRACT-02 | pass | Details are kept out of the text; results end with "Backup saved: `<id>`" (#794 WT-07). |
| CONTRACT-03 | pass | Blocked write lists the existing data, with the B-18 wording. |
| CONTRACT-04 | pass | `Sheet "Nope" not found. Sheets in this workbook: …`. |
| CONTRACT-05 | pass | Bad address: `…incorrect format.` |
| READ-01 | pass | Displayed text plus the stored value in brackets. |
| READ-02 | pass | `0`. |
| READ-03 | pass | Escaped. |
| READ-04 | pass | `### Number Formats`. |
| READ-05 | pass | CSV pages; the page header names the rows. |
| READ-06 | pass | `Merged: A1, B4`. |
| READ-07 | pass | `Hidden rows (filtered or hidden): 3, 5, 7`. |
| READ-08 | pass | `Spills: A1 → A1:A5`. |
| READ-09 | pass | Quoted and unquoted forms give the same result. |
| READ-10 | pass | Paged. |
| READ-11 | pass | Clipped to the data. |
| READ-12 | pass | 500-row pages instead of truncation. |
| READ-13 | pass | Cell-edit mode: waits, then reads (10.2 s). |
| OVER-01 | pass | 11 sheets. |
| OVER-02 | pass | Spill sheet: `Headers: 1` plus `Sheet state: Spills: A1 → A1:A5`. |
| OVER-03 | pass | `Table: "Sales" (6 rows × 3 cols)`. |
| OVER-04 | pass | Preview, as displayed. |
| OVER-05 | pass | The Big sheet works (274 ms). |
| SEARCH-01 | pass | 3 matches. |
| SEARCH-02 | pass | `Big!B200000`, 2.4 s. |
| TRACE-01 | pass | A1 and A2. |
| TRACE-02 | pass | Both references listed. |

### writing.md

| ID | Result | Note |
| --- | --- | --- |
| WRITE-01 | pass | |
| WRITE-02 | pass | Overwrite block, with the B-18 wording. On `Spill!A3` the spill block comes first. |
| WRITE-03 | pass | `=SUM(A1` blocked before Excel. |
| WRITE-04 | pass | `#SPILL!` reported with a hint. |
| WRITE-05 | pass | In Auto mode the agent sets `allow_overwrite` when asked (B-18). In Confirm mode it asks in chat and then the dialog asks too (see observations). |
| WRITE-06 | pass | |
| WRITE-07 | pass | Preview 1×6 of 1×8. |
| WRITE-08 | pass | Blocked. |
| WRITE-09 | pass | Blocked: C4, B5, C5 would be dropped. |
| WRITE-10 | pass | Blocked; the spill is intact. |
| WRITE-11 | pass | Protection named, `isError`. |
| WRITE-12 | pass | Violation reported. |
| WRITE-13 | pass | `'My Sheet'!B2`, named in the result. |
| WRITE-14 | pass | Success, partial verification. |
| WRITE-15 | pass | Conversions reported. |
| WRITE-16 | pass | Real Confirm-mode turn, declined: "The user declined this workbook change in the approval dialog (Confirm mode)." H3 is unchanged. |
| FILL-01 | pass | |
| FILL-02 | pass | Now reported: "Excel added the written cells to table Sales … Column1". |
| FILL-03 | pass | Single contiguous range error. |
| FILL-04 | pass | **a2b3f2d: fail** (Timeout; 4,500 of 50,000 formulas left; "nothing was written"). **e48d83c: pass** (0.39 s, sampled verification). |
| FILL-05 | pass | e48d83c: 200,000 rows in 0.75 s. Overwriting 200k occupied cells is refused. |
| FILL-06 | pass | Too-large reason. |
| FMT-01 | pass | |
| FMT-02 | pass | |
| FMT-03 | pass | Protection named. **a2b3f2d** stored a 0-change backup anyway; **e48d83c** stores none. |
| FMT-04 | pass | |
| STRUCT-01 | pass | Rerun on 8099cd6 (#791 @ aa17e28), both variants, each on fresh sheets with the fixture rows (hidden rows 3, 5, 7 when filtered). **Single hidden row, `delete_rows` at row 3:** with a table filter, an AutoFilter, or row 3 hidden by hand, Excel on the web deleted the hidden "South, b" row in all three cases. Pi now says so: "⚠️ 1 deleted row(s) were hidden (a filter is on): row 3. The rows were in table **FiltTbl** (filtered). Deleted hidden data: row 3 (\"South\", \"b\", 20). Deleting by position removes hidden rows too. If the user meant only the visible rows, restore this change with `workbook_history` …", with `details.hiddenRowsDeleted: 1` and a backup. The hand-hidden case says "were hidden" without "(a filter is on)". **Mixed range, rows 2–3 (visible a + hidden b):** with a table filter or an AutoFilter, Excel kept the filtered-out row and deleted only the visible one. Pi reports "Deleted 1 of 2 row(s) … ⚠️ Excel kept 1 hidden row(s) (a filter is on) and deleted only the others … The kept rows (originally row 3) now start at row 2", saves no backup and says why. Read-back matched (b kept, now hidden at row 2). Hand-hidden, the same range deleted both rows, and Pi listed the hidden one. Rows 4–6 of the filtered table (hidden d, visible e, hidden f) gave "Deleted 1 of 3 row(s)", with e deleted. So #791's finding holds only for a mixed range: a range made only of filtered-out rows is deleted. Pi reports both cases correctly. **Earlier note (e48d83c):** the "South, b" row deleted then was hidden by the table filter (Region = North, from the fixture), not by hand. Read-back on 8099cd6 confirms it is gone from table Sales. |
| STRUCT-02 | pass | |
| STRUCT-03 | pass | "the workbook was not changed. Sheet "Protected" is protected …" |
| STRUCT-04 | pass | |
| CHART-01 | pass | `Chart 'Chart 1' created`. Chart types are lower-case (`column`); `ColumnClustered` is rejected with a clear error. |
| CHART-02 | pass | `name is required`. |
| CHART-03 | pass | 231 ms, 600×360. |
| COMMENT-01 | pass | |
| COMMENT-02 | pass | Clear `range is required` error. |
| OJS-01 | pass | |
| OJS-02 | pass | |
| OJS-03 | pass | Fixed on 8099cd6 (#791 @ aa17e28, 32f5e1a). A write to `Protected!A1` returns `isError: true` with "Error executing Office.js: Protected sheet(s): \"Protected\". Excel refused the change, most likely because it touches one of them. Pi cannot unprotect a sheet …", plus "Steps in the code before the failing call may already have changed the workbook; read back before retrying." That warning is true: with an unlocked `B2` write synced first, B2 changed before A1 failed. Other Office.js errors are also `isError` with a specific reason (a size mismatch gives InvalidArgument; a missing sheet gives the sheet list). An error thrown by the code itself keeps its message. |
| OJS-04 | pass | Compact JSON. |
| HIST-01 | pass | Full ids. |
| HIST-02 | pass | `delete` asks even in Auto mode ("Let Pi delete backups?"). Declined in a real turn. |
| PY-01 | pass | |
| PY-02 | pass | stdout `emscripten`; success results do not name the backend. |
| PY-03 | pass | Exact error with the backend named. |
| PY-04 | pass | 504 after 2.0 s, no second run. |
| PY-05 | pass | Single traceback plus stdout. |
| PY-06 | pass | stdout and stderr in the result. |
| PY-07 | pass | Padding warning, `paddedRowCount`. |
| PY-08 | pass | Blocked. |
| PY-09 | pass | Input still raw (`45292`, `0.05`); the write reports conversions. |
| PY-10 | pass | Refused with a row estimate. |
| PY-11 | pass | Wording differs from B-18 (N-3). |
| PY-12 | pass | `Converted /tmp/lo/t.csv to XLSX`. |

### hostility.md

| ID | Result | Note |
| --- | --- | --- |
| EDIT-01 | pass | Waits instead of failing. |
| EDIT-02 | pass | After 58 s: "nothing was written … ask the user". |
| EDIT-03 | pass | Real turn: about 17 s wait between calls, then completed with three reads. |
| EDIT-04 | pass | Raw `onChanged.add` still fails in edit mode, as expected. The tracker registers at start. |
| EDIT-05 | pass | desktop (macOS), Excel for Mac, `8099cd6`: a cell held in edit mode by a real keystroke, then a turn asked for `read_range` on A1:A3. The dock showed "Waiting for you to finish editing a cell" / "Press Enter, Tab or Esc in Excel and Pi will continue" and the turn kept running; after Escape in Excel the read ran and the turn finished ("Sheet1!A1:A3 are all empty."). On `main` the read failed at once (hostility.md). |
| MOVE-01 | pass | User edit `Unknown`, Pi `ThisLocalAddin`, co-author `Remote`. |
| MOVE-02 | pass | Product call: the result names the sheet (B-17). |
| MOVE-03 | pass | Agent answered "You changed UserAdded2!F11 …" from the stored note. |
| LOCK-01 | pass | Two coordinated tool sets with different session ids: writes serialized (1.7 s and 3.4 s, each Excel call delayed 1.5 s). |
| LOCK-02 | pass | Read returned `new`. |
| LOCK-05 | pass | Queue: no auto-decline. |
| LOCK-06 | pass | Real Confirm turn: three dialogs one at a time, all applied. |
| SIZE-01 | pass | Rerun on 8099cd6 (#787 @ ba0e22c) in a fresh tab: "What's the total Qty on the Big sheet for rows 2 to 1000?" The agent read C1:C500 and C501:C1000. Each page now ends with "Column sums on this page (numbers only): C 4,464 (499 numbers)" (then 4,514), plus "compute them … do not add up values yourself". It answered **8,978**, which is correct (`functions.sum(Big!C2:C1000)` = 8,978). On a2b3f2d and e48d83c it answered 8,965. One run. |
| RECALC-01 | pass | Verified as 3, with a note. |
| RECALC-02 | pass | read_range warns "Calculation mode is Manual". |
| CTX-01 | pass | Paging instead. |
| CTX-02 | pass | `[Compacted tool result] read_range … run the tool again if you need it` (see B-16 for the moving boundary). |
| CTX-03 | pass | "Nothing to compact" below the 20k floor (17k session). |
| CTX-04 | pass | See B-24. |
| COAUTH-01 | pass | Remote edit labelled "By a co-author". |
| COAUTH-02 | pass | Pi's writes arrived as `ThisLocalAddin` here, even with the co-author present, and were left out. |
| COAUTH-03 | pass | Overwrite reported on the next request. |
| COAUTH-04 | pass | Row insert reported with "may have moved". |
| COAUTH-05 | pass | The co-author's edit mode did not block Pi's write or read. |
| SAVE-01 | pass | Backups listed after about 40 min on the web. |
| STOP-01 | pass | Stop before the batch: J1:J10 unchanged, turn ended 2.7 s after Stop. The result says "outcome unknown … may still be applied", which is conservative but not exact. |
| STOP-02 | pass | Stop after the batch: K11:K20 written, backup saved, turn ended 2.6 s after Stop. The result says "outcome unknown", not "the change stands" (N-4). **8099cd6:** see N-4; the result now says the change stands and gives the backup id. |
| STOP-03 | pass | Hung `Excel.run`: "Stopped by the user before Excel answered", 2.6 s. |
| TIME-01 | pass | Rerun on e48d83c as a small burst. Probe calls make no model requests, so the shared login and its rate limit were not involved and no separate login was needed. Results: 30 parallel `read_range` of 500×5 pages in 5.1 s; 20 parallel `write_cells` serialized through the lock in 5.1 s, all 20 landed; 100 parallel raw 5,000-cell reads in 4.6 s; no errors. A `read_range` straight after the burst took 4.0 s. Overload shows as delay, not as a throttling error, as the claim says. |
| TIME-02 | pass (simulated) | A real host hang can't be produced on purpose. The first pass got one only by firing 150 parallel 10,000-cell volatile-formula batches; that is not repeatable, needs a page reload to recover, and stalls the shared Excel on the web host and guest workbook for the other orbs. So `Excel.run` in the taskpane was replaced with a never-settling promise, which is the symptom of a hang as the code sees it. Rerun on e48d83c: `write_cells` returned `excel_timeout` with `isError` after 60.0 s. In a real turn, Stop ended a hung `read_range` 2.7 s after it was pressed ("Stopped by the user before Excel answered"). Once `Excel.run` was restored, the next read and the next turn worked, so the lock had been released. The same behavior was seen on a2b3f2d (B-20). |

### Prompt-cache prefix (`prefixChangeReasons`)

- **a2b3f2d:** 14 calls in 2 sessions, `prefixChangeReasons` `[]` on every call, `prefixChanges` 0. A Confirm/Auto mode switch arrived as an appended developer message ("Updated system prompt section …"), not a prefix change.
- **e48d83c:** 57 calls, `prefixChanges` 2, both `[systemPrompt, tools]`. Both come from `/compact`: the summarizer request (310-char instructions, no tools) is sent with the same `prompt_cache_key` as the chat, so the tracker counts it and the next call as changes. Every other call was `[]`.
- **Message-level prefix:** each request's input was the previous request's minus the trailing selection note in every consecutive pair, except two B-16 shaping rewrites and the compaction.
- **8099cd6:** 23 calls (one 20-read loop plus the turns around it), `prefixChanges` 0, `prefixChangeReasons` `[]` on every call. The message-level prefix broke only at the two step rewrites (B-16).

## Analyst tasks (GPT-6 Luna, real turns, fresh tab each)

1. **Build a small model** (5-year P&L on a new sheet, inputs plus formulas, formatted). Calls: get_workbook_overview, modify_structure, read_range, write_cells (16×6), 10 format_cells, read_range. Assumptions are in `Model!A4:B8`. Every projection cell is a formula referencing them. Values checked: 2029 revenue 1,632,587 = 1.2M×1.08⁴, 2029 EBIT 560,270. Currency and percent formats applied; it re-read and reported the changed ranges. **Verdict: succeeded.**
2. **Clean a messy import** (trim and proper-case names, mixed date formats in UK style, bracket negatives, drop blanks and duplicates, leave Import untouched). It used python_run on the bridge (one bridge-consent dialog, approved) to compute date serials. Clean!A1:D7 is correct: 05/02/2024 read as 5 Feb, `(250.00)` → -250, duplicate and blank removed. It kept `n/a` and `13/13/2024` as text and said exactly that ("Could not parse: Clean!B6 … Clean!C5 …"). Import is unchanged. **Verdict: succeeded, honestly.**
3. **Reconcile two sheets** (Ledger vs Bank by reference). It wrote dynamic-array formulas (`SORT(UNIQUE(VSTACK(…)))` and `XLOOKUP` spills) in Recon!A1:E2 and read the spill back. It reported all four exceptions correctly: INV-002 -90, INV-007 +9, INV-008 missing in bank, INV-009 missing in ledger. It noted the formulas cover source rows 2–1000. **Verdict: succeeded.**
4. **Chart a trend** (line chart plus a 3-month moving average, checked visually). It added C4:C25 `=AVERAGE(B2:B4)` filled, created the chart and checked it with get_image. It described the trend correctly. Unprompted, it flagged that the fixture's month dates repeat and skip months (my 30.44-day spacing), so the "3-month" average is really three rows. **Verdict: succeeded, honestly.** Low point: it labelled the column "3-Month Moving Average" and called it "3-observation" only in chat.
5. **Fix broken formulas** (#REF! discount, SUM ranges missing a row, #DIV/0!, a price stored as text). Fixed: E2:E5 (`=D2*(1-20%)`), B6 and D6 extended to row 5, D8 changed to `=D6/B6`; re-read. Missed: C7 `=AVERAGE(C2:C5)` = 4.17, because C3 holds the text "12" and AVERAGE skips it (true average 6.125). The detailed read showed `text (@) → C3`, but the grid renders the text "12" the same as a number. It also hard-coded 20% instead of putting it in an input cell, although the prompt's workflow rule says to use input cells. Its reply says "Fixed the formula problems", which overclaims. **Verdict on a2b3f2d: partial.** One misleading number was left in place and reported as fixed.
   **Rerun on 8099cd6** (#787 @ ba0e22c), on a freshly rebuilt Broken sheet: "The Broken sheet's formulas look wrong in places. Find and fix the problems, then tell me what you changed and anything you couldn't fix." Both read_range results carried "⚠️ Numbers stored as text (SUM, AVERAGE and other aggregates skip them …): C3 \"12\"". The agent wrote B6 `=SUM(B2:B5)`, D6 `=SUM(D2:D5)` and D8 `=D6/B6`. It rewrote C3 as the number 12 with General format, so C7 = 6.125, and it re-read the sheet. It searched for "discount" and "#REF!", found no source for the deleted rate, and left E2:E5 as #REF!, saying "the workbook doesn't indicate the intended discount rate … Those four errors remain until that rate is supplied." Checked in the workbook: all values are as stated. **Verdict on 8099cd6: succeeded, honestly.**

Also partial on a2b3f2d and e48d83c, from the SIZE-01 turns: a wrong Qty sum (8,965 vs 8,978) stated confidently. On 8099cd6 the same question was answered correctly (SIZE-01).

## Regressions and new bugs

### Fixed in e48d83c (found on a2b3f2d)

**R-1. fill_formula times out above about 4,500 cells and misreports what landed.** Suspected owner: #794 (R1C1 per-cell fill), plus the integration resolution.
- Repro on a2b3f2d: `fill_formula {range:"Big!G2:G50001", formula:"=C2*E2"}`. After about 39 s it returns `Error filling formula — nothing was written … Excel error Timeout`, but rows 2–4501 hold formulas, and there is no backup or audit entry.
- Raw timing for 4,000 cells: `formulasR1C1` grid 34 s vs `autoFill` 0.6 s.
- Code: `src/tools/fill-formula.ts:211-226` at a2b3f2d. The chunk loop ran before `fillInFlight`.
- Fixed by #794 @ 9f5558f (`copyFrom`, formulas only): 50k rows in 0.39 s, 200k in 0.75 s, formats not copied.

**R-2. A refused format_cells stored a backup.** Suspected owner: #794 with #791 (integration resolution).
- Repro: `format_cells {range:"Protected!A1", bold:true}`. The result says nothing was applied, but a 0-change backup tops `workbook_history list`.
- Code: `src/tools/format-cells.ts:618-631` at a2b3f2d.
- Fixed on e48d83c: no backup is stored.

### Found on a2b3f2d (all low severity), status on 8099cd6

**N-1. Manual-calculation note only when formulas are written.** Owner: #791.
- Repro: Manual calculation, B1 `=A1+A2`. Run `write_cells Calc!A1 = 10` with `allow_overwrite`. The result is clean, but B1 shows a stale value.
- Code: `src/tools/mutation/range-write.ts:546` gates on `options.sentFormulas`.
- **8099cd6: fixed** (#791 99d8242). Under Manual calculation, `write_cells Calc!A1 = 10` now ends with "ℹ️ Calculation is set to Manual. Formulas that depend on them still show old results until the user recalculates (F9) …", with `details.calculationMode: "Manual"`. A formula write says the written cells were calculated.

**N-2. Timeout result contradicts itself.** Owner: #791.
- Repro: `Excel.run` never settles; call `write_cells Calc!H7`. The result says "nothing was written to Calc!H7. Excel did not answer (…), so whether the change was applied is unknown."
- Code: `range-write.ts:900` plus `src/tools/mutation/excel-failure.ts:175`.
- **8099cd6: partly fixed** (#791 99d8242). Same repro: after 60.0 s, `isError` with `excel_timeout`, "Error writing cells — nothing was written to Calc!H7. Excel did not answer (Excel did not answer within 60 s. The outcome is unknown: the request may or may not have been applied. Read the affected range back before retrying a change. If this happens again, ask the user to reload the Excel page.) before the change was sent. Retry once Excel responds." The conclusion is now right (H7 was empty on read-back). But the quoted timeout message still says the outcome is unknown and may have been applied, because `src/tools/mutation/excel-failure.ts:188` interpolates the raw message from `src/excel/helpers.ts:97`. Sent to the #791 thread.
- **cea7588: fixed** (#791 e1c3f29). Same repro: "Error writing cells — nothing was written to TO!F1. Excel did not answer before the change was sent, so nothing was written. Retry once Excel responds; if it keeps happening, ask the user to reload the Excel page." No "outcome is unknown" quote. The held batch was released 10 s after the result, and the sheet was still empty 75 s after the call.

**N-3. Python wording leftovers.** Owner: #797.
- `python_run`'s description promises "numpy, pandas, etc. work automatically", but Pyodide returns "pandas … is not installed". Code: `src/tools/python-run.ts:512`; the system prompt's Python section says the same.
- The transform's overwrite block says "confirm with the user and retry", unlike B-18's rule. Code: `src/tools/python-transform-range.ts:645`.
- **8099cd6: fixed** (#797 fb4d645). With the bridge stopped, `python_run` ran on Pyodide (`sys.platform` = `emscripten`), and `import pandas as pd; import numpy as np` worked (pandas 2.2.3, 4.0 s cold). The transform's overwrite text now uses B-18's wording ("If the user asked to change these cells, retry with `allow_overwrite: true`. Otherwise show the user the existing data and ask …"); checked in code, not run.

**N-4. Stop reporting masked in the combination.** Owners: #792 with #791 (combination).
- Repro: delay `Excel.run`'s return by 6 s after the batch; Stop during `write_cells Calc!K11:K20`.
- The cells are written and a backup is saved, but the agent receives "outcome unknown … may still be applied" with no backup id. #791's "⏹ Stopped after writing … the change stands" never shows for a batch slower than 500 ms.
- Code: `src/tools/with-workbook-coordinator.ts:336` (`settleOnStop`, 500 ms).
- **8099cd6: fixed** (#792 @ d31bf82). In a real turn, every `Excel.run` was delayed 6 s and Stop was pressed 1.5 s after the `write_cells` card for Calc!K11:K30 appeared. The turn ended 6.9 s after Stop. The tool result the agent keeps: "Written to **Calc!K11:K30** (20×1) … Changed cell(s): 20. ⏹ Stop was pressed while this ran; the change above was already applied and stands. Backup saved: `6d668314-…`". The cells hold 1–20. A 6,000×5 probe write aborted 150 ms in returned "⏹ **Stopped after writing Calc!M1:Q6000** (6000×5) — the change stands. It was not verified because Stop was pressed …", with `stopped: true` and `writtenCount: 30000`; it saved no backup and said why (over the 20,000-cell limit; the range was empty, so clearing it undoes the change).

**N-5. Clearing a formula cell is reported as "replaced by values".** Owner: #791.
- Repro: write `""` over a formula. Seen in a real turn clearing `Co!A9`.
- Code: `range-write.ts:864`.
- **8099cd6: fixed** (#791 8b17deb). Writing `""` over `Calc!B2` (`=SUM(A1:A2)`) gives no "replaced by values" warning; `details.changes` keeps `beforeFormula`. Writing `5` over the same formula still warns ("1 formula(s) on sheet \"Calc\" replaced by values: B2").

**N-6. History shaping still moves its boundary on every call (B-16 remainder).** Owner: #792. Existed on main.
- Repro: four turns each with one or two ~13k-char reads in one tab.
- Calls 8 and 10 rewrite an earlier tool output into `[Compacted tool result]`. cacheRead drops from 38.5k to 9.34k, and `prefixChangeReasons` stays `[]`.
- Code: `src/messages/tool-result-shaping.ts` (`shapeToolResultsForLlm`).
- **8099cd6: fixed** (#792, 8cc8bf1). The boundary moves every 6 results (see B-16).

### Observations, not filed

- **Confirm mode asks twice.** The agent asks for confirmation in chat (rule 7 and the "Execution mode: Ask for explicit user confirmation" text), and then the dialog asks again. Seen in two Confirm-mode turns, including "no need to ask first".
- **The compaction summarizer shares the chat's `prompt_cache_key`.** Telemetry counts two `[systemPrompt, tools]` prefix changes per `/compact`.
- **Text numbers look like numbers in the compact read.** A number stored as text (`"12"`) renders the same as a number, which led to task 5's missed AVERAGE bug. Only detailed mode shows `text (@)`. On 8099cd6 every read lists them ("⚠️ Numbers stored as text …"), and task 5 then succeeded.
- **Failed `python_run` and read_range results keep `isError: false`.**
- **A row delete by position in a filtered table removes a hidden row without saying so** (STRUCT-01). The agent cannot tell from the result that it deleted a row the user could not see. On 8099cd6 the result lists the hidden rows deleted or kept (see STRUCT-01).
- **Structural changes made outside a Pi write window are attributed to the user.** Excel reports `RowInserted` and `RowDeleted` with `triggerSource: "Unknown"`, even from the add-in. Pi's own row insert was correctly left out of its own tab's note and listed under "By Pi in another chat tab" in the other tab. A probe delete made outside any tool call was listed "By the user". In the product this affects only add-in code with no write window, such as extensions.
- **Routed elsewhere:** OJS-03 to #791; SIZE-01 and task 5's AVERAGE miss to #787. All three pass on 8099cd6.

### Rechecked on 8099cd6, outside the rows above

- **#794 @ 515072d, approval dialog and live regions:** with an approval open, the four app children of `body` were `inert`, and `div.pi-announcer` (`aria-live=polite`) and the dialog's `p.pi-approval__waiting` were not inside an inert subtree. On abort the dialog resolved `false` and nothing stayed inert. Pass (DOM check).
- **#793 @ aff33d6, tabs:** with 13 tabs in the 349 px pane, titles read in full ("Chat 10", "Chat 11", "Chat 12"), the strip scrolls, and inactive tabs' × buttons have opacity 0 while the active tab's has opacity 1. After a reload the active tab was fully in view. Once, before a reload, the strip stopped 24 px short of its end (scrollLeft 531 of 555), so the active tab's × was clipped behind the › arrow. A new tab and a later reload both showed it in full, and it did not recur. Pass, with that one unreproduced glitch.
- **#789 @ a66868d, a partial answer kept when retries run out (SEND-13):** not exercised live. It needs a provider failure after streaming has started, which can't be caused on demand against the shared login.
- **Spot checks of touched tools:** `fill_formula Calc!R2:R6` wrote and verified 5 formulas with a backup. A refused `format_cells Protected!A1` returned `isError`, named the protection and stored no backup (R-2 still fixed).

### Rechecked on cea7588 (timeouts)

Method: `Excel.run` in the taskpane was wrapped so that one chosen `context.sync()` of one run is held for 70 s before it is sent, then sent for real to Excel on the web. That is longer than the 60 s `excelRun` timeout, so the tool reports first and the held batch lands afterwards, as a slow host would. The sheet was read 15–20 s after the held batch landed. Every case is a probe on a fresh sheet `TO`, deleted afterwards.

| Check | Result | Note |
| --- | --- | --- |
| N-2, write_cells timeout before anything is sent | pass | Sync 1 (the target read) held. Result: "nothing was written … Excel did not answer before the change was sent". Sheet empty at +75 s. |
| write_cells timeout mid-write | pass | 24,000×4 write (≈4.6 M chars, 3 chunks; 3.2 s unheld). Sync 3 (chunk 2) held. Result, `isError`: "rows 1–10538 … written and stands; rows 10539–20847 may or may not have been written; the remaining 3153 row(s) were not written", no backup (over the limit), "Read TO!A1:D20847 before retrying". Sheet at +80 s: exactly A1:D20847, last row r20847. Chunk 3 never reached Excel. |
| fill_formula timeout before the fill | pass | `TO!H1:H20000 = ROW()*3`, sync 2 (inspection) held: "nothing was written … before the change was sent". Sheet empty at +75 s; the fill was never sent. |
| fill_formula timeout on the fill | pass | Sync 3 (the paste) held: "⚠️ Outcome unknown — the fill … may or may not have been applied … Read TO!H1:H20000 before retrying." The held paste then landed (H1:H20000 filled), which the wording allows. Its quoted timeout text said "Excel has not answered 2 requests in a row", counting the previous probe's timeout because the cleanup in between bypassed `excelRun`; not a product issue. |
| python_transform_range timeout before the write | pass | Python output 24,000×4, write-run sync 1 held: "Error: Nothing was written to TO!J1:M24000. Excel did not answer before the change was sent …". Sheet empty at +75 s. |
| python_transform_range timeout mid-write (real web host, Python on Pyodide) | pass | Write-run sync 3 held: same partial wording as write_cells, `details.rowsWritten: 20847`, `verification: "failed"`. Sheet at +76 s: exactly J1:M20847. Before 47e8799 the abandoned chunks kept landing. |
| execute_office_js timeout | pass | Code writes P1, sync; P2, sync; P3, sync in try with a catch that writes P5 and syncs. Sync 1 held: `isError`, plus "Pi cannot stop Office.js code once it runs. It blocked this code's later Excel calls, so nothing after the unanswered step reached Excel, but that step may still be applied …". Sheet at +75 s: only P1 ("step1", the held step). P2, P3 and the catch-path P5 never reached Excel. Unheld, the same code wrote P1:P3. Code that calls `Excel.run` itself is refused before running ("Do not call Excel.run() in code"), so it cannot open a run outside `runMutation`. |

Not rechecked on cea7588: #792 c8593dc, #794 81a17f6 (restores) and the stopped-write card, which the integration thread smoke-tested.
