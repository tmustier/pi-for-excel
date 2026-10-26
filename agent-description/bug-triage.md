# Bug triage

The agent-experience problems the documents raised, deduplicated by root cause and ranked by how much they hurt task success: how often an agent doing ordinary spreadsheet work meets them, and how badly they mislead it when it does. Each entry names where the agent meets it, what it receives, how to reproduce it, the cause in the code, and a proposed fix. Twenty-three of the twenty-five were confirmed in Excel on the web on 2026-10-26 by calling the real tools in the taskpane (wrapped as the runtime wraps them) and, for twelve of them, by real model turns (GPT-6.1 Sol in the first pass, GPT-6 Luna in the second) whose requests were captured from the network; those carry a **Status** line. The second pass added a co-author in a second browser profile, a real Confirm-mode turn, Stop during a delayed write, `/compact`, an Excel burst that hung the host, and the Python tools against the repo's native bridge (`PYTHON_BRIDGE_MODE=real`) and Pyodide. The rest are read from code. None has been filed as an issue. None has been filed as an issue.

Ranking principle: a result that lies (says success when the workbook disagrees, or failure when the change was made) ranks above one that fails loudly, because the agent has no reason to look again. Among loud failures, the ones that every large or busy workbook hits rank above the rare ones.

## Summary

Twenty-five entries after merging about fifty raw observations. Ten are high. They fall into four clusters. **Results that lie** (B-01, B-02, B-03): writes into merged areas, oversized writes and fills, and writes that break a spill, land under manual calculation or violate data validation all produce a result that contradicts the workbook. **Size** (B-04, B-05): nothing bounds what the agent reads or is sent; one large selection costs ~150k tokens per model call, and one large sheet disables `search_workbook` for the whole workbook. **Excel state the agent cannot see or survive** (B-06, B-07, B-11): cell-edit mode ends the task, values arrive raw (dates as serials, zero as `0.0000e+0`), and protection, merges, filters, spills and calculation mode are invisible, so the agent leans on `execute_office_js` (six of eleven real turns). **Context the agent is given wrongly** (B-08, B-09, B-12): the change list reports Pi's own writes as the user's and disappears after one model call; parallel approvals decline each other; the blueprint goes stale. **A host that stops answering** (B-20): when Excel on the web hangs, the tool call, the lock and the turn hang with it, and Stop does not end the turn. The common thread of the high entries is that each tool trusts Excel's success and reports only on the cells it was asked about.

| ID | Title | Severity | Area | Decision needed |
| --- | --- | --- | --- | --- |
| B-01 | Writes into merged areas report success while values are dropped | high | writing | fix |
| B-02 | Large writes and fills report failure after the change is applied | high | writing, limits | fix |
| B-03 | Verification checks only the written cells, so broken spills, stale formulas and invalid entries pass | high | writing, verification | fix |
| B-04 | The selection context has no size limit, and disappears for whole-column selections | high | auto-context | fix |
| B-05 | Large sheets break reads, search and sheet overviews with a raw payload error | high | reading, limits | fix |
| B-06 | Cell-edit mode fails every call at once and ends the task | high | hostility | fix |
| B-07 | Values reach the agent raw and mis-rendered, and writes coerce text | high | reading, writing | fix |
| B-08 | The change list reports Pi's own writes, misses new sheets and lasts one model call | high | auto-context | fix |
| B-09 | Parallel Confirm-mode approvals decline each other | high | approvals | fix |
| B-20 | Excel calls have no timeout, and a hung host hangs the turn past Stop | high | hostility | fix |
| B-10 | Excel failures arrive as generic text in results marked successful | medium | tool contract | fix |
| B-11 | The agent cannot see workbook state that changes what its calls do | medium | reading, prompt | fix |
| B-12 | The blueprint goes stale and is re-sent on every model call | medium | auto-context | fix |
| B-13 | Dependency traces show a precedent range as its first cell | medium | inspecting | fix |
| B-14 | A read in the same batch as a write can see the old value | medium | parallel calls | fix |
| B-15 | Truncation drops the summaries that matter, and CSV has no row numbers | medium | truncation | fix |
| B-16 | History shaping hides earlier results behind a preview that points nowhere | medium | context | fix |
| B-17 | Unqualified addresses follow the user's active sheet | medium | addresses | product call |
| B-18 | Overwrite protection forces a confirmation round trip on explicit requests | medium | prompt | product call |
| B-19 | Backup notes give the wrong reason when a backup is skipped | low | backups | fix |
| B-22 | A missing Python module on the native bridge falls back to Pyodide and returns `Error: ` with no message | medium | python | fix |
| B-23 | A co-author's edits overwrite or move Pi's verified writes with no notice | medium | co-authoring | fix |
| B-24 | Compaction summarizes finished work as still to do, and frees little | medium | context | fix |
| B-21 | Small result-format slips | low | results | fix |
| B-25 | The Python transform hides stdout, pads ragged output and hits bridge size limits | low | python | fix |

## High

### B-01: Writes into merged areas report success while values are dropped

- **Where the agent meets it:** any `write_cells` (or `python_transform_range`, `fill_formula`) whose target overlaps a merged area other than at its top-left cell. Merged title rows and merged label blocks are common in finance workbooks.
- **What happens / what was expected:** Excel accepts the write and silently keeps only the top-left value of each merged area. The tool reads back the same cells, sees blanks, and still reports success. Writing `hello` to `Merged!B1` inside a merge `A1:D1` returned `Written to **Merged!B1** (1×1)` with an empty "Verified values" table and no change line; writing a 2×2 block to `B4` over a merged `B4:C6` returned `Written to **Merged!B4:C5** (2×2)` and `Changed cell(s): 1.` — three of four values were lost. Expected: the write is blocked or the result says which values did not land and why.
- **Reproduce:** merge `A1:D1`; call `write_cells {start_cell: "B1", values: [["hello"]]}`; read `A1:D1`.
- **Why (from the code):** `src/tools/write-cells.ts:193` assigns the whole grid and `:197-201` reads it back, but `formatSuccess` (`:410-496`) never compares the read-back with what was written; it reports errors (`#…` values) only. Merge areas are never loaded.
- **Severity:** `high`. Silent data loss with a success message; the agent tells the user the work is done.
- **Decision needed:** `fix`. After the read-back, compare each written value with the read-back value (allowing for type coercion, see B-07) and report mismatches as "not written" with addresses; before writing, load `getMergedAreasOrNullObject()` for the target and block or warn when the target cuts through a merge.
- **Raised by:** [tools/writing-cells.md](tools/writing-cells.md), [tools/filling-a-formula.md](tools/filling-a-formula.md), [tools/python.md](tools/python.md), [foundations/the-tool-contract.md](foundations/the-tool-contract.md).
- **Status:** confirmed 2026-10-26 on Excel on the web (probe calls; checklist WRITE-08, WRITE-09). In a real turn asked to put "Done" into `Merged!B1`, the agent avoided the trap only by spending two `execute_office_js` calls to discover the merge.

### B-02: Large writes and fills report failure after the change is applied

- **Where the agent meets it:** `fill_formula` over a long column, or `write_cells` with a large grid — the ordinary way to add a calculated column to a data sheet.
- **What happens / what was expected:** the write is sent and applied, then the read-back exceeds the response payload limit, and the tool returns the payload error. `fill_formula` over `Big!F2:F200001` returned `Error filling formula: The response payload size has exceeded the limit. …` and all 200,000 formulas were in the sheet. `write_cells` of 60,000 × 4 returned `Error writing cells: The response payload size has exceeded the limit. …` and row 60,000 was written. No backup and no audit entry were made, because those happen only on the success path. Expected: a successful result (with a bounded verification), or, if the write itself failed, a result that says nothing was changed.
- **Reproduce:** a sheet with 200,000 data rows; `fill_formula {range: "Big!F2:F200001", formula: "=C2*E2"}`; read `F200001`.
- **Why (from the code):** `src/tools/fill-formula.ts:122-126` queues the formula, the autofill and a load of `values,formulas` for the whole range into one batch; the response carries 400,000 cells and is refused after Excel has executed the fill. `src/tools/write-cells.ts:193-201` writes in one sync and reads back the whole range in the next; the second sync fails. The catch blocks (`fill-formula.ts:253-257`, `write-cells.ts:257-262`) report a generic error with no hint that the write may have landed.
- **Severity:** `high`. The agent believes the workbook is unchanged; it retries (now blocked by overwrite protection on its own data), gives up, or tells the user nothing happened, while an unbacked-up change sits in the workbook.
- **Decision needed:** `fix`. Separate the write from the verification: sync the write on its own, then verify with a bounded read (a sample, or `getSpecialCellsOrNullObject(Errors)` for errors) instead of loading the whole range. If a payload error happens after the write sync succeeded, say "written, but verification was skipped: the range is too large to read back". Cap the backup snapshot read the same way (it is already skipped above 20,000 cells).
- **Raised by:** [tools/writing-cells.md](tools/writing-cells.md), [tools/filling-a-formula.md](tools/filling-a-formula.md), [hostility/payload-limits-and-large-ranges.md](hostility/payload-limits-and-large-ranges.md).
- **Status:** confirmed 2026-10-26 on Excel on the web (FILL-05, WRITE-14).

### B-03: Verification checks only the written cells, so broken spills, stale formulas and invalid entries pass

- **Where the agent meets it:** every write whose effect reaches beyond the target, or whose verification cannot be trusted: overwriting a cell inside a spill range, writing formulas under manual calculation, writing into cells with data validation.
- **What happens / what was expected:** the result says "**Verified values:**" and lists only the written cells. Overwriting `Spill!A3` (a child of `=SEQUENCE(5)` in A1) returned a clean result with `x`; A1 became `#SPILL!`. Under manual calculation, `=A1+A2` was "verified" as `0.0000e+0` (the correct value was 3); the agent was not told calculation is manual. Writing `Maybe` into a cell restricted to Yes/No was accepted and verified. Expected: the result reports new errors caused in the anchor or dependents, says calculation is manual and the values are not current, and flags values that violate validation.
- **Reproduce:** spill: `=SEQUENCE(5)` in A1, `write_cells {start_cell: "A3", values: [["x"]], allow_overwrite: true}`, read A1. Manual: set `application.calculationMode = "Manual"`, write `=A1+A2`.
- **Why (from the code):** `src/tools/write-cells.ts:197-201` reads back only `rangeAddr`; `findErrors` (`src/utils/format.ts:79-100`) scans only those values. Neither the calculation mode nor `dataValidation.valid` is loaded anywhere in `src/tools/`. The system prompt's rule "**Verify writes.** write_cells auto-verifies and reports errors" (`src/prompt/system-prompt.ts`, WORKFLOW) tells the agent to trust it.
- **Severity:** `high`. A broken spill or a wrong total is reported as verified.
- **Decision needed:** `fix`. In the write batch also load `application.calculationMode` and, when it is not Automatic, either call `range.calculate()` before reading back or say "calculation is Manual; values shown are not recalculated". Before overwriting, detect spill membership (`getSpillParentOrNullObject()`) and block like overwrite protection does, naming the anchor. After writing, load `dataValidation.valid` for the target and report invalid cells. Optionally report new errors in the direct dependents' first level (`getDirectDependents`).
- **Raised by:** [tools/writing-cells.md](tools/writing-cells.md), [hostility/recalculation.md](hostility/recalculation.md), [foundations/the-tool-contract.md](foundations/the-tool-contract.md).
- **Status:** confirmed 2026-10-26 (WRITE-10, WRITE-12, RECALC-01). In a real turn under manual calculation the agent noticed the 0, spent two `execute_office_js` calls, found `calculationMode: "Manual"` and calculated the cell itself.

### B-04: The selection context has no size limit, and disappears for whole-column selections

- **Where the agent meets it:** before every model call, whenever the user has a large block selected (they often select the data they are asking about) or a whole column or row.
- **What happens / what was expected:** with `Big!A1:A5000` selected, the auto-context note was 311,761 characters, sent on every model call of the turn; the context meter went from 3% to 32% of a 500k window for "Reply with just OK.". A 20,000-row selection measures 873,849 characters, more than a 200k-token window. With a whole column (`B:B`, or `A:A` on a four-row sheet) or the whole sheet selected, the selection context is silently absent: the agent is not told what is selected at all. Expected: a bounded note (the address, the size, a sample of rows) whatever the selection.
- **Reproduce:** select `A1:A5000` on a 5-column, 200,000-row sheet; send any message; inspect the request (or the context meter).
- **Why (from the code):** `src/context/selection.ts:63-64` widens the window to the whole selection plus five rows, bounded only by the used range; `:35` loads `values,formulas` for the whole selection, which for a whole column exceeds the payload limit, and `:100` swallows the error. `src/taskpane/context-injection.ts:173-178` injects the text whatever its length. The injection is rebuilt for every model call (`transformContext`), so the cost repeats.
- **Severity:** `high`. It can make every request of a turn overflow, and when it fails it fails silently on the most common "look at this column" request.
- **Decision needed:** `fix`. Load the selection's address and size first; read values only for the selection's intersection with the used range, capped (for example 50 rows × 20 columns), and say "{n} more rows not shown; use read_range". For whole columns or rows, describe the selection by address and the used part. Include row numbers and column letters in the window (see B-21).
- **Raised by:** [foundations/auto-context.md](foundations/auto-context.md), [hostility/payload-limits-and-large-ranges.md](hostility/payload-limits-and-large-ranges.md), [hostility/context-pressure.md](hostility/context-pressure.md).
- **Status:** confirmed 2026-10-26 (AUTO-06, AUTO-07; real turn and direct measurement).

### B-05: Large sheets break reads, search and sheet overviews with a raw payload error

- **Where the agent meets it:** any workbook with one large sheet (tens of thousands of rows). Search works again only if the agent restricts it to a small sheet with `sheet`.
- **What happens / what was expected:** `read_range` of 50,000 × 5 and of any whole column fails with `Error reading "…": The response payload size has exceeded the limit. Please refer to the documentation: …`. `search_workbook` fails for every query in the whole workbook (`Error searching: The response payload size has exceeded the limit…`) because it loads every sheet's used range. `get_workbook_overview` with `sheet: "Big"` fails because it loads the whole used range to show five preview rows. The error does not say how much would fit or how to page. Expected: tools that page or bound their reads, and an error that tells the agent the range's size and a safe chunk size.
- **Reproduce:** a 200,000 × 5 sheet; call each of the three tools as above.
- **Why (from the code):** `src/tools/read-range.ts:82` loads values, formulas and number formats for the whole range in every mode, with no size check. `src/tools/search-workbook.ts:262-263` loads `values,formulas,address` of each sheet's used range. `src/tools/get-workbook-overview.ts:175-176` loads `values` of the whole used range for the preview.
- **Severity:** `high`. Large sheets are where agents are most useful; search, the obvious first step, stops working for the whole workbook.
- **Decision needed:** `fix`. `read_range`: load `rowCount/columnCount` first and refuse (or page) above a cell budget with "Range is 200,001 × 5; read at most ~10,000 rows per call, e.g. A1:E10000"; load number formats only in detailed mode. `search_workbook`: use `Worksheet.findAllOrNullObject` (ExcelApi 1.9) or scan the used range in row chunks. Overview: load only the first five rows of the used range.
- **Raised by:** [tools/reading-a-range.md](tools/reading-a-range.md), [tools/searching.md](tools/searching.md), [tools/the-workbook-overview.md](tools/the-workbook-overview.md), [hostility/payload-limits-and-large-ranges.md](hostility/payload-limits-and-large-ranges.md).
- **Status:** confirmed 2026-10-26 (READ-10, READ-11, SEARCH-02, OVER-05). In a real turn asked to count rows on the big sheet, the agent sidestepped the limit by writing its own aggregation in `execute_office_js`.

### B-06: Cell-edit mode fails every call at once and ends the task

- **Where the agent meets it:** whenever the user clicks into a cell and types (or leaves the formula bar open) while Pi is working. Users keep working while Pi runs.
- **What happens / what was expected:** every Excel batch fails in 3–12 ms with `Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.`, wrapped as "Error reading …" or "Error writing cells: …". The auto-context's selection note disappears silently for the same reason. In a real turn the agent read one range, failed the second, and ended the turn: "Reading Dates!A1:D4 failed because Excel is in cell-editing mode. Stopped; Calc!D1 was not changed. Press Enter or Tab to finish editing, then I can continue." Expected: the tool waits (with a bounded retry) for the user to leave edit mode, and the agent is told it waited.
- **Reproduce:** start a multi-step turn; click a cell in the grid and type without pressing Enter.
- **Why (from the code):** `src/excel/helpers.ts:27-29` (`excelRun`) passes `Excel.run` through with no retry. No tool checks for the `InvalidOperationInCellEditMode` error code.
- **Severity:** `high`. A normal user action aborts the agent's task partway, possibly after some writes and before others.
- **Decision needed:** `fix`. In `excelRun`, catch the cell-edit-mode error (code `InvalidOperationInCellEditMode`) and retry with backoff for up to, say, 30 seconds, reporting progress; if it still fails, return a result that says "the user is editing a cell; nothing was changed by this call". Have the auto-context say "(selection unavailable: Excel is in cell-edit mode)" instead of omitting it.
- **Raised by:** [hostility/cell-edit-mode.md](hostility/cell-edit-mode.md), [tools/reading-a-range.md](tools/reading-a-range.md), [tools/writing-cells.md](tools/writing-cells.md), [foundations/auto-context.md](foundations/auto-context.md).
- **Status:** confirmed 2026-10-26 on Excel on the web (EDIT-01 to EDIT-04; probe calls and a real turn). Desktop Excel is documented to queue rather than fail; not observed.

### B-07: Values reach the agent raw and mis-rendered, and writes coerce text

- **Where the agent meets it:** every read of dates, percentages, currencies, zeros, and every write of codes, fractions or ranges like "3-4".
- **What happens / what was expected:** reads: `01/01/2024` arrives as `45292`, `5.0%` as `0.05`, `£1,234.50` as `1234.5`, in every mode; formats appear only in `detailed` mode and only as format strings. Every zero is rendered `0.0000e+0` (and 1e-7 as `1.0000e-7`). A `|` or a line break inside a cell breaks the markdown grid. Writes: `"00123"` became 123, `"1/2"` became the date 2 January (46024), `"3-4"` became 3 April, `"50%"` became 0.5; the result shows the coerced numbers (`46024`) under "Verified values", which the agent rarely questions. Expected: reads that include displayed text for formatted cells; zero as `0`; escaped cells; writes that keep text as text when the agent sends a string, or a warning when Excel converted it.
- **Reproduce:** read a formatted block; write `[["00123","1/2","3-4"]]` and read it back in detailed mode.
- **Why (from the code):** `src/tools/read-range.ts:82` loads `values`, never `text`. `src/utils/format.ts` `formatAsMarkdownTable`: the number branch writes plain only when `Math.abs(v) < 1e15 && Math.abs(v) > 1e-6`, so 0 falls through to `toExponential(4)`; strings are joined without escaping `|` or newlines. `src/tools/write-cells.ts:193` assigns `values`, which Excel parses like typed input.
- **Severity:** `high`. Wrong dates and percentages in answers, corrupted IDs in writes, and a misleading grid on every read.
- **Decision needed:** `fix`. Load `text` alongside `values` and render the displayed text when it differs (for example `45292 (01/01/2024)`, or a "displayed" column in detailed mode); render 0 as `0`; escape `|` and replace newlines with `⏎` in the grid. For writes, compare read-back types with sent types and report "Excel converted '1/2' to a date in C1; prefix with ' to keep text"; or set the number format to `@` before writing strings that would be coerced.
- **Raised by:** [tools/reading-a-range.md](tools/reading-a-range.md), [tools/writing-cells.md](tools/writing-cells.md), [foundations/the-workbook-as-the-agent-sees-it.md](foundations/the-workbook-as-the-agent-sees-it.md).
- **Status:** confirmed 2026-10-26 (READ-01 to READ-05, WRITE-15). In a real turn asked for the dates "exactly as they appear", the agent recognised the serials and used `execute_office_js` to read `text`.

### B-08: The change list reports Pi's own writes, misses new sheets and lasts one model call

- **Where the agent meets it:** whenever the user asks "what did I change?" or edits cells between messages and expects Pi to notice.
- **What happens / what was expected:** four faults, three confirmed. (1) Pi's own writes are reported as `**User changes since last message:** Sheet1: H1` — Excel reports them with `triggerSource: "ThisLocalAddin"` and the tracker does not check. (2) Edits on sheets added after the taskpane opened are never reported (the user's edit to `Dates!G6` was absent). (3) The line appears in the first model call after the change and nowhere else: in a real turn it was present in request 1 of 6; when the agent answered in request 6 it said "I wasn't given an edit log". (4) There is one list per taskpane, so whichever tab sends next consumes the user's edits (read from code). Expected: only the user's edits, on every sheet, kept in the conversation once reported, per tab.
- **Reproduce:** have Pi write a cell, add a sheet, edit a cell on it, send "Which cells did I change since my last message?", capture the requests.
- **Why (from the code):** `src/context/change-tracker.ts:45-57` records every `onChanged` event without reading `event.triggerSource`; `:37-43` registers handlers only for sheets visible at `start()`, called once (`src/taskpane/init.ts:326`). `flush()` (`:69-104`) empties the list; `src/taskpane/context-injection.ts:180-181` flushes on every model call and the note is not stored (pi-agent-core `transformContext` output is discarded after the call). `init.ts:222` creates one tracker shared by every runtime (`:920`).
- **Severity:** `high`. The agent either attributes its own work to the user or cannot see the user's edits, and loses them mid-turn.
- **Decision needed:** `fix`. Skip events whose `triggerSource` is `ThisLocalAddin`; register `worksheets.onChanged` at workbook level (ExcelApi 1.9) or `onAdded` for new sheets; flush per tab and append the flushed line to that tab's stored user message (or keep it until the turn ends) so later model calls see it.
- **Raised by:** [foundations/auto-context.md](foundations/auto-context.md), [hostility/the-user-moves-or-edits.md](hostility/the-user-moves-or-edits.md).
- **Status:** confirmed 2026-10-26 for (1), (2), (3) (AUTO-10 to AUTO-12, real turns); (4) read from code. Second pass (COAUTH-01, COAUTH-02): a co-author's edit to `Calc!Z3` arrived with `source: "Remote"`, `triggerSource: "Unknown"`, and the tracker records it as the user's. Once the co-author had joined, Pi's own writes also arrived with `triggerSource: "Unknown"` (they were `ThisLocalAddin` before), and stayed that way after the co-author left. So filtering on `triggerSource` alone is not enough: also skip ranges Pi has just written (the coordinator knows them), and label `source: "Remote"` entries as a co-author's.

### B-09: Parallel Confirm-mode approvals decline each other

- **Where the agent meets it:** in Confirm mode, whenever one model call asks for two or more mutations at once (formatting several ranges, writing two blocks). Models batch independent calls by default (`parallel_tool_calls: true`).
- **What happens / what was expected:** each mutating call opens an approval dialog with the same overlay id; opening the second closes the first, which resolves as declined. The agent receives `Mutation cancelled by user (Confirm mode).` for every call but the last, though the user never declined anything. Expected: approvals queue, or one dialog lists all pending mutations.
- **Reproduce:** Confirm mode; ask for two independent formatting changes in one message.
- **Why (from the code):** `src/ui/confirm-dialog.ts:44-45` calls `closeOverlayById(overlayId)` with the shared default id before opening; closing settles the earlier promise as false. `src/tools/with-workbook-coordinator.ts:118-140` asks for approval before taking the lock, so all calls in a batch ask at once.
- **Severity:** `high` in Confirm mode. The agent is told the user refused, and by the system prompt's rules it should not retry or should ask again.
- **Decision needed:** `fix`. Serialize approvals (take the workbook lock before asking, or queue approval requests), or give each request its own overlay id and stack them.
- **Raised by:** [hostility/the-lock-and-parallel-calls.md](hostility/the-lock-and-parallel-calls.md), [foundations/the-tool-contract.md](foundations/the-tool-contract.md).
- **Status:** confirmed 2026-10-26 at the dialog level (LOCK-05) and in a real Confirm-mode turn (LOCK-06, GPT-6 Luna). Asked for "make A1:B1 bold, give C1 a yellow fill, and write the word checked in H2", the model issued all three calls in one response. One dialog appeared (for `write_cells`, the last), and it was approved. Both `format_cells` calls returned `Mutation cancelled by user (Confirm mode).` without the user seeing them. The agent ended: "The formatting changes to Co!A1:B1 and Co!C1 were cancelled, so they weren't applied." The dialog's title also says "Allow workbook mutation in Safe mode?" while its body says "Confirm mode".

### B-20: Excel calls have no timeout, and a hung host hangs the turn past Stop

- **Where the agent meets it:** whenever Excel on the web stops answering: after a heavy burst (its own or another add-in's), a lost connection, or a host stuck "Saving...". Rare, but when it happens the whole session is stuck until the user reloads the page.
- **What happens / what was expected:** 150 parallel batches each writing 10,000 volatile formulas (`=RAND()*SUM($A$1:$E$2000)`) were sent from the taskpane. Excel never answered any of them. Only 10,800 of 1.5 million formulas landed. From then on every `Excel.run` waited: a 2 × 2 `read_range` was still waiting after 50 s. A real turn ("Read Co!A1:B3 and tell me what is in it.") showed the `read_range` card spinning for 90 s; pressing Stop did not end the turn, and pressing it again 10 s later still did not. Excel's own UI showed only "Saving..." with no error. Reloading the page recovered Excel in 22 s; Pi restored both chats. In the next turn the agent said the read had returned "No result provided". Expected: the call gives up after a bounded time, releases the lock, and tells the agent "Excel did not answer in 60 s; the change may or may not have been applied, read back before retrying"; Stop ends the turn at once.
- **Reproduce:** fire ~150 large formula writes in parallel through `Excel.run` from the taskpane console, then send a turn that reads a small range; press Stop.
- **Why (from the code):** `src/excel/helpers.ts:27-29` (`excelRun`) passes `Excel.run` through with no timeout and no abort. The coordinator checks the signal only before running (`with-workbook-coordinator.ts:63-68`); the tool's promise is awaited by the agent loop, so the turn cannot end until it settles.
- **Severity:** `high`. Rare, but total: the agent, the lock and the user's Stop are all held hostage, and nothing tells anyone why.
- **Decision needed:** `fix`. Race every `Excel.run` against a timeout (say 60 s for reads, longer for mutations) and the abort signal; on either, release the lock and return an error-shaped result that says the outcome is unknown. Make Stop settle pending tool calls with "cancelled by Stop; the Excel batch may still complete".
- **Raised by:** [hostility/throttling-and-timeouts.md](hostility/throttling-and-timeouts.md), [hostility/stop.md](hostility/stop.md).
- **Status:** confirmed 2026-10-26 on Excel on the web (TIME-01, TIME-02, STOP-03). Lighter bursts did not throttle: 60 parallel `read_range` calls finished in 8.7 s, 40 parallel `write_cells` serialized through the lock in 34 s, 300 raw 5,000-cell reads in 45 s and 100 raw 10,000-cell writes in 96 s, all with no error. A `read_range` issued straight after the 100-write burst waited 15 s. No throttling error was ever returned; overload shows up as delay, then as a hang.

## Medium

### B-10: Excel failures arrive as generic text in results marked successful

- **Where the agent meets it:** every Excel-side failure of every workbook tool.
- **What happens / what was expected:** failures are returned as ordinary results (`isError` false) whose text is Office's message. For a protected sheet that message is `You cannot perform the requested operation.` — for `write_cells`, `format_cells`, `modify_structure` and `execute_office_js` alike — with no mention of protection. The Office error code (`AccessDenied`, `InvalidOperationInCellEditMode`, `ItemNotFound`) and debug info are dropped. In a real turn the agent needed two extra `execute_office_js` calls to discover the protection. The compaction summarizer labels tool results "(error)" only by the flag, so failures are summarized as successes. Expected: thrown or flagged errors with the code and a hint ("Sheet 'Protected' is protected; cells A2 are locked").
- **Why (from the code):** every tool's catch builds a text result (`write-cells.ts:257-262`, `read-range.ts:158-162`, `format-cells.ts:620`, `modify-structure.ts:507`, …); `getErrorMessage` (`src/utils/errors.ts:7-23`) keeps only `message`. No tool sets `isError`.
- **Severity:** `medium`. Recoverable, but every failure costs round trips and some are misreported in summaries.
- **Decision needed:** `fix`. A shared `describeExcelError(error, {range})` that maps codes to agent-actionable hints (protection: load `worksheet.protection.protected`; edit mode; payload: include sizes) and returns `isError: true`.
- **Raised by:** [foundations/the-tool-contract.md](foundations/the-tool-contract.md), [tools/writing-cells.md](tools/writing-cells.md), [tools/formatting.md](tools/formatting.md), [tools/structure-and-view.md](tools/structure-and-view.md), [tools/office-js.md](tools/office-js.md).
- **Status:** confirmed 2026-10-26 (WRITE-11, FMT-03, STRUCT-03, OJS-03; real turn).

### B-11: The agent cannot see workbook state that changes what its calls do

- **Where the agent meets it:** planning any task on a real workbook.
- **What happens / what was expected:** no structured tool or auto-context reports sheet protection, merged areas, filters and hidden rows, data validation, spill ranges, the calculation mode, where the used range starts, or which rows are table headers. The overview's "Headers" is whatever row 1 holds (`#SPILL!, #SPILL!`; `Locked, 1`; `100, pi wrote this, row 0, 0, …`). The system prompt names none of these hazards. The agent compensates with `execute_office_js` (six of eleven real turns) — which in Confirm mode costs an approval each time.
- **Why (from the code):** `src/tools/get-workbook-overview.ts:83-88` loads row 1 and the used-range size only; `read-range.ts` loads no state; `src/prompt/system-prompt.ts` (WORKFLOW, CONVENTIONS) has no hazard guidance.
- **Severity:** `medium`. Each gap alone is recoverable; together they make the escape hatch the agent's main tool.
- **Decision needed:** `fix`. Add a compact "sheet state" line to the overview and to `get_workbook_overview {sheet}`: used range address, protection, calc mode (workbook), freeze panes, filters/hidden rows count, merged areas count, tables with header row and range, spill anchors. Mark merged, hidden and spill cells in `read_range` detailed mode. Add three sentences to the system prompt: values are raw; check merges/protection when a write reports fewer changes than expected; never read whole columns.
- **Raised by:** [foundations/the-workbook-as-the-agent-sees-it.md](foundations/the-workbook-as-the-agent-sees-it.md), [foundations/what-the-agent-is-told.md](foundations/what-the-agent-is-told.md), [tools/the-workbook-overview.md](tools/the-workbook-overview.md), [tools/reading-a-range.md](tools/reading-a-range.md), [tools/office-js.md](tools/office-js.md).
- **Status:** confirmed 2026-10-26 (OVER-02, READ-06 to READ-08; real turns).

### B-12: The blueprint goes stale and is re-sent on every model call

- **Where the agent meets it:** after the user (or Pi's `write_cells`) adds sheets, renames them or grows a sheet; and on every model call.
- **What happens / what was expected:** the blueprint is cached per workbook and rebuilt only after Pi's structural calls; a sheet added by the user (or by anything else) is missing from it until then (confirmed: `UserAdded` absent after `worksheets.add`). Separately, because the injected note is never stored, the "already in history" check always fails, and the full blueprint is re-sent on every model call labelled `[Workbook context refresh: context missing from history]`. Expected: a blueprint refreshed on workbook change events and sent once per change.
- **Why (from the code):** `src/context/blueprint.ts:50-66` caches; `src/taskpane/init.ts:860-863` invalidates only on `impact === "structure"`. `src/taskpane/context-injection.ts:53-68` looks for the note in stored messages, where it never is; `context-refresh-decision.ts` then returns `context_missing` every time.
- **Severity:** `medium`. Stale sheet lists mislead planning; the re-send costs tokens on every call and moves within the prompt.
- **Decision needed:** `fix`. Invalidate on `worksheets.onAdded/onDeleted/onNameChanged` and on used-range growth; either store the injected note in the session when it changes, or drop the "missing" check and label the note simply `[Workbook overview]`.
- **Raised by:** [foundations/auto-context.md](foundations/auto-context.md), [tools/the-workbook-overview.md](tools/the-workbook-overview.md).
- **Status:** confirmed 2026-10-26 (AUTO-03, AUTO-04).

### B-13: Dependency traces show a precedent range as its first cell

- **Where the agent meets it:** `trace_dependencies` on any formula whose precedents are adjacent cells (every `SUM(A1:A100)`, every `=A1+A2`).
- **What happens / what was expected:** Excel returns precedents as coalesced ranges (`Calc!A1:A2`); the tool turns each into its first cell, so `=A1+A2` traced as depending on A1 only. Expected: the range, or each cell.
- **Why (from the code):** `src/tools/trace-dependencies-logic.ts:52-56` (`normalizeTraversalAddress`) parses a range address as a single cell; `src/tools/trace-dependencies.ts:275-283` uses it for every API precedent.
- **Severity:** `medium`. Wrong lineage in an auditing tool, reported confidently.
- **Decision needed:** `fix`. Keep range nodes as ranges (with a cell count) or expand small ranges into cells.
- **Raised by:** [tools/tracing-and-explaining.md](tools/tracing-and-explaining.md).
- **Status:** confirmed 2026-10-26 (TRACE-01).

### B-14: A read in the same batch as a write can see the old value

- **Where the agent meets it:** when the model issues a write and a verifying read together.
- **What happens / what was expected:** reads never wait for the workbook lock, so `read_range Calc!Z1` issued with `write_cells Calc!Z1 = "new"` returned `old`. Expected: reads of a range with a pending or running Pi write wait for it, or the agent is told calls in one batch are unordered.
- **Why (from the code):** `src/tools/with-workbook-coordinator.ts:164-169` runs reads immediately; `src/workbook/coordinator.ts` has no read barrier.
- **Severity:** `medium`.
- **Decision needed:** `fix`. Make reads wait for writes already queued in the same workbook (a read barrier), or document it in the prompt.
- **Raised by:** [hostility/the-lock-and-parallel-calls.md](hostility/the-lock-and-parallel-calls.md), [tools/reading-a-range.md](tools/reading-a-range.md).
- **Status:** confirmed 2026-10-26 (LOCK-02).

### B-15: Truncation drops the summaries that matter, and CSV has no row numbers

- **Where the agent meets it:** any read over 50 KB.
- **What happens / what was expected:** head truncation keeps the grid and cuts the `**Formulas:**` and `⚠️ **Errors:**` lines that come after it, so a truncated read hides that errors exist. CSV output has no row numbers, so the agent must count lines to find where the cut fell. The notice says lines, not cells. Expected: summaries first (or kept), and a notice that names the last row returned and the next range to read.
- **Why (from the code):** `src/tools/read-range.ts` (`formatCompact`) appends summaries after the grid; `src/tools/output-truncation.ts:95-173` keeps the head; `buildTruncationNotice` (`:247-269`) knows nothing about ranges.
- **Severity:** `medium`.
- **Decision needed:** `fix`. Put the error and formula summaries (capped) before the grid; have `read_range` truncate itself by rows and say "Rows 1–1,117 of 10,000 shown; continue with A1118:E10000".
- **Raised by:** [tools/reading-a-range.md](tools/reading-a-range.md), [hostility/context-pressure.md](hostility/context-pressure.md).
- **Status:** confirmed 2026-10-26 (READ-12).

### B-16: History shaping hides earlier results behind a preview that points nowhere

- **Where the agent meets it:** in any turn with more than six tool calls, and in every later turn.
- **What happens / what was expected:** every tool result older than the six most recent and longer than 1,200 characters is replaced, for the model, by its first 500 characters under `[Compacted tool result]`, ending "Full output remains visible in chat history; this compact version is model-facing only." The model cannot see chat history; nothing tells it how to get the data back (re-read which range). The replacement changes which result is shaped on every call, so the prompt prefix changes every call. Expected: a preview that names the call's arguments ("read_range Big!A1:E10000 — re-run to see it again") and a shaping boundary that moves in steps.
- **Why (from the code):** `src/messages/tool-result-shaping.ts:133-163` builds the preview; `:175-205` applies it by recency; `src/messages/convert-to-llm.ts:87-89` runs it on every model call.
- **Severity:** `medium`. Agents re-read or, worse, answer from a 500-character fragment.
- **Decision needed:** `fix`. Include the tool arguments in the preview, drop the chat-history sentence, and shape in blocks (for example only when 6 more results have accumulated).
- **Raised by:** [hostility/context-pressure.md](hostility/context-pressure.md), [foundations/the-tool-contract.md](foundations/the-tool-contract.md).
- **Status:** read from code.

### B-17: Unqualified addresses follow the user's active sheet

- **Where the agent meets it:** whenever the agent omits the sheet (the schemas allow it: "If no sheet is specified, uses the active sheet") and the user changes sheets during the turn.
- **What happens / what was expected:** `write_cells {start_cell: "B2"}` wrote to `'My Sheet'!B2` because the user had just activated that sheet. The result names the sheet, so a careful agent can notice. Expected: a product call between keeping the convenience and requiring qualified addresses for mutations.
- **Why (from the code):** `src/excel/helpers.ts:50-56` (`getRange`) falls back to `getActiveWorksheet()` at batch time.
- **Severity:** `medium`.
- **Decision needed:** `product call`. Option A: require a sheet for mutating tools (reject unqualified addresses with "qualify the address, e.g. Sheet1!B2"). Option B: resolve the active sheet once per turn and use it consistently. Option C: keep, and add the prompt rule "always qualify addresses".
- **Raised by:** [hostility/the-user-moves-or-edits.md](hostility/the-user-moves-or-edits.md), [foundations/the-workbook-as-the-agent-sees-it.md](foundations/the-workbook-as-the-agent-sees-it.md).
- **Status:** confirmed 2026-10-26 (MOVE-02).

### B-18: Overwrite protection forces a confirmation round trip on explicit requests

- **Where the agent meets it:** "change A2 to say Output" when A2 has a value.
- **What happens / what was expected:** the prompt says "write_cells blocks if the target has data. Ask the user before setting allow_overwrite=true." In a real turn the agent read A2 and asked "Confirm replacing it with Output?" before doing exactly what the user had asked. Expected: a product call on whether an explicit instruction counts as confirmation.
- **Why (from the code):** `src/prompt/system-prompt.ts` WORKFLOW rule 4; `src/tools/write-cells.ts` schema description of `allow_overwrite`.
- **Severity:** `medium`. A wasted turn on many edit requests; harmless otherwise.
- **Decision needed:** `product call`. Reword rule 4: "If the user explicitly asked to change those cells, set allow_overwrite=true; otherwise ask."
- **Raised by:** [tools/writing-cells.md](tools/writing-cells.md), [foundations/what-the-agent-is-told.md](foundations/what-the-agent-is-told.md).
- **Status:** observed 2026-10-26 in a real turn (one model, one prompt).

### B-22: A missing Python module on the native bridge falls back to Pyodide and returns `Error: ` with no message

- **Where the agent meets it:** `python_run` with the native bridge running, whenever the code imports a package the local Python lacks (pandas and numpy are the usual ones), or the bridge times out.
- **What happens / what was expected:** with the bridge in real mode and no pandas installed, `python_run {code: "import pandas as pd\nresult = 1"}` returned exactly `Error: ` and details `{ok: false, error: ""}`. The bridge had answered 400 with `ModuleNotFoundError: No module named 'pandas'`, but the tool then retried in Pyodide, where pandas was also not loaded; Pyodide's traceback went to `stderr`, which the error path drops. Any unknown module (`import notamodule_xyz`) gives the same empty error. A bridge timeout also re-runs the code in Pyodide: `while True: pass` with `timeout_ms: 2000` ran on the bridge, then in the browser, and returned `Pyodide execution timed out after 2000ms` after 4 s. Other bridge errors are clear: `1/0` returned the full traceback. Expected: the bridge's own error, and no silent second run.
- **Reproduce:** `PORT=3340 PYTHON_BRIDGE_MODE=real node scripts/python-bridge-server.mjs --https` on a machine without pandas; call `python_run` with `import pandas`.
- **Why (from the code):** `shouldFallbackToPyodideAfterBridgeError` (`src/tools/python-run.ts:415-436`) matches `message.includes("enotfound")`, which `modulenotfounderror` contains, and `"timed out"`. The Pyodide branch throws `response.error ?? "Pyodide execution failed."` (`:543-545`); Pyodide sets `error: ""`, so the message is empty and `stderr` is lost. The tool description promises "numpy, pandas, etc. work automatically" in Pyodide, but a direct Pyodide call returned "The module 'pandas' is included in the Pyodide distribution, but it is not installed."
- **Severity:** `medium`. The agent cannot tell what failed and usually abandons Python.
- **Decision needed:** `fix`. Match network errors by error code or a word-boundary pattern, not substrings; never fall back on a bridge timeout; on a Pyodide failure, return `stderr` (and say which backend ran).
- **Raised by:** [tools/python.md](tools/python.md).
- **Status:** confirmed 2026-10-26 (PY-03, PY-04, PY-05).

### B-23: A co-author's edits overwrite or move Pi's verified writes with no notice

- **Where the agent meets it:** any shared workbook on the web (or desktop with AutoSave on OneDrive) where someone else is editing while Pi works.
- **What happens / what was expected:** three observations with a second guest editing the same workbook in another browser profile. (1) The co-author started typing in `Co!B4` (without Enter). Pi wrote `555` there and got `Written to **Co!B4** … Verified values: 555`. The co-author then pressed Enter, and `B4` became `9`: last writer wins, and the verified result is now false. (2) The co-author inserted a row above row 2 while Pi wrote `[["Total", "=SUM(B2:B5)"]]` at `Co!A6`. The result said `Written to **Co!A6:B6**`, but the cells ended in row 7 as `=SUM(B3:B6)`, and every address the agent read earlier now points one row off. (3) A co-author editing a cell does not put Pi into cell-edit mode: Pi's reads and writes succeeded while the co-author was mid-edit. Remote values reached Pi within about 4 s; a sheet Pi added took over 20 s to appear for the co-author. Expected: the agent learns, at least on its next model call, that a co-author changed cells or structure it is working with.
- **Reproduce:** open the same folder-link workbook in a second profile (`PI_EXCEL_WEB_PROFILE=/tmp/coauthor-profile PI_EXCEL_WEB_CDP_PORT=9224 scripts/orb/excel-web.sh "<workbook url>"`); edit while Pi writes.
- **Why (from the code):** Remote changes arrive through `onChanged` with `source: "Remote"`, but the change tracker labels them as the user's (B-08) and keeps them for one model call; nothing reports a remote structural change (the blueprint is invalidated only by Pi's own writes, B-12). Verification reads back once, immediately.
- **Severity:** `medium`. Co-authoring is common on the web, and the agent's model of the sheet silently drifts.
- **Decision needed:** `fix`. Track `source: "Remote"` changes and row/column insertions (`onChanged` change types, `worksheets.onRowInserted` where available) per tab, report them as "A co-author changed …", and invalidate the blueprint on remote structural change. Optionally mention in write results when the target overlaps a remote change seen during the call.
- **Raised by:** [hostility/co-authoring-and-autosave.md](hostility/co-authoring-and-autosave.md).
- **Status:** confirmed 2026-10-26 (COAUTH-01 to COAUTH-05).

### B-24: Compaction summarizes finished work as still to do, and frees little

- **Where the agent meets it:** after `/compact` or auto-compaction (near 85% of a 272k window, 425k of a 500k one).
- **What happens / what was expected:** on a 63k-token session, `/compact` summarized 21 messages. The summary's Goal and "In Progress" were the request before last ("Read `Calc!A1:D2500` … then `Calc!A2501:D5000` …"), which had already been done and answered, and its Next Steps told the agent to call `read_range` on both ranges again. The agent did not redo them on "continue", but only because the next kept message showed a later answer. The context meter went from 62,934 to 58,752 tokens: the kept tail must hold at least 20,000 tokens and cannot start with a tool result, so a single 22k-token read result stayed verbatim. Below 20,000 tokens, `/compact` reports "Nothing to compact" (a 14k-token session). Expected: completed requests summarized as done with their answer, and compaction that frees most of the window.
- **Reproduce:** a session with a few large reads; `/compact`; inspect `agent.state.messages` for the `compactionSummary`.
- **Why (from the code):** `src/commands/builtins/export.ts` cuts at `findCutIndex` (`:302-323`) with `keepRecentTokens` = 20,000 (`src/compaction/defaults.ts`), walking back past tool results; the summarization prompt is given the summarized slice only, so the last request in that slice looks unfinished even when its answer is in the kept tail.
- **Severity:** `medium`. A summary that says "next, do X" invites repeating mutations.
- **Decision needed:** `fix`. Tell the summarizer which requests were answered (pass the first kept user message, or mark each summarized request's final assistant reply); shrink oversized tool results in the kept tail to their shaped preview before measuring the tail.
- **Raised by:** [hostility/context-pressure.md](hostility/context-pressure.md).
- **Status:** confirmed 2026-10-26 (CTX-03, CTX-04).

## Low

### B-19: Backup notes give the wrong reason when a backup is skipped

- **What happens:** a fill of 50,000 cells returned "ℹ️ Backup not created for this mutation. `workbook_history` tracks … when backup capture succeeds." with the reason (in details) "Backup capture was skipped (workbook identity unavailable or snapshot data unsupported)". The real reason was the 20,000-cell cap. The agent cannot tell the user why or what to do (`/backup` first).
- **Why:** `src/tools/recovery-metadata.ts:11-15` uses one reason for every skip.
- **Decision needed:** `fix`. Distinguish "too large to back up (N cells > 20,000)" from "workbook identity unknown".
- **Raised by:** [tools/backups.md](tools/backups.md), [tools/filling-a-formula.md](tools/filling-a-formula.md).
- **Status:** confirmed 2026-10-26 (FILL-06).

### B-21: Small result-format slips

- **What happens:** grouped.
    - The selection context's grid has no row numbers or column letters and treats its first row as a markdown header; the selected cell can lie outside the window shown (confirmed: selection `Sheet1!C3`, context `Sheet1!B2`).
    - `findErrors` treats any string starting with `#` as an Excel error, so text like `#1 priority` is reported as an error (read from code, `src/utils/format.ts:91`).
    - `execute_office_js` pretty-prints its JSON with two-space indentation before cutting at 8,000 characters, so a result holds about 130 rows of a 5-column range (confirmed).
    - Filling a formula into the column right of a table extends the table and names the new column `Column1`; the result does not say so (confirmed).
    - Overview sheet previews label rows 1–5 and columns from A even when the used range starts elsewhere (read from code, `get-workbook-overview.ts:278-298`).
    - Office.js errors lose their `code` and `debugInfo` (see B-10).
    - Passing a wrong optional parameter name (`cell` for `comments`) gives an internal JavaScript error, `Cannot read properties of undefined (reading 'includes')`, instead of "range is required" (confirmed).
    - Stop after a write had landed shows `Error: signal is aborted without reason` under the successful result (confirmed, STOP-02).
    - The Confirm-mode approval dialog is titled "Allow workbook mutation in Safe mode?" while its body and the mode button say "Confirm" (confirmed, LOCK-06).
- **Decision needed:** `fix` each.
- **Raised by:** [foundations/auto-context.md](foundations/auto-context.md), [tools/office-js.md](tools/office-js.md), [tools/filling-a-formula.md](tools/filling-a-formula.md), [tools/the-workbook-overview.md](tools/the-workbook-overview.md).

### B-25: The Python transform hides stdout, pads ragged output and hits bridge size limits

- **What happens:** grouped, all confirmed 2026-10-26 on the native bridge (PY-06 to PY-10).
    - `print()` output and stderr from the transform code are not in the result; a transform that printed diagnostics returned only `✅ Transformed …`.
    - A ragged result (`[[1,2,3],[4]]`) is written as 2 × 3 with blanks and reported as success.
    - `input_data` is `{range, values}`, and values are raw: an echo of formatted dates, percentages and text codes wrote `45292`, `0.05` and `123` in General format (the text `00123` became a number; B-07).
    - Writing into a merged destination drops values with a success result (`Changed cell(s): 3` for 4 values; B-01).
    - Reading `Big!A1:E20000` failed with `Python bridge request failed (413): Request body too large (max 524288 bytes).`; returning 50,000 rows failed with `Process output exceeded 262144 bytes.`. Neither says how many rows would fit.
- **Why:** `src/tools/python-transform-range.ts` formats only the write; the bridge's `MAX_JSON_BODY_BYTES` and `MAX_OUTPUT_BYTES` (`scripts/python-bridge-server.mjs`).
- **Decision needed:** `fix`. Include stdout/stderr previews as `python_run` does; reject ragged output; state the byte limits and a row estimate in the error.
- **Raised by:** [tools/python.md](tools/python.md).

## Items raised elsewhere, checked here

The product-description drafts raised these; this pass checked what it could.

- **Workbook identity is null for anonymous-link workbooks on the web** — consistent: the blueprint cache key here was the "unknown" key (`getWorkbookContext().workbookId === null`). Backups were nevertheless created and listed (`workbook_history list` showed ten entries from 17:55–18:03 at 18:10), so "automatic backups may be cleared within seconds" was not reproduced in this session.
- **Pi never retries failed provider requests** — out of this project's scope (provider behavior); the `openai-codex` path used here has its own retry loop (`node_modules/@earendil-works/pi-ai/dist/api/openai-codex-responses.js:259-289`).
- **A length-stopped response fails all its tool calls with "was not executed"** — confirmed in the agent loop (`agent-loop.js`, `failToolCallsFromTruncatedMessage`); the agent receives `Tool call "…" was not executed: the response hit the output token limit, so its arguments may be truncated. Re-issue the tool call with complete arguments.` as an error. Large `write_cells` grids are the likely trigger; see [tools/writing-cells.md](tools/writing-cells.md).
- **`workbook_history` delete and clear are classified as reads** — confirmed in `src/tools/execution-policy.ts:71-74`; they skip the lock and Confirm-mode approval. See [tools/backups.md](tools/backups.md).
