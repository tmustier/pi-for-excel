# Tool Behavior Decisions (Pi for Excel)

Concise record of recent tool behavior choices to avoid regressions. Update this as we tweak tooling.

## WPS registry seam and Phase 2 overrides (NEXSELL-370)
- **Tool list:** keep `CORE_TOOL_NAMES` as the single ordered source of truth and keep core tool schemas/metadata stable across hosts.
- **WPS Phase 2 override mechanism:** `host-selection.ts` owns `WPS_CORE_TOOL_EXECUTE_OVERRIDES`, a narrow execute-only map for WPS-backed core tools. Supported WPS tools are composed as `{ ...officeTool, execute: wpsExecute }`, so `name`/`label`/`description`/`parameters` remain byte-for-byte identical to the Office tool object for prompt-cache stability.
- **Supported WPS core slice:** `get_workbook_overview`, `read_range`, and `write_cells` run against the synchronous WPS ET JSAPI. Other workbook tools keep the typed fail-fast wrapper.
- **Local-only core tools:** `instructions`, `conventions`, and `skills` remain available because they operate on settings/skills storage. Workbook-scoped instructions use the WPS workbook identity when `ActiveWorkbook.FullName` is available.
- **No backups on WPS:** `write_cells` preserves overwrite protection and read-back verification, but WPS automatic workbook backups/snapshots are not implemented. The WPS write result says this explicitly and sets recovery metadata to `not_available`; `workbook_history` remains fail-fast on WPS.
- **Office-coupled non-core tools:** `execute_office_js` and `python_transform_range` drive Office.js/Excel directly, so `createAllTools()` wraps them with the same fail-fast handler on WPS. Local-bridge tools (`tmux`, `python_run`, `libreoffice_convert`, `files`, extensions manager) are host-independent and stay unwrapped.
- **WPS direct JS tool:** `execute_wps_js` is registered only when `hostKind === "wps"`. It uses the same direct-JS approval gate and JSON serialization policy as `execute_office_js`, but runs synchronous WPS JSAPI code with `Application` in scope.
- **Typed failure:** fail-fast wrappers throw `UnsupportedHostToolError` (`code: "unsupported_host_tool"`, with `hostKind` + `toolName`), so callers/tests/UI can distinguish "unsupported on this host" from implementation failures.
- **Rationale:** preserving the tool list avoids prompt-cache/schema churn and keeps disclosure/UI mappings deterministic, while honest runtime failures prevent pretending unsupported WPS workbook features exist.

## Column width (`format_cells.column_width`)
- **User-facing unit:** Excel character-width units (same as Excel UI).
- **Conversion:** assume **Arial 10** and convert to points with `1 char ≈ 7.2 points`.
- **Application:** apply to **entire columns** via `getEntireColumn()`.
- **Verification:** read back `columnWidth` and warn if applied width differs.
- **Warnings:** if `font_name` or `font_size` is set and not Arial 10, we warn that widths may differ.
- **Rationale:** Excel column width is font-dependent and Office.js `columnWidth` is in points. A fixed Arial 10 baseline is predictable and simpler than per-sheet calibration.

## Borders (`format_cells.borders`)
- **Accepted values:** canonical `thin | medium | thick | none` (weight, not style).
- **Input normalization:** `borders` / `border_*` values are normalized case-insensitively and accept `Border*`/`Borders*` prefixes (`None`, `BordersNone`, etc.) before applying.
- **Implementation:**
  - `none` → `border.style = "None"`
  - others → `border.style = "Continuous"` + `border.weight = Thin|Medium|Thick`
- **Rationale:** Office.js `BorderLineStyle` does not include Thin/Medium/Thick; those are weights.

## Multi-range formatting (`format_cells.range`)
- **Supported syntax:** comma/semicolon separated ranges **on a single sheet**.
- **Implementation:** uses `worksheet.getRanges()` (RangeAreas).
- **Limitations:** multi-sheet ranges are rejected.
- **Rationale:** reduces repetitive calls for non-contiguous header styling.

## Overwrite protection (`write_cells.allow_overwrite`)
- **Blocks only on existing data:** values or formulas.
- **Does NOT block** on formatting, conditional formats, or data validation rules.
- **Rationale:** formatting-only cells are not meaningful "content" and shouldn't block writes.
- **Explicit requests count as confirmation:** the system prompt and the `allow_overwrite` description tell the agent to set it when the user asked to change those cells, and to ask only otherwise. Asking "Confirm replacing A2?" after "change A2 to Output" cost a turn and added nothing.

## Write results tell the truth (`write_cells`, `fill_formula`)
- **Shared path:** `src/tools/mutation/range-write.ts` (`inspectWriteTarget`, `verifyWrittenRange`); failure reasons in `src/tools/mutation/excel-failure.ts`.
- **Before writing (blocked, `isError: true`, nothing written):** locked cells on a protected sheet; target cells inside a merged area other than its top-left (Excel silently drops them); target cells inside a spilled range other than its anchor (Excel turns the anchor into `#SPILL!`). Merges are queried from A1 to the target's bottom-right because Excel on the web reports a merge only when its top-left cell is inside the queried range.
- **Writes are chunked** under Excel's request payload limit and synced on their own; `fill_formula` syncs the autofill before any read-back. A read-back failure after the write therefore reports "written, but could not be verified", never a failure for a change that landed.
- **Verification is bounded:** up to `MAX_RECOVERY_CELLS` (20,000) cells every cell is read back and compared with what was sent. Above that, the first and last ~1,000 cells are read back and the whole range is scanned with `getSpecialCellsOrNullObject` for formula errors and blank cells; the result says "Verification is partial" and what was checked (`details.verification: "sampled"`).
- **Compared, not trusted:** cells that do not hold the value sent are listed as not written (`isError: true`, `details.notWrittenCount`); text Excel converted to a number/date/boolean is listed with its displayed text and the apostrophe workaround (`'00123` stores text; `details.convertedCount`); values stored differently (duplicate table headers renamed `Qty2`) are listed; formula errors, including `#SPILL!`, are listed with formulas.
- **Data validation:** written values that violate the cells' rule are listed with the rule (`details.invalidCount`). Not blocked: Office.js does not enforce validation and the agent may be fixing data in bulk.
- **Manual calculation:** when `calculationMode` is not Automatic, every write result says dependents still show old results until the user recalculates (`details.calculationMode`); changing an input is the common case. When formulas were written, the target is also calculated (`range.calculate()`) before the read-back so their own values are current.
- **Backups for large writes:** a write above the backup limit into an empty target proceeds and says "No backup… was empty before, so clearing it undoes the change". A large write that would overwrite existing data is refused up front ("Write in blocks of at most 20,000 cells"), because it could not be undone.
- **Timeouts cannot cancel Excel.run:** `excelRun`'s timeout (#792) ends the tool call, but the batch keeps running. `write_cells` and `fill_formula` mark the run abandoned when the call fails and check that before every chunk or fill, so the abandoned run never writes more than the result allowed for. The result snapshots the written and in-flight rows before its first `await`.
- **Read-back size:** a full read-back is used only up to 20,000 cells and about 1.5 M characters sent; long text above that is verified in part, since reading values and formulas back in one call would exceed Excel's response limit.
- **Partial chunked failure:** if a later chunk fails, the result names the rows that were written, backs up and audits them, and is an error.
- **Formulas replaced by values:** when a write (`write_cells`, `python_transform_range`, WPS `write_cells`) leaves a constant where a formula was, the result says "⚠️ N formula(s) on sheet … replaced by values: D5, D6. D5 was `=…`" with up to three formulas (`details.formulasReplacedCount`). Not blocked: hard-coding may be what the user asked for, but it must never be silent. Clearing a formula cell (writing `""`) is not reported as a replacement.
- **Tables extended by Excel:** when a write next to a table makes Excel extend it, the result names the table, its new range and any header Excel made up (`Column1`).
- **Backup notes give the real reason:** "No backup was saved, so `workbook_history` cannot restore this change." followed by the cause from the tool (too large, with the 20,000-cell limit and `/backup` as the alternative; or Pi could not identify the workbook / save the backup), not one generic reason.
- **Stop and the coordinator:** `write_cells` and `fill_formula` set `reportsStop: true` (`StopReporting`). A wrapper that settles calls on Stop must let these finish: they check Stop between chunks, each sync is bounded by `excelRun`, and only their own result knows which rows landed and the backup id.
- **Stop:** a mutation stopped before it ran says "Stopped before running; nothing changed". `write_cells` and `fill_formula` check Stop between chunks and say "Stopped before writing — nothing changed in X" or "Stopped after writing X (rows …) — the change stands", backing up and auditing what landed. Their details carry `stopped: true` and `writtenCount` (cells confirmed written; `0` = nothing changed), so the UI need not parse the text. Any other mutation that finishes after Stop gets "Stop was pressed while this ran; the change above was already applied and stands." A timeout while the tool was still reading the target (before any write was sent) says "nothing was written" and does not call the outcome unknown. A failure whose outcome Excel never confirmed (no Office error code, or `outcomeUnknown`, e.g. #792's `Excel.run` timeout) marks the in-flight rows as "may or may not have been written" rather than unwritten.

## Tool failures are errors with a reason
- Every workbook mutation tool (`write_cells`, `fill_formula`, `format_cells`, `modify_structure`, `conditional_format`, `comments`, `charts`, `view_settings`, `python_transform_range`) returns failures with `isError: true`, states that nothing was changed, and maps the Office.js code to a reason: `AccessDenied` → probe sheet/workbook protection and name the sheet; `InvalidOperationInCellEditMode` → ask the user to press Enter/Esc; payload limits → split the request; `ItemNotFound` on a sheet → list the sheets that exist; `InvalidArgument` on `getRange` → invalid address. `execute_office_js` failures are also `isError: true`; Office.js codes get the same reasons (with no named target, protection names every protected sheet), and the result says earlier steps in the code may already have run.
- A failed mutation records an audit entry but no "backup not created" note: nothing changed.
- **Cell-edit mode:** one reason (`cellEditModeReason` in `excel-failure.ts`), worded like `excelRun`'s `ExcelCellEditModeError` (#792) without its "what was applied" clause, because each tool states what it changed. The write tools never assume `excelRun` retries: `write_cells` counts the chunks whose sync succeeded and reports those rows as written (backed up and audited); `fill_formula` reports whether the fill sync succeeded. Other mutation tools say "part of this change may already be in the workbook" instead of "not changed" when the error says an earlier step reached Excel (`earlierStepReachedExcel`).

## Deleting rows that are hidden (`modify_structure delete_rows`)
- Before deleting, the tool finds hidden rows in the range (one probe per row, up to 500) and keeps their cells. After deleting it reads back to see which ones Excel kept, because the result differs: Excel on the web deletes only the visible rows of a filtered range (table filter or AutoFilter), and the kept rows move up, but it does delete hand-hidden rows (confirmed 2026-10-27).
- Hidden rows deleted: "⚠️ N deleted row(s) were hidden: row … Deleted hidden data: row 3 ("South", "b")", plus the table and whether a filter is on (`details.hiddenRowsDeleted`). Not refused: a backup restores them, and the warning tells the agent how.
- Hidden rows kept: "Deleted 1 of 2 row(s) … Excel kept 1 hidden row(s)", and no backup is saved, since a backup of all the rows would restore the kept ones twice.
- Refused up front only when hidden rows hold data and the rows are too large to back up.

## Reads in a batch see earlier writes
- `withWorkbookCoordinator` tracks mutations that have started and not finished. A read waits for mutations started before it (the agent loop starts a batch's calls in order); mutations never wait for reads. A read issued before a write in the same batch is not delayed.

## Fill formulas (`fill_formula`)
- **Purpose:** avoid large 2D formula arrays by using Excel AutoFill.
- **Behavior:** sets formula in top-left cell, then `autoFill` across the range. A 2D range fills across the first row, then down: Excel on the web leaves the first row and column blank when one cell autofills a 2D range.
- **Validation:** uses `validateFormula()` (same as `write_cells`).
- **Overwrite protection:** blocks only if values/formulas exist (same policy as `write_cells`).
- **Rationale:** major productivity win for large formula blocks.

## Tool consolidation (14 → 10)
- `get_range_as_csv` merged into `read_range` as `mode: 'csv'`
- `read_selection` removed - auto-context already reads the selection every turn
- `get_all_objects` absorbed into `get_workbook_overview` via optional `sheet` param
- `get_recent_changes` removed - auto-context already injects changes every turn
- `find_by_label` (#7) absorbed into `search_workbook` via `context_rows` param
- `get_sheet_summary` (#8) absorbed into `get_workbook_overview` via `sheet` param
- **Rationale:** one tool per distinct verb, modes over multiplied tools. Progressive disclosure for future tools (charts, tables, etc.)

## Range reading (`read_range`)
- **Compact/detailed tables:** render an Excel-style markdown grid with **column letters** and **row numbers** (instead of treating the first data row as a table header).
- **Empty ranges:** if a range has **no values, formulas, or errors**, return `_All cells are empty._` (omit the table) to avoid confusing "blank header" visuals.
- **Rationale:** improves readability in the sidebar UI and avoids ambiguous tables for 1-row or empty ranges.
- **Bounded reads:** a page holds at most 500 rows / 2,500 cells (100 columns). Whole-column/row addresses (`A:A`, `3:5`) and any range larger than a page are clipped to the sheet's used range (values only) first. A larger target returns its first page plus the exact range to read next (`Next page: read_range "Big!A501:E1000"`); no `offset` parameter, so the tool schema stays as it was. A load that still hits Excel's response-size limit (long text) halves its rows and retries.
- **Displayed values:** compact/detailed tables show each cell as the user sees it (`range.text`). When the display hides the stored value (dates, rounded currency), the stored value follows in brackets: `01/01/2024 [45292]`, `£98,765.43 [98765.4321]`; a legend line appears once. Displays that denote the value exactly (`12.5%`, `£1,234.50`), General-format cells and `####` overflow show one value. CSV mode stays raw.
- **Numbers and text:** numbers render with `String()` (0 is `0`; no scientific notation for small or large values). `formatAsMarkdownTable` escapes `\`, `|` and line breaks (`\n`) in every cell, so cell text cannot split the grid.
- **Sheet state notes (`sheet-state.ts`):** for the page read, one line each, only when present: manual calculation mode, sheet protection with the cells' locked state, hidden (filtered) rows and columns, tables touching the range and whether a filter is applied, sheet AutoFilter, merged areas, spill ranges (formula cells that can spill, up to 100 checked), data validation rules. A state read that fails adds an "(Some sheet state could not be read …)" line instead of failing the read.
- **Summaries first:** notes, formulas and errors come before the table, so head truncation cuts table rows rather than the summary. Formulas filled down a column collapse to one entry by R1C1 (`F2:F833: =C2*2 (filled, 832 cells)`); formula, error and number-format lists are capped with a `+N more` count.
- **Value notes:** numbers stored as text (strings Excel would read as numbers, e.g. `"12"`, `"1,200"`, `"40%"`; not zero-padded codes like `"00123"`) are listed, because the grid shows them like numbers and SUM/AVERAGE skip them. A page with 50+ numbers in a column gets exact column sums for the page (up to 10 columns) and a line telling the agent to compute other aggregates (execute_office_js `context.workbook.functions`, python_run) instead of adding values itself. Real turns on GPT-6 Luna summed 1,000 paged rows wrong (8,965 vs 8,978) without it. Result text only; tool schemas unchanged.
- **Sheet resolution:** the sheet is resolved once, before any range work. An address without a sheet uses the active sheet and the result says so (`Sheet: "Model" (the active sheet; …)`); an unknown sheet fails with the list of sheet names.

## Workbook search (`search_workbook`)
- **Chunked scan:** each sheet's used range is loaded in row chunks of about 100k cells (halved on a response-size error) instead of one used-range load, which failed for the whole workbook once any sheet was large. Each chunk loads only the searched grid; the other grid (formulas or values) is loaded only for chunks with a match. Context rows are loaded across chunk boundaries.
- **Bound:** one search scans at most 5M cells and says where it stopped.
- **Unknown sheet:** `sheet` naming a missing sheet fails with the list of sheet names instead of reporting no matches.

## Sheet overview (`get_workbook_overview {sheet}`)
- **No used-range values:** the preview reads only its first 5 rows × 20 columns (as displayed), so a 200k-row sheet previews like a small one.
- **Labels follow the used range:** the preview's row numbers and column letters start where the used range starts, and `Headers` is the first row of the used range (not row 1), named when it is not row 1: `Headers (row 5): Region, Q1, Q2`. Same in the workbook-level overview.
- **Sheet state:** the same notes as `read_range` for the used range, plus frozen panes. The workbook-level tool result appends a `### Current state` section (calculation mode when not Automatic, protected sheets), read fresh per call. These stay out of `buildOverview()`: that text is the cached auto-context blueprint, which only refreshes on structure changes, and protection or calculation-mode changes are not structure changes.

## Dependency ranges (`trace_dependencies`)
- **Ranges are not collapsed:** a precedent/dependent range of up to 16 cells is traced cell by cell (`SUM(E2:E3)` → E2 and E3); a larger range becomes one range node whose values are not loaded (`Mid50k!C2:C3000 (range, values not traced)`). Previously only the first cell was kept.

## Default formatting assumption
- **System prompt:** "Default font for formatting is Arial 10 unless user specifies otherwise."
- **Rationale:** keeps column width conversions consistent with the chosen baseline.

## Named styles and format presets (`format_cells.style`)
- **Style param:** `string | string[]` — single name or composable array (left-to-right merge).
- **Built-in format styles:** `number` (2dp), `integer` (0dp), `currency` ($, 2dp), `percent` (1dp), `ratio` (1dp, x suffix), `text`.
- **Built-in structural styles:** `header`, `total-row`, `subtotal`, `input`, `blank-section`.
- **Composition:** format + structural styles are orthogonal (no property overlap), so composing is always clean.
- **Override with params:** individual params (bold, fill_color, etc.) always win over style properties (CSS inline specificity).
- **`number_format` accepts preset names:** `number_format: "currency"` is equivalent to `style: "currency"` — backward compatible, raw format strings still accepted.
- **`number_format_dp`:** override decimal places for any numeric preset.
- **`currency_symbol`:** override the currency symbol (only applies to `currency` preset; warned and ignored otherwise).
- **Type-checking warnings:** integer + dp > 0, currency_symbol on non-currency, dp on text → warning in tool output.
- **Date formats:** no preset — too many variations. Use raw `number_format` string (e.g. "dd-mmm-yyyy").
- **Source of truth:** `src/conventions/defaults.ts` — format strings, styles, and house-style conventions are defined once and imported by tools + prompt.
- **Rationale:** agents say `"currency"` instead of pasting fragile 40-char format strings. Composition reduces multi-param calls. See `.research/conventions-design.md` for full design.

## Individual border edges (`format_cells.border_top/bottom/left/right`)
- **New params:** `border_top`, `border_bottom`, `border_left`, `border_right` — each accepts `thin | medium | thick | none`.
- **Priority:** individual edge params > style-resolved edges > `borders` shorthand.
- **Shorthand preserved:** `borders` still applies to all edges + inside (existing behavior, backward compatible).
- **Rationale:** enables `total-row` style (top border only) and other edge-specific formatting without the all-edges shorthand.

## `view_settings` scope boundary (view vs print)
- **Included (view/navigation):** gridlines, headings, freeze panes, tab color, sheet visibility (`Visible/Hidden/VeryHidden`), activate sheet, standard column width.
- **Excluded (print/page layout):** zoom, margins, orientation, print area, and other `pageLayout` concerns.
- **Rationale:** keep `view_settings` focused on what the user sees/navigates in-sheet. Print concerns belong in a separate future `page_layout` tool.

## Conventions tool (`conventions`)
- **Actions:** `get` (view current), `set` (partial update), `reset` (restore defaults).
- **Storage:** `SettingsStore` key `conventions.v1` (user-level only for now).
- **Schema:** `StoredConventions` now stores:
  - built-in preset format strings (`presetFormats.<preset>.format`)
  - optional builder metadata (`builderParams`) for quick-toggle regeneration
  - custom presets (`customPresets`)
  - visual defaults (`visualDefaults.fontName/fontSize`)
  - font-color conventions (`colorConventions`)
  - header style (`headerStyle`)
- **Format-string-first model:** `format` is the source of truth. `builderParams` are auxiliary metadata and may be absent for hand-authored/custom formats.
- **Resolution:** stored overrides merge over defaults for preset formats, colors, header style, and default font. `format_cells` loads resolved conventions each call.
- **System prompt:** non-default values are injected as "Active convention overrides"; configured custom preset names are listed for agent use.
- **Execution policy:** classified as read/none (mutates local config, not workbook).
- **Validation:** nested sections are validated on read/write. Colors accept hex or `rgb(...)` input and are normalized to hex.
- **Rationale:** power users can set exact Excel format strings while still keeping optional quick-toggle ergonomics for generated presets.

## Instructions tool (`instructions`)
- **Scopes:** `user` (global, local machine) and `workbook` (scoped by workbook identity hash).
- **Actions:** `append` and `replace`.
- **Storage:** `SettingsStore` keys:
  - `user.instructions`
  - `workbook.instructions.v1.<workbookId>`
- **Execution policy:** classified as read/none for workbook coordinator purposes (it mutates prompt metadata, not workbook cells/structure).
- **Rationale:** AGENTS.md-style persistent guidance without creating a separate workbook mutation path.

## Tool details are schemas, decoded once at the UI seam
- **Contract:** each `details` payload is a TypeBox schema in `src/tools/tool-details.ts`; the exported type is `Static<>` of the schema, so tools that construct the payload are checked against the same shape the UI validates at runtime.
- **Seam:** `ToolResultMessage.details` reaches the renderer as `unknown` (older persisted sessions may carry any shape). `decodeToolDetails()` runs once per tool card and returns the `ExcelToolDetails` union or `undefined`; everything inward takes the union and switches on `kind`. A payload that fails its schema renders generically rather than partially.
- **Extras:** schemas allow additional properties so cross-cutting metadata (`outputTruncation`) can be merged into any payload without each schema knowing about it.
- **Cell values:** `read_range_csv.values` and `trace_dependencies` node values stay `unknown` until the host adapters type `range.values`; that is the host-boundary item in `docs/clean-base-followups.md`.
- **No compiled validators:** `Value.Check` only. `typebox/compile` generates code at runtime, which the Office WebView CSP forbids. A decode costs roughly 10–20 µs.

## Recovery snapshots are schemas; a persisted entry is decoded or dropped
- **Contract:** every recovery state (`src/workbook/recovery/schemas.ts`) and the persisted snapshot envelope (`PersistedWorkbookRecoverySnapshotSchema` in `log-codec.ts`) is a TypeBox schema with its type as `Static<>`. The snapshot is a union keyed by `snapshotKind`; each kind requires its own state field, and `snapshotKind` may be absent only for the original `range_values` entries.
- **Invariants live in the schema:** grid dimensions against `rowCount`/`columnCount` (structure data ranges, format `numberFormat`/`columnWidths`/`rowHeights`) are `Type.Refine` checks, so a persisted entry that breaks one fails the same `Value.Check` as one with a wrong field type.
- **Drop, do not repair:** an entry that fails its schema is skipped (unknown kind, unknown tool name, wrong field type, dimension mismatch). The writer is this module, so a mismatch is version skew or corruption, neither of which is a restorable checkpoint. Missing `id`/`at`/`cellCount`/`changedCount` are still filled in, since early entries lacked them.
- **Rule variants:** `RecoveryConditionalFormatRule` is a union keyed by `type` with the per-type required fields (`custom.formula`, `cell_value.operator`+`formula1`, …). Apply handlers receive their own variant and no longer re-validate; `applyConditionalFormatRule` dispatches with a `switch` so each case narrows.
- **Clone:** `cloneRecoveryState(schema, value)` (`Value.Clone` + `Value.Clean`) replaces the field-by-field clone functions; it cannot drift from the schema and also strips properties the schema does not declare.
- **Grids stay off the generic checker:** cell grids are `Type.Unsafe` with a `Type.Refine` row walk. A per-cell `Type.Array(Type.Array(...))` check costs about 20 ms per 20k cells and `Value.Clean` on the snapshot union re-checks every variant, which put the worst-case boot load (120 full-size entries) at 560 ms against 20 ms before; with the row walk and per-state cleaning it is 30 ms.

## Global tool output truncation (Pi-style guardrail)
- **Scope:** applied as a runtime wrapper around all registered tools (core + integrations + extensions) before tool results are persisted to message history.
- **Limits:** **50KB** UTF-8 bytes and **2000 lines** (whichever is hit first), aligned with pi-coding-agent defaults. For models with context windows **below 128k**, limits scale linearly with the window (floors: **8KB** / **200 lines**) — resolved per execution from the active model via `src/context/window-budgets.ts` (#566).
- **Strategy:**
  - default: **head** truncation (read/search style outputs)
  - specific log/terminal style tools (`python_run`, `tmux`, `mcp`, `execute_office_js`): **tail** truncation
- **Metadata:** truncated results include stable `details.outputTruncation` with strategy, hit reason, total/output sizes, and limits.
- **Overflow persistence:** best-effort full-output save to Files workspace under `.tool-output/...` for truncated payloads within save budget.
- **Rationale:** enforce predictable context-safe bounds independent of per-tool implementation details; keep `shapeToolResultsForLlm` as a secondary history-shaping layer.

## Tool card input/output humanization (UI)
- **Input:** tool parameters are rendered as a clean key-value list instead of raw JSON. Each tool has a per-tool humanizer in `src/ui/humanize-params.ts` that maps params to readable labels (e.g. "Range", "Fill ● White", "Font ● Gray, italic").
- **Output:** hex color codes (`#RRGGBB`) in tool result text are replaced with human-readable names via nearest-match against a ~45-color palette (`src/ui/color-names.ts`). Section label changed from "Output" to "Result".
- **Color chips:** inline colored circles (`pi-color-chip`) shown next to fill/font colors.
- **Data preview:** `write_cells` values shown as a mini table (up to 3 rows × 6 columns) instead of raw JSON arrays.
- **Fallback:** unknown tools (non-Excel) still get the raw JSON code-block.
- **Rationale:** raw JSON and hex codes are unintuitive for Excel-savvy, less-technical users. The humanized view keeps all info but presents it in Excel vocabulary.

## Format humanization (`read_range` detailed mode)
- **Behavior:** known format strings are displayed with human-readable labels alongside the raw string (e.g. `**currency (£, 2dp)** (\`£* #,##0.00...\`)`).
- **Unknown formats:** displayed as raw strings (no change from before).
- **Implementation:** `src/conventions/humanize.ts` pre-generates a lookup table from all preset+dp+symbol combinations.
- **Rationale:** raw format strings in read-back are opaque; labels make them immediately understandable.

## CSV table rendering (`read_range` mode=csv)
- **UI:** CSV results are rendered as an HTML table with Excel-style column letters (A, B, …) and row numbers, plus a "Copy CSV" button.
- **Agent text:** unchanged — still the markdown code-fenced CSV block.
- **Implementation:** `ReadRangeCsvDetails` passes `values[][]`, `startCol`, `startRow`, and `csv` string to the UI. `src/ui/render-csv-table.ts` renders the table.
- **Rationale:** the syntax-highlighted code block (language "csv") produced garbled output with numbers in red and keywords in blue. A proper table with row/column headers is immediately readable.

## Dependency tree rendering (`trace_dependencies`)
- **Modes:** `trace_dependencies` supports both `mode: "precedents"` (upstream inputs) and `mode: "dependents"` (downstream impact).
- **UI:** dependency trees are rendered as structured HTML with clickable cell refs, code-styled formulas, and collapsible branch nodes for on-demand deep expansion.
- **Agent text:** unchanged — still the ASCII tree with `├──`/`└──`/`│` connectors.
- **Implementation:** `TraceDependenciesDetails` carries `mode` + tree metadata; `src/ui/render-dep-tree.ts` renders the visual tree and supports branch expand/collapse.
- **Fallback behavior:** tool prefers Office.js direct precedent/dependent APIs and falls back to formula parsing/scanning when APIs are unavailable.
- **Rationale:** ASCII art via `<markdown-block>` lacked interactivity and visual hierarchy. Clickable addresses + collapsible branches make formula navigation significantly more usable.

## Formula explanation workflow (`explain_formula`)
- **Scope:** `explain_formula` targets a single cell and returns a concise natural-language explanation, current value preview, formula text, and direct reference citations.
- **Reference preview policy:** loads and previews a bounded set of direct references (`max_references`, default 8, max 20) to keep response latency predictable.
- **Fallback behavior:** if the target is not a formula cell, returns an explicit static-value explanation instead of failing silently.
- **UI:** explanation card renders clickable reference citations with value previews.
- **Rationale:** users need plain-English interpretation without losing inspectability; bounded reference previews preserve responsiveness on dense workbooks.

## Tmux bridge tool (`tmux`)
- **Availability:** non-core tool, always registered via `createAllTools()`; execution is gated by `applyExperimentalToolGates()`.
- **Gate model:** requires a healthy bridge URL (`tmux.bridge.url` override, else default `https://localhost:3341`) and successful `/health` probe.
- **Gate failure contract:** blocked gate checks return structured `AgentToolResult` payloads (`details.gateReason`, `details.skillHint`) instead of throwing, enabling inline setup UX and deterministic agent recovery.
- **Execution policy:** classified as `read/none` in workbook coordinator (no workbook lock writes or blueprint invalidation).
- **Bridge implementation:** local helper script `scripts/tmux-bridge-server.mjs`.
  - one-command helper (`npx pi-for-excel-tmux-bridge`) defaults to real `tmux` mode
  - raw server script default remains `stub` for local development/tests
- **Bridge contract:** POST JSON to `https://localhost:<port>/v1/tmux` with actions:
  - `list_sessions`
  - `create_session`
  - `send_keys`
  - `capture_pane`
  - `send_and_capture`
  - `kill_session`
  - capture actions support optional `wait_ms` delay (0..120000) to avoid tight polling loops for long-running commands
- **Security posture:** local opt-in only; bridge URL validated via `validateOfficeProxyUrl`; tool execution re-checks gate before every call; bridge enforces loopback+origin checks and optional bearer token (`TMUX_BRIDGE_TOKEN` / setting `tmux.bridge.token`, managed via `/experimental tmux-bridge-token ...`).
- **Diagnostics UX:** `/experimental tmux-status` reports URL/token config, gate result, and bridge health details for quick troubleshooting.
- **Rationale:** stable local adapter contract now (issue #3) with one-command real execution and incremental hardening.

## Python / LibreOffice execution tools (`python_run`, `libreoffice_convert`, `python_transform_range`)
- **Availability:** always registered via `createAllTools()`.
- **Default runtime (Pyodide):** `python_run` and `python_transform_range` run in-browser via Pyodide (WebAssembly) with zero setup. Standard library and pure-Python packages (numpy, pandas, scipy, etc.) auto-install via micropip. ~15MB cold-start on first use, cached by the browser thereafter.
- **Power-user upgrade (native bridge):** when a bridge URL is configured (Settings → Experimental) — or when a bridge is reachable at the default URL (`https://localhost:3340`) — Python tools use the local `python3` process. This unlocks C extensions, filesystem access, and long-running scripts. `libreoffice_convert` requires the native bridge (no Pyodide equivalent).
- **Fallback order:** native bridge (configured URL or reachable default URL) → Pyodide (if WebAssembly + Workers available) → error.
- **System prompt awareness:** the agent knows Python is available by default via Pyodide and only mentions the native bridge when the task requires native-only capabilities.
- **Gate model:**
  - no `python-bridge` experiment flag.
  - effective bridge URL = configured override or default `https://localhost:3340`.
  - first execution requires user confirmation when the effective bridge URL is reachable (cached once per bridge URL).
  - bridge-only tool (`libreoffice_convert`) is blocked when no reachable bridge URL is available.
  - blocked gate checks return structured `AgentToolResult` payloads (`details.gateReason`, `details.skillHint`) rather than thrown errors.
- **Execution policy:**
  - `python_run` + `libreoffice_convert` → `read/none` (no direct workbook mutation)
  - `python_transform_range` → `mutate/content` (writes transformed values into workbook)
- **Bridge implementation:** local helper script `scripts/python-bridge-server.mjs`.
  - one-command helper (`npx pi-for-excel-python-bridge`) defaults to real mode
  - raw server script default remains `stub` for local development/tests
- **Bridge contract:**
  - `POST /v1/python-run` — execute Python snippet with optional `input_json`, return stdout/stderr/result JSON
  - `POST /v1/libreoffice-convert` — convert files across `csv|pdf|xlsx`
- **Security posture:** bridge URL validated via `validateOfficeProxyUrl`; approval prompt protects native bridge execution; bridge enforces loopback+origin checks and optional bearer token (`PYTHON_BRIDGE_TOKEN` / setting `python.bridge.token`, managed via `/experimental python-bridge-token ...`).
- **Overwrite perf guard (`python_transform_range`):** pre-write `values/formulas` reads are skipped for large `allow_overwrite: true` outputs (> `MAX_RECOVERY_CELLS`) since those snapshots would be dropped anyway.
- **Rationale:** Python works out of the box for most users via Pyodide. The native bridge is now plug-and-play when the local helper is running, while Pyodide remains a resilient fallback. The system prompt makes the agent aware of both tiers so it can use Python confidently without suggesting unnecessary setup.

## External tool integrations (`web_search`, `fetch_page`, `mcp`)
- **Packaging:** exposed as opt-in **integrations** instead of always-on core tools.
- **Scopes:** integrations can be enabled per-**session** and/or per-**workbook**; effective integrations are the union (ordered by catalog).
- **Global gate:** `external.tools.enabled` defaults to **off** and blocks all external integration tools until explicitly enabled.
- **Web search providers:** Jina (default, zero-config), Serper.dev, Tavily, and Brave Search (`web_search`) with optional proxy routing and explicit "Sent" attribution in results.
- **Fallback behavior:** when a configured keyed provider fails with auth/rate-limit/server/network errors, `web_search` automatically retries with Jina for that call and includes a warning in tool output/details.
- **Page retrieval companion:** `fetch_page` fetches URL content and returns extracted markdown for source grounding workflows (`web_search` → `fetch_page`).
- **MCP integration:** configurable server registry (`mcp.servers.v1`) + bearer token secrets in connection store (`connections.store.v1` / `builtin.mcp.servers`), UI add/remove/test, and a single `mcp` gateway tool for list/search/describe/call flows.
- **Rationale:** satisfy issue #24 with explicit consent, clear attribution, and minimal overlap with the extension system.

## Direct Office.js tool (`execute_office_js`)
- **Availability:** always available (not behind `/experimental`), always registered via `createAllTools()`.
- **Contract:** accepts `code` (async function body receiving `context: Excel.RequestContext`) plus a short user-facing `explanation`.
- **Safety guards:** blocks nested `Excel.run(...)` usage (host already provides context), enforces explanation/code length limits, and fails closed if confirmation UI is unavailable when approval is required.
- **Approval model:** Confirm mode prompts on every call. Auto mode runs pure Excel API code without a prompt (#347), but a static ambient-authority lint (`experimental-tool-gates/office-js-risk.ts`) forces per-call approval — in any mode — when code references browser capabilities beyond the Excel API (network egress, storage, DOM/global handles, workers, dynamic evaluation, `Office.*`, or realm escape via `constructor`). Flagged approvals include the detected identifiers in the prompt.
- **Threat model:** submitted code executes in the taskpane page realm (blob module import) with ambient access to the realm that holds provider credentials, and workbook content is untrusted input injected into context every turn. The lint is a heuristic tripwire (raw-text scan, false positives cost one dialog); the durable fix is an isolated-realm runner with only a guarded workbook API bridged in (#605).
- **Result policy:** tool output must be JSON-serializable; non-serializable results are returned as deterministic errors.
- **Execution policy:** treated as `mutate/structure` to force conservative workbook-context refresh after execution.
- **Rationale:** unlock advanced Office.js scenarios when structured tools are insufficient while preserving explicit consent.

## Extension manager tool (`extensions_manager`)
- **Availability:** always registered via `createAllTools()`.
- **Purpose:** lets the agent manage extension lifecycle from chat (`list`, `install_code`, `set_enabled`, `reload`, `uninstall`).
- **Default install policy:** `install_code` replaces existing extensions with the same name unless `replace_existing=false` is provided.
- **Execution policy:** treated as `read/none` for workbook coordination (mutates local extension registry/runtime only, not workbook cells/structure).
- **Rationale:** supports non-engineer extension authoring by allowing users to ask Pi to generate + install an extension directly.

## Extension sandbox UI bridge (default-on for untrusted)
- **Default routing:** inline-code + remote-url extensions run in iframe sandbox runtime by default; built-in/local modules remain host-side.
- **Rollback switch:** maintainers can temporarily route untrusted extensions back to host runtime via `/experimental on extension-sandbox-rollback`.
- **Surface:** sandbox runtime bridges command/tool/event/UI calls through explicit host contracts rather than exposing host internals directly.
- **UI model:** sandbox may only send a structured UI tree (allowed tag set, sanitized class names/action ids), never raw HTML.
- **Interactivity:** host supports explicit action callbacks via `data-pi-action` markers mapped to click dispatch inside the sandbox.
- **Rationale:** graduate sandbox hardening into default behavior while preserving a guarded rollback path.

## Extension host capability bridge expansion (`ExcelExtensionAPI`)
- **New mediated APIs:** `llm.complete`, `http.fetch`, `storage`, `clipboard`, `agent.injectContext/steer/followUp`, `skills`, and `download`.
- **Permission model:** capability gates now include `llm.complete`, `http.fetch`, `storage.readwrite`, `clipboard.write`, `agent.context.write`, `agent.steer`, `agent.followup`, `skills.read`, `skills.write`, and `download.file`.
- **Dynamic tools:** extensions can now remove tools at runtime via `unregisterTool(name)`; runtime refreshes toolset after dynamic add/remove.
- **Storage lifecycle:** extension-scoped storage is persisted under `extensions.storage.v1` and cleared on extension uninstall.
- **HTTP safety:** outbound extension HTTP calls enforce local/private-network host blocking, timeout caps, and response-size limits.
- **Rationale:** unlock practical extension workflows (sub-agents, external APIs, persistence, skill install) while keeping sandbox mediation and permission controls.

## Feature-flagged extension widget API v2 (`extension-widget-v2`)
- **Activation:** opt-in via `/experimental on extension-widget-v2`; default behavior stays on legacy `widget.show/dismiss` semantics.
- **API:** additive `widget.upsert/remove/clear` methods with stable widget ids.
- **Placement/order:** widgets sort deterministically by `(order asc, createdAt asc, id asc)` within `above-input` / `below-input` buckets.
- **Ownership model:** widgets are extension-owned (`ownerId`) and auto-cleared on extension teardown/reload/uninstall.
- **Header behavior (slice B):** `collapsible: true` renders host-owned expand/collapse controls with predictable labels and keyboard focus semantics.
- **Sizing behavior (slice B):** `minHeightPx` / `maxHeightPx` are clamped to safe host bounds (`72..640`), `max < min` is normalized to `max = min`, and `null` clears an existing bound.
- **Upsert semantics (slice B):** omitted optional metadata preserves existing widget state; updates can focus on content without restating layout fields.
- **Compatibility:** legacy `widget.show/dismiss` remains supported and maps to a reserved legacy widget id when v2 is enabled.
- **Rationale:** establish predictable multi-widget lifecycle semantics before richer layout controls.

## Feature-flagged files tool (`files`)
- **Availability:** non-core tool, always registered. `list`/`read` stay available even when `files-workspace` is off; `write`/`delete` remain gated by `files-workspace`.
- **Built-in assistant docs:** a read-only `assistant-docs/` namespace ships with the app (README + key docs) and is always visible to both UI and tool.
- **Backend strategy:** native folder handle (when permitted) → OPFS → in-memory fallback.
- **Workbook tagging:** files are **not segregated** by workbook; each file stores an optional workbook tag (`workbookId` + label) based on the active workbook when last written/imported.
- **Audit trail:** persisted locally (list/read/write/delete/rename/import/backend switches) including actor, source, timestamp, and workbook label. Not shown in the Files dialog UI — available via `/export audit` for debugging.
- **Download fix:** uses `window.open(blobUrl)` instead of `<a download>` + `anchor.click()` for reliable binary file downloads in Office Add-in WebView (WKWebView on macOS silently ignores programmatic anchor clicks).
- **Detail actions:** file detail view exposes `Open ↗` + `Download` for all files, and writable files also expose `Rename` + `Delete`.
- **Open safety:** `Open ↗` and file downloads both sanitize script-capable MIME types (`text/html`, `image/svg+xml`, JavaScript, etc.) to `application/octet-stream` via shared blob-URL safety helpers.
- **Rename safety:** if a rename input omits an extension, preserve the current extension (`report.xlsx` + `report-final` → `report-final.xlsx`) unless the user explicitly types a dotted name.
- **Preview UX:** Files dialog supports inline text editing plus image/PDF preview; other binaries fall back to metadata + download. Built-in docs are marked read-only in the UI.
- **Filter UX:** Files dialog includes workbook-tag filtering (`all`, `current workbook`, `untagged`, and per-tag options) without changing underlying shared storage.
- **Input drop UX:** dropping files onto the chat input imports them directly into workspace (and auto-enables `files-workspace` if needed).
- **Naming:** user-facing UI uses "Files" (not "Files workspace"). Internal flag id remains `files_workspace`; slug `files-workspace` and alias `files` are both accepted.
- **Rationale:** keep one shared artifact space while preserving workbook context/transparency, while making core assistant docs available without extra setup.

## Workbook mutation change previews + audit log (slice)
- **Cell-diff scope:** `write_cells`, `fill_formula`, and `python_transform_range` compute before/after cell diffs.
- **Structured details:** these tools return `changes` metadata (`changedCount` + sampled cell-level before/after, including formula deltas) for tool-card rendering.
- **UI rendering:** tool cards include a dedicated **Changes** section with clickable cell addresses.
- **Compact status receipts:** mutation card headers include changed/error counts when available (e.g., `— 24 changed, 1 error`) for at-a-glance comprehension.
- **Context efficiency:**
  - diff samples are intentionally bounded (default sample limit = 12 changed cells)
  - `write_cells` verification output shows a bounded preview for large writes instead of dumping full tables
- **Audit coverage extension:** `format_cells`, `conditional_format`, `modify_structure`, mutating `comments` actions, mutating `view_settings` actions, and `workbook_history` restore now append structured entries to `workbook.change-audit.v1` (operation-focused summaries, not per-cell value diffs).
- **Export option:** `/export audit` writes the persisted workbook mutation audit log as JSON (download by default, `clipboard` optional).
- **Optional explanation UX:** mutation tool cards expose an on-demand **Explain these changes** drawer that synthesizes a concise explanation + clickable citations from structured audit metadata, with bounded payload/text limits.
- **Rationale:** improve user trust with concrete, navigable deltas while keeping implementation incremental and low-risk.

## Charts tool (`charts`)
- **Actions:** `list`, `create`, `update`, `delete`, `get_image` in one action-based core tool.
- **Chart type scope:** v1 exposes common ExcelApi 1.1 chart families only (column/bar/line/area/pie/doughnut/scatter/radar), mapped from friendly names to Office.js `ChartType` values.
- **Image contract:** `get_image` returns a short text receipt plus an image content block; base64 PNG bytes and dimensions live in `details.image`, not in text output.
- **Backups:** `create` stores `chart_absent` (restore deletes the created chart); `update` stores `chart_present` with pre-update chart properties. If `update` changes `source_range`, restore cannot revert the prior data source because Office.js does not expose it, so the tool result explicitly notes this limitation. `delete` has no automatic backup in v1 for the same source-range reason and reports `no backup`.
- **Restore identity:** `chart_absent` checkpoints capture the stable Office.js chart id at creation (best-effort — hosts without `Chart.id` still create charts). Restore locates the chart by id (rename-proof, never deletes an unrelated chart that reused the name). If a stored id cannot be verified at restore time, the restore fails safely without deleting anything; unique-name matching applies only to legacy states with no captured id. Chart restores return the post-restore identity so inverse (rollback) snapshots are addressed where the chart actually is after a rename-restore.
- **Downgrade safety:** persisted `chart_state` snapshots omit the `beforeValues`/`beforeFormulas` grids so codecs predating the kind fail their range-grid check and drop the entry instead of misreading it as an empty range backup.
- **Context impact:** `list`/`get_image` are read-only; `create`/`delete` invalidate structure context, while `update` is content impact.
- **Rationale:** cover common chart workflows with typed params and visual verification while avoiding a false promise of faithful deleted-chart recreation.

## Workbook backups (`workbook_history`)
- **Goal:** prefer low-friction workflows over pre-execution approval selectors by making rollback easy and reliable.
- **Execution mode toggle:** `/yolo` switches between:
  - `YOLO` (default): mutate tools run without extra pre-execution confirmations.
  - `Safe`: mutate tools require pre-execution confirmation via runtime gate.
- **Automatic backups:** successful `write_cells`, `fill_formula`, `python_transform_range`, `format_cells`, `conditional_format`, mutating `comments` actions, `charts` create/update, and supported `modify_structure` actions (`rename_sheet`, `hide_sheet`, `unhide_sheet`, `insert_rows`, `insert_columns`, `add_sheet`, and `duplicate_sheet` when the duplicate has no value data) store pre-mutation snapshots in local `workbook.recovery-snapshots.v1`.
- **Manual full-workbook fallback (`/backup`):** explicit user-triggered full-file captures are stored in Files workspace under `manual-backups/full-workbook/v1/...` and downloaded for restore. This remains separate from automatic per-mutation checkpoints.
- **Safety limits:** backup capture is skipped for very large writes (> `MAX_RECOVERY_CELLS`) to avoid oversized local state.
- **Workbook identity guardrails:** append/list/delete/clear/restore paths are scoped to the active workbook identity; restore rejects identity-less or cross-workbook backups.
- **Never-saved workbooks:** a workbook with no path-derived identity gets backups scoped to a random token stored in the document's add-in settings (`doc_instance:<uuid>`, see `src/workbook/recovery-scope.ts`). The token is minted only when the first backup is captured (the workbook is already dirty by then), never by list/restore/polling, and never on hosts that cannot store document settings (WPS, browser). It is not an alias: once the workbook has a path-derived identity, token-scoped backups are no longer visible, because Save As / Save a Copy duplicate the token. Manual full-workbook backups (`/backup`) stay canonical-identity-only for the same reason. Workbook identity is read live through `Office.context.document.getFilePropertiesAsync` (`src/host/office-document-url.ts`), because the static `Office.context.document.url` is captured at add-in init and not updated by Save As; the first save of a new workbook therefore switches to the path identity on the next read, and the save-boundary monitor clears the token-scoped backups at that transition so nothing lingers under the token.
- **Save boundary behavior:** backups are intended as "in between saves" recovery points and are cleared after the workbook transitions from dirty → saved.
- **Restore UX:** `workbook_history` can list/restore/delete/clear backups; restores also create an inverse backup (`restore_snapshot`) so users can undo a mistaken restore.
- **Coverage signaling:** `modify_structure`, mutating `view_settings`, and `charts delete` actions explicitly report when no backup was created.
- **Current `modify_structure` backup behavior:** captures/restores all `modify_structure` actions, including value-preserving checkpoints for destructive deletes (`delete_rows`, `delete_columns`, `delete_sheet`) by storing deleted value/formula data when capture is within recovery size limits.
- **Restore safety gate for structure absence states:** restoring `sheet_absent` / `rows_absent` / `columns_absent` checkpoints remains blocked when target data exists, unless the checkpoint was explicitly generated with data-delete intent during a prior restore inversion (`allowDataDelete`).
- **Current `format_cells` backup scope:** captures/restores core range-format properties (font/fill/number format/alignment/wrap/borders), row/column dimensions (`column_width`, `row_height`, `auto_fit`), and merge state (`merge`/`unmerge`).
- **Current `conditional_format` backup scope:** captures/restores all current rule families (`custom`, `cell_value`, `contains_text`, `top_bottom`, `preset_criteria`, `data_bar`, `color_scale`, `icon_set`) including per-rule applies-to ranges.
- **Quick affordance:** users can restore via `/history` (Backups overlay) or `/revert` (latest backup). The Backups overlay also exposes a **Full backup** button for explicit manual full-workbook capture.
- **Rationale:** addresses #27 by shifting from cumbersome up-front approvals to versioned recovery with explicit user-controlled rollback.
