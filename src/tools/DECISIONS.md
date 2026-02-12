# Tool Behavior Decisions (Pi for Excel)

Concise record of recent tool behavior choices to avoid regressions. Update this as we tweak tooling.

## Column width (`format_cells.column_width`)
- **User-facing unit:** Excel character-width units (same as Excel UI).
- **Conversion:** assume **Arial 10** and convert to points with `1 char ≈ 7.2 points`.
- **Application:** apply to **entire columns** via `getEntireColumn()`.
- **Verification:** read back `columnWidth` and warn if applied width differs.
- **Warnings:** if `font_name` or `font_size` is set and not Arial 10, we warn that widths may differ.
- **Rationale:** Excel column width is font-dependent and Office.js `columnWidth` is in points. A fixed Arial 10 baseline is predictable and simpler than per-sheet calibration.

## Borders (`format_cells.borders`)
- **Accepted values:** `thin | medium | thick | none` (weight, not style).
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

## Fill formulas (`fill_formula`)
- **Purpose:** avoid large 2D formula arrays by using Excel AutoFill.
- **Behavior:** sets formula in top-left cell, then `autoFill` across the range.
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
- **Schema:** `StoredConventions` — all fields optional. Omitted = hardcoded default.
- **Configurable fields:**
  - `currency_symbol` — default `$`
  - `negative_style` — `parens` (default) or `minus`
  - `zero_style` — `dash` (default), `zero`, or `blank`
  - `thousands_separator` — `true` (default)
  - `accounting_padding` — `true` (default)
  - `number_dp`, `currency_dp`, `percent_dp`, `ratio_dp` — default dp per preset
- **Resolution:** stored overrides merge over `DEFAULT_CONVENTIONS` / `DEFAULT_CURRENCY_SYMBOL` / `PRESET_DEFAULT_DP`. `format_cells` loads resolved config each call.
- **System prompt:** non-default values injected as "Active convention overrides" section.
- **Execution policy:** classified as read/none (mutates local config, not workbook).
- **Validation:** stored values are validated on read (invalid values silently dropped). dp constrained to integer 0–10.
- **Rationale:** users shouldn't have to repeat "I use pounds" or "no parentheses" every session. Structured config is cleaner than free-text instructions for this.

## Instructions tool (`instructions`)
- **Scopes:** `user` (global, local machine) and `workbook` (scoped by workbook identity hash).
- **Actions:** `append` and `replace`.
- **Storage:** `SettingsStore` keys:
  - `user.instructions`
  - `workbook.instructions.v1.<workbookId>`
- **Execution policy:** classified as read/none for workbook coordinator purposes (it mutates prompt metadata, not workbook cells/structure).
- **Rationale:** AGENTS.md-style persistent guidance without creating a separate workbook mutation path.

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
- **UI:** dependency trees are rendered as structured HTML with clickable cell refs, code-styled formulas, and thread-style left-border indentation.
- **Agent text:** unchanged — still the ASCII tree with `├──`/`└──`/`│` connectors.
- **Implementation:** `TraceDependenciesDetails` passes the `DepNodeDetail` tree to the UI. `src/ui/render-dep-tree.ts` renders the visual tree.
- **Rationale:** ASCII art rendered via `<markdown-block>` lacked interactivity and visual hierarchy. Clickable addresses + clean CSS indentation is much more usable.

## Experimental tmux bridge tool (`tmux`)
- **Availability:** non-core experimental tool, always registered via `createAllTools()`; execution is gated by `applyExperimentalToolGates()`.
- **Gate model:** requires `tmux-bridge` experiment enabled, configured `tmux.bridge.url`, and successful bridge `/health` probe.
- **Execution policy:** classified as `read/none` in workbook coordinator (no workbook lock writes or blueprint invalidation).
- **Bridge implementation:** local helper script `scripts/tmux-bridge-server.mjs`.
  - default mode: `stub` (in-memory simulator)
  - real mode: `TMUX_BRIDGE_MODE=tmux` (subprocess-backed tmux bridge)
- **Bridge contract:** POST JSON to `https://localhost:<port>/v1/tmux` with actions:
  - `list_sessions`
  - `create_session`
  - `send_keys`
  - `capture_pane`
  - `send_and_capture`
  - `kill_session`
- **Security posture:** local opt-in only; bridge URL validated via `validateOfficeProxyUrl`; tool execution re-checks gate before every call; bridge enforces loopback+origin checks and optional bearer token (`TMUX_BRIDGE_TOKEN` / setting `tmux.bridge.token`, managed via `/experimental tmux-bridge-token ...`).
- **Diagnostics UX:** `/experimental tmux-status` reports feature flag, URL/token config, gate result, and bridge health details for quick troubleshooting.
- **Rationale:** stable local adapter contract now (issue #3) with safe stub-first rollout and incremental hardening.

## Experimental Python / LibreOffice bridge tools (`python_run`, `libreoffice_convert`, `python_transform_range`)
- **Availability:** non-core experimental tools, always registered via `createAllTools()`; execution is gated by `applyExperimentalToolGates()`.
- **Gate model:** requires `python-bridge` experiment enabled, configured `python.bridge.url`, successful bridge `/health` probe, and user confirmation once per configured bridge URL.
- **Execution policy:**
  - `python_run` + `libreoffice_convert` → `read/none` (no direct workbook mutation)
  - `python_transform_range` → `mutate/content` (writes transformed values into workbook)
- **Bridge implementation:** local helper script `scripts/python-bridge-server.mjs`.
  - default mode: `stub` (deterministic simulated responses)
  - real mode: `PYTHON_BRIDGE_MODE=real` (local subprocess execution)
- **Bridge contract:**
  - `POST /v1/python-run` — execute Python snippet with optional `input_json`, return stdout/stderr/result JSON
  - `POST /v1/libreoffice-convert` — convert files across `csv|pdf|xlsx`
- **Security posture:** local opt-in only; bridge URL validated via `validateOfficeProxyUrl`; tool execution re-checks gate before every call; bridge enforces loopback+origin checks and optional bearer token (`PYTHON_BRIDGE_TOKEN` / setting `python.bridge.token`, managed via `/experimental python-bridge-token ...`).
- **Rationale:** unblock heavier offline analysis/conversion workflows for issue #25 while keeping workbook writes explicit/auditable and adding an approval checkpoint for local execution.

## External tools skills (`web_search`, `mcp`)
- **Packaging:** exposed as opt-in **skills** instead of always-on core tools.
- **Scopes:** skills can be enabled per-**session** and/or per-**workbook**; effective skills are the union (ordered by catalog).
- **Global gate:** `external.tools.enabled` defaults to **off** and blocks all external skill tools until explicitly enabled.
- **Web search provider:** Brave Search (`web_search`) with optional proxy routing and explicit "Sent" attribution in results.
- **MCP integration:** configurable server registry (`mcp.servers.v1`), UI add/remove/test, and a single `mcp` gateway tool for list/search/describe/call flows.
- **Rationale:** satisfy issue #24 with explicit consent, clear attribution, and minimal overlap with the extension system.

## Experimental files workspace tool (`files`)
- **Availability:** non-core experimental tool, always registered; execution hard-gated by `files-workspace` flag.
- **Backend strategy:** native folder handle (when permitted) → OPFS → in-memory fallback.
- **Workbook tagging:** files are **not segregated** by workbook; each file stores an optional workbook tag (`workbookId` + label) based on the active workbook when last written/imported.
- **Audit trail:** workspace keeps a local activity log (list/read/write/delete/rename/import/backend switches) including actor (`assistant`/`user`), source, timestamp, and workbook label when known.
- **Preview UX:** Files dialog supports inline text editing plus image/PDF preview; other binaries fall back to metadata + download.
- **Filter UX:** Files dialog includes workbook-tag filtering (`all`, `current workbook`, `untagged`, and per-tag options) without changing underlying shared storage.
- **Input drop UX:** dropping files onto the chat input imports them directly into workspace (and auto-enables `files-workspace` if needed).
- **Rationale:** keep one shared artifact space while preserving workbook context and transparency on who accessed/changed files.

## Workbook mutation change previews + audit log (slice)
- **Scope in this slice:** `write_cells`, `fill_formula`, and `python_transform_range` now compute before/after cell diffs.
- **Structured details:** each tool returns `changes` metadata (`changedCount` + sampled cell-level before/after, including formula deltas) for tool-card rendering.
- **UI rendering:** tool cards include a dedicated **Changes** section with clickable cell addresses.
- **Persistence:** successful/blocked mutations in the scoped tools are appended to local `workbook.change-audit.v1` for future export/history UX.
- **Rationale:** improve user trust with concrete, navigable deltas while keeping implementation incremental and low-risk.
