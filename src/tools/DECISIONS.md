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

## Feature-flagged tmux bridge tool (`tmux`)
- **Availability:** non-core feature-flagged tool, always registered via `createAllTools()`; execution is gated by `applyExperimentalToolGates()`.
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

## Feature-flagged Python / LibreOffice bridge tools (`python_run`, `libreoffice_convert`, `python_transform_range`)
- **Availability:** non-core feature-flagged tools, always registered via `createAllTools()`; execution is gated by `applyExperimentalToolGates()`.
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
- **Overwrite perf guard (`python_transform_range`):** pre-write `values/formulas` reads are skipped for large `allow_overwrite: true` outputs (> `MAX_RECOVERY_CELLS`) since those snapshots would be dropped anyway.
- **Rationale:** unblock heavier offline analysis/conversion workflows for issue #25 while keeping workbook writes explicit/auditable and adding an approval checkpoint for local execution.

## External tool integrations (`web_search`, `fetch_page`, `mcp`)
- **Packaging:** exposed as opt-in **integrations** instead of always-on core tools.
- **Scopes:** integrations can be enabled per-**session** and/or per-**workbook**; effective integrations are the union (ordered by catalog).
- **Global gate:** `external.tools.enabled` defaults to **off** and blocks all external integration tools until explicitly enabled.
- **Web search providers:** Serper.dev (default), Tavily, and Brave Search (`web_search`) with optional proxy routing and explicit "Sent" attribution in results.
- **Page retrieval companion:** `fetch_page` fetches URL content and returns extracted markdown for source grounding workflows (`web_search` → `fetch_page`).
- **MCP integration:** configurable server registry (`mcp.servers.v1`), UI add/remove/test, and a single `mcp` gateway tool for list/search/describe/call flows.
- **Rationale:** satisfy issue #24 with explicit consent, clear attribution, and minimal overlap with the extension system.

## Direct Office.js tool (`execute_office_js`)
- **Availability:** always available (not behind `/experimental`), always registered via `createAllTools()`.
- **Contract:** accepts `code` (async function body receiving `context: Excel.RequestContext`) plus a short user-facing `explanation`.
- **Safety guards:** blocks nested `Excel.run(...)` usage (host already provides context), enforces explanation/code length limits, requires explicit user confirmation on every execution, and fails closed if confirmation UI is unavailable.
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

## Workbook backups (`workbook_history`)
- **Goal:** prefer low-friction workflows over pre-execution approval selectors by making rollback easy and reliable.
- **Execution mode toggle:** `/yolo` switches between:
  - `YOLO` (default): mutate tools run without extra pre-execution confirmations.
  - `Safe`: mutate tools require pre-execution confirmation via runtime gate.
- **Automatic backups:** successful `write_cells`, `fill_formula`, `python_transform_range`, `format_cells`, `conditional_format`, mutating `comments` actions, and supported `modify_structure` actions (`rename_sheet`, `hide_sheet`, `unhide_sheet`, `insert_rows`, `insert_columns`, `add_sheet`, and `duplicate_sheet` when the duplicate has no value data) store pre-mutation snapshots in local `workbook.recovery-snapshots.v1`.
- **Manual full-workbook fallback (`/backup`):** explicit user-triggered full-file captures are stored in Files workspace under `manual-backups/full-workbook/v1/...` and downloaded for restore. This remains separate from automatic per-mutation checkpoints.
- **Safety limits:** backup capture is skipped for very large writes (> `MAX_RECOVERY_CELLS`) to avoid oversized local state.
- **Workbook identity guardrails:** append/list/delete/clear/restore paths are scoped to the active workbook identity; restore rejects identity-less or cross-workbook backups.
- **Save boundary behavior:** backups are intended as "in between saves" recovery points and are cleared after the workbook transitions from dirty → saved.
- **Restore UX:** `workbook_history` can list/restore/delete/clear backups; restores also create an inverse backup (`restore_snapshot`) so users can undo a mistaken restore.
- **Coverage signaling:** `modify_structure` and mutating `view_settings` actions explicitly report when no backup was created.
- **Current `modify_structure` backup behavior:** captures/restores all `modify_structure` actions, including value-preserving checkpoints for destructive deletes (`delete_rows`, `delete_columns`, `delete_sheet`) by storing deleted value/formula data when capture is within recovery size limits.
- **Restore safety gate for structure absence states:** restoring `sheet_absent` / `rows_absent` / `columns_absent` checkpoints remains blocked when target data exists, unless the checkpoint was explicitly generated with data-delete intent during a prior restore inversion (`allowDataDelete`).
- **Current `format_cells` backup scope:** captures/restores core range-format properties (font/fill/number format/alignment/wrap/borders), row/column dimensions (`column_width`, `row_height`, `auto_fit`), and merge state (`merge`/`unmerge`).
- **Current `conditional_format` backup scope:** captures/restores all current rule families (`custom`, `cell_value`, `contains_text`, `top_bottom`, `preset_criteria`, `data_bar`, `color_scale`, `icon_set`) including per-rule applies-to ranges.
- **Quick affordance:** users can restore via `/history` (Backups overlay) or `/revert` (latest backup). The Backups overlay also exposes a **Full backup** button for explicit manual full-workbook capture.
- **Rationale:** addresses #27 by shifting from cumbersome up-front approvals to versioned recovery with explicit user-controlled rollback.
