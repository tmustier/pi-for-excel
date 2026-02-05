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
- **Rationale:** formatting-only cells are not meaningful “content” and shouldn’t block writes.

## Fill formulas (`fill_formula`)
- **Purpose:** avoid large 2D formula arrays by using Excel AutoFill.
- **Behavior:** sets formula in top-left cell, then `autoFill` across the range.
- **Validation:** uses `validateFormula()` (same as `write_cells`).
- **Overwrite protection:** blocks only if values/formulas exist (same policy as `write_cells`).
- **Rationale:** major productivity win for large formula blocks.

## Selection & change tools
- **`read_selection`:** explicit tool version of auto-injected selection context.
- **`get_recent_changes`:** flushes `ChangeTracker` and returns user edits since last message.
- **Rationale:** makes context features discoverable/useful to the agent when it needs them explicitly.

## Default formatting assumption
- **System prompt:** “Default font for formatting is Arial 10 unless user specifies otherwise.”
- **Rationale:** keeps column width conversions consistent with the chosen baseline.
