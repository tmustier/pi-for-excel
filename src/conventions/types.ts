/**
 * Type definitions for the conventions / composable cell styles system.
 */

// ── Border weight ────────────────────────────────────────────────────

export type BorderWeight = "thin" | "medium" | "thick" | "none";

// ── Number format presets ────────────────────────────────────────────

/** Presets that the format-builder knows how to turn into Excel format strings. */
export type NumberPreset = "number" | "integer" | "currency" | "percent" | "ratio" | "text";

// ── House-style conventions ──────────────────────────────────────────

/** Configurable formatting conventions (defaults can be overridden per-user later). */
export interface NumberFormatConventions {
  /** How negatives are shown: parentheses `(1,234)` or minus `-1,234`. */
  negativeStyle: "parens" | "minus";
  /** How zeros are shown: `--` dash, literal `0`, or blank. */
  zeroStyle: "dash" | "zero" | "blank";
  /** Whether to include thousands separator (`,`). */
  thousandsSeparator: boolean;
  /** Whether to add trailing `_)` accounting padding. */
  accountingPadding: boolean;
}

// ── Cell style properties ────────────────────────────────────────────

/** Any formatting property that can be applied to a cell/range. */
export interface CellStyle {
  // Number format
  numberFormat?: string;
  numberFormatDp?: number;
  currencySymbol?: string;

  // Font
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fontSize?: number;
  fontName?: string;

  // Fill
  fillColor?: string;

  // Borders (individual edges)
  borderTop?: BorderWeight;
  borderBottom?: BorderWeight;
  borderLeft?: BorderWeight;
  borderRight?: BorderWeight;

  // Alignment
  horizontalAlignment?: "Left" | "Center" | "Right" | "General";
  verticalAlignment?: "Top" | "Center" | "Bottom";
  wrapText?: boolean;
}

// ── Named styles ─────────────────────────────────────────────────────

/** A built-in named style (shipped with the app). */
export interface NamedStyle {
  name: string;
  description: string;
  properties: CellStyle;
  builtIn: true;
}

// ── Stored conventions (user-configurable) ───────────────────────────

/**
 * User-configurable convention overrides persisted in SettingsStore.
 * All fields optional — omitted fields fall back to hardcoded defaults.
 */
export interface StoredConventions {
  currencySymbol?: string;
  negativeStyle?: "parens" | "minus";
  zeroStyle?: "dash" | "zero" | "blank";
  thousandsSeparator?: boolean;
  accountingPadding?: boolean;
  /** Override default decimal places per preset. */
  presetDp?: Partial<Record<NumberPreset, number>>;
}

// ── Resolved conventions (defaults + stored merged) ──────────────────

/** Fully resolved conventions — hardcoded defaults merged with stored overrides. */
export interface ResolvedConventions {
  /** Number format conventions (negative style, zero style, etc.). */
  conventions: NumberFormatConventions;
  /** Default currency symbol for the "currency" preset. */
  currencySymbol: string;
  /** Default decimal places per preset (all presets filled). */
  presetDp: Record<NumberPreset, number | null>;
}

// ── Resolved output ──────────────────────────────────────────────────

/** The fully-resolved result of composing styles + overrides. */
export interface ResolvedCellStyle {
  /** Flat style properties ready to apply to a range. */
  properties: CellStyle;
  /**
   * The final Excel format string (built from preset + dp + symbol + conventions).
   * Undefined if no number format was specified.
   */
  excelNumberFormat?: string;
  /** Warnings generated during resolution (type-check issues, unknown styles, etc.). */
  warnings: string[];
}
