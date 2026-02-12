/**
 * Built-in defaults — format presets, structural styles, and house-style conventions.
 *
 * This is the single source of truth. Tools, prompt, and read-back all import from here.
 */

import type { NamedStyle, NumberFormatConventions, NumberPreset, ResolvedConventions } from "./types.js";

// ── Default conventions ──────────────────────────────────────────────

export const DEFAULT_CONVENTIONS: NumberFormatConventions = {
  negativeStyle: "parens",
  thousandsSeparator: true,
  zeroStyle: "dash",
  accountingPadding: true,
};

export const DEFAULT_CURRENCY_SYMBOL = "$";

/** Default decimal places per preset. */
export const PRESET_DEFAULT_DP: Record<NumberPreset, number | null> = {
  number: 2,
  integer: 0,
  currency: 2,
  percent: 1,
  ratio: 1,
  text: null,
  // date intentionally absent — dropped from presets
};

/** Default resolved conventions (all hardcoded defaults, no overrides). */
export const DEFAULT_CONVENTION_CONFIG: ResolvedConventions = {
  conventions: DEFAULT_CONVENTIONS,
  currencySymbol: DEFAULT_CURRENCY_SYMBOL,
  presetDp: { ...PRESET_DEFAULT_DP },
};

// ── Format styles (number format only — no visual properties) ────────

const FORMAT_STYLES: NamedStyle[] = [
  {
    name: "number",
    description: "Standard number (2dp, thousands separator)",
    properties: { numberFormat: "number" },
    builtIn: true,
  },
  {
    name: "integer",
    description: "Whole number (0dp, thousands separator)",
    properties: { numberFormat: "integer" },
    builtIn: true,
  },
  {
    name: "currency",
    description: "Currency (2dp, accounting-aligned)",
    properties: { numberFormat: "currency" },
    builtIn: true,
  },
  {
    name: "percent",
    description: "Percentage (1dp)",
    properties: { numberFormat: "percent" },
    builtIn: true,
  },
  {
    name: "ratio",
    description: 'Multiple / ratio with "x" suffix (1dp)',
    properties: { numberFormat: "ratio" },
    builtIn: true,
  },
  {
    name: "text",
    description: "Plain text (no number formatting)",
    properties: { numberFormat: "text" },
    builtIn: true,
  },
];

// ── Structural styles (visual only — no number format) ───────────────

const STRUCTURAL_STYLES: NamedStyle[] = [
  {
    name: "header",
    description: "Column heading: bold, blue fill, white font, wrap",
    properties: {
      bold: true,
      fillColor: "#4472C4",
      fontColor: "#FFFFFF",
      wrapText: true,
    },
    builtIn: true,
  },
  {
    name: "total-row",
    description: "Total row: bold, thin top border",
    properties: {
      bold: true,
      borderTop: "thin",
    },
    builtIn: true,
  },
  {
    name: "subtotal",
    description: "Subtotal row: bold",
    properties: {
      bold: true,
    },
    builtIn: true,
  },
  {
    name: "input",
    description: "User-input cell: yellow fill",
    properties: {
      fillColor: "#FFFD78",
    },
    builtIn: true,
  },
  {
    name: "blank-section",
    description: "Intentionally blank area: light grey fill",
    properties: {
      fillColor: "#F2F2F2",
    },
    builtIn: true,
  },
];

// ── All built-in styles ──────────────────────────────────────────────

export const BUILTIN_STYLES: ReadonlyMap<string, NamedStyle> = new Map(
  [...FORMAT_STYLES, ...STRUCTURAL_STYLES].map((s) => [s.name, s]),
);

/** All built-in style names (for validation). */
export const BUILTIN_STYLE_NAMES: ReadonlySet<string> = new Set(BUILTIN_STYLES.keys());

/** Format preset names (for detecting preset names in number_format param). */
export const FORMAT_PRESET_NAMES: ReadonlySet<NumberPreset> = new Set<NumberPreset>([
  "number",
  "integer",
  "currency",
  "percent",
  "ratio",
  "text",
]);
