/**
 * Built-in defaults for conventions + named styles.
 */

import { buildFormatString } from "./format-builder.js";
import {
  DEFAULT_CURRENCY_SYMBOL,
  DEFAULT_FORMAT_CONVENTIONS,
  PRESET_DEFAULT_DP,
} from "./format-defaults.js";
import type {
  NamedStyle,
  NumberPreset,
  ResolvedConventions,
  StoredFormatPreset,
} from "./types.js";

export { DEFAULT_CURRENCY_SYMBOL, PRESET_DEFAULT_DP };

/** Backward-compatible alias. */
export const DEFAULT_CONVENTIONS = DEFAULT_FORMAT_CONVENTIONS;

export const DEFAULT_VISUAL_DEFAULTS = {
  fontName: "Arial",
  fontSize: 10,
} as const;

export const DEFAULT_COLOR_CONVENTIONS = {
  hardcodedValueColor: "#0000FF",
  crossSheetLinkColor: "#008000",
} as const;

export const DEFAULT_HEADER_STYLE = {
  fillColor: "#4472C4",
  fontColor: "#FFFFFF",
  bold: true,
  wrapText: true,
} as const;

function buildDefaultPresetFormat(preset: NumberPreset): StoredFormatPreset {
  const defaultDp = PRESET_DEFAULT_DP[preset];
  const defaultSymbol = preset === "currency" ? DEFAULT_CURRENCY_SYMBOL : undefined;

  const built = buildFormatString(
    preset,
    defaultDp,
    defaultSymbol,
    DEFAULT_FORMAT_CONVENTIONS,
  );

  const presetFormat: StoredFormatPreset = {
    format: built.format,
  };

  if (preset !== "text") {
    presetFormat.builderParams = {
      ...(defaultDp !== null ? { dp: defaultDp } : {}),
      negativeStyle: DEFAULT_FORMAT_CONVENTIONS.negativeStyle,
      zeroStyle: DEFAULT_FORMAT_CONVENTIONS.zeroStyle,
      thousandsSeparator: DEFAULT_FORMAT_CONVENTIONS.thousandsSeparator,
      ...(defaultSymbol !== undefined ? { currencySymbol: defaultSymbol } : {}),
    };
  }

  return presetFormat;
}

export const DEFAULT_PRESET_FORMATS: Record<NumberPreset, StoredFormatPreset> = {
  number: buildDefaultPresetFormat("number"),
  integer: buildDefaultPresetFormat("integer"),
  currency: buildDefaultPresetFormat("currency"),
  percent: buildDefaultPresetFormat("percent"),
  ratio: buildDefaultPresetFormat("ratio"),
  text: buildDefaultPresetFormat("text"),
};

export const DEFAULT_CONVENTION_CONFIG: ResolvedConventions = {
  presetFormats: { ...DEFAULT_PRESET_FORMATS },
  customPresets: {},
  visualDefaults: { ...DEFAULT_VISUAL_DEFAULTS },
  colorConventions: { ...DEFAULT_COLOR_CONVENTIONS },
  headerStyle: { ...DEFAULT_HEADER_STYLE },
};

// ── Named style defaults ─────────────────────────────────────────────

const FORMAT_STYLES: NamedStyle[] = [
  {
    name: "number",
    description: "Standard number preset",
    properties: { numberFormat: "number" },
    builtIn: true,
  },
  {
    name: "integer",
    description: "Whole number preset",
    properties: { numberFormat: "integer" },
    builtIn: true,
  },
  {
    name: "currency",
    description: "Currency preset",
    properties: { numberFormat: "currency" },
    builtIn: true,
  },
  {
    name: "percent",
    description: "Percentage preset",
    properties: { numberFormat: "percent" },
    builtIn: true,
  },
  {
    name: "ratio",
    description: "Ratio / multiple preset",
    properties: { numberFormat: "ratio" },
    builtIn: true,
  },
  {
    name: "text",
    description: "Plain text preset",
    properties: { numberFormat: "text" },
    builtIn: true,
  },
];

function createStructuralStyles(conventions: ResolvedConventions): NamedStyle[] {
  return [
    {
      name: "header",
      description: "Column heading style",
      properties: {
        bold: conventions.headerStyle.bold,
        fillColor: conventions.headerStyle.fillColor,
        fontColor: conventions.headerStyle.fontColor,
        wrapText: conventions.headerStyle.wrapText,
      },
      builtIn: true,
    },
    {
      name: "total-row",
      description: "Total row: bold + top border",
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
      description: "Blank separator area",
      properties: {
        fillColor: "#F2F2F2",
      },
      builtIn: true,
    },
  ];
}

export function getBuiltinStyles(conventions: ResolvedConventions = DEFAULT_CONVENTION_CONFIG): ReadonlyMap<string, NamedStyle> {
  const styles = [...FORMAT_STYLES, ...createStructuralStyles(conventions)];
  return new Map(styles.map((style) => [style.name, style]));
}

export const BUILTIN_STYLES: ReadonlyMap<string, NamedStyle> = getBuiltinStyles(DEFAULT_CONVENTION_CONFIG);

export const BUILTIN_STYLE_NAMES: ReadonlySet<string> = new Set(BUILTIN_STYLES.keys());

export const FORMAT_PRESET_NAMES: ReadonlySet<NumberPreset> = new Set<NumberPreset>([
  "number",
  "integer",
  "currency",
  "percent",
  "ratio",
  "text",
]);
