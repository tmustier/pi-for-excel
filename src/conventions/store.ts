function isConventionsStorePayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Persistent conventions storage + resolution.
 */

import {
  DEFAULT_COLOR_CONVENTIONS,
  DEFAULT_HEADER_STYLE,
  DEFAULT_PRESET_FORMATS,
  DEFAULT_VISUAL_DEFAULTS,
} from "./defaults.js";
import type {
  FormatBuilderParams,
  NumberPreset,
  ResolvedConventions,
  StoredColorConventions,
  StoredConventions,
  StoredCustomPreset,
  StoredFormatPreset,
  StoredHeaderStyle,
  StoredVisualDefaults,
} from "./types.js";

const CONVENTIONS_KEY = "conventions.v1";

export interface ConventionsStore {
  get: (key: string) => Promise<DynamicValue>;
  set: (key: string, value: DynamicValue) => Promise<void>;
}

export interface ConventionDiff {
  field: string;
  label: string;
  value: string;
}

const PRESET_NAMES: NumberPreset[] = [
  "number",
  "integer",
  "currency",
  "percent",
  "ratio",
  "text",
];

function cloneBuilderParams(value: FormatBuilderParams | undefined): FormatBuilderParams | undefined {
  if (!value) {
    return undefined;
  }

  const cloned: FormatBuilderParams = {};
  if (value.dp !== undefined) cloned.dp = value.dp;
  if (value.negativeStyle !== undefined) cloned.negativeStyle = value.negativeStyle;
  if (value.zeroStyle !== undefined) cloned.zeroStyle = value.zeroStyle;
  if (value.thousandsSeparator !== undefined) cloned.thousandsSeparator = value.thousandsSeparator;
  if (value.currencySymbol !== undefined) cloned.currencySymbol = value.currencySymbol;

  return cloned;
}

function cloneFormatPreset(value: StoredFormatPreset): StoredFormatPreset {
  const cloned: StoredFormatPreset = {
    format: value.format,
  };

  const builderParams = cloneBuilderParams(value.builderParams);
  if (builderParams !== undefined) {
    cloned.builderParams = builderParams;
  }

  return cloned;
}

function cloneCustomPreset(value: StoredCustomPreset): StoredCustomPreset {
  const cloned: StoredCustomPreset = {
    format: value.format,
  };

  if (value.description !== undefined) {
    cloned.description = value.description;
  }

  const builderParams = cloneBuilderParams(value.builderParams);
  if (builderParams !== undefined) {
    cloned.builderParams = builderParams;
  }

  return cloned;
}

function normalizeNumberInRange(value: DynamicValue, min: number, max: number): number | null {
  if (typeof value !== "number" || Number.isNaN(value) || !Number.isFinite(value)) {
    return null;
  }

  if (value < min || value > max) {
    return null;
  }

  return value;
}

function normalizeIntegerInRange(value: DynamicValue, min: number, max: number): number | null {
  const n = normalizeNumberInRange(value, min, max);
  if (n === null || !Number.isInteger(n)) {
    return null;
  }

  return n;
}

function normalizeHexColor(value: string): string | null {
  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const shortMatch = /^#([0-9a-fA-F]{3})$/u.exec(trimmed);
  const shortHex = shortMatch?.[1];
  if (shortHex) {
    const [r, g, b] = shortHex.split("");
    if (r === undefined || g === undefined || b === undefined) {
      return null;
    }
    return `#${r}${r}${g}${g}${b}${b}`.toUpperCase();
  }

  const fullMatch = /^#([0-9a-fA-F]{6})$/u.exec(trimmed);
  const fullHex = fullMatch?.[1];
  if (fullHex) {
    return `#${fullHex.toUpperCase()}`;
  }

  return null;
}

function normalizeRgbColor(value: string): string | null {
  const match = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/u.exec(value.trim());
  if (!match) {
    return null;
  }

  const r = Number(match[1]);
  const g = Number(match[2]);
  const b = Number(match[3]);
  const channels = [r, g, b];
  if (channels.some((channel) => Number.isNaN(channel) || channel < 0 || channel > 255)) {
    return null;
  }

  const toHex = (channel: number) => channel.toString(16).padStart(2, "0").toUpperCase();
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

export function normalizeConventionColor(value: DynamicValue): string | null {
  if (typeof value !== "string") {
    return null;
  }

  return normalizeHexColor(value) ?? normalizeRgbColor(value);
}

function normalizeBuilderParams(raw: DynamicValue): FormatBuilderParams | undefined {
  if (!isConventionsStorePayloadShape(raw)) {
    return undefined;
  }

  const result: FormatBuilderParams = {};

  const dp = normalizeIntegerInRange(raw.dp, 0, 10);
  if (dp !== null) {
    result.dp = dp;
  }

  if (raw.negativeStyle === "parens" || raw.negativeStyle === "minus") {
    result.negativeStyle = raw.negativeStyle;
  }

  if (raw.zeroStyle === "dash" || raw.zeroStyle === "single-dash" || raw.zeroStyle === "zero" || raw.zeroStyle === "blank") {
    result.zeroStyle = raw.zeroStyle;
  }

  if (typeof raw.thousandsSeparator === "boolean") {
    result.thousandsSeparator = raw.thousandsSeparator;
  }

  if (typeof raw.currencySymbol === "string" && raw.currencySymbol.trim().length > 0) {
    result.currencySymbol = raw.currencySymbol.trim();
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

function normalizeStoredFormatPreset(raw: DynamicValue): StoredFormatPreset | null {
  if (!isConventionsStorePayloadShape(raw)) {
    return null;
  }

  const format = typeof raw.format === "string" ? raw.format.trim() : "";
  if (format.length === 0) {
    return null;
  }

  const preset: StoredFormatPreset = { format };
  const builderParams = normalizeBuilderParams(raw.builderParams);
  if (builderParams !== undefined) {
    preset.builderParams = builderParams;
  }

  return preset;
}

function normalizeStoredCustomPreset(raw: DynamicValue): StoredCustomPreset | null {
  const base = normalizeStoredFormatPreset(raw);
  if (!base || !isConventionsStorePayloadShape(raw)) {
    return null;
  }

  const description = typeof raw.description === "string"
    ? raw.description.trim()
    : undefined;

  const preset: StoredCustomPreset = { ...base };
  if (description && description.length > 0) {
    preset.description = description;
  }

  return preset;
}

function normalizePresetFormats(raw: DynamicValue): Partial<Record<NumberPreset, StoredFormatPreset>> | undefined {
  if (!isConventionsStorePayloadShape(raw)) {
    return undefined;
  }

  const result: Partial<Record<NumberPreset, StoredFormatPreset>> = {};

  for (const preset of PRESET_NAMES) {
    const normalized = normalizeStoredFormatPreset(raw[preset]);
    if (normalized) {
      result[preset] = normalized;
    }
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

function normalizeCustomPresetName(name: string): string | null {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    return null;
  }

  return trimmed;
}

function normalizeCustomPresets(raw: DynamicValue): Record<string, StoredCustomPreset> | undefined {
  if (!isConventionsStorePayloadShape(raw)) {
    return undefined;
  }

  const result: Record<string, StoredCustomPreset> = {};

  for (const [name, value] of Object.entries(raw)) {
    const normalizedName = normalizeCustomPresetName(name);
    if (!normalizedName) {
      continue;
    }

    const normalizedPreset = normalizeStoredCustomPreset(value);
    if (!normalizedPreset) {
      continue;
    }

    result[normalizedName] = normalizedPreset;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

function normalizeVisualDefaults(raw: DynamicValue): StoredVisualDefaults | undefined {
  if (!isConventionsStorePayloadShape(raw)) {
    return undefined;
  }

  const result: StoredVisualDefaults = {};

  if (typeof raw.fontName === "string" && raw.fontName.trim().length > 0) {
    result.fontName = raw.fontName.trim();
  }

  const fontSize = normalizeNumberInRange(raw.fontSize, 6, 72);
  if (fontSize !== null) {
    result.fontSize = fontSize;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

function normalizeColorConventions(raw: DynamicValue): StoredColorConventions | undefined {
  if (!isConventionsStorePayloadShape(raw)) {
    return undefined;
  }

  const result: StoredColorConventions = {};

  const hardcoded = normalizeConventionColor(raw.hardcodedValueColor);
  if (hardcoded) {
    result.hardcodedValueColor = hardcoded;
  }

  const crossSheet = normalizeConventionColor(raw.crossSheetLinkColor);
  if (crossSheet) {
    result.crossSheetLinkColor = crossSheet;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

function normalizeHeaderStyle(raw: DynamicValue): StoredHeaderStyle | undefined {
  if (!isConventionsStorePayloadShape(raw)) {
    return undefined;
  }

  const result: StoredHeaderStyle = {};

  const fillColor = normalizeConventionColor(raw.fillColor);
  if (fillColor) {
    result.fillColor = fillColor;
  }

  const fontColor = normalizeConventionColor(raw.fontColor);
  if (fontColor) {
    result.fontColor = fontColor;
  }

  if (typeof raw.bold === "boolean") {
    result.bold = raw.bold;
  }

  if (typeof raw.wrapText === "boolean") {
    result.wrapText = raw.wrapText;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

function validateStoredConventions(raw: DynamicValue): StoredConventions {
  if (!isConventionsStorePayloadShape(raw)) {
    return {};
  }

  const result: StoredConventions = {};

  const presetFormats = normalizePresetFormats(raw.presetFormats);
  if (presetFormats) {
    result.presetFormats = presetFormats;
  }

  const customPresets = normalizeCustomPresets(raw.customPresets);
  if (customPresets) {
    result.customPresets = customPresets;
  }

  const visualDefaults = normalizeVisualDefaults(raw.visualDefaults);
  if (visualDefaults) {
    result.visualDefaults = visualDefaults;
  }

  const colorConventions = normalizeColorConventions(raw.colorConventions);
  if (colorConventions) {
    result.colorConventions = colorConventions;
  }

  const headerStyle = normalizeHeaderStyle(raw.headerStyle);
  if (headerStyle) {
    result.headerStyle = headerStyle;
  }

  return result;
}

export async function getStoredConventions(store: ConventionsStore): Promise<StoredConventions> {
  const raw = await store.get(CONVENTIONS_KEY);
  return validateStoredConventions(raw);
}

export async function setStoredConventions(
  store: ConventionsStore,
  value: StoredConventions,
): Promise<void> {
  await store.set(CONVENTIONS_KEY, validateStoredConventions(value));
}

export function resolveConventions(stored: StoredConventions): ResolvedConventions {
  const normalized = validateStoredConventions(stored);

  const presetFormats: Record<NumberPreset, StoredFormatPreset> = {
    number: cloneFormatPreset(DEFAULT_PRESET_FORMATS.number),
    integer: cloneFormatPreset(DEFAULT_PRESET_FORMATS.integer),
    currency: cloneFormatPreset(DEFAULT_PRESET_FORMATS.currency),
    percent: cloneFormatPreset(DEFAULT_PRESET_FORMATS.percent),
    ratio: cloneFormatPreset(DEFAULT_PRESET_FORMATS.ratio),
    text: cloneFormatPreset(DEFAULT_PRESET_FORMATS.text),
  };

  for (const preset of PRESET_NAMES) {
    const override = normalized.presetFormats?.[preset];
    if (override) {
      presetFormats[preset] = cloneFormatPreset(override);
    }
  }

  const customPresets: Record<string, StoredCustomPreset> = {};
  for (const [name, preset] of Object.entries(normalized.customPresets ?? {})) {
    customPresets[name] = cloneCustomPreset(preset);
  }

  return {
    presetFormats,
    customPresets,
    visualDefaults: {
      fontName: normalized.visualDefaults?.fontName ?? DEFAULT_VISUAL_DEFAULTS.fontName,
      fontSize: normalized.visualDefaults?.fontSize ?? DEFAULT_VISUAL_DEFAULTS.fontSize,
    },
    colorConventions: {
      hardcodedValueColor: normalized.colorConventions?.hardcodedValueColor
        ?? DEFAULT_COLOR_CONVENTIONS.hardcodedValueColor,
      crossSheetLinkColor: normalized.colorConventions?.crossSheetLinkColor
        ?? DEFAULT_COLOR_CONVENTIONS.crossSheetLinkColor,
    },
    headerStyle: {
      fillColor: normalized.headerStyle?.fillColor ?? DEFAULT_HEADER_STYLE.fillColor,
      fontColor: normalized.headerStyle?.fontColor ?? DEFAULT_HEADER_STYLE.fontColor,
      bold: normalized.headerStyle?.bold ?? DEFAULT_HEADER_STYLE.bold,
      wrapText: normalized.headerStyle?.wrapText ?? DEFAULT_HEADER_STYLE.wrapText,
    },
  };
}

export async function getResolvedConventions(
  store: ConventionsStore,
): Promise<ResolvedConventions> {
  const stored = await getStoredConventions(store);
  return resolveConventions(stored);
}

function mergeSection<T extends object>(current: T | undefined, updates: T | undefined): T | undefined {
  if (!updates) {
    return current;
  }

  return {
    ...(current ?? {}),
    ...updates,
  };
}

export function mergeStoredConventions(
  current: StoredConventions,
  updates: StoredConventions,
): StoredConventions {
  const normalizedCurrent = validateStoredConventions(current);
  const normalizedUpdates = validateStoredConventions(updates);

  const merged: StoredConventions = {};
  const presetFormats = mergeSection(normalizedCurrent.presetFormats, normalizedUpdates.presetFormats);
  const customPresets = mergeSection(normalizedCurrent.customPresets, normalizedUpdates.customPresets);
  const visualDefaults = mergeSection(normalizedCurrent.visualDefaults, normalizedUpdates.visualDefaults);
  const colorConventions = mergeSection(normalizedCurrent.colorConventions, normalizedUpdates.colorConventions);
  const headerStyle = mergeSection(normalizedCurrent.headerStyle, normalizedUpdates.headerStyle);

  if (presetFormats !== undefined) merged.presetFormats = presetFormats;
  if (customPresets !== undefined) merged.customPresets = customPresets;
  if (visualDefaults !== undefined) merged.visualDefaults = visualDefaults;
  if (colorConventions !== undefined) merged.colorConventions = colorConventions;
  if (headerStyle !== undefined) merged.headerStyle = headerStyle;

  return validateStoredConventions(merged);
}

export function removeCustomPresets(
  current: StoredConventions,
  presetNames: readonly string[],
): StoredConventions {
  const normalizedCurrent = validateStoredConventions(current);
  const custom = { ...(normalizedCurrent.customPresets ?? {}) };

  for (const name of presetNames) {
    const normalized = normalizeCustomPresetName(name);
    if (!normalized) {
      continue;
    }

    delete custom[normalized];
  }

  const next: StoredConventions = { ...normalizedCurrent };
  if (Object.keys(custom).length > 0) {
    next.customPresets = custom;
  } else {
    delete next.customPresets;
  }

  return validateStoredConventions(next);
}

function formatBoolean(value: boolean): string {
  return value ? "yes" : "no";
}

function formatPresetDiffLabel(preset: string): string {
  return `${preset} format`;
}

export function diffFromDefaults(resolved: ResolvedConventions): ConventionDiff[] {
  const diffs: ConventionDiff[] = [];

  for (const preset of PRESET_NAMES) {
    const current = resolved.presetFormats[preset]?.format;
    const fallback = DEFAULT_PRESET_FORMATS[preset].format;
    if (current !== fallback) {
      diffs.push({
        field: `presetFormats.${preset}`,
        label: formatPresetDiffLabel(preset),
        value: current,
      });
    }
  }

  for (const [name, preset] of Object.entries(resolved.customPresets)) {
    const suffix = preset.description ? ` — ${preset.description}` : "";
    diffs.push({
      field: `customPresets.${name}`,
      label: `custom preset ${name}`,
      value: `${preset.format}${suffix}`,
    });
  }

  if (resolved.visualDefaults.fontName !== DEFAULT_VISUAL_DEFAULTS.fontName) {
    diffs.push({
      field: "visualDefaults.fontName",
      label: "Default font",
      value: resolved.visualDefaults.fontName,
    });
  }

  if (resolved.visualDefaults.fontSize !== DEFAULT_VISUAL_DEFAULTS.fontSize) {
    diffs.push({
      field: "visualDefaults.fontSize",
      label: "Default font size",
      value: `${resolved.visualDefaults.fontSize}`,
    });
  }

  if (resolved.colorConventions.hardcodedValueColor !== DEFAULT_COLOR_CONVENTIONS.hardcodedValueColor) {
    diffs.push({
      field: "colorConventions.hardcodedValueColor",
      label: "Hardcoded value font color",
      value: resolved.colorConventions.hardcodedValueColor,
    });
  }

  if (resolved.colorConventions.crossSheetLinkColor !== DEFAULT_COLOR_CONVENTIONS.crossSheetLinkColor) {
    diffs.push({
      field: "colorConventions.crossSheetLinkColor",
      label: "Cross-sheet link font color",
      value: resolved.colorConventions.crossSheetLinkColor,
    });
  }

  if (resolved.headerStyle.fillColor !== DEFAULT_HEADER_STYLE.fillColor) {
    diffs.push({
      field: "headerStyle.fillColor",
      label: "Header fill",
      value: resolved.headerStyle.fillColor,
    });
  }

  if (resolved.headerStyle.fontColor !== DEFAULT_HEADER_STYLE.fontColor) {
    diffs.push({
      field: "headerStyle.fontColor",
      label: "Header font color",
      value: resolved.headerStyle.fontColor,
    });
  }

  if (resolved.headerStyle.bold !== DEFAULT_HEADER_STYLE.bold) {
    diffs.push({
      field: "headerStyle.bold",
      label: "Header bold",
      value: formatBoolean(resolved.headerStyle.bold),
    });
  }

  if (resolved.headerStyle.wrapText !== DEFAULT_HEADER_STYLE.wrapText) {
    diffs.push({
      field: "headerStyle.wrapText",
      label: "Header wrap text",
      value: formatBoolean(resolved.headerStyle.wrapText),
    });
  }

  return diffs;
}

export function normalizePresetName(name: string): string {
  return name.trim();
}

export function isBuiltinPresetName(value: string): value is NumberPreset {
  return value === "number"
    || value === "integer"
    || value === "currency"
    || value === "percent"
    || value === "ratio"
    || value === "text";
}

export function isPresetName(value: string, resolved: ResolvedConventions): boolean {
  if (isBuiltinPresetName(value)) {
    return true;
  }

  return Object.prototype.hasOwnProperty.call(resolved.customPresets, value);
}

export function getPresetFormat(
  resolved: ResolvedConventions,
  presetName: string,
): StoredFormatPreset | StoredCustomPreset | null {
  if (isBuiltinPresetName(presetName)) {
    return resolved.presetFormats[presetName];
  }

  return resolved.customPresets[presetName] ?? null;
}
