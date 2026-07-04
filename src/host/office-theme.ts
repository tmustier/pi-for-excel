/** Office theme detection helpers used by OfficeHost. */

import { isRecord } from "../utils/type-guards.js";

interface RgbColor {
  r: number;
  g: number;
  b: number;
}

function parseHexColor(input: string): RgbColor | null {
  const raw = input.trim();
  const normalized = raw.startsWith("#") ? raw.slice(1) : raw;

  if (normalized.length === 3) {
    const r = Number.parseInt(normalized[0].repeat(2), 16);
    const g = Number.parseInt(normalized[1].repeat(2), 16);
    const b = Number.parseInt(normalized[2].repeat(2), 16);
    if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
      return null;
    }

    return { r, g, b };
  }

  if (normalized.length !== 6) {
    return null;
  }

  const r = Number.parseInt(normalized.slice(0, 2), 16);
  const g = Number.parseInt(normalized.slice(2, 4), 16);
  const b = Number.parseInt(normalized.slice(4, 6), 16);

  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
    return null;
  }

  return { r, g, b };
}

function toLinearSrgb(channel: number): number {
  const normalized = channel / 255;
  if (normalized <= 0.04045) {
    return normalized / 12.92;
  }

  return ((normalized + 0.055) / 1.055) ** 2.4;
}

function relativeLuminance(rgb: RgbColor): number {
  return (
    0.2126 * toLinearSrgb(rgb.r)
    + 0.7152 * toLinearSrgb(rgb.g)
    + 0.0722 * toLinearSrgb(rgb.b)
  );
}

function isDarkColor(rgb: RgbColor): boolean {
  return relativeLuminance(rgb) < 0.35;
}

function resolveThemeDarkFromColor(input: unknown): boolean | null {
  if (typeof input !== "string") {
    return null;
  }

  const parsed = parseHexColor(input);
  if (!parsed) {
    return null;
  }

  return isDarkColor(parsed);
}

export function resolveOfficeThemeDark(): boolean | null {
  const officeRoot = Reflect.get(globalThis, "Office");
  if (!isRecord(officeRoot)) {
    return null;
  }

  const context = officeRoot.context;
  if (!isRecord(context)) {
    return null;
  }

  const officeTheme = context.officeTheme;
  if (!isRecord(officeTheme)) {
    return null;
  }

  if (typeof officeTheme.isDarkTheme === "boolean") {
    return officeTheme.isDarkTheme;
  }

  const backgroundCandidates = [
    officeTheme.bodyBackgroundColor,
    officeTheme.controlBackgroundColor,
  ];

  for (const color of backgroundCandidates) {
    const isDark = resolveThemeDarkFromColor(color);
    if (isDark !== null) {
      return isDark;
    }
  }

  const foregroundCandidates = [
    officeTheme.bodyForegroundColor,
    officeTheme.controlForegroundColor,
  ];

  for (const color of foregroundCandidates) {
    const isDark = resolveThemeDarkFromColor(color);
    if (isDark !== null) {
      return !isDark;
    }
  }

  return null;
}
