/**
 * style-resolver — compose named styles + overrides into a flat ResolvedCellStyle.
 *
 * Resolution order:
 *   1. For each style name (left → right), merge its properties.
 *   2. Apply individual param overrides (inline > class, like CSS).
 *   3. Build the final Excel format string from preset + dp + symbol.
 */

import type { CellStyle, ResolvedCellStyle, ResolvedConventions } from "./types.js";
import { BUILTIN_STYLES, DEFAULT_CONVENTION_CONFIG } from "./defaults.js";
import { buildFormatString, isPresetName } from "./format-builder.js";

/**
 * Resolve an array of style names + individual overrides into a flat CellStyle
 * with a ready-to-use Excel format string.
 *
 * @param styles    - Style name or array of style names (left-to-right composition).
 * @param overrides - Individual param overrides (always win).
 * @param config    - Resolved conventions (defaults merged with stored overrides).
 */
export function resolveStyles(
  styles: string | string[] | undefined,
  overrides?: Partial<CellStyle>,
  config: ResolvedConventions = DEFAULT_CONVENTION_CONFIG,
): ResolvedCellStyle {
  const warnings: string[] = [];
  const merged: CellStyle = {};

  // ── 1. Compose named styles (left → right) ────────────────────────
  if (styles) {
    const names = Array.isArray(styles) ? styles : [styles];
    for (const name of names) {
      const style = BUILTIN_STYLES.get(name);
      if (!style) {
        warnings.push(`Unknown style "${name}" — ignored.`);
        continue;
      }
      mergeCellStyle(merged, style.properties);
    }
  }

  // ── 2. Apply individual overrides ──────────────────────────────────
  if (overrides) {
    mergeCellStyle(merged, overrides);
  }

  // ── 3. Resolve number format ───────────────────────────────────────
  let excelNumberFormat: string | undefined;

  if (merged.numberFormat) {
    if (isPresetName(merged.numberFormat)) {
      // It's a preset name — build the Excel format string.
      // Resolve default dp + currency from config before calling builder.
      const effectiveDp = merged.numberFormatDp ?? config.presetDp[merged.numberFormat] ?? null;
      const effectiveSymbol = merged.currencySymbol ??
        (merged.numberFormat === "currency" ? config.currencySymbol : undefined);
      const result = buildFormatString(
        merged.numberFormat,
        effectiveDp,
        effectiveSymbol,
        config.conventions,
      );
      excelNumberFormat = result.format;
      warnings.push(...result.warnings);
    } else {
      // It's a raw Excel format string — pass through
      excelNumberFormat = merged.numberFormat;

      // Warn if preset-only params were set alongside a raw string
      if (merged.numberFormatDp != null) {
        warnings.push("number_format_dp ignored — only applies to preset names, not raw format strings.");
      }
      if (merged.currencySymbol) {
        warnings.push("currency_symbol ignored — only applies to currency preset, not raw format strings.");
      }
    }
  } else {
    // No number format specified, but check for orphaned params
    if (merged.numberFormatDp != null) {
      warnings.push("number_format_dp ignored — no number format or style specified.");
    }
    if (merged.currencySymbol) {
      warnings.push("currency_symbol ignored — no currency format or style specified.");
    }
  }

  return { properties: merged, excelNumberFormat, warnings };
}

// ── Helpers ──────────────────────────────────────────────────────────

/** Shallow-merge `source` into `target`, only overriding defined values. */
function mergeCellStyle(target: CellStyle, source: Partial<CellStyle>): void {
  for (const key of Object.keys(source) as Array<keyof CellStyle>) {
    const value = source[key];
    if (value !== undefined) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any -- generic merge over union-typed properties
      (target as Record<string, any>)[key] = value;
    }
  }
}
