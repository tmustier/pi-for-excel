import type { NumberFormatConventions, NumberPreset } from "./types.js";
import {
  DEFAULT_CURRENCY_SYMBOL,
  DEFAULT_FORMAT_CONVENTIONS,
  PRESET_DEFAULT_DP,
} from "./format-defaults.js";

export interface FormatBuildResult {
  format: string;
  warnings: string[];
}

export function buildFormatString(
  preset: NumberPreset,
  dp?: number | null,
  symbol?: string | null,
  conventions: NumberFormatConventions = DEFAULT_FORMAT_CONVENTIONS,
): FormatBuildResult {
  const warnings: string[] = [];

  if (preset === "text") {
    if (dp != null) {
      warnings.push("number_format_dp ignored — not applicable to text preset.");
    }
    if (symbol) {
      warnings.push("currency_symbol ignored — only applies to currency preset.");
    }
    return { format: "@", warnings };
  }

  if (preset === "integer" && dp != null && dp > 0) {
    warnings.push(
      `integer preset has 0dp by definition; use "number" with dp:${dp} instead.`,
    );
  }
  if (symbol && preset !== "currency") {
    warnings.push("currency_symbol ignored — only applies to currency preset.");
  }

  const defaultDp = PRESET_DEFAULT_DP[preset];
  const resolvedDp = dp ?? defaultDp ?? 0;

  const intPart = conventions.thousandsSeparator ? "#,##0" : "0";
  const decPart = resolvedDp > 0 ? "." + "0".repeat(resolvedDp) : "";
  const core = intPart + decPart;

  const suffix = preset === "percent" ? "%" : preset === "ratio" ? "x" : "";

  const padChars = conventions.accountingPadding
    ? "_)"
    : "";
  const suffixPad = suffix
    ? `_${suffix}`
    : "";

  const currSymbol = preset === "currency" ? (symbol ?? DEFAULT_CURRENCY_SYMBOL) : "";
  const currPrefix = currSymbol ? `${currSymbol}* ` : "";

  const positive = `${currPrefix}${core}${suffix}${padChars}`;

  let negative: string;
  if (conventions.negativeStyle === "parens") {
    negative = `${currPrefix}(${core}${suffix})`;
  } else {
    negative = `${currPrefix}-${core}${suffix}${padChars}`;
  }

  let zero: string;
  if (conventions.zeroStyle === "dash") {
    zero = `${currPrefix}--${suffixPad}${padChars}`;
  } else if (conventions.zeroStyle === "single-dash") {
    zero = `${currPrefix}-${suffixPad}${padChars}`;
  } else if (conventions.zeroStyle === "blank") {
    zero = "";
  } else {
    zero = positive;
  }

  const format = `${positive};${negative};${zero}`;
  return { format, warnings };
}

export function isPresetName(value: string): value is NumberPreset {
  return (
    value === "number" ||
    value === "integer" ||
    value === "currency" ||
    value === "percent" ||
    value === "ratio" ||
    value === "text"
  );
}
