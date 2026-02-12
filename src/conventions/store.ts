/**
 * Persistent conventions storage.
 *
 * Reads/writes user-level formatting conventions from SettingsStore.
 * All fields are optional — omitted values fall back to hardcoded defaults.
 */

import type { NumberPreset, StoredConventions, ResolvedConventions } from "./types.js";
import { DEFAULT_CONVENTIONS, DEFAULT_CURRENCY_SYMBOL, PRESET_DEFAULT_DP } from "./defaults.js";

const CONVENTIONS_KEY = "conventions.v1";

/** Presets that support dp overrides (excludes "text"). */
const DP_OVERRIDABLE_PRESETS: readonly NumberPreset[] = [
  "number", "integer", "currency", "percent", "ratio",
];

export interface ConventionsStore {
  get: (key: string) => Promise<unknown>;
  set: (key: string, value: unknown) => Promise<void>;
}

// ── Read / write ─────────────────────────────────────────────────────

export async function getStoredConventions(store: ConventionsStore): Promise<StoredConventions> {
  const raw = await store.get(CONVENTIONS_KEY);
  if (!raw || typeof raw !== "object") return {};
  return validateStoredConventions(raw as Record<string, unknown>);
}

export async function setStoredConventions(
  store: ConventionsStore,
  value: StoredConventions,
): Promise<void> {
  await store.set(CONVENTIONS_KEY, value);
}

// ── Resolve (merge stored over defaults) ─────────────────────────────

export function resolveConventions(stored: StoredConventions): ResolvedConventions {
  const conventions = {
    negativeStyle: stored.negativeStyle ?? DEFAULT_CONVENTIONS.negativeStyle,
    zeroStyle: stored.zeroStyle ?? DEFAULT_CONVENTIONS.zeroStyle,
    thousandsSeparator: stored.thousandsSeparator ?? DEFAULT_CONVENTIONS.thousandsSeparator,
    accountingPadding: stored.accountingPadding ?? DEFAULT_CONVENTIONS.accountingPadding,
  };

  const currencySymbol = stored.currencySymbol ?? DEFAULT_CURRENCY_SYMBOL;

  const presetDp: Record<NumberPreset, number | null> = { ...PRESET_DEFAULT_DP };
  if (stored.presetDp) {
    for (const key of DP_OVERRIDABLE_PRESETS) {
      const val = stored.presetDp[key];
      if (typeof val === "number") {
        presetDp[key] = val;
      }
    }
  }

  return { conventions, currencySymbol, presetDp };
}

/** Load stored conventions and resolve against defaults in one call. */
export async function getResolvedConventions(
  store: ConventionsStore,
): Promise<ResolvedConventions> {
  const stored = await getStoredConventions(store);
  return resolveConventions(stored);
}

// ── Merge helper (for partial "set" updates) ─────────────────────────

/** Merge partial updates into existing stored conventions. */
export function mergeStoredConventions(
  current: StoredConventions,
  updates: StoredConventions,
): StoredConventions {
  const result: StoredConventions = { ...current };

  if (updates.currencySymbol !== undefined) result.currencySymbol = updates.currencySymbol;
  if (updates.negativeStyle !== undefined) result.negativeStyle = updates.negativeStyle;
  if (updates.zeroStyle !== undefined) result.zeroStyle = updates.zeroStyle;
  if (updates.thousandsSeparator !== undefined) result.thousandsSeparator = updates.thousandsSeparator;
  if (updates.accountingPadding !== undefined) result.accountingPadding = updates.accountingPadding;

  if (updates.presetDp !== undefined) {
    result.presetDp = { ...current.presetDp, ...updates.presetDp };
  }

  return result;
}

// ── Diff helper (for system prompt / UI) ─────────────────────────────

export interface ConventionDiff {
  field: string;
  label: string;
  value: string;
}

/** Return list of fields that differ from hardcoded defaults. */
export function diffFromDefaults(resolved: ResolvedConventions): ConventionDiff[] {
  const diffs: ConventionDiff[] = [];

  if (resolved.currencySymbol !== DEFAULT_CURRENCY_SYMBOL) {
    diffs.push({ field: "currencySymbol", label: "Currency", value: resolved.currencySymbol });
  }
  if (resolved.conventions.negativeStyle !== DEFAULT_CONVENTIONS.negativeStyle) {
    const label = resolved.conventions.negativeStyle === "parens" ? "parentheses" : "minus sign";
    diffs.push({ field: "negativeStyle", label: "Negatives", value: label });
  }
  if (resolved.conventions.zeroStyle !== DEFAULT_CONVENTIONS.zeroStyle) {
    const labels: Record<string, string> = { dash: "dash (--)", zero: "literal 0", blank: "blank" };
    diffs.push({ field: "zeroStyle", label: "Zeros", value: labels[resolved.conventions.zeroStyle] ?? resolved.conventions.zeroStyle });
  }
  if (resolved.conventions.thousandsSeparator !== DEFAULT_CONVENTIONS.thousandsSeparator) {
    diffs.push({ field: "thousandsSeparator", label: "Thousands sep", value: resolved.conventions.thousandsSeparator ? "yes" : "no" });
  }
  if (resolved.conventions.accountingPadding !== DEFAULT_CONVENTIONS.accountingPadding) {
    diffs.push({ field: "accountingPadding", label: "Accounting padding", value: resolved.conventions.accountingPadding ? "yes" : "no" });
  }

  for (const key of DP_OVERRIDABLE_PRESETS) {
    if (resolved.presetDp[key] !== PRESET_DEFAULT_DP[key]) {
      const val = resolved.presetDp[key];
      if (val !== null) {
        diffs.push({ field: `presetDp.${key}`, label: `${key} dp`, value: `${val}` });
      }
    }
  }

  return diffs;
}

// ── Validation ───────────────────────────────────────────────────────

function validateStoredConventions(raw: Record<string, unknown>): StoredConventions {
  const result: StoredConventions = {};

  if (typeof raw.currencySymbol === "string" && raw.currencySymbol.trim().length > 0) {
    result.currencySymbol = raw.currencySymbol.trim();
  }
  if (raw.negativeStyle === "parens" || raw.negativeStyle === "minus") {
    result.negativeStyle = raw.negativeStyle;
  }
  if (raw.zeroStyle === "dash" || raw.zeroStyle === "zero" || raw.zeroStyle === "blank") {
    result.zeroStyle = raw.zeroStyle;
  }
  if (typeof raw.thousandsSeparator === "boolean") {
    result.thousandsSeparator = raw.thousandsSeparator;
  }
  if (typeof raw.accountingPadding === "boolean") {
    result.accountingPadding = raw.accountingPadding;
  }
  if (raw.presetDp && typeof raw.presetDp === "object") {
    const dpRaw = raw.presetDp as Record<string, unknown>;
    const dp: Partial<Record<NumberPreset, number>> = {};
    for (const key of DP_OVERRIDABLE_PRESETS) {
      const val = dpRaw[key];
      if (typeof val === "number" && Number.isInteger(val) && val >= 0 && val <= 10) {
        dp[key] = val;
      }
    }
    if (Object.keys(dp).length > 0) {
      result.presetDp = dp;
    }
  }

  return result;
}
