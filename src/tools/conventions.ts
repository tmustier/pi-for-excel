/**
 * conventions — read/write persistent formatting conventions.
 *
 * Structured key-value config for number format defaults (currency symbol,
 * negative style, zero display, decimal places, etc.).
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  getStoredConventions,
  setStoredConventions,
  resolveConventions,
  mergeStoredConventions,
  diffFromDefaults,
} from "../conventions/store.js";
import { PRESET_DEFAULT_DP } from "../conventions/defaults.js";
import type { StoredConventions, NumberPreset } from "../conventions/types.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  action: Type.Union([
    Type.Literal("get"),
    Type.Literal("set"),
    Type.Literal("reset"),
  ], {
    description: "get = show current conventions, set = update one or more fields, reset = restore all defaults",
  }),

  // ── Set params (all optional — only provided fields are updated) ───
  currency_symbol: Type.Optional(
    Type.String({ description: 'Default currency symbol, e.g. "£", "€", "$", "CHF".' }),
  ),
  negative_style: Type.Optional(
    Type.Union([Type.Literal("parens"), Type.Literal("minus")], {
      description: "How negatives display: parens = (1,234), minus = -1,234.",
    }),
  ),
  zero_style: Type.Optional(
    Type.Union([Type.Literal("dash"), Type.Literal("zero"), Type.Literal("blank")], {
      description: 'How zeros display: dash = "--", zero = "0", blank = empty.',
    }),
  ),
  thousands_separator: Type.Optional(
    Type.Boolean({ description: "Include thousands separator (comma)." }),
  ),
  accounting_padding: Type.Optional(
    Type.Boolean({ description: "Add trailing accounting padding for alignment." }),
  ),
  number_dp: Type.Optional(
    Type.Number({ description: 'Default decimal places for "number" preset.' }),
  ),
  currency_dp: Type.Optional(
    Type.Number({ description: 'Default decimal places for "currency" preset.' }),
  ),
  percent_dp: Type.Optional(
    Type.Number({ description: 'Default decimal places for "percent" preset.' }),
  ),
  ratio_dp: Type.Optional(
    Type.Number({ description: 'Default decimal places for "ratio" preset.' }),
  ),
});

type Params = Static<typeof schema>;

function emitConventionsUpdatedEvent(): void {
  if (typeof document === "undefined") return;
  document.dispatchEvent(new CustomEvent("pi:conventions-updated"));
  document.dispatchEvent(new CustomEvent("pi:status-update"));
}

/** Build the dp overrides object from flat tool params. */
function extractDpOverrides(params: Params): Partial<Record<NumberPreset, number>> | undefined {
  const dp: Partial<Record<NumberPreset, number>> = {};
  let hasAny = false;

  if (params.number_dp !== undefined) { dp.number = params.number_dp; hasAny = true; }
  if (params.currency_dp !== undefined) { dp.currency = params.currency_dp; hasAny = true; }
  if (params.percent_dp !== undefined) { dp.percent = params.percent_dp; hasAny = true; }
  if (params.ratio_dp !== undefined) { dp.ratio = params.ratio_dp; hasAny = true; }

  return hasAny ? dp : undefined;
}

/** Format the resolved conventions as readable markdown. */
function formatConventions(
  resolved: ReturnType<typeof resolveConventions>,
  stored: StoredConventions,
): string {
  const negLabel = resolved.conventions.negativeStyle === "parens"
    ? "parentheses (1,234)" : "minus sign -1,234";
  const zeroLabels: Record<string, string> = { dash: 'dash "--"', zero: "literal 0", blank: "blank" };
  const zeroLabel = zeroLabels[resolved.conventions.zeroStyle] ?? resolved.conventions.zeroStyle;

  const diffs = diffFromDefaults(resolved);
  const marker = (field: string): string =>
    diffs.some((d) => d.field === field) ? " ★" : "";

  const lines = [
    "**Formatting conventions**" + (diffs.length > 0 ? " (★ = customized)" : ""),
    "",
    `- Currency symbol: ${resolved.currencySymbol}${marker("currencySymbol")}`,
    `- Negatives: ${negLabel}${marker("negativeStyle")}`,
    `- Zeros: ${zeroLabel}${marker("zeroStyle")}`,
    `- Thousands separator: ${resolved.conventions.thousandsSeparator ? "yes" : "no"}${marker("thousandsSeparator")}`,
    `- Accounting padding: ${resolved.conventions.accountingPadding ? "yes" : "no"}${marker("accountingPadding")}`,
    `- Default dp: number ${resolved.presetDp.number ?? 0}${marker("presetDp.number")}, ` +
      `integer ${resolved.presetDp.integer ?? 0}, ` +
      `currency ${resolved.presetDp.currency ?? 0}${marker("presetDp.currency")}, ` +
      `percent ${resolved.presetDp.percent ?? 0}${marker("presetDp.percent")}, ` +
      `ratio ${resolved.presetDp.ratio ?? 0}${marker("presetDp.ratio")}`,
  ];

  // Show stored (non-default) values for transparency
  const storedKeys = Object.keys(stored).filter((k) => {
    if (k === "presetDp") return stored.presetDp && Object.keys(stored.presetDp).length > 0;
    return stored[k as keyof StoredConventions] !== undefined;
  });
  if (storedKeys.length === 0) {
    lines.push("", "_All defaults — nothing customized._");
  }

  return lines.join("\n");
}

/** Format a change summary line, e.g. "currency_symbol: $ → £". */
function formatChange(label: string, oldVal: string, newVal: string): string {
  return oldVal === newVal ? `- ${label}: ${newVal} (unchanged)` : `- ${label}: ${oldVal} → ${newVal}`;
}

export function createConventionsTool(): AgentTool<typeof schema, undefined> {
  return {
    name: "conventions",
    label: "Conventions",
    description:
      "Read or update persistent formatting conventions (currency symbol, negative style, " +
      "zero display, decimal places). Changes apply to all future format_cells calls.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const storage = getAppStorage();
        const settings = storage.settings;

        // ── GET ──────────────────────────────────────────────
        if (params.action === "get") {
          const stored = await getStoredConventions(settings);
          const resolved = resolveConventions(stored);
          return {
            content: [{ type: "text", text: formatConventions(resolved, stored) }],
            details: undefined,
          };
        }

        // ── RESET ────────────────────────────────────────────
        if (params.action === "reset") {
          await setStoredConventions(settings, {});
          emitConventionsUpdatedEvent();
          const resolved = resolveConventions({});
          return {
            content: [{
              type: "text",
              text: "Reset all formatting conventions to defaults.\n\n" +
                formatConventions(resolved, {}),
            }],
            details: undefined,
          };
        }

        // ── SET ──────────────────────────────────────────────
        const currentStored = await getStoredConventions(settings);
        const currentResolved = resolveConventions(currentStored);

        // Build the updates
        const updates: StoredConventions = {};
        if (params.currency_symbol !== undefined) updates.currencySymbol = params.currency_symbol;
        if (params.negative_style !== undefined) updates.negativeStyle = params.negative_style;
        if (params.zero_style !== undefined) updates.zeroStyle = params.zero_style;
        if (params.thousands_separator !== undefined) updates.thousandsSeparator = params.thousands_separator;
        if (params.accounting_padding !== undefined) updates.accountingPadding = params.accounting_padding;
        updates.presetDp = extractDpOverrides(params);

        const hasUpdates = Object.values(updates).some((v) => v !== undefined);
        if (!hasUpdates) {
          return {
            content: [{ type: "text", text: "No changes specified. Use action: \"get\" to view current conventions." }],
            details: undefined,
          };
        }

        const merged = mergeStoredConventions(currentStored, updates);
        await setStoredConventions(settings, merged);
        emitConventionsUpdatedEvent();

        const newResolved = resolveConventions(merged);

        // Build change summary
        const changes: string[] = [];
        if (params.currency_symbol !== undefined) {
          changes.push(formatChange("Currency symbol", currentResolved.currencySymbol, newResolved.currencySymbol));
        }
        if (params.negative_style !== undefined) {
          const old = currentResolved.conventions.negativeStyle === "parens" ? "parentheses" : "minus";
          const neo = newResolved.conventions.negativeStyle === "parens" ? "parentheses" : "minus";
          changes.push(formatChange("Negatives", old, neo));
        }
        if (params.zero_style !== undefined) {
          changes.push(formatChange("Zeros", currentResolved.conventions.zeroStyle, newResolved.conventions.zeroStyle));
        }
        if (params.thousands_separator !== undefined) {
          changes.push(formatChange("Thousands sep", String(currentResolved.conventions.thousandsSeparator), String(newResolved.conventions.thousandsSeparator)));
        }
        if (params.accounting_padding !== undefined) {
          changes.push(formatChange("Accounting padding", String(currentResolved.conventions.accountingPadding), String(newResolved.conventions.accountingPadding)));
        }
        const dpKeys: Array<{ param: keyof Params; preset: NumberPreset; label: string }> = [
          { param: "number_dp", preset: "number", label: "number dp" },
          { param: "currency_dp", preset: "currency", label: "currency dp" },
          { param: "percent_dp", preset: "percent", label: "percent dp" },
          { param: "ratio_dp", preset: "ratio", label: "ratio dp" },
        ];
        for (const { param, preset, label } of dpKeys) {
          if (params[param] !== undefined) {
            changes.push(formatChange(label, String(currentResolved.presetDp[preset] ?? DEFAULT_CONVENTIONS_DP(preset)), String(newResolved.presetDp[preset] ?? 0)));
          }
        }

        return {
          content: [{
            type: "text",
            text: `Updated formatting conventions:\n${changes.join("\n")}\n\n` +
              formatConventions(newResolved, merged),
          }],
          details: undefined,
        };
      } catch (error: unknown) {
        return {
          content: [{ type: "text", text: `Error updating conventions: ${getErrorMessage(error)}` }],
          details: undefined,
        };
      }
    },
  };
}

function DEFAULT_CONVENTIONS_DP(preset: NumberPreset): number {
  return PRESET_DEFAULT_DP[preset] ?? 0;
}
