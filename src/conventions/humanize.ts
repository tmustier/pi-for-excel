import type { NumberPreset, ResolvedConventions } from "./types.js";
import { buildFormatString } from "./format-builder.js";
import { DEFAULT_CURRENCY_SYMBOL } from "./defaults.js";

const FORMAT_TO_LABEL = new Map<string, string>();
const COMMON_CURRENCIES = ["$", "£", "€", "¥", "CHF", "kr", "R", "A$", "C$"];
const DP_RANGE = [0, 1, 2, 3, 4];

function register(format: string, label: string): void {
  FORMAT_TO_LABEL.set(format, label);
}

function init(): void {
  const simplePresets: Array<{ preset: NumberPreset; label: string }> = [
    { preset: "number", label: "number" },
    { preset: "percent", label: "percent" },
    { preset: "ratio", label: "ratio" },
  ];

  for (const { preset, label } of simplePresets) {
    for (const dp of DP_RANGE) {
      const { format } = buildFormatString(preset, dp);
      register(format, `${label} (${dp}dp)`);
    }
  }

  const { format: intFmt } = buildFormatString("integer", 0);
  register(intFmt, "integer");

  for (const sym of COMMON_CURRENCIES) {
    for (const dp of DP_RANGE) {
      const { format } = buildFormatString("currency", dp, sym);
      const symLabel = sym === DEFAULT_CURRENCY_SYMBOL ? "" : ` ${sym},`;
      const dpLabel = `${dp}dp`;
      register(format, `currency (${symLabel}${dpLabel})`.replace("( ", "("));
    }
  }

  register("@", "text");
}

init();

export function buildResolvedFormatLabels(resolved: ResolvedConventions): Map<string, string> {
  const labels = new Map<string, string>();

  for (const [presetName, preset] of Object.entries(resolved.presetFormats)) {
    labels.set(preset.format, String(presetName));
  }

  for (const [presetName, preset] of Object.entries(resolved.customPresets)) {
    const suffix = preset.description ? ` — ${preset.description}` : "";
    labels.set(preset.format, `${presetName}${suffix}`);
  }

  return labels;
}

export function humanizeFormat(excelFormat: string, resolvedLabels?: Map<string, string>): string {
  if (!excelFormat || excelFormat === "General") {
    return excelFormat;
  }

  if (resolvedLabels && resolvedLabels.has(excelFormat)) {
    return resolvedLabels.get(excelFormat) ?? excelFormat;
  }

  return FORMAT_TO_LABEL.get(excelFormat) ?? excelFormat;
}
