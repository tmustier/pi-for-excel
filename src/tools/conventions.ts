/**
 * conventions — read/write persistent formatting conventions.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

import {
  diffFromDefaults,
  getStoredConventions,
  mergeStoredConventions,
  normalizeConventionColor,
  removeCustomPresets,
  resolveConventions,
  setStoredConventions,
} from "../conventions/store.js";
import { getErrorMessage } from "../utils/errors.js";
import type { StoredConventions, StoredFormatPreset } from "../conventions/types.js";

const builderParamsSchema = Type.Object({
  dp: Type.Optional(Type.Number({ description: "Decimal places (0-10)." })),
  negative_style: Type.Optional(Type.Union([Type.Literal("parens"), Type.Literal("minus")])),
  zero_style: Type.Optional(Type.Union([
    Type.Literal("dash"),
    Type.Literal("single-dash"),
    Type.Literal("zero"),
    Type.Literal("blank"),
  ])),
  thousands_separator: Type.Optional(Type.Boolean()),
  currency_symbol: Type.Optional(Type.String()),
});

const formatPresetSchema = Type.Object({
  format: Type.String({ description: "Exact Excel format string (Format Cells > Custom)." }),
  builder_params: Type.Optional(builderParamsSchema),
});

const schema = Type.Object({
  action: Type.Union([
    Type.Literal("get"),
    Type.Literal("set"),
    Type.Literal("reset"),
  ]),

  preset_formats: Type.Optional(Type.Object({
    number: Type.Optional(formatPresetSchema),
    integer: Type.Optional(formatPresetSchema),
    currency: Type.Optional(formatPresetSchema),
    percent: Type.Optional(formatPresetSchema),
    ratio: Type.Optional(formatPresetSchema),
    text: Type.Optional(formatPresetSchema),
  })),

  custom_presets: Type.Optional(Type.Record(Type.String(), Type.Object({
    format: Type.String(),
    description: Type.Optional(Type.String()),
    builder_params: Type.Optional(builderParamsSchema),
  }))),
  remove_custom_presets: Type.Optional(Type.Array(Type.String())),

  visual_defaults: Type.Optional(Type.Object({
    font_name: Type.Optional(Type.String()),
    font_size: Type.Optional(Type.Number()),
  })),

  color_conventions: Type.Optional(Type.Object({
    hardcoded_value_color: Type.Optional(Type.String()),
    cross_sheet_link_color: Type.Optional(Type.String()),
  })),

  header_style: Type.Optional(Type.Object({
    fill_color: Type.Optional(Type.String()),
    font_color: Type.Optional(Type.String()),
    bold: Type.Optional(Type.Boolean()),
    wrap_text: Type.Optional(Type.Boolean()),
  })),
});

type Params = Static<typeof schema>;

function emitConventionsUpdatedEvent(): void {
  if (typeof document === "undefined") return;
  document.dispatchEvent(new CustomEvent("pi:conventions-updated"));
  document.dispatchEvent(new CustomEvent("pi:status-update"));
}

function formatPresetLine(name: string, preset: StoredFormatPreset): string {
  return `- ${name}: \`${preset.format}\``;
}

function formatConventionsMarkdown(stored: StoredConventions): string {
  const resolved = resolveConventions(stored);
  const diffs = diffFromDefaults(resolved);

  const lines: string[] = [
    "**Formatting conventions**" + (diffs.length > 0 ? " (★ = customized)" : ""),
    "",
    "### Built-in preset formats",
  ];

  lines.push(formatPresetLine("number", resolved.presetFormats.number));
  lines.push(formatPresetLine("integer", resolved.presetFormats.integer));
  lines.push(formatPresetLine("currency", resolved.presetFormats.currency));
  lines.push(formatPresetLine("percent", resolved.presetFormats.percent));
  lines.push(formatPresetLine("ratio", resolved.presetFormats.ratio));
  lines.push(formatPresetLine("text", resolved.presetFormats.text));

  lines.push("", "### Custom presets");
  const customEntries = Object.entries(resolved.customPresets);
  if (customEntries.length === 0) {
    lines.push("- _none_");
  } else {
    for (const [name, preset] of customEntries) {
      const desc = preset.description ? ` — ${preset.description}` : "";
      lines.push(`- ${name}${desc}: \`${preset.format}\``);
    }
  }

  lines.push(
    "",
    "### Visual defaults",
    `- Font: ${resolved.visualDefaults.fontName}`,
    `- Font size: ${resolved.visualDefaults.fontSize}`,
    "",
    "### Font color conventions",
    `- Hardcoded value font color: ${resolved.colorConventions.hardcodedValueColor}`,
    `- Cross-sheet link font color: ${resolved.colorConventions.crossSheetLinkColor}`,
    "",
    "### Header style",
    `- Fill: ${resolved.headerStyle.fillColor}`,
    `- Font color: ${resolved.headerStyle.fontColor}`,
    `- Bold: ${resolved.headerStyle.bold ? "yes" : "no"}`,
    `- Wrap text: ${resolved.headerStyle.wrapText ? "yes" : "no"}`,
  );

  if (diffs.length === 0) {
    lines.push("", "_All defaults — nothing customized._");
  } else {
    lines.push("", "### Active overrides");
    for (const diff of diffs) {
      lines.push(`- ${diff.label}: ${diff.value}`);
    }
  }

  return lines.join("\n");
}

function mapFormatPreset(input: {
  format: string;
  builder_params?: {
    dp?: number;
    negative_style?: "parens" | "minus";
    zero_style?: "dash" | "single-dash" | "zero" | "blank";
    thousands_separator?: boolean;
    currency_symbol?: string;
  };
}): StoredFormatPreset {
  return {
    format: input.format,
    builderParams: input.builder_params
      ? {
        dp: input.builder_params.dp,
        negativeStyle: input.builder_params.negative_style,
        zeroStyle: input.builder_params.zero_style,
        thousandsSeparator: input.builder_params.thousands_separator,
        currencySymbol: input.builder_params.currency_symbol,
      }
      : undefined,
  };
}

function buildUpdates(params: Params): StoredConventions {
  const updates: StoredConventions = {};

  if (params.preset_formats) {
    updates.presetFormats = {};
    if (params.preset_formats.number) updates.presetFormats.number = mapFormatPreset(params.preset_formats.number);
    if (params.preset_formats.integer) updates.presetFormats.integer = mapFormatPreset(params.preset_formats.integer);
    if (params.preset_formats.currency) updates.presetFormats.currency = mapFormatPreset(params.preset_formats.currency);
    if (params.preset_formats.percent) updates.presetFormats.percent = mapFormatPreset(params.preset_formats.percent);
    if (params.preset_formats.ratio) updates.presetFormats.ratio = mapFormatPreset(params.preset_formats.ratio);
    if (params.preset_formats.text) updates.presetFormats.text = mapFormatPreset(params.preset_formats.text);

    if (Object.keys(updates.presetFormats).length === 0) {
      updates.presetFormats = undefined;
    }
  }

  if (params.custom_presets) {
    updates.customPresets = {};

    for (const [name, preset] of Object.entries(params.custom_presets)) {
      updates.customPresets[name] = {
        format: preset.format,
        description: preset.description,
        builderParams: preset.builder_params
          ? {
            dp: preset.builder_params.dp,
            negativeStyle: preset.builder_params.negative_style,
            zeroStyle: preset.builder_params.zero_style,
            thousandsSeparator: preset.builder_params.thousands_separator,
            currencySymbol: preset.builder_params.currency_symbol,
          }
          : undefined,
      };
    }

    if (Object.keys(updates.customPresets).length === 0) {
      updates.customPresets = undefined;
    }
  }

  if (params.visual_defaults) {
    updates.visualDefaults = {
      fontName: params.visual_defaults.font_name,
      fontSize: params.visual_defaults.font_size,
    };
  }

  if (params.color_conventions) {
    updates.colorConventions = {
      hardcodedValueColor: params.color_conventions.hardcoded_value_color,
      crossSheetLinkColor: params.color_conventions.cross_sheet_link_color,
    };
  }

  if (params.header_style) {
    updates.headerStyle = {
      fillColor: params.header_style.fill_color,
      fontColor: params.header_style.font_color,
      bold: params.header_style.bold,
      wrapText: params.header_style.wrap_text,
    };
  }

  return updates;
}

function hasUpdates(value: StoredConventions): boolean {
  return (
    value.presetFormats !== undefined
    || value.customPresets !== undefined
    || value.visualDefaults !== undefined
    || value.colorConventions !== undefined
    || value.headerStyle !== undefined
  );
}

function validateSetParams(params: Params): string[] {
  const errors: string[] = [];

  const validateColor = (label: string, value: string | undefined): void => {
    if (value === undefined) {
      return;
    }

    if (!normalizeConventionColor(value)) {
      errors.push(`${label} must be #RRGGBB/#RGB or rgb(r,g,b).`);
    }
  };

  if (params.color_conventions) {
    validateColor("color_conventions.hardcoded_value_color", params.color_conventions.hardcoded_value_color);
    validateColor("color_conventions.cross_sheet_link_color", params.color_conventions.cross_sheet_link_color);
  }

  if (params.header_style) {
    validateColor("header_style.fill_color", params.header_style.fill_color);
    validateColor("header_style.font_color", params.header_style.font_color);
  }

  return errors;
}

export function createConventionsTool(): AgentTool<typeof schema, undefined> {
  return {
    name: "conventions",
    label: "Conventions",
    description:
      "Read/update formatting conventions: built-in/custom format presets, default font, "
      + "font-color conventions, and header style.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const settings = getAppStorage().settings;

        if (params.action === "get") {
          const stored = await getStoredConventions(settings);
          return {
            content: [{ type: "text", text: formatConventionsMarkdown(stored) }],
            details: undefined,
          };
        }

        if (params.action === "reset") {
          await setStoredConventions(settings, {});
          emitConventionsUpdatedEvent();
          return {
            content: [{
              type: "text",
              text: `Reset formatting conventions to defaults.\n\n${formatConventionsMarkdown({})}`,
            }],
            details: undefined,
          };
        }

        const validationErrors = validateSetParams(params);
        if (validationErrors.length > 0) {
          return {
            content: [{ type: "text", text: `Invalid conventions update:\n- ${validationErrors.join("\n- ")}` }],
            details: undefined,
          };
        }

        const updates = buildUpdates(params);
        const removeList = params.remove_custom_presets ?? [];

        if (!hasUpdates(updates) && removeList.length === 0) {
          return {
            content: [{
              type: "text",
              text: "No changes specified. Use action=\"get\" to view current conventions.",
            }],
            details: undefined,
          };
        }

        const current = await getStoredConventions(settings);
        let merged = mergeStoredConventions(current, updates);
        if (removeList.length > 0) {
          merged = removeCustomPresets(merged, removeList);
        }

        await setStoredConventions(settings, merged);
        emitConventionsUpdatedEvent();

        return {
          content: [{
            type: "text",
            text: `Updated formatting conventions.\n\n${formatConventionsMarkdown(merged)}`,
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
