/**
 * conditional_format — Add or clear conditional formatting rules.
 *
 * Supports:
 * - Custom formula rules
 * - Cell value rules (greater/less/equal/between)
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";

type CellValueOperator =
  | "Between"
  | "NotBetween"
  | "EqualTo"
  | "NotEqualTo"
  | "GreaterThan"
  | "LessThan"
  | "GreaterThanOrEqual"
  | "LessThanOrEqual";

const schema = Type.Object({
  action: Type.Union([Type.Literal("add"), Type.Literal("clear")], {
    description: '"add" to create a rule, "clear" to remove all rules in the range.',
  }),
  range: Type.String({
    description: 'Target range, e.g. "A1:D10" or "Sheet2!B2:B50".',
  }),
  type: Type.Optional(
    Type.Union([Type.Literal("formula"), Type.Literal("cell_value")], {
      description: 'Rule type for "add": "formula" or "cell_value".',
    }),
  ),
  formula: Type.Optional(
    Type.String({
      description: 'Custom formula for "formula" rules, e.g. "=A1>0".',
    }),
  ),
  operator: Type.Optional(
    Type.Union(
      [
        Type.Literal("Between"),
        Type.Literal("NotBetween"),
        Type.Literal("EqualTo"),
        Type.Literal("NotEqualTo"),
        Type.Literal("GreaterThan"),
        Type.Literal("LessThan"),
        Type.Literal("GreaterThanOrEqual"),
        Type.Literal("LessThanOrEqual"),
      ],
      {
        description:
          "Cell value operator (required for cell_value rules).",
      },
    ),
  ),
  value: Type.Optional(
    Type.Union([Type.String(), Type.Number()], {
      description:
        "Cell value comparison target (required for cell_value rules). Use numbers or formulas like \"=$B$2\".",
    }),
  ),
  value2: Type.Optional(
    Type.Union([Type.String(), Type.Number()], {
      description:
        "Second value for Between/NotBetween operators (optional).",
    }),
  ),
  fill_color: Type.Optional(
    Type.String({ description: 'Fill color hex, e.g. "#FFFDE0".' }),
  ),
  font_color: Type.Optional(
    Type.String({ description: 'Font color hex, e.g. "#000000".' }),
  ),
  bold: Type.Optional(Type.Boolean({ description: "Bold text." })),
  italic: Type.Optional(Type.Boolean({ description: "Italic text." })),
  underline: Type.Optional(Type.Boolean({ description: "Underline text." })),
  stop_if_true: Type.Optional(
    Type.Boolean({ description: "Stop evaluating later rules if true." }),
  ),
});

type Params = Static<typeof schema>;

export function createConditionalFormatTool(): AgentTool<typeof schema> {
  return {
    name: "conditional_format",
    label: "Conditional Format",
    description:
      "Add or clear conditional formatting rules. Supports custom formula and cell value rules.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        if (params.action === "clear") {
          return await clearFormats(params);
        }

        if (!params.type) {
          return {
            content: [
              {
                type: "text",
                text: "Error: type is required when action is \"add\".",
              },
            ],
            details: undefined,
          };
        }

        return await addFormat(params);
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error: ${e.message}` }],
          details: undefined,
        };
      }
    },
  };
}

async function clearFormats(params: Params): Promise<AgentToolResult<undefined>> {
  const result = await excelRun(async (context: any) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    const formats = range.conditionalFormats;
    const countResult = formats.getCount();
    await context.sync();

    const existing = countResult.value;
    formats.clearAll();
    await context.sync();

    return { sheetName: sheet.name, address: range.address, existing };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);
  return {
    content: [
      {
        type: "text",
        text: `✅ Cleared ${result.existing} conditional format(s) from **${fullAddr}**.`,
      },
    ],
    details: undefined,
  };
}

async function addFormat(params: Params): Promise<AgentToolResult<undefined>> {
  validateAddParams(params);

  const result = await excelRun(async (context: any) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    await context.sync();

    const cfType =
      params.type === "formula"
        ? Excel.ConditionalFormatType.custom
        : Excel.ConditionalFormatType.cellValue;

    const cf = range.conditionalFormats.add(cfType);

    if (params.stop_if_true !== undefined) {
      cf.stopIfTrue = params.stop_if_true;
    }

    if (params.type === "formula") {
      cf.custom.rule.formula = params.formula as string;
      applyFormat(cf.custom.format, params);
    } else {
      const operator = params.operator as CellValueOperator;
      const formula1 = stringifyValue(params.value);
      const rule: any = { formula1, operator };
      if (params.value2 !== undefined) {
        rule.formula2 = stringifyValue(params.value2);
      }
      cf.cellValue.rule = rule;
      applyFormat(cf.cellValue.format, params);
    }

    await context.sync();

    return { sheetName: sheet.name, address: range.address };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);
  const details =
    params.type === "formula"
      ? `formula rule (${params.formula})`
      : `cell value rule (${params.operator} ${params.value}${
          params.value2 !== undefined ? ` and ${params.value2}` : ""
        })`;

  return {
    content: [
      {
        type: "text",
        text: `✅ Added conditional format to **${fullAddr}** — ${details}.`,
      },
    ],
    details: undefined,
  };
}

function validateAddParams(params: Params): void {
  if (params.type === "formula") {
    if (!params.formula) {
      throw new Error("formula is required for type=\"formula\".");
    }
    return;
  }

  if (!params.operator) {
    throw new Error("operator is required for type=\"cell_value\".");
  }

  if (params.value === undefined || params.value === null || params.value === "") {
    throw new Error("value is required for type=\"cell_value\".");
  }

  if (
    (params.operator === "Between" || params.operator === "NotBetween") &&
    (params.value2 === undefined || params.value2 === null || params.value2 === "")
  ) {
    throw new Error("value2 is required for Between/NotBetween operators.");
  }
}

function stringifyValue(value: string | number | undefined): string {
  if (value === undefined || value === null) return "";
  return typeof value === "number" ? value.toString() : value;
}

function applyFormat(format: any, params: Params): void {
  if (params.fill_color) {
    format.fill.color = params.fill_color;
  }
  if (params.font_color) {
    format.font.color = params.font_color;
  }
  if (params.bold !== undefined) {
    format.font.bold = params.bold;
  }
  if (params.italic !== undefined) {
    format.font.italic = params.italic;
  }
  if (params.underline !== undefined) {
    format.font.underline = params.underline ? "Single" : "None";
  }
}
