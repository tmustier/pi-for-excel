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
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { dispatchWorkbookSnapshotCreated } from "../workbook/recovery-events.js";
import { captureConditionalFormatState } from "../workbook/recovery-states.js";
import { getWorkbookRecoveryLog } from "../workbook/recovery-log.js";
import { getErrorMessage } from "../utils/errors.js";
import type { ConditionalFormatDetails } from "./tool-details.js";
import {
  CHECKPOINT_SKIPPED_NOTE,
  CHECKPOINT_SKIPPED_REASON,
  recoveryCheckpointCreated,
  recoveryCheckpointUnavailable,
} from "./recovery-metadata.js";

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

export function createConditionalFormatTool(): AgentTool<typeof schema, ConditionalFormatDetails> {
  return {
    name: "conditional_format",
    label: "Conditional Format",
    description:
      "Add or clear conditional formatting rules. Supports custom formula and cell value rules.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ConditionalFormatDetails>> => {
      try {
        if (params.action === "clear") {
          return await clearFormats(toolCallId, params);
        }

        if (!params.type) {
          const message = "type is required when action is \"add\".";

          await getWorkbookChangeAuditLog().append({
            toolName: "conditional_format",
            toolCallId,
            blocked: true,
            outputAddress: params.range,
            changedCount: 0,
            changes: [],
            summary: `error: ${message}`,
          });

          return {
            content: [
              {
                type: "text",
                text: `Error: ${message}`,
              },
            ],
            details: {
              kind: "conditional_format",
              action: params.action,
            },
          };
        }

        return await addFormat(toolCallId, params);
      } catch (e: unknown) {
        const message = getErrorMessage(e);

        await getWorkbookChangeAuditLog().append({
          toolName: "conditional_format",
          toolCallId,
          blocked: true,
          outputAddress: params.range,
          changedCount: 0,
          changes: [],
          summary: `error: ${message}`,
        });

        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "conditional_format",
            action: params.action,
          },
        };
      }
    },
  };
}

async function clearFormats(
  toolCallId: string,
  params: Params,
): Promise<AgentToolResult<ConditionalFormatDetails>> {
  const checkpointState = await captureConditionalFormatState(params.range);

  const result = await excelRun(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address,rowCount,columnCount");
    const formats = range.conditionalFormats;
    const countResult = formats.getCount();
    await context.sync();

    const existing = countResult.value;
    const cellCount = range.rowCount * range.columnCount;
    formats.clearAll();
    await context.sync();

    return { sheetName: sheet.name, address: range.address, existing, cellCount };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);

  await getWorkbookChangeAuditLog().append({
    toolName: "conditional_format",
    toolCallId,
    blocked: false,
    outputAddress: fullAddr,
    changedCount: result.cellCount,
    changes: [],
    summary: `cleared ${result.existing} rule(s) across ${result.cellCount} cell(s)`,
  });

  const output: AgentToolResult<ConditionalFormatDetails> = {
    content: [
      {
        type: "text",
        text: `Cleared ${result.existing} conditional format(s) from **${fullAddr}**.`,
      },
    ],
    details: {
      kind: "conditional_format",
      action: "clear",
      address: fullAddr,
    },
  };

  if (!checkpointState.supported) {
    output.details.recovery = recoveryCheckpointUnavailable(checkpointState.reason ?? CHECKPOINT_SKIPPED_REASON);
    appendResultNote(output, CHECKPOINT_SKIPPED_NOTE);
    return output;
  }

  const checkpoint = await getWorkbookRecoveryLog().appendConditionalFormat({
    toolName: "conditional_format",
    toolCallId,
    address: fullAddr,
    changedCount: result.cellCount,
    cellCount: result.cellCount,
    conditionalFormatRules: checkpointState.rules,
  });

  if (!checkpoint) {
    output.details.recovery = recoveryCheckpointUnavailable(CHECKPOINT_SKIPPED_REASON);
    appendResultNote(output, CHECKPOINT_SKIPPED_NOTE);
    return output;
  }

  output.details.recovery = recoveryCheckpointCreated(checkpoint.id);
  dispatchWorkbookSnapshotCreated({
    snapshotId: checkpoint.id,
    toolName: checkpoint.toolName,
    address: checkpoint.address,
    changedCount: checkpoint.changedCount,
  });

  return output;
}

async function addFormat(
  toolCallId: string,
  params: Params,
): Promise<AgentToolResult<ConditionalFormatDetails>> {
  validateAddParams(params);

  const checkpointState = await captureConditionalFormatState(params.range);

  const result = await excelRun(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address,rowCount,columnCount");
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
      if (typeof params.formula !== "string") {
        throw new Error("formula is required for type=\"formula\".");
      }

      cf.custom.rule.formula = params.formula;
      applyFormat(cf.custom.format, params);
    } else {
      if (typeof params.operator !== "string") {
        throw new Error("operator is required for type=\"cell_value\".");
      }

      const formula1 = stringifyValue(params.value);
      const rule: Excel.ConditionalCellValueRule = {
        formula1,
        operator: params.operator,
      };

      if (params.value2 !== undefined) {
        rule.formula2 = stringifyValue(params.value2);
      }

      cf.cellValue.rule = rule;
      applyFormat(cf.cellValue.format, params);
    }

    await context.sync();

    return {
      sheetName: sheet.name,
      address: range.address,
      cellCount: range.rowCount * range.columnCount,
    };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);
  const details =
    params.type === "formula"
      ? `formula rule (${params.formula ?? ""})`
      : `cell value rule (${params.operator ?? ""} ${params.value ?? ""}${
          params.value2 !== undefined ? ` and ${String(params.value2)}` : ""
        })`;

  await getWorkbookChangeAuditLog().append({
    toolName: "conditional_format",
    toolCallId,
    blocked: false,
    outputAddress: fullAddr,
    changedCount: result.cellCount,
    changes: [],
    summary: `added ${params.type ?? "rule"} across ${result.cellCount} cell(s)`,
  });

  const output: AgentToolResult<ConditionalFormatDetails> = {
    content: [
      {
        type: "text",
        text: `Added conditional format to **${fullAddr}** — ${details}.`,
      },
    ],
    details: {
      kind: "conditional_format",
      action: "add",
      address: fullAddr,
    },
  };

  if (!checkpointState.supported) {
    output.details.recovery = recoveryCheckpointUnavailable(checkpointState.reason ?? CHECKPOINT_SKIPPED_REASON);
    appendResultNote(output, CHECKPOINT_SKIPPED_NOTE);
    return output;
  }

  const checkpoint = await getWorkbookRecoveryLog().appendConditionalFormat({
    toolName: "conditional_format",
    toolCallId,
    address: fullAddr,
    changedCount: result.cellCount,
    cellCount: result.cellCount,
    conditionalFormatRules: checkpointState.rules,
  });

  if (!checkpoint) {
    output.details.recovery = recoveryCheckpointUnavailable(CHECKPOINT_SKIPPED_REASON);
    appendResultNote(output, CHECKPOINT_SKIPPED_NOTE);
    return output;
  }

  output.details.recovery = recoveryCheckpointCreated(checkpoint.id);
  dispatchWorkbookSnapshotCreated({
    snapshotId: checkpoint.id,
    toolName: checkpoint.toolName,
    address: checkpoint.address,
    changedCount: checkpoint.changedCount,
  });

  return output;
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

function applyFormat(format: Excel.ConditionalRangeFormat, params: Params): void {
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

function appendResultNote(result: AgentToolResult<ConditionalFormatDetails>, note: string): void {
  const first = result.content[0];
  if (!first || first.type !== "text") return;

  first.text = `${first.text}\n\n${note}`;
}
