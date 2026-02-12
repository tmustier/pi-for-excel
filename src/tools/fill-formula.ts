/**
 * fill_formula — Fill a single formula across a range with Excel's AutoFill.
 *
 * This avoids constructing large 2D formula arrays. References adjust automatically.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import type { FillFormulaDetails } from "./tool-details.js";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";
import { buildWorkbookCellChangeSummary } from "../audit/cell-diff.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { dispatchWorkbookSnapshotCreated } from "../workbook/recovery-events.js";
import { getWorkbookRecoveryLog } from "../workbook/recovery-log.js";
import { countOccupiedCells, validateFormula } from "./write-cells.js";
import { findErrors } from "../utils/format.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  range: Type.String({
    description: 'Target range to fill, e.g. "B2:B20" or "Sheet1!C3:F20". Single contiguous range only.',
  }),
  formula: Type.String({
    description: 'Formula to fill, starting with "=", e.g. "=SUM(B2:B10)". Relative refs will adjust.',
  }),
  allow_overwrite: Type.Optional(
    Type.Boolean({
      description:
        "Set to true to overwrite existing data. Default: false. " +
        "If false and the target range contains data, the fill is blocked.",
    }),
  ),
});

type Params = Static<typeof schema>;

type FillFormulaResult =
  | {
    blocked: true;
    sheetName: string;
    address: string;
    existingCount: number;
    existingValues: unknown[][];
  }
  | {
    blocked: false;
    sheetName: string;
    address: string;
    rowCount: number;
    columnCount: number;
    beforeValues: unknown[][];
    beforeFormulas: unknown[][];
    readBackValues: unknown[][];
    readBackFormulas: unknown[][];
  };

export function createFillFormulaTool(): AgentTool<typeof schema, FillFormulaDetails> {
  return {
    name: "fill_formula",
    label: "Fill Formula",
    description:
      "Fill a single formula across a range using Excel's AutoFill. " +
      "Relative references adjust automatically. Use this instead of building large formula arrays.",
    parameters: schema,
    execute: async (toolCallId: string, params: Params): Promise<AgentToolResult<FillFormulaDetails>> => {
      try {
        if (!params.formula.startsWith("=")) {
          return {
            content: [{ type: "text", text: "Error: formula must start with '='." }],
            details: { kind: "fill_formula", blocked: false },
          };
        }

        const invalid = validateFormula(params.formula);
        if (invalid) {
          return {
            content: [{ type: "text", text: `Error: invalid formula (${invalid}).` }],
            details: { kind: "fill_formula", blocked: false },
          };
        }

        if (/[;,]/.test(params.range)) {
          return {
            content: [{ type: "text", text: "Error: fill_formula only supports a single contiguous range." }],
            details: { kind: "fill_formula", blocked: false },
          };
        }

        const result = await excelRun<FillFormulaResult>(async (context) => {
          const { sheet, range } = getRange(context, params.range);
          sheet.load("name");
          range.load("address,rowCount,columnCount,values,formulas");
          await context.sync();

          const beforeValues = range.values;
          const beforeFormulas = range.formulas;

          if (!params.allow_overwrite) {
            const occupiedCount = countOccupiedCells(beforeValues, beforeFormulas);
            if (occupiedCount > 0) {
              return {
                blocked: true,
                sheetName: sheet.name,
                address: range.address,
                existingCount: occupiedCount,
                existingValues: beforeValues,
              };
            }
          }

          const topLeft = range.getCell(0, 0);
          topLeft.formulas = [[params.formula]];
          topLeft.autoFill(range, "FillDefault");

          range.load("values,formulas,address,rowCount,columnCount");
          await context.sync();

          return {
            blocked: false,
            sheetName: sheet.name,
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            beforeValues,
            beforeFormulas,
            readBackValues: range.values,
            readBackFormulas: range.formulas,
          };
        });

        if (result.blocked) {
          const fullAddr = qualifiedAddress(result.sheetName, result.address);
          const blockedResult: AgentToolResult<FillFormulaDetails> = {
            content: [
              {
                type: "text",
                text:
                  `⛔ **Fill blocked** — ${fullAddr} contains ${result.existingCount} non-empty cell(s).\n\n` +
                  "To overwrite, confirm with the user and retry with `allow_overwrite: true`.",
              },
            ],
            details: {
              kind: "fill_formula",
              blocked: true,
              address: fullAddr,
              existingCount: result.existingCount,
            },
          };

          await getWorkbookChangeAuditLog().append({
            toolName: "fill_formula",
            toolCallId,
            blocked: true,
            outputAddress: fullAddr,
            changedCount: 0,
            changes: [],
          });

          return blockedResult;
        }

        const fullAddr = qualifiedAddress(result.sheetName, result.address);
        const lines: string[] = [];
        lines.push(`Filled formula across **${fullAddr}** (${result.rowCount}×${result.columnCount})`);
        lines.push(`**Formula pattern:** \`${params.formula}\``);

        const topLeftRaw: unknown = result.readBackFormulas?.[0]?.[0] ?? "";
        const bottomRightRaw: unknown = result.readBackFormulas?.[result.rowCount - 1]?.[result.columnCount - 1] ?? "";
        const topLeft = typeof topLeftRaw === "string" ? topLeftRaw : JSON.stringify(topLeftRaw);
        const bottomRight = typeof bottomRightRaw === "string" ? bottomRightRaw : JSON.stringify(bottomRightRaw);
        if (topLeft && bottomRight && (result.rowCount > 1 || result.columnCount > 1)) {
          lines.push(`**Example formulas:** top-left \`${topLeft}\`, bottom-right \`${bottomRight}\``);
        }

        const cellPart = result.address.includes("!") ? result.address.split("!")[1] : result.address;
        const startCell = cellPart.split(":")[0];
        const errors = findErrors(result.readBackValues, startCell);
        if (errors.length > 0) {
          lines.push("");
          lines.push(`⚠️ **${errors.length} formula error(s):**`);
          for (const e of errors) {
            lines.push(`- ${e.address}: ${e.error}`);
          }
        }

        const changes = buildWorkbookCellChangeSummary({
          sheetName: result.sheetName,
          startCell,
          beforeValues: result.beforeValues,
          beforeFormulas: result.beforeFormulas,
          afterValues: result.readBackValues,
          afterFormulas: result.readBackFormulas,
        });

        if (changes.changedCount > 0) {
          lines.push("");
          lines.push(`Changed cell(s): ${changes.changedCount}.`);
        }

        const details: FillFormulaDetails = {
          kind: "fill_formula",
          blocked: false,
          address: fullAddr,
          formulaErrorCount: errors.length,
          changes,
        };

        await getWorkbookChangeAuditLog().append({
          toolName: "fill_formula",
          toolCallId,
          blocked: false,
          outputAddress: fullAddr,
          changedCount: changes.changedCount,
          changes: changes.sample,
        });

        const checkpoint = await getWorkbookRecoveryLog().append({
          toolName: "fill_formula",
          toolCallId,
          address: fullAddr,
          changedCount: changes.changedCount,
          beforeValues: result.beforeValues,
          beforeFormulas: result.beforeFormulas,
        });

        if (checkpoint) {
          dispatchWorkbookSnapshotCreated({
            snapshotId: checkpoint.id,
            toolName: checkpoint.toolName,
            address: checkpoint.address,
            changedCount: checkpoint.changedCount,
          });
        }

        return { content: [{ type: "text", text: lines.join("\n") }], details };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error filling formula: ${getErrorMessage(e)}` }],
          details: { kind: "fill_formula", blocked: false },
        };
      }
    },
  };
}
