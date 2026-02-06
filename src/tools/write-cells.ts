/**
 * write_cells — Write values and formulas to Excel cells.
 *
 * Features:
 * - Overwrite protection (blocks by default if target has data)
 * - Auto-verify: reads back after writing, reports formula errors
 * - Supports formulas (strings starting with "=")
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import {
  excelRun, getRange, qualifiedAddress, parseCell,
  colToLetter, computeRangeAddress, padValues,
} from "../excel/helpers.js";
import { formatAsMarkdownTable, findErrors } from "../utils/format.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  start_cell: Type.String({
    description:
      'Top-left cell to write from, e.g. "A1", "Sheet2!B3". ' +
      "If no sheet is specified, uses the active sheet.",
  }),
  values: Type.Array(Type.Array(Type.Any()), {
    description:
      "2D array of values. Each inner array is a row. " +
      'Strings starting with "=" are formulas. ' +
      'Example: [["Name", "Total"], ["Alice", "=SUM(B2:B10)"]]',
  }),
  allow_overwrite: Type.Optional(
    Type.Boolean({
      description:
        "Set to true to overwrite existing data. Default: false. " +
        "If false and the target range contains values or formulas, the write is blocked " +
        "and the existing data is returned so you can ask the user.",
    }),
  ),
});

type Params = Static<typeof schema>;

interface InvalidFormula {
  address: string;
  formula: string;
  reason: string;
}

type WriteCellsResult =
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
    readBackValues: unknown[][];
    readBackFormulas: unknown[][];
  };

type BlockedWriteCellsResult = Extract<WriteCellsResult, { blocked: true }>;
type SuccessWriteCellsResult = Extract<WriteCellsResult, { blocked: false }>;

export function createWriteCellsTool(): AgentTool<typeof schema> {
  return {
    name: "write_cells",
    label: "Write Cells",
    description:
      "Write values and formulas to Excel cells. Provide a start cell and a 2D array. " +
      'Strings starting with "=" are treated as formulas. ' +
      "By default, blocks if the target range already contains data — " +
      "set allow_overwrite=true after confirming with the user. " +
      "After writing, automatically verifies results and reports any formula errors.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        if (!params.values || params.values.length === 0) {
          return {
            content: [{ type: "text", text: "Error: values array is empty." }],
            details: undefined,
          };
        }

        const { padded, rows, cols } = padValues(params.values);

        const startCellRef = params.start_cell.includes("!")
          ? params.start_cell.split("!")[1]
          : params.start_cell;

        if (startCellRef.includes(":")) {
          return {
            content: [{ type: "text", text: "Error: start_cell must be a single cell (e.g. \"A1\")." }],
            details: undefined,
          };
        }

        let invalidFormulas: InvalidFormula[] = [];
        try {
          invalidFormulas = findInvalidFormulas(padded, startCellRef);
        } catch {
          return {
            content: [{ type: "text", text: `Error: invalid start_cell "${params.start_cell}".` }],
            details: undefined,
          };
        }

        if (invalidFormulas.length > 0) {
          const lines: string[] = [];
          lines.push("⛔ **Write blocked** — invalid formula syntax detected:");
          for (const invalid of invalidFormulas) {
            lines.push(`- ${invalid.address}: ${invalid.formula} (${invalid.reason})`);
          }
          lines.push("");
          lines.push("Fix the formulas and retry.");
          return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
        }

        const result = await excelRun<WriteCellsResult>(async (context) => {
          const { sheet } = getRange(context, params.start_cell);
          sheet.load("name");

          const rangeAddr = computeRangeAddress(startCellRef, rows, cols);
          const targetRange = sheet.getRange(rangeAddr);

          // Overwrite protection: check if target has existing data (values or formulas)
          if (!params.allow_overwrite) {
            targetRange.load("values,formulas");
            await context.sync();

            const occupiedCount = countOccupiedCells(targetRange.values, targetRange.formulas);
            if (occupiedCount > 0) {
              return {
                blocked: true,
                sheetName: sheet.name,
                address: rangeAddr,
                existingCount: occupiedCount,
                existingValues: targetRange.values,
              };
            }
          }

          // Write
          targetRange.values = padded;
          await context.sync();

          // Read back to verify
          const verify = sheet.getRange(rangeAddr);
          verify.load("values,formulas,address");
          await context.sync();

          return {
            blocked: false,
            sheetName: sheet.name,
            address: verify.address,
            readBackValues: verify.values,
            readBackFormulas: verify.formulas,
          };
        });

        if (result.blocked) {
          return formatBlocked(result);
        }
        return formatSuccess(result, rows, cols);
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error writing cells: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

function findInvalidFormulas(values: unknown[][], startCell: string): InvalidFormula[] {
  const start = parseCell(startCell);
  const invalid: InvalidFormula[] = [];

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const value = values[r][c];
      if (typeof value === "string" && value.startsWith("=")) {
        const reason = validateFormula(value);
        if (reason) {
          invalid.push({
            address: `${colToLetter(start.col + c)}${start.row + r}`,
            formula: value,
            reason,
          });
        }
      }
    }
  }

  return invalid;
}

export function validateFormula(formula: string): string | null {
  if (!formula.startsWith("=")) return null;
  const body = formula.slice(1);

  if (body.trim().length === 0) return "Empty formula";

  const quoteCount = (body.match(/"/g) || []).length;
  if (quoteCount % 2 !== 0) return "Unbalanced quotes";

  let depth = 0;
  let inString = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "(") depth += 1;
    if (ch === ")") {
      depth -= 1;
      if (depth < 0) return "Unbalanced parentheses";
    }
  }
  if (depth !== 0) return "Unbalanced parentheses";

  const trimmed = body.trim();
  if (/[+\-*/^&,]$/.test(trimmed)) return "Formula ends with an operator";

  return null;
}

export function countOccupiedCells(values: unknown[][], formulas: unknown[][]): number {
  let count = 0;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const value = values[r][c];
      const formula = formulas?.[r]?.[c];
      const hasValue = value !== null && value !== undefined && value !== "";
      const hasFormula = typeof formula === "string" && formula.startsWith("=");
      if (hasValue || hasFormula) count += 1;
    }
  }
  return count;
}

function formatBlocked(result: BlockedWriteCellsResult): AgentToolResult<undefined> {
  const fullAddr = qualifiedAddress(result.sheetName, result.address);
  const lines: string[] = [];

  lines.push(`⛔ **Write blocked** — ${fullAddr} contains ${result.existingCount} non-empty cell(s).`);
  lines.push("");

  if (result.existingCount > 0) {
    lines.push("**Existing data:**");
    lines.push(formatAsMarkdownTable(result.existingValues));
    lines.push("");
  } else {
    lines.push("**Existing data:** (empty)");
    lines.push("");
  }

  lines.push(
    "To overwrite, confirm with the user and retry with `allow_overwrite: true`.",
  );
  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}

function formatSuccess(result: SuccessWriteCellsResult, rows: number, cols: number): AgentToolResult<undefined> {
  const fullAddr = qualifiedAddress(result.sheetName, result.address);
  const cellPart = result.address.includes("!") ? result.address.split("!")[1] : result.address;
  const startCell = cellPart.split(":")[0];

  const lines: string[] = [];
  lines.push(`✅ Written to **${fullAddr}** (${rows}×${cols})`);

  // Check for formula errors
  const errors = findErrors(result.readBackValues, startCell);
  if (errors.length > 0) {
    // Attach formula info to errors
    const start = parseCell(startCell);
    for (const err of errors) {
      const errCell = parseCell(err.address);
      const r = errCell.row - start.row;
      const c = errCell.col - start.col;
      if (r >= 0 && c >= 0 && r < result.readBackFormulas.length && c < result.readBackFormulas[r].length) {
        const f = result.readBackFormulas[r][c];
        if (typeof f === "string") {
          err.formula = f;
        }
      }
    }

    lines.push("");
    lines.push(`⚠️ **${errors.length} formula error(s):**`);
    for (const e of errors) {
      lines.push(`- ${e.address}: ${e.error}${e.formula ? ` (formula: ${e.formula})` : ""}`);
    }
    lines.push("");
    lines.push("Review and fix with another write_cells call.");
  } else {
    lines.push("");
    lines.push("**Verified values:**");
    lines.push(formatAsMarkdownTable(result.readBackValues));
  }

  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}
