/**
 * read_range — Read cell values, formulas, and optionally formatting.
 *
 * Two modes:
 * - "compact" (default): Markdown table of values. Token-efficient.
 * - "detailed": Full JSON with formulas, number formats. For debugging.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, getRange, qualifiedAddress, parseCell, colToLetter } from "../excel/helpers.js";
import { formatAsMarkdownTable, extractFormulas, findErrors } from "../utils/format.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  range: Type.String({
    description:
      'Cell range in A1 notation, e.g. "A1:D10", "Sheet2!A1:B5". ' +
      "If no sheet is specified, uses the active sheet.",
  }),
  mode: Type.Optional(
    Type.Union([Type.Literal("compact"), Type.Literal("detailed")], {
      description:
        '"compact" (default): markdown table of values. ' +
        '"detailed": includes formulas and number formats.',
    }),
  ),
});

type Params = Static<typeof schema>;

interface ReadRangeResult {
  sheetName: string;
  address: string;
  rows: number;
  cols: number;
  values: unknown[][];
  formulas: unknown[][];
  numberFormats: unknown[][];
}

export function createReadRangeTool(): AgentTool<typeof schema> {
  return {
    name: "read_range",
    label: "Read Range",
    description:
      "Read cell values from an Excel range. Returns a markdown table by default (compact mode). " +
      'Use mode "detailed" to also see formulas and number formats. ' +
      "Always read before modifying — never guess what's in the spreadsheet.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await excelRun<ReadRangeResult>(async (context) => {
          const { sheet, range } = getRange(context, params.range);
          range.load("values,formulas,numberFormat,address,rowCount,columnCount");
          sheet.load("name");
          await context.sync();
          return {
            sheetName: sheet.name,
            address: range.address,
            rows: range.rowCount,
            cols: range.columnCount,
            values: range.values,
            formulas: range.formulas,
            numberFormats: range.numberFormat,
          };
        });

        const fullAddress = qualifiedAddress(result.sheetName, result.address);
        // Extract just the cell part (without sheet!) for offset calculations
        const cellPart = result.address.includes("!") ? result.address.split("!")[1] : result.address;
        const startCell = cellPart.split(":")[0];

        const mode = params.mode || "compact";

        if (mode === "compact") {
          return formatCompact(fullAddress, result, startCell);
        } else {
          return formatDetailed(fullAddress, result, startCell);
        }
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error reading "${params.range}": ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

function formatCompact(
  address: string,
  result: ReadRangeResult,
  startCell: string,
): AgentToolResult<undefined> {
  const lines: string[] = [];
  lines.push(`**${address}** (${result.rows}×${result.cols})`);
  lines.push("");
  lines.push(formatAsMarkdownTable(result.values));

  // Append formulas if any exist
  const formulas = extractFormulas(result.formulas, startCell);
  if (formulas.length > 0) {
    lines.push("");
    lines.push(`**Formulas:** ${formulas.join(", ")}`);
  }

  // Append errors if any
  const errors = findErrors(result.values, startCell);
  if (errors.length > 0) {
    lines.push("");
    lines.push(`⚠️ **Errors:** ${errors.map((e) => `${e.address}=${e.error}`).join(", ")}`);
  }

  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}

function formatDetailed(
  address: string,
  result: ReadRangeResult,
  startCell: string,
): AgentToolResult<undefined> {
  const lines: string[] = [];
  lines.push(`**${address}** (${result.rows}×${result.cols})`);
  lines.push("");

  // Values table
  lines.push("### Values");
  lines.push(formatAsMarkdownTable(result.values));

  // All formulas
  const formulas = extractFormulas(result.formulas, startCell);
  if (formulas.length > 0) {
    lines.push("");
    lines.push("### Formulas");
    for (const f of formulas) {
      lines.push(`- ${f}`);
    }
  }

  // Number formats (deduplicated)
  const formatMap = new Map<string, string[]>();
  const start = parseCell(startCell);
  for (let r = 0; r < result.numberFormats.length; r++) {
    for (let c = 0; c < result.numberFormats[r].length; c++) {
      const fmt = result.numberFormats[r][c];
      if (typeof fmt === "string" && fmt !== "" && fmt !== "General") {
        const addr = `${colToLetter(start.col + c)}${start.row + r}`;
        const existing = formatMap.get(fmt) || [];
        existing.push(addr);
        formatMap.set(fmt, existing);
      }
    }
  }
  if (formatMap.size > 0) {
    lines.push("");
    lines.push("### Number Formats");
    for (const [fmt, cells] of formatMap) {
      lines.push(`- \`${fmt}\` → ${cells.join(", ")}`);
    }
  }

  // Errors
  const errors = findErrors(result.values, startCell);
  if (errors.length > 0) {
    lines.push("");
    lines.push("### ⚠️ Errors");
    for (const e of errors) {
      lines.push(`- ${e.address}: ${e.error}`);
    }
  }

  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}
