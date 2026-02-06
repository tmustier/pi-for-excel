/**
 * get_range_as_csv — Read a range and return values as CSV.
 *
 * Intended for compact, analysis-friendly reads.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";

const schema = Type.Object({
  range: Type.String({
    description:
      'Cell range in A1 notation, e.g. "A1:D10", "Sheet2!A1:B5". ' +
      "If no sheet is specified, uses the active sheet.",
  }),
});

type Params = Static<typeof schema>;

export function createGetRangeAsCsvTool(): AgentTool<typeof schema> {
  return {
    name: "get_range_as_csv",
    label: "Get Range as CSV",
    description:
      "Read cell values from a range and return them as CSV (values only).",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await excelRun(async (context) => {
          const { sheet, range } = getRange(context, params.range);
          range.load("values,address,rowCount,columnCount");
          sheet.load("name");
          await context.sync();
          return {
            sheetName: sheet.name,
            address: range.address,
            rows: range.rowCount,
            cols: range.columnCount,
            values: range.values,
          };
        });

        const fullAddr = qualifiedAddress(result.sheetName, result.address);
        const csv = formatCsv(result.values);

        const lines: string[] = [];
        lines.push(`**${fullAddr}** (${result.rows}×${result.cols})`);
        lines.push("");
        if (!csv) {
          lines.push("(empty)");
        } else {
          lines.push("```csv");
          lines.push(csv);
          lines.push("```");
        }

        return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error reading CSV: ${e.message}` }],
          details: undefined,
        };
      }
    },
  };
}

function formatCsv(values: any[][]): string {
  if (!values || values.length === 0) return "";

  return values
    .map((row) => row.map((value) => toCsvField(value)).join(","))
    .join("\n");
}

function toCsvField(value: any): string {
  if (value === null || value === undefined) return "";
  let str = typeof value === "string" ? value : String(value);
  if (/[",\n\r]/.test(str)) {
    str = str.replace(/"/g, '""');
    return `"${str}"`;
  }
  return str;
}
