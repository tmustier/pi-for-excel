/**
 * get_all_objects — List charts, pivot tables, and shapes on a sheet.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  sheet: Type.Optional(
    Type.String({
      description: "Target sheet name. If omitted, uses the active sheet.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createGetAllObjectsTool(): AgentTool<typeof schema> {
  return {
    name: "get_all_objects",
    label: "Get All Objects",
    description:
      "List charts, pivot tables, and shapes on a sheet, including their names.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await excelRun(async (context) => {
          const sheet = params.sheet
            ? context.workbook.worksheets.getItem(params.sheet)
            : context.workbook.worksheets.getActiveWorksheet();

          sheet.load("name");

          const charts = sheet.charts;
          charts.load("items/name,count");

          const pivotTables = sheet.pivotTables;
          pivotTables.load("items/name");
          const pivotCount = pivotTables.getCount();

          let shapes: Excel.ShapeCollection | null = null;
          try {
            shapes = sheet.shapes;
            shapes.load("items/name");
          } catch {
            shapes = null;
          }

          await context.sync();

          return {
            sheetName: sheet.name,
            charts: charts.items.map((c) => c.name),
            pivotTables: pivotTables.items.map((p) => p.name),
            pivotCount: pivotCount.value,
            shapes: shapes ? shapes.items.map((s) => s.name) : [],
          };
        });

        const lines: string[] = [];
        lines.push(`**Sheet: ${result.sheetName}**`);

        lines.push(
          `- Charts (${result.charts.length}): ${
            result.charts.length > 0 ? result.charts.join(", ") : "(none)"
          }`,
        );
        lines.push(
          `- Pivot tables (${result.pivotCount}): ${
            result.pivotTables.length > 0 ? result.pivotTables.join(", ") : "(none)"
          }`,
        );
        lines.push(
          `- Shapes (${result.shapes.length}): ${
            result.shapes.length > 0 ? result.shapes.join(", ") : "(none)"
          }`,
        );

        return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error listing objects: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}
