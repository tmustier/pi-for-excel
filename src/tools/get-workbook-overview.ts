/**
 * get_workbook_overview — Returns a structural blueprint of the workbook.
 *
 * Includes: workbook name, all sheets with dimensions, header rows,
 * named ranges, and table inventory. This is injected at session start
 * and available on-demand.
 *
 * Pushes rich structural metadata (headers, named ranges, tables) — not just sheet names + dimensions.
 */

import { Type } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({});

export function createGetWorkbookOverviewTool(): AgentTool<typeof schema> {
  return {
    name: "get_workbook_overview",
    label: "Workbook Overview",
    description:
      "Get a structural overview of the workbook: sheet names, dimensions, " +
      "header rows, named ranges, tables, and object counts. Use this at the start of a " +
      "conversation or when you need to understand the workbook's structure " +
      "before reading specific ranges.",
    parameters: schema,
    execute: async (): Promise<AgentToolResult<undefined>> => {
      try {
        const overview = await buildOverview();
        return {
          content: [{ type: "text", text: overview }],
          details: undefined,
        };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error getting workbook overview: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

/** Build the full workbook overview. Also used by context injection. */
export async function buildOverview(): Promise<string> {
  return excelRun(async (context) => {
    const wb = context.workbook;
    wb.load("name");

    const sheets = wb.worksheets;
    sheets.load("items/name,items/id,items/position,items/visibility");

    const names = wb.names;
    names.load("items/name,items/type,items/value,items/visible");

    await context.sync();

    const lines: string[] = [];
    lines.push(`## Workbook: ${wb.name}`);
    lines.push("");
    lines.push(`### Sheets (${sheets.items.length})`);

    for (const sheet of sheets.items) {
      const used = sheet.getUsedRangeOrNullObject();
      used.load("rowCount,columnCount,address");

      // Get header row (first populated row)
      const headerRange = sheet.getRange("1:1").getUsedRangeOrNullObject();
      headerRange.load("values");

      // Get tables on this sheet
      const tables = sheet.tables;
      tables.load("items/name,items/columns/count,items/rows/count");

      // Get object counts
      const charts = sheet.charts;
      charts.load("count");

      const pivotTables = sheet.pivotTables;
      const pivotCount = pivotTables.getCount();

      let shapes: any | null = null;
      try {
        shapes = sheet.shapes;
        shapes.load("items");
      } catch {
        shapes = null;
      }

      await context.sync();

      const dims = used.isNullObject
        ? "empty"
        : `${used.rowCount} rows × ${used.columnCount} cols`;

      const visibility = sheet.visibility === "Visible" ? "" : ` (${sheet.visibility})`;

      const headers = headerRange.isNullObject
        ? []
        : headerRange.values[0].filter((v: any) => v !== null && v !== undefined && v !== "");

      lines.push(
        `${sheet.position + 1}. **${sheet.name}**${visibility} — ${dims}`,
      );

      if (headers.length > 0) {
        const display = headers.length > 8
          ? headers.slice(0, 8).join(", ") + `, … (+${headers.length - 8} more)`
          : headers.join(", ");
        lines.push(`   Headers: ${display}`);
      }

      // List tables
      if (tables.items.length > 0) {
        for (const table of tables.items) {
          lines.push(`   Table: "${table.name}" (${table.rows.count} rows × ${table.columns.count} cols)`);
        }
      }

      const chartCount = charts.count || 0;
      const pivotTotal = pivotCount.value || 0;
      const shapeCount = shapes ? shapes.items.length : 0;
      const objectTotal = chartCount + pivotTotal + shapeCount;

      if (objectTotal > 0) {
        lines.push(
          `   Objects: ${chartCount} chart(s), ${pivotTotal} pivot table(s), ${shapeCount} shape(s)`,
        );
      }
    }

    // Named ranges
    const visibleNames = names.items.filter((n: any) => n.visible);
    if (visibleNames.length > 0) {
      lines.push("");
      lines.push(`### Named Ranges (${visibleNames.length})`);
      for (const n of visibleNames) {
        lines.push(`- **${n.name}** = ${n.value}`);
      }
    }

    return lines.join("\n");
  });
}
