/**
 * get_workbook_overview — Returns a structural blueprint of the workbook.
 *
 * Includes: workbook name, all sheets with dimensions, header rows,
 * named ranges, and table inventory. This is injected at session start
 * and available on-demand.
 *
 * Pushes rich structural metadata (headers, named ranges, tables) — not just sheet names + dimensions.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { excelRun, colToLetter } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  sheet: Type.Optional(
    Type.String({
      description:
        "If provided, return detailed info for this specific sheet " +
        "(dimensions, headers, tables, named ranges, objects, and a data preview). " +
        "If omitted, return the workbook-level overview.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createGetWorkbookOverviewTool(): AgentTool<typeof schema> {
  return {
    name: "get_workbook_overview",
    label: "Workbook Overview",
    description:
      "Get a structural overview of the workbook: sheet names, dimensions, " +
      "header rows, named ranges, tables, and object counts. Use this at the start of a " +
      "conversation or when you need to understand the workbook's structure " +
      "before reading specific ranges. Optionally pass a sheet name for detailed " +
      "sheet-level info including objects, tables, named ranges, and a data preview.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = params.sheet
          ? await buildSheetDetail(params.sheet)
          : await buildOverview();
        return {
          content: [{ type: "text", text }],
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

      let shapes: Excel.ShapeCollection | null = null;
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
        : headerRange.values[0].filter((v) => v !== null && v !== undefined && v !== "");

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
    const visibleNames = names.items.filter((n) => n.visible);
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

// ============================================================================
// Sheet-level detail (absorbs former get_all_objects logic)
// ============================================================================

/** Build detailed info for a single sheet including objects, tables, named ranges, and a data preview. */
async function buildSheetDetail(sheetName: string): Promise<string> {
  return excelRun(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name,visibility");

    const used = sheet.getUsedRangeOrNullObject();
    used.load("rowCount,columnCount,address,values");

    // Header row
    const headerRange = sheet.getRange("1:1").getUsedRangeOrNullObject();
    headerRange.load("values");

    // Tables
    const tables = sheet.tables;
    tables.load("items/name,items/columns/count,items/rows/count");

    // Named ranges (workbook-level; we filter to this sheet below)
    const names = context.workbook.names;
    names.load("items/name,items/value,items/visible");

    // Objects — charts, pivot tables, shapes
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

    const lines: string[] = [];

    // ── Header ──
    const visibility = sheet.visibility === "Visible" ? "" : ` (${sheet.visibility})`;
    const dims = used.isNullObject
      ? "empty"
      : `${used.rowCount} rows × ${used.columnCount} cols`;
    lines.push(`## Sheet: ${sheet.name}${visibility}`);
    lines.push(`Dimensions: ${dims}`);

    // ── Headers ──
    const headers = headerRange.isNullObject
      ? []
      : (headerRange.values[0] as unknown[]).filter(
          (v) => v !== null && v !== undefined && v !== "",
        );
    if (headers.length > 0) {
      lines.push(`Headers: ${headers.join(", ")}`);
    }

    // ── Tables ──
    if (tables.items.length > 0) {
      lines.push("");
      lines.push(`### Tables (${tables.items.length})`);
      for (const table of tables.items) {
        lines.push(
          `- **${table.name}** (${table.rows.count} rows × ${table.columns.count} cols)`,
        );
      }
    }

    // ── Named ranges referencing this sheet ──
    const sheetPrefix = `${sheet.name}!`.toLowerCase();
    const sheetQuotedPrefix = `'${sheet.name}'!`.toLowerCase();
    const relevantNames = names.items.filter((n) => {
      if (!n.visible) return false;
      const rawVal: unknown = n.value;
      const val = typeof rawVal === "string" ? rawVal.toLowerCase() : "";
      return val.startsWith(sheetPrefix) || val.startsWith(sheetQuotedPrefix);
    });
    if (relevantNames.length > 0) {
      lines.push("");
      lines.push(`### Named Ranges (${relevantNames.length})`);
      for (const n of relevantNames) {
        lines.push(`- **${n.name}** = ${n.value}`);
      }
    }

    // ── Objects (charts, pivot tables, shapes) ──
    const chartNames = charts.items.map((c) => c.name);
    const pivotNames = pivotTables.items.map((p) => p.name);
    const shapeNames = shapes ? shapes.items.map((s) => s.name) : [];
    const objectTotal = chartNames.length + pivotNames.length + shapeNames.length;

    if (objectTotal > 0) {
      lines.push("");
      lines.push("### Objects");
      lines.push(
        `- Charts (${chartNames.length}): ${chartNames.length > 0 ? chartNames.join(", ") : "(none)"}`,
      );
      lines.push(
        `- Pivot tables (${pivotCount.value}): ${pivotNames.length > 0 ? pivotNames.join(", ") : "(none)"}`,
      );
      lines.push(
        `- Shapes (${shapeNames.length}): ${shapeNames.length > 0 ? shapeNames.join(", ") : "(none)"}`,
      );
    }

    // ── Data preview (first 5 rows as markdown table) ──
    if (!used.isNullObject && used.rowCount > 0 && used.columnCount > 0) {
      const previewRowCount = Math.min(5, used.rowCount);
      const allValues = used.values as unknown[][];
      const previewRows = allValues.slice(0, previewRowCount);
      const colCount = used.columnCount;

      // Column letters as header
      const colLetters = Array.from({ length: colCount }, (_, i) => colToLetter(i));
      const headerRow = `| | ${colLetters.join(" | ")} |`;
      const separator = `|---|${colLetters.map(() => "---").join("|")}|`;

      lines.push("");
      lines.push(`### Preview (first ${previewRowCount} rows)`);
      lines.push(headerRow);
      lines.push(separator);
      for (let r = 0; r < previewRows.length; r++) {
        const cells = previewRows[r].map((v) =>
          v === null || v === undefined || v === "" ? "" : typeof v === "string" ? v : typeof v === "number" || typeof v === "boolean" ? String(v) : JSON.stringify(v),
        );
        lines.push(`| ${r + 1} | ${cells.join(" | ")} |`);
      }
    }

    return lines.join("\n");
  });
}
