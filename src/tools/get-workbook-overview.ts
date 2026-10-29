import { t } from "../language/index.js";
import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { blockAddress, excelRun, colToLetter, parseRangeRef, type ExcelRunOptions } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import { cellValueGridSchema, formatAsMarkdownTable, formatDisplayedCell, textGridSchema } from "../utils/format.js";
import { collectRangeState, parseBlock, renderRangeState } from "./sheet-state.js";

const PREVIEW_ROWS = 5;
const PREVIEW_COLUMNS = 20;
const MAX_OVERVIEW_SPILL_CHECKS = 100;

function formulaCellAddresses(address: string, limit: number): string[] {
  const cells: string[] = [];
  for (const part of address.split(",")) {
    const block = parseBlock(part.trim());
    if (!block) continue;
    for (let r = 0; r < block.rowCount; r++) {
      for (let c = 0; c < block.columnCount; c++) {
        if (cells.length >= limit) return cells;
        cells.push(blockAddress({ rowIndex: block.rowIndex + r, columnIndex: block.columnIndex + c, rowCount: 1, columnCount: 1 }));
      }
    }
  }
  return cells;
}

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
    label: t("tools.workbookOverview"),
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
          : `${await buildOverview()}${await buildWorkbookStateNotes()}`;
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error getting workbook overview: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

const MAX_HEADER_COLUMNS = 200;

async function loadHeaderRow(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  used: Excel.Range,
): Promise<{ labels: string[]; rowNote: string }> {
  if (used.isNullObject) return { labels: [], rowNote: "" };
  const row = sheet.getRangeByIndexes(used.rowIndex, used.columnIndex, 1, Math.min(used.columnCount, MAX_HEADER_COLUMNS));
  row.load("values");
  await context.sync();
  const values = Value.Parse(cellValueGridSchema, row.values)[0] ?? [];
  const labels = values.filter((v) => v !== null && v !== undefined && v !== "").map((v) => String(v));
  return { labels, rowNote: used.rowIndex === 0 ? "" : ` (row ${used.rowIndex + 1})` };
}

async function buildWorkbookStateNotes(): Promise<string> {
  return excelRun(async (context) => {
    const application = context.workbook.application;
    application.load("calculationMode");
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const protections = sheets.items.map((sheet) => {
      sheet.protection.load("protected");
      return { name: sheet.name, protection: sheet.protection };
    });
    await context.sync();

    const lines: string[] = [];
    if (application.calculationMode !== "Automatic") {
      lines.push(`Calculation mode: ${application.calculationMode} (formula results may be stale until recalculated)`);
    }
    const protectedSheets = protections.filter((item) => item.protection.protected).map((item) => item.name);
    if (protectedSheets.length > 0) {
      lines.push(`Protected sheets: ${protectedSheets.join(", ")}`);
    }
    return lines.length > 0 ? `\n\n### Current state\n${lines.map((line) => `- ${line}`).join("\n")}` : "";
  });
}

export async function buildOverview(runOptions?: ExcelRunOptions): Promise<string> {
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
      used.load("rowCount,columnCount,rowIndex,columnIndex,address");

      const tables = sheet.tables;
      tables.load("items/name,items/columns/count,items/rows/count");

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

      const header = await loadHeaderRow(context, sheet, used);

      lines.push(
        `${sheet.position + 1}. **${sheet.name}**${visibility} — ${dims}`,
      );

      if (header.labels.length > 0) {
        const display = header.labels.length > 8
          ? header.labels.slice(0, 8).join(", ") + `, … (+${header.labels.length - 8} more)`
          : header.labels.join(", ");
        lines.push(`   Headers${header.rowNote}: ${display}`);
      }

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

    const visibleNames = names.items.filter((n) => n.visible);
    if (visibleNames.length > 0) {
      lines.push("");
      lines.push(`### Named Ranges (${visibleNames.length})`);
      for (const n of visibleNames) {
        lines.push(`- **${n.name}** = ${n.value}`);
      }
    }

    return lines.join("\n");
  }, runOptions);
}

async function buildSheetDetail(sheetName: string): Promise<string> {
  return excelRun(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load("name,visibility");
    await context.sync();
    if (sheet.isNullObject) {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const existing = sheets.items.map((item) => `"${item.name}"`).join(", ");
      throw new Error(`Sheet "${sheetName}" not found. Sheets in this workbook: ${existing}.`);
    }

    const used = sheet.getUsedRangeOrNullObject();
    used.load("rowCount,columnCount,rowIndex,columnIndex,address");
    const formulaCells = used.getSpecialCellsOrNullObject("Formulas");
    formulaCells.load("address");
    const freeze = sheet.freezePanes.getLocationOrNullObject();
    freeze.load("address");

    const tables = sheet.tables;
    tables.load("items/name,items/columns/count,items/rows/count");

    const names = context.workbook.names;
    names.load("items/name,items/value,items/visible");

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

    const visibility = sheet.visibility === "Visible" ? "" : ` (${sheet.visibility})`;
    const dims = used.isNullObject
      ? "empty"
      : `${used.rowCount} rows × ${used.columnCount} cols`;
    lines.push(`## Sheet: ${sheet.name}${visibility}`);
    lines.push(`Dimensions: ${dims}`);

    const header = await loadHeaderRow(context, sheet, used);
    if (header.labels.length > 0) {
      lines.push(`Headers${header.rowNote}: ${header.labels.join(", ")}`);
    }

    if (tables.items.length > 0) {
      lines.push("");
      lines.push(`### Tables (${tables.items.length})`);
      for (const table of tables.items) {
        lines.push(
          `- **${table.name}** (${table.rows.count} rows × ${table.columns.count} cols)`,
        );
      }
    }

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

    if (!used.isNullObject) {
      const usedBlock = {
        rowIndex: used.rowIndex,
        columnIndex: used.columnIndex,
        rowCount: used.rowCount,
        columnCount: used.columnCount,
      };
      const state = await collectRangeState(
        context,
        sheet,
        usedBlock,
        formulaCells.isNullObject ? [] : formulaCellAddresses(formulaCells.address, MAX_OVERVIEW_SPILL_CHECKS),
      );
      const notes = renderRangeState(state, colToLetter);
      if (!freeze.isNullObject) notes.push(`Frozen panes: ${parseRangeRef(freeze.address).address}`);
      if (notes.length > 0) {
        lines.push("");
        lines.push("### Sheet state");
        for (const note of notes) lines.push(`- ${note}`);
      }

      const previewRowCount = Math.min(PREVIEW_ROWS, used.rowCount);
      const previewColumnCount = Math.min(PREVIEW_COLUMNS, used.columnCount);
      const preview = sheet.getRangeByIndexes(used.rowIndex, used.columnIndex, previewRowCount, previewColumnCount);
      preview.load("values,text,numberFormat");
      await context.sync();

      const values = Value.Parse(cellValueGridSchema, preview.values);
      const text = Value.Parse(textGridSchema, preview.text);
      const numberFormats = Value.Parse(textGridSchema, preview.numberFormat);
      const header: Array<string | number> = [""];
      for (let c = 0; c < previewColumnCount; c++) header.push(colToLetter(used.columnIndex + c));
      const rows: Array<Array<string | number>> = [header];
      for (const [r, row] of values.entries()) {
        rows.push([
          used.rowIndex + r + 1,
          ...row.map((value, c) => formatDisplayedCell(value, text[r]?.[c], numberFormats[r]?.[c]).text),
        ]);
      }
      const columnsNote = previewColumnCount < used.columnCount
        ? `, first ${previewColumnCount} of ${used.columnCount} columns`
        : "";
      lines.push("");
      lines.push(`### Preview (first ${previewRowCount} rows${columnsNote}, as displayed)`);
      lines.push(formatAsMarkdownTable(rows));
    }

    return lines.join("\n");
  });
}
