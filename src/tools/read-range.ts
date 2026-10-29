import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  blockAddress,
  colToLetter,
  excelRun,
  intersectBlocks,
  isCellInRange,
  isResponsePayloadLimitMessage,
  qualifiedAddress,
  resolveSheetRef,
  type CellBlock,
  type SheetSource,
} from "../excel/helpers.js";
import {
  findErrors,
  formatAsMarkdownTable,
  formatDisplayedCell,
  cellValueGridSchema,
  textGridSchema,
  type CellValue,
} from "../utils/format.js";
import { getErrorMessage } from "../utils/errors.js";
import { buildResolvedFormatLabels, getResolvedConventions, humanizeFormat } from "../conventions/index.js";
import { getAppStorage } from "../storage/local/app-storage.js";
import type { ReadRangeCsvDetails } from "./tool-details.js";
import { canSpill, collectRangeState, renderRangeState, type RangeState } from "./sheet-state.js";

const READ_RANGE_PAGE_CELLS = 2_500;
const MAX_PAGE_ROWS = 500;
const MAX_PAGE_COLUMNS = 100;
const MAX_LISTED_FORMULAS = 40;
const MAX_LISTED_ERRORS = 40;
const MAX_LISTED_FORMAT_AREAS = 20;

const schema = Type.Object({
  range: Type.String({
    description:
      'Cell range in A1 notation, e.g. "A1:D10", "Sheet2!A1:B5", "Data!A:C". ' +
      "Include the sheet name; without one the active sheet is used and the result says so.",
  }),
  mode: Type.Optional(
    Type.Union([Type.Literal("compact"), Type.Literal("csv"), Type.Literal("detailed")], {
      description:
        '"compact" (default): markdown table of values as displayed. ' +
        '"csv": raw values only as CSV. ' +
        '"detailed": includes formulas and number formats.',
    }),
  ),
});

type Params = Static<typeof schema>;
type Mode = "compact" | "csv" | "detailed";

interface CommentSummary {
  cell: string;
  content: string;
  author: string;
  resolved: boolean;
  replyCount: number;
}

interface ReadRangePage {
  sheetName: string;
  sheetSource: SheetSource;
  requested: string;
  target: CellBlock | null;
  clippedToUsedRange: boolean;
  page: CellBlock | null;
  values: CellValue[][];
  formulas: CellValue[][];
  formulasR1C1: CellValue[][];
  numberFormats: string[][];
  text: string[][];
  comments: CommentSummary[];
  state: RangeState | null;
}

const WHOLE_LINE_PATTERN = /^\$?(?:[A-Z]+:\$?[A-Z]+|\d+:\$?\d+)$/iu;

async function loadPage(context: Excel.RequestContext, ref: string, mode: Mode): Promise<ReadRangePage> {
  const resolved = await resolveSheetRef(context, ref);
  const { sheet } = resolved;
  const requestedRange = sheet.getRange(resolved.address);
  requestedRange.load("rowIndex,columnIndex,rowCount,columnCount");
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();

  const requested: CellBlock = {
    rowIndex: requestedRange.rowIndex,
    columnIndex: requestedRange.columnIndex,
    rowCount: requestedRange.rowCount,
    columnCount: requestedRange.columnCount,
  };
  const base = {
    sheetName: resolved.sheetName,
    sheetSource: resolved.sheetSource,
    requested: resolved.address,
  };

  const clip = WHOLE_LINE_PATTERN.test(resolved.address) || requested.rowCount * requested.columnCount > READ_RANGE_PAGE_CELLS;
  let target: CellBlock | null = requested;
  if (clip) {
    target = used.isNullObject
      ? null
      : intersectBlocks(requested, {
        rowIndex: used.rowIndex,
        columnIndex: used.columnIndex,
        rowCount: used.rowCount,
        columnCount: used.columnCount,
      });
  }
  const empty = { values: [], formulas: [], formulasR1C1: [], numberFormats: [], text: [], comments: [], state: null };
  if (target === null) {
    return { ...base, target: null, clippedToUsedRange: clip, page: null, ...empty };
  }

  const columnCount = Math.min(target.columnCount, MAX_PAGE_COLUMNS);
  let rowCount = Math.min(target.rowCount, MAX_PAGE_ROWS, Math.max(1, Math.floor(READ_RANGE_PAGE_CELLS / columnCount)));
  for (;;) {
    const page = { ...target, rowCount, columnCount };
    const range = sheet.getRangeByIndexes(page.rowIndex, page.columnIndex, page.rowCount, page.columnCount);
    range.load(mode === "csv" ? "values,formulas,numberFormat" : "values,formulas,formulasR1C1,numberFormat,text");
    try {
      await context.sync();
    } catch (error) {
      if (page.rowCount > 1 && isResponsePayloadLimitMessage(getErrorMessage(error))) {
        rowCount = Math.ceil(rowCount / 2);
        continue;
      }
      throw error;
    }

    const values = Value.Parse(cellValueGridSchema, range.values);
    const formulas = Value.Parse(cellValueGridSchema, range.formulas);
    const numberFormats = Value.Parse(textGridSchema, range.numberFormat);
    const text = mode === "csv" ? [] : Value.Parse(textGridSchema, range.text);
    const formulasR1C1 = mode === "csv" ? [] : Value.Parse(cellValueGridSchema, range.formulasR1C1);
    const spillCells: string[] = [];
    for (const [r, row] of formulas.entries()) {
      for (const [c, formula] of row.entries()) {
        if (typeof formula === "string" && formula.startsWith("=") && canSpill(formula)) {
          spillCells.push(`${colToLetter(page.columnIndex + c)}${page.rowIndex + r + 1}`);
        }
      }
    }
    const state = await collectRangeState(context, sheet, page, spillCells);
    const comments = mode === "detailed" ? await loadComments(context, sheet, blockAddress(page)) : [];

    return {
      ...base,
      target,
      clippedToUsedRange: clip,
      page,
      values,
      formulas,
      formulasR1C1,
      numberFormats,
      text,
      comments,
      state,
    };
  }
}

async function loadComments(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  pageAddress: string,
): Promise<CommentSummary[]> {
  const commentsCol = sheet.comments;
  commentsCol.load("items");
  await context.sync();
  if (commentsCol.items.length === 0) return [];

  const entries = commentsCol.items.map((comment) => {
    comment.load("content,authorName,resolved");
    const location = comment.getLocation();
    location.load("address");
    const replyCount = comment.replies.getCount();
    return { comment, location, replyCount };
  });
  await context.sync();

  const comments: CommentSummary[] = [];
  for (const { comment, location, replyCount } of entries) {
    const locBangIndex = location.address.indexOf("!");
    const locCell = locBangIndex >= 0 ? location.address.slice(locBangIndex + 1) : location.address;
    if (isCellInRange(locCell, pageAddress)) {
      comments.push({
        cell: locCell,
        content: comment.content,
        author: comment.authorName,
        resolved: comment.resolved,
        replyCount: replyCount.value,
      });
    }
  }
  return comments;
}

export function createReadRangeTool(): AgentTool<typeof schema> {
  return {
    name: "read_range",
    label: "Read Range",
    description:
      "Read cell values from an Excel range. Returns a markdown table of values as the user sees them " +
      "(compact mode), with notes on protection, merges, hidden/filtered rows, tables, spills, validation " +
      "and manual calculation. Large ranges and whole columns (A:A) are clipped to the data and paged; " +
      "the result names the range to read next. " +
      'Use mode "csv" for raw CSV values (analysis-friendly). ' +
      'Use mode "detailed" to also see formulas and number formats. ' +
      "Always read before modifying — never guess what's in the spreadsheet.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ReadRangeCsvDetails | undefined>> => {
      try {
        const mode: Mode = params.mode ?? "compact";
        const result = await excelRun((context) => loadPage(context, params.range, mode));

        const lines = renderPreamble(result);
        if (result.page === null) {
          lines.push("", "_All cells are empty._");
          return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
        }

        if (mode === "csv") return formatCsvOutput(lines, result, result.page);

        let resolvedLabels: Map<string, string> | undefined;
        if (mode === "detailed") {
          try {
            const conventions = await getResolvedConventions(getAppStorage().settings);
            resolvedLabels = buildResolvedFormatLabels(conventions);
          } catch {
          }
        }
        return formatTable(lines, result, result.page, mode, resolvedLabels);
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error reading "${params.range}": ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

function renderPreamble(result: ReadRangePage): string[] {
  const lines: string[] = [];
  const { page, target } = result;
  const shownAddress = page ? blockAddress(page) : result.requested;
  const size = page ? ` (${page.rowCount}×${page.columnCount})` : "";
  lines.push(`**${qualifiedAddress(result.sheetName, shownAddress)}**${size}`);

  if (result.sheetSource === "active") {
    lines.push(`Sheet: "${result.sheetName}" (the active sheet; the address named no sheet).`);
  }

  if (target === null) {
    lines.push(`Requested ${result.requested}: the sheet has no data there (used range checked).`);
    return lines;
  }
  if (result.clippedToUsedRange && blockAddress(target) !== result.requested.replace(/\$/gu, "").toUpperCase()) {
    lines.push(`Requested ${result.requested}; clipped to the data: ${blockAddress(target)} (${target.rowCount}×${target.columnCount}).`);
  }
  if (page && (page.rowCount < target.rowCount || page.columnCount < target.columnCount)) {
    lines.push(...renderPagingNote(result.sheetName, target, page));
  }
  if (result.state) {
    lines.push(...renderRangeState(result.state, colToLetter));
  }
  if (page) lines.push(...renderValueNotes(result.values, page));
  return lines;
}

const NUMERIC_TEXT_PATTERN = /^[-+]?(?:\d{1,3}(?:,\d{3})+|0|[1-9]\d*)?(?:\.\d+)?%?$/u;
const MAX_LISTED_TEXT_NUMBERS = 10;
const AGGREGATE_HINT_MIN_NUMBERS = 50;
const MAX_SUMMED_COLUMNS = 10;

function renderValueNotes(values: CellValue[][], page: CellBlock): string[] {
  const lines: string[] = [];
  const textNumbers: string[] = [];
  const numericColumns = new Map<number, { count: number; sum: number }>();
  for (const [r, row] of values.entries()) {
    for (const [c, value] of row.entries()) {
      if (typeof value === "string" && /\d/u.test(value.trim()) && NUMERIC_TEXT_PATTERN.test(value.trim())) {
        textNumbers.push(`${colToLetter(page.columnIndex + c)}${page.rowIndex + r + 1} "${value}"`);
      } else if (typeof value === "number") {
        const column = numericColumns.get(c) ?? { count: 0, sum: 0 };
        column.count += 1;
        column.sum += value;
        numericColumns.set(c, column);
      }
    }
  }
  if (textNumbers.length > 0) {
    lines.push(
      `⚠️ Numbers stored as text (SUM, AVERAGE and other aggregates skip them; they look like numbers in the table): ` +
      listCapped(textNumbers, MAX_LISTED_TEXT_NUMBERS, "cells"),
    );
  }
  const longColumns = [...numericColumns.entries()]
    .filter(([, column]) => column.count >= AGGREGATE_HINT_MIN_NUMBERS)
    .slice(0, MAX_SUMMED_COLUMNS);
  if (longColumns.length > 0) {
    const sums = longColumns.map(([c, { count, sum }]) => {
      const shown = Number(sum.toPrecision(12)).toLocaleString("en-US", { maximumFractionDigits: 10 });
      return `${colToLetter(page.columnIndex + c)} ${shown} (${count} numbers)`;
    });
    lines.push(
      `Column sums on this page (numbers only): ${sums.join(", ")}. ` +
      "For other totals, counts or averages over many cells, compute them (execute_office_js with " +
      "`context.workbook.functions`, or python_run); do not add up values yourself.",
    );
  }
  return lines;
}

function renderPagingNote(sheetName: string, target: CellBlock, page: CellBlock): string[] {
  const lines: string[] = [];
  const firstRow = page.rowIndex + 1;
  const lastRow = page.rowIndex + page.rowCount;
  const targetLastRow = target.rowIndex + target.rowCount;
  if (page.rowCount < target.rowCount) {
    const nextRows = Math.min(page.rowCount, targetLastRow - lastRow);
    const next = blockAddress({ ...page, rowIndex: lastRow, rowCount: nextRows });
    lines.push(
      `Page: rows ${firstRow}-${lastRow} of ${target.rowIndex + 1}-${targetLastRow} ` +
      `(pages hold at most ${MAX_PAGE_ROWS} rows or ${READ_RANGE_PAGE_CELLS.toLocaleString("en-US")} cells). ` +
      `Next page: read_range "${qualifiedAddress(sheetName, next)}". ` +
      "For whole-sheet analysis prefer search_workbook or python over paging.",
    );
  }
  if (page.columnCount < target.columnCount) {
    const nextColumn = page.columnIndex + page.columnCount;
    const nextColumns = Math.min(page.columnCount, target.columnIndex + target.columnCount - nextColumn);
    const next = blockAddress({ ...page, columnIndex: nextColumn, columnCount: nextColumns });
    lines.push(
      `Columns ${colToLetter(page.columnIndex)}-${colToLetter(nextColumn - 1)} of ` +
      `${colToLetter(target.columnIndex)}-${colToLetter(target.columnIndex + target.columnCount - 1)}. ` +
      `Next columns: read_range "${qualifiedAddress(sheetName, next)}".`,
    );
  }
  return lines;
}

function hasAnyNonEmptyCell(values: CellValue[][]): boolean {
  return values.some((row) => row.some((v) => v !== null && v !== undefined && v !== ""));
}

function listCapped(items: string[], max: number, noun: string): string {
  if (items.length <= max) return items.join(", ");
  return `${items.slice(0, max).join(", ")}, … (+${items.length - max} more ${noun})`;
}

function formatAsExcelMarkdownTable(values: string[][], page: CellBlock): string {
  const header: string[] = [""];
  for (let c = 0; c < page.columnCount; c++) header.push(colToLetter(page.columnIndex + c));
  const rows: Array<Array<string | number>> = [header];
  for (let r = 0; r < values.length; r++) {
    const row: Array<string | number> = [page.rowIndex + r + 1, ...(values[r] ?? [])];
    while (row.length < page.columnCount + 1) row.push("");
    rows.push(row);
  }
  return formatAsMarkdownTable(rows);
}

function formatCellRuns(cells: Array<{ row: number; col: number }>): string[] {
  const byColumn = new Map<number, number[]>();
  for (const { row, col } of cells) {
    const rows = byColumn.get(col) ?? [];
    rows.push(row);
    byColumn.set(col, rows);
  }
  const runs: string[] = [];
  for (const [col, rows] of [...byColumn.entries()].sort((a, b) => a[0] - b[0])) {
    rows.sort((a, b) => a - b);
    let start = rows[0];
    let prev = start;
    for (const row of [...rows.slice(1), undefined]) {
      if (row !== undefined && prev !== undefined && row === prev + 1) {
        prev = row;
        continue;
      }
      if (start !== undefined && prev !== undefined) {
        const letter = colToLetter(col);
        runs.push(start === prev ? `${letter}${start}` : `${letter}${start}:${letter}${prev}`);
      }
      start = row;
      prev = row;
    }
  }
  return runs;
}

function summarizeFormulas(result: ReadRangePage, page: CellBlock): string[] {
  interface Run { col: number; startRow: number; endRow: number; formula: string; r1c1: string }
  const runs: Run[] = [];
  const open = new Map<number, Run>();
  for (const [r, row] of result.formulas.entries()) {
    for (const [c, formula] of row.entries()) {
      const current = open.get(c);
      if (typeof formula !== "string" || !formula.startsWith("=")) {
        open.delete(c);
        continue;
      }
      const r1c1Value = result.formulasR1C1[r]?.[c];
      const r1c1 = typeof r1c1Value === "string" ? r1c1Value : formula;
      if (current && current.r1c1 === r1c1 && current.endRow === r - 1) {
        current.endRow = r;
        continue;
      }
      const run = { col: c, startRow: r, endRow: r, formula, r1c1 };
      runs.push(run);
      open.set(c, run);
    }
  }
  return runs.map((run) => {
    const column = colToLetter(page.columnIndex + run.col);
    const first = `${column}${page.rowIndex + run.startRow + 1}`;
    if (run.endRow === run.startRow) return `${first}: ${run.formula}`;
    const count = run.endRow - run.startRow + 1;
    return `${first}:${column}${page.rowIndex + run.endRow + 1}: ${run.formula} (filled, ${count} cells)`;
  });
}

function formatTable(
  lines: string[],
  result: ReadRangePage,
  page: CellBlock,
  mode: "compact" | "detailed",
  resolvedLabels?: Map<string, string>,
): AgentToolResult<undefined> {
  const startCell = `${colToLetter(page.columnIndex)}${page.rowIndex + 1}`;
  const formulas = summarizeFormulas(result, page);
  const errors = findErrors(result.values, startCell);

  if (!hasAnyNonEmptyCell(result.values) && formulas.length === 0 && errors.length === 0) {
    lines.push("", "_All cells are empty._");
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  if (mode === "compact") {
    if (formulas.length > 0) lines.push("", `**Formulas:** ${listCapped(formulas, MAX_LISTED_FORMULAS, "formulas")}`);
    if (errors.length > 0) {
      const listed = errors.map((e) => `${e.address}=${e.error}`);
      lines.push("", `⚠️ **Errors:** ${listCapped(listed, MAX_LISTED_ERRORS, "errors")}`);
    }
  } else {
    if (formulas.length > 0) {
      lines.push("", "### Formulas");
      for (const f of formulas.slice(0, MAX_LISTED_FORMULAS)) lines.push(`- ${f}`);
      if (formulas.length > MAX_LISTED_FORMULAS) lines.push(`- … +${formulas.length - MAX_LISTED_FORMULAS} more formulas`);
    }
    lines.push(...renderNumberFormats(result, page, resolvedLabels));
    if (errors.length > 0) {
      lines.push("", "### ⚠️ Errors");
      for (const e of errors.slice(0, MAX_LISTED_ERRORS)) lines.push(`- ${e.address}: ${e.error}`);
      if (errors.length > MAX_LISTED_ERRORS) lines.push(`- … +${errors.length - MAX_LISTED_ERRORS} more errors`);
    }
    if (result.comments.length > 0) {
      lines.push("", "### Comments");
      for (const c of result.comments) {
        const resolved = c.resolved ? " ✓" : "";
        const replies = c.replyCount > 0 ? ` (${c.replyCount} ${c.replyCount === 1 ? "reply" : "replies"})` : "";
        lines.push(`- **${c.cell}**: "${c.content}" — *${c.author}*${resolved}${replies}`);
      }
    }
  }

  let usedRaw = false;
  const grid = result.values.map((row, r) => row.map((value, c) => {
    const cell = formatDisplayedCell(value, result.text[r]?.[c], result.numberFormats[r]?.[c]);
    usedRaw ||= cell.usedRaw;
    return cell.text;
  }));
  lines.push("");
  if (mode === "detailed") lines.push("### Values");
  if (usedRaw) lines.push("_Cells show the displayed text; [brackets] hold the stored value where the display differs._");
  lines.push(formatAsExcelMarkdownTable(grid, page));

  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}

function renderNumberFormats(
  result: ReadRangePage,
  page: CellBlock,
  resolvedLabels?: Map<string, string>,
): string[] {
  const formatCells = new Map<string, Array<{ row: number; col: number }>>();
  for (const [r, row] of result.numberFormats.entries()) {
    for (const [c, fmt] of row.entries()) {
      if (fmt === "" || fmt === "General") continue;
      const cells = formatCells.get(fmt) ?? [];
      cells.push({ row: page.rowIndex + r + 1, col: page.columnIndex + c });
      formatCells.set(fmt, cells);
    }
  }
  if (formatCells.size === 0) return [];
  const lines = ["", "### Number Formats"];
  for (const [fmt, cells] of formatCells) {
    const label = humanizeFormat(fmt, resolvedLabels);
    const display = label !== fmt ? `**${label}** (\`${fmt}\`)` : `\`${fmt}\``;
    lines.push(`- ${display} → ${listCapped(formatCellRuns(cells), MAX_LISTED_FORMAT_AREAS, "areas")}`);
  }
  return lines;
}

function toCsvField(value: CellValue): string {
  if (value === null || value === undefined) return "";
  const str = String(value);
  return /[",\n\r]/u.test(str) ? `"${str.replace(/"/g, '""')}"` : str;
}

function formatCsvOutput(
  lines: string[],
  result: ReadRangePage,
  page: CellBlock,
): AgentToolResult<ReadRangeCsvDetails | undefined> {
  if (!hasAnyNonEmptyCell(result.values)) {
    lines.push("", "(empty)");
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  const csv = result.values.map((row) => row.map(toCsvField).join(",")).join("\n");
  lines.push("", "```csv", csv, "```");

  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details: {
      kind: "read_range_csv",
      startCol: page.columnIndex,
      startRow: page.rowIndex + 1,
      values: result.values,
      csv,
    },
  };
}
