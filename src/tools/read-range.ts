/**
 * read_range — Read cell values, formulas, and optionally formatting.
 *
 * Three modes:
 * - "compact" (default): Markdown table of displayed values. Token-efficient.
 * - "csv": Raw CSV values only. Analysis-friendly.
 * - "detailed": Formulas, number formats and comments as well.
 *
 * Reads are bounded: whole-column/row addresses and ranges larger than one
 * page are clipped to the sheet's used range and returned one page at a time,
 * with the exact range to read next. Summaries (sheet state, formulas,
 * errors) come before the table so output truncation only cuts the bulk.
 */

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

/**
 * Cells and rows returned per call. Compact mode costs roughly 8 bytes per
 * cell plus 10 per row, so a full page is about 25 KB, half the output limit.
 */
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
  /** The address as given, without the sheet. */
  requested: string;
  /** Requested block after clipping to the used range; null when that leaves nothing. */
  target: CellBlock | null;
  clippedToUsedRange: boolean;
  page: CellBlock | null;
  values: CellValue[][];
  formulas: CellValue[][];
  /** R1C1 formulas, to group filled formulas; empty in csv mode. */
  formulasR1C1: CellValue[][];
  numberFormats: string[][];
  text: string[][];
  comments: CommentSummary[];
  state: RangeState | null;
}

const WHOLE_LINE_PATTERN = /^\$?(?:[A-Z]+:\$?[A-Z]+|\d+:\$?\d+)$/iu;

function cellCount(block: CellBlock): number {
  return block.rowCount * block.columnCount;
}

function firstPage(target: CellBlock): CellBlock {
  const columnCount = Math.min(target.columnCount, MAX_PAGE_COLUMNS);
  const rowCount = Math.min(target.rowCount, MAX_PAGE_ROWS, Math.max(1, Math.floor(READ_RANGE_PAGE_CELLS / columnCount)));
  return { ...target, rowCount, columnCount };
}

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

  // Whole columns/rows and anything bigger than a page only make sense where data is.
  const clip = WHOLE_LINE_PATTERN.test(resolved.address) || cellCount(requested) > READ_RANGE_PAGE_CELLS;
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

  const pageShape = firstPage(target);
  let rowCount = pageShape.rowCount;
  for (;;) {
    const page = { ...pageShape, rowCount };
    const range = sheet.getRangeByIndexes(page.rowIndex, page.columnIndex, page.rowCount, page.columnCount);
    range.load(mode === "csv" ? "values,formulas,numberFormat" : "values,formulas,formulasR1C1,numberFormat,text");
    try {
      await context.sync();
    } catch (error) {
      // Long text can exceed the host's response limit even within the cell budget.
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
    const state = await collectRangeState(context, sheet, page, spillCandidates(formulas, page));
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

function spillCandidates(formulas: CellValue[][], page: CellBlock): string[] {
  const cells: string[] = [];
  for (const [r, row] of formulas.entries()) {
    for (const [c, formula] of row.entries()) {
      if (typeof formula === "string" && formula.startsWith("=") && canSpill(formula)) {
        cells.push(`${colToLetter(page.columnIndex + c)}${page.rowIndex + r + 1}`);
      }
    }
  }
  return cells;
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
            // Ignore conventions lookup failures in read-only output.
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

/** Header, sheet note, paging note and sheet-state notes: the lines truncation must not cut. */
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

/** Values as displayed, with the stored value in brackets when the display hides it. */
function displayGrid(result: ReadRangePage): { grid: string[][]; usedRaw: boolean } {
  let usedRaw = false;
  const grid = result.values.map((row, r) => row.map((value, c) => {
    const cell = formatDisplayedCell(value, result.text[r]?.[c], result.numberFormats[r]?.[c]);
    usedRaw ||= cell.usedRaw;
    return cell.text;
  }));
  return { grid, usedRaw };
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

/** Group a format's cells into per-column runs: "A2:A6, E2:E3, K2". */
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

/**
 * Formula summary with filled blocks collapsed: cells in one column with the
 * same R1C1 formula in consecutive rows read "F2:F833: =C2*2 (filled, 832 cells)".
 */
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

  const { grid, usedRaw } = displayGrid(result);
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

/* ── CSV helpers (migrated from get-range-as-csv) ──────────────────── */

function toCsvField(value: CellValue): string {
  if (value === null || value === undefined) return "";
  const str = String(value);
  return /[",\n\r]/u.test(str) ? `"${str.replace(/"/g, '""')}"` : str;
}

function valuesToCsv(values: CellValue[][]): string {
  return values.map((row) => row.map((v) => toCsvField(v)).join(",")).join("\n");
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

  const csv = valuesToCsv(result.values);
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
