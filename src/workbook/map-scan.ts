/**
 * Workbook map scan — classifies used-range cells into formula / input /
 * label / error blocks using Office.js "special cells" queries.
 *
 * The scan never reads cell values or formulas: it only asks Excel for the
 * *addresses* of each cell class (the same engine behind "Go To Special"),
 * which keeps it fast even on very large sheets.
 *
 * Pure helpers (address-list splitting, rect parsing, count building) are
 * side-effect free so they can be unit tested without an Office host.
 */

import { excelRun, letterToCol } from "../excel/helpers.js";

// ============================================================================
// Types
// ============================================================================

/** Painted cell classes, in paint order (later classes paint over earlier). */
export const MAP_PAINT_ORDER = ["label", "input", "formula", "error"] as const;

export type MapCellClass = (typeof MAP_PAINT_ORDER)[number];

/** A rectangular block of cells in absolute 0-based sheet coordinates (inclusive). */
export interface CellRect {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** Used-range bounds of a sheet in absolute 0-based coordinates. */
export interface SheetMapBounds {
  startRow: number;
  startCol: number;
  rowCount: number;
  colCount: number;
}

export interface SheetMapCounts {
  formula: number;
  input: number;
  label: number;
  error: number;
  empty: number;
  total: number;
}

export interface SheetMapScan {
  sheetName: string;
  /** `null` when the sheet has no used range (empty sheet). */
  bounds: SheetMapBounds | null;
  rects: Record<MapCellClass, CellRect[]>;
  counts: SheetMapCounts;
}

/** Raw special-cells query results before class merging. */
export interface RawSpecialCellGroups {
  /** All formula cells (includes formula cells currently showing errors). */
  formulasAll: CellRect[];
  /** Formula cells whose current value is an error. */
  formulaErrors: CellRect[];
  /** Constant cells holding numbers or booleans (hardcoded inputs). */
  constantNumbers: CellRect[];
  /** Constant cells holding text (labels/headers). */
  constantText: CellRect[];
  /** Constant cells holding literal error values. */
  constantErrors: CellRect[];
}

export interface WorkbookMapSheets {
  sheetNames: string[];
  activeSheetName: string;
}

/** Thrown when the host Excel lacks the special-cells API (ExcelApi 1.9+). */
export class WorkbookMapUnsupportedError extends Error {
  constructor() {
    super("Workbook map requires the special-cells API (ExcelApi 1.9+).");
    this.name = "WorkbookMapUnsupportedError";
  }
}

// ============================================================================
// Pure helpers
// ============================================================================

/**
 * Split a RangeAreas-style comma-separated address list into individual
 * range addresses. Commas inside quoted sheet names ('P&L, FY26'!A1) are
 * preserved; doubled quotes inside quoted names are handled naturally
 * because each `'` toggles quote state twice.
 */
export function splitAddressList(list: string): string[] {
  const parts: string[] = [];
  let current = "";
  let inQuote = false;

  for (const ch of list) {
    if (ch === "'") {
      inQuote = !inQuote;
      current += ch;
      continue;
    }
    if (ch === "," && !inQuote) {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += ch;
  }
  parts.push(current.trim());

  return parts.filter((part) => part.length > 0);
}

/** Strip a leading sheet-name prefix ("Sheet1!" or "'My, Sheet'!") from an address. */
function stripSheetPrefix(address: string): string {
  const trimmed = address.trim();

  if (trimmed.startsWith("'")) {
    // Walk to the closing quote (doubled quotes escape a literal quote).
    let index = 1;
    while (index < trimmed.length) {
      if (trimmed[index] === "'") {
        if (trimmed[index + 1] === "'") {
          index += 2;
          continue;
        }
        // Closing quote — expect "!" right after it.
        if (trimmed[index + 1] === "!") {
          return trimmed.slice(index + 2);
        }
        break;
      }
      index += 1;
    }
    return trimmed;
  }

  const bang = trimmed.indexOf("!");
  return bang >= 0 ? trimmed.slice(bang + 1) : trimmed;
}

const CELL_TOKEN_PATTERN = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

function parseCellToken(token: string): { row: number; col: number } | null {
  const match = CELL_TOKEN_PATTERN.exec(token.trim());
  if (!match) return null;

  const letters = match[1];
  const digits = match[2];
  if (letters === undefined || digits === undefined) return null;

  const row = Number.parseInt(digits, 10) - 1;
  if (!Number.isSafeInteger(row) || row < 0) return null;

  return { row, col: letterToCol(letters.toUpperCase()) };
}

/**
 * Parse a single range address ("B3", "A1:C4", "Sheet1!$B$2:$D$10") into a
 * cell rect. Returns `null` for shapes the map cannot paint (full-row/column
 * references, malformed tokens).
 */
export function rectFromAddress(address: string): CellRect | null {
  const local = stripSheetPrefix(address);
  const tokens = local.split(":");
  if (tokens.length > 2) return null;

  const startToken = tokens[0];
  if (startToken === undefined) return null;

  const start = parseCellToken(startToken);
  if (!start) return null;

  const endToken = tokens[1];
  const end = endToken === undefined ? start : parseCellToken(endToken);
  if (!end) return null;

  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

/** Parse a comma-separated address list into rects, skipping unpaintable parts. */
export function rectsFromAddressList(list: string): CellRect[] {
  const rects: CellRect[] = [];
  for (const part of splitAddressList(list)) {
    const rect = rectFromAddress(part);
    if (rect) rects.push(rect);
  }
  return rects;
}

export function rectCellCount(rect: CellRect): number {
  return (rect.endRow - rect.startRow + 1) * (rect.endCol - rect.startCol + 1);
}

export function sumRectCells(rects: readonly CellRect[]): number {
  let total = 0;
  for (const rect of rects) total += rectCellCount(rect);
  return total;
}

export function emptyRawGroups(): RawSpecialCellGroups {
  return {
    formulasAll: [],
    formulaErrors: [],
    constantNumbers: [],
    constantText: [],
    constantErrors: [],
  };
}

/**
 * Merge raw special-cells groups into the painted class model.
 *
 * Error cells are a subset of formula/constant cells, so the formula count
 * excludes formula-error cells and error rects are painted last.
 */
export function buildSheetMapScan(
  sheetName: string,
  bounds: SheetMapBounds | null,
  raw: RawSpecialCellGroups,
): SheetMapScan {
  const errorRects = [...raw.formulaErrors, ...raw.constantErrors];
  const errorCount = sumRectCells(errorRects);
  const formulaCount = Math.max(0, sumRectCells(raw.formulasAll) - sumRectCells(raw.formulaErrors));
  const inputCount = sumRectCells(raw.constantNumbers);
  const labelCount = sumRectCells(raw.constantText);
  const total = bounds ? bounds.rowCount * bounds.colCount : 0;
  const empty = Math.max(0, total - (formulaCount + inputCount + labelCount + errorCount));

  return {
    sheetName,
    bounds,
    rects: {
      formula: raw.formulasAll,
      input: raw.constantNumbers,
      label: raw.constantText,
      error: errorRects,
    },
    counts: {
      formula: formulaCount,
      input: inputCount,
      label: labelCount,
      error: errorCount,
      empty,
      total,
    },
  };
}

// ============================================================================
// Office.js scan
// ============================================================================

/** Scan one sheet (active sheet when `sheetName` is omitted) into a map model. */
export async function scanSheetMap(sheetName?: string): Promise<SheetMapScan> {
  return excelRun(async (context) => {
    const sheet =
      sheetName !== undefined && sheetName.length > 0
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("rowIndex, columnIndex, rowCount, columnCount, isNullObject");

    // Older hosts (ExcelApi < 1.9) lack the special-cells API entirely.
    const usedRangeApi: Partial<Excel.Range> = usedRange;
    if (typeof usedRangeApi.getSpecialCellsOrNullObject !== "function") {
      throw new WorkbookMapUnsupportedError();
    }

    const formulasAll = usedRange.getSpecialCellsOrNullObject("Formulas", "All");
    const formulaErrors = usedRange.getSpecialCellsOrNullObject("Formulas", "Errors");
    const constantNumbers = usedRange.getSpecialCellsOrNullObject("Constants", "LogicalNumbers");
    const constantText = usedRange.getSpecialCellsOrNullObject("Constants", "Text");
    const constantErrors = usedRange.getSpecialCellsOrNullObject("Constants", "Errors");
    for (const areas of [formulasAll, formulaErrors, constantNumbers, constantText, constantErrors]) {
      areas.load("address, isNullObject");
    }

    await context.sync();

    if (usedRange.isNullObject) {
      return buildSheetMapScan(sheet.name, null, emptyRawGroups());
    }

    const bounds: SheetMapBounds = {
      startRow: usedRange.rowIndex,
      startCol: usedRange.columnIndex,
      rowCount: usedRange.rowCount,
      colCount: usedRange.columnCount,
    };

    const toRects = (areas: Excel.RangeAreas): CellRect[] =>
      areas.isNullObject ? [] : rectsFromAddressList(areas.address);

    return buildSheetMapScan(sheet.name, bounds, {
      formulasAll: toRects(formulasAll),
      formulaErrors: toRects(formulaErrors),
      constantNumbers: toRects(constantNumbers),
      constantText: toRects(constantText),
      constantErrors: toRects(constantErrors),
    });
  });
}

/** List visible sheet names plus the currently active sheet. */
export async function listVisibleSheets(): Promise<WorkbookMapSheets> {
  return excelRun(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name, items/visibility");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");

    await context.sync();

    const sheetNames = sheets.items
      .filter((item) => item.visibility === "Visible")
      .map((item) => item.name);

    return { sheetNames, activeSheetName: active.name };
  });
}

/** Activate a sheet and select a cell/range on it (map click-through). */
export async function selectSheetCell(sheetName: string, address: string): Promise<void> {
  return excelRun(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    sheet.getRange(address).select();
    await context.sync();
  });
}
