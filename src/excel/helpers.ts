/**
 * Office.js wrappers — thin abstraction over Excel.run() with
 * address parsing, error handling, and guarded API calls.
 */

// ============================================================================
// Types
// ============================================================================

// Office.js is loaded at runtime by Excel, but its types are available via
// `@types/office-js` (global `Excel` namespace/value).

/** Result of parsing a range reference like "Sheet1!A1:B5" */
export interface RangeRef {
  sheet?: string;   // undefined = active sheet
  address: string;  // e.g. "A1:B5"
}

// ============================================================================
// Excel.run wrapper
// ============================================================================

/**
 * Run an Office.js Excel operation with error handling.
 * Wraps Excel.run() and provides typed context.
 */
export async function excelRun<T>(fn: (context: Excel.RequestContext) => Promise<T>): Promise<T> {
  return Excel.run(fn);
}

// ============================================================================
// Address parsing
// ============================================================================

/** Parse "Sheet1!A1:B5" → { sheet: "Sheet1", address: "A1:B5" } */
export function parseRangeRef(ref: string): RangeRef {
  if (ref.includes("!")) {
    const idx = ref.lastIndexOf("!");
    const sheet = ref
      .substring(0, idx)
      .replace(/^'|'$/g, "")
      .replace(/''/g, "'"); // strip quotes + unescape
    return { sheet, address: ref.substring(idx + 1) };
  }
  return { address: ref };
}

/** Get a Range object from a context, resolving sheet name if present */
export function getRange(context: Excel.RequestContext, ref: string): { sheet: Excel.Worksheet; range: Excel.Range } {
  const parsed = parseRangeRef(ref);
  const sheet = parsed.sheet
    ? context.workbook.worksheets.getItem(parsed.sheet)
    : context.workbook.worksheets.getActiveWorksheet();
  return { sheet, range: sheet.getRange(parsed.address) };
}

/** Whether a sheet came from the address itself or from the active sheet. */
export type SheetSource = "explicit" | "active";

export interface ResolvedSheetRef {
  sheet: Excel.Worksheet;
  sheetName: string;
  /** Address without the sheet prefix, e.g. "A1:B5" or "A:A". */
  address: string;
  sheetSource: SheetSource;
}

/**
 * Resolve a range reference to one named sheet before any range work, so a
 * tool reads the sheet it reports even if the user switches tabs mid-call.
 * An unknown sheet name fails with the list of existing sheets.
 */
export async function resolveSheetRef(context: Excel.RequestContext, ref: string): Promise<ResolvedSheetRef> {
  const parsed = parseRangeRef(ref.trim());
  if (parsed.sheet === undefined) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    return { sheet, sheetName: sheet.name, address: parsed.address, sheetSource: "active" };
  }

  const sheet = context.workbook.worksheets.getItemOrNullObject(parsed.sheet);
  sheet.load("name");
  await context.sync();
  if (sheet.isNullObject) {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const names = sheets.items.map((item) => `"${item.name}"`).join(", ");
    throw new Error(`Sheet "${parsed.sheet}" not found. Sheets in this workbook: ${names}.`);
  }
  return { sheet, sheetName: sheet.name, address: parsed.address, sheetSource: "explicit" };
}

/** A rectangular block of cells by zero-based indexes, as Office.js reports them. */
export interface CellBlock {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

/** A1 address of a block, e.g. "A1:E500" (or "B3" for one cell). */
export function blockAddress(block: CellBlock): string {
  const start = cellAddress(block.columnIndex, block.rowIndex + 1);
  if (block.rowCount === 1 && block.columnCount === 1) return start;
  const end = cellAddress(block.columnIndex + block.columnCount - 1, block.rowIndex + block.rowCount);
  return `${start}:${end}`;
}

export function intersectBlocks(a: CellBlock, b: CellBlock): CellBlock | null {
  const rowStart = Math.max(a.rowIndex, b.rowIndex);
  const colStart = Math.max(a.columnIndex, b.columnIndex);
  const rowEnd = Math.min(a.rowIndex + a.rowCount, b.rowIndex + b.rowCount);
  const colEnd = Math.min(a.columnIndex + a.columnCount, b.columnIndex + b.columnCount);
  if (rowEnd <= rowStart || colEnd <= colStart) return null;
  return { rowIndex: rowStart, columnIndex: colStart, rowCount: rowEnd - rowStart, columnCount: colEnd - colStart };
}

/** Excel's "response payload size has exceeded the limit" (about 5 MB per sync on the web). */
export function isResponsePayloadLimitMessage(message: string): boolean {
  return /payload size has exceeded|ResponsePayloadSizeLimitExceeded/iu.test(message);
}

/** Build a fully-qualified address like "Sheet1!A1:B5" */
export function qualifiedAddress(sheetName: string, address: string): string {
  // Strip sheet prefix if already in the address
  const clean = address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
  const escaped = sheetName.replace(/'/g, "''");
  const needsQuote = /[\s'!]/.test(sheetName);
  const quoted = needsQuote ? `'${escaped}'` : sheetName;
  return `${quoted}!${clean}`;
}

// ============================================================================
// Column / address math
// ============================================================================

/** Convert 0-indexed column to letter (0=A, 25=Z, 26=AA) */
export function colToLetter(col: number): string {
  let letter = "";
  let c = col;
  while (c >= 0) {
    letter = String.fromCharCode((c % 26) + 65) + letter;
    c = Math.floor(c / 26) - 1;
  }
  return letter;
}

/** Convert column letter to 0-indexed number (A=0, Z=25, AA=26) */
export function letterToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1; // 0-indexed
}

/** Parse cell address "B3" → { col: 1, row: 3 } (col is 0-indexed, row is 1-indexed) */
export function parseCell(cell: string): { col: number; row: number } {
  const clean = cell.includes("!") ? cell.slice(cell.indexOf("!") + 1) : cell;
  const match = clean.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) throw new Error(`Invalid cell address: ${cell}`);

  const colLetters = match[1];
  const rowDigits = match[2];
  if (colLetters === undefined || rowDigits === undefined) {
    throw new Error(`Invalid cell address: ${cell}`);
  }

  return { col: letterToCol(colLetters.toUpperCase()), row: parseInt(rowDigits, 10) };
}

/** Build cell address from 0-indexed col and 1-indexed row */
export function cellAddress(col: number, row: number): string {
  return `${colToLetter(col)}${row}`;
}

/** Check if a cell address falls within a range address. */
export function isCellInRange(cellAddress: string, rangeAddress: string): boolean {
  const bangIndex = rangeAddress.indexOf("!");
  const cleanRange = bangIndex >= 0 ? rangeAddress.slice(bangIndex + 1) : rangeAddress;
  const colonIndex = cleanRange.indexOf(":");
  const startAddress = colonIndex >= 0 ? cleanRange.slice(0, colonIndex) : cleanRange;
  const endAddress = colonIndex >= 0 ? cleanRange.slice(colonIndex + 1) : cleanRange;
  const start = parseCell(startAddress);
  const end = parseCell(endAddress);
  const cell = parseCell(cellAddress);
  return (
    cell.col >= start.col &&
    cell.col <= end.col &&
    cell.row >= start.row &&
    cell.row <= end.row
  );
}

/**
 * Compute the end address for a 2D array starting at a given cell.
 * E.g. startCell="B2", rows=3, cols=4 → "B2:E4"
 */
export function computeRangeAddress(startCell: string, rows: number, cols: number): string {
  const { col, row } = parseCell(startCell);
  return `${startCell}:${colToLetter(col + cols - 1)}${row + rows - 1}`;
}

/**
 * Compute cell address at offset from a range's start address.
 * E.g. rangeStart="B2", row offset 1, col offset 2 → "D3"
 */
export function cellAtOffset(rangeStart: string, rowOffset: number, colOffset: number): string {
  const { col, row } = parseCell(rangeStart);
  return cellAddress(col + colOffset, row + rowOffset);
}

// ============================================================================
// Guarded API calls
// ============================================================================

function isOfficeItemNotFound(error: unknown): boolean {
  if (typeof error !== "object" || error === null) return false;

  if ("code" in error && typeof error.code === "string") {
    return error.code === "ItemNotFound";
  }

  if ("message" in error && typeof error.message === "string") {
    return /item\s*not\s*found/iu.test(error.message);
  }

  return false;
}

/**
 * Safely call getDirectPrecedents().
 *
 * Returns:
 * - `string[][]` when API succeeds
 * - `[]` when the API is available but there are no precedents
 * - `null` when the API is unavailable/failing (so callers can use fallback logic)
 */
export async function getDirectPrecedentsSafe(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<string[][] | null> {
  try {
    const precedents = range.getDirectPrecedents();
    precedents.load("addresses");
    await context.sync();
    // WorkbookRangeAreas.addresses is string[]; each entry may itself contain
    // a comma-separated list of address blocks.
    return precedents.addresses
      .map((s) => s.split(",").map((x) => x.trim()).filter(Boolean));
  } catch (error) {
    if (isOfficeItemNotFound(error)) {
      return [];
    }
    return null;
  }
}

/**
 * Safely call getDirectDependents().
 *
 * Returns:
 * - `string[][]` when API succeeds
 * - `[]` when the API is available but there are no dependents
 * - `null` when the API is unavailable/failing (so callers can use fallback logic)
 */
export async function getDirectDependentsSafe(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<string[][] | null> {
  try {
    const dependents = range.getDirectDependents();
    dependents.load("addresses");
    await context.sync();
    return dependents.addresses
      .map((s) => s.split(",").map((x) => x.trim()).filter(Boolean));
  } catch (error) {
    if (isOfficeItemNotFound(error)) {
      return [];
    }
    return null;
  }
}

/** Pad a 2D array so all rows have the same length */
export function padValues(values: unknown[][]): { padded: unknown[][]; rows: number; cols: number } {
  const rows = values.length;
  const cols = Math.max(...values.map((r) => r.length));
  const padded = values.map((row) => {
    const r = [...row];
    while (r.length < cols) r.push("");
    return r;
  });
  return { padded, rows, cols };
}
