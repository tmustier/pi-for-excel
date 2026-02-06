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
    const idx = ref.indexOf("!");
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

/** Build a fully-qualified address like "Sheet1!A1:B5" */
export function qualifiedAddress(sheetName: string, address: string): string {
  // Strip sheet prefix if already in the address
  const clean = address.includes("!") ? address.split("!")[1] : address;
  const escaped = sheetName.replace(/'/g, "''");
  const needsQuote = /[\s']/.test(sheetName);
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
  const clean = cell.includes("!") ? cell.split("!")[1] : cell;
  const match = clean.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) throw new Error(`Invalid cell address: ${cell}`);
  return { col: letterToCol(match[1].toUpperCase()), row: parseInt(match[2], 10) };
}

/** Build cell address from 0-indexed col and 1-indexed row */
export function cellAddress(col: number, row: number): string {
  return `${colToLetter(col)}${row}`;
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

/**
 * Safely call getDirectPrecedents() — returns null if it throws
 * (fails on empty cells, preview API).
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
  } catch {
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
