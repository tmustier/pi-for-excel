/**
 * Formatting utilities for LLM-facing output.
 * Converts Excel data into compact, token-efficient text.
 */

import { Type } from "typebox";

import { colToLetter, parseCell } from "../excel/helpers.js";

/** A cell value as Office.js and WPS return it in `range.values`. */
export type CellValue = string | number | boolean | null | undefined;

/** Parse `range.values` / `range.formulas` with `Value.Parse`. */
export const cellValueGridSchema = Type.Refine(
  Type.Unsafe<CellValue[][]>({}),
  (raw) => Array.isArray(raw) && raw.every((row) => Array.isArray(row) && row.every((cell) =>
    cell === null || typeof cell === "string" || typeof cell === "number" || typeof cell === "boolean")),
  () => "cell values must be a grid of strings, numbers, booleans or nulls",
);

/** Parse `range.text` / `range.numberFormat` with `Value.Parse`. */
export const textGridSchema = Type.Refine(
  Type.Unsafe<string[][]>({}),
  (raw) => Array.isArray(raw) && raw.every((row) => Array.isArray(row) && row.every((cell) => typeof cell === "string")),
  () => "cell text must be a grid of strings",
);

/**
 * Render a raw cell value as text. Numbers use the shortest exact form
 * (`String`), so 0 is "0" and no digits are lost.
 */
export function formatCellValue(v: CellValue): string {
  if (v === null || v === undefined) return "";
  return typeof v === "string" ? v : String(v);
}

/**
 * Escape text for one markdown table cell: backslashes, pipes and line breaks
 * would otherwise split or merge grid cells. Line breaks become a literal `\n`.
 */
export function escapeMarkdownCell(text: string): string {
  return text
    .replace(/\\/g, "\\\\")
    .replace(/\|/g, "\\|")
    .replace(/\r\n|\r|\n/g, "\\n");
}

/** True when displayed text such as "£1,234.50", "12.5%" or "(3.00)" denotes exactly `value`. */
function displayDenotesValue(text: string, value: number): boolean {
  let s = text.trim();
  let sign = 1;
  if (/^\(.*\)$/u.test(s)) {
    sign = -1;
    s = s.slice(1, -1);
  }
  const percent = s.endsWith("%");
  const numeric = s.replace(/[^0-9.eE+-]/gu, "");
  if (numeric === "" || !/\d/u.test(numeric)) return false;
  const parsed = Number(numeric);
  if (!Number.isFinite(parsed)) return false;
  const shown = sign * (percent ? parsed / 100 : parsed);
  return Math.abs(shown - value) <= 1e-9 * Math.max(1, Math.abs(value));
}

/**
 * Render a cell the way the user sees it, keeping the underlying value when the
 * display hides it: `01/01/2024 [45292]`, `£98,765.43 [98765.4321]`. Displays
 * that denote the value exactly (`12.5%`, `£1,234.50`) are shown alone. Cells
 * in General format, and "####" overflow displays, show the raw value.
 *
 * Returns `usedRaw` when a bracketed raw value was appended, so callers can
 * add a legend once.
 */
export function formatDisplayedCell(
  value: CellValue,
  text: string | undefined,
  numberFormat: string | undefined,
): { text: string; usedRaw: boolean } {
  const raw = formatCellValue(value);
  if (typeof value !== "number" || text === undefined || numberFormat === undefined || numberFormat === "" || numberFormat === "General") {
    return { text: raw, usedRaw: false };
  }
  const shown = text.trim();
  if (shown === "" || shown === raw || /^#+$/u.test(shown)) return { text: raw, usedRaw: false };
  if (displayDenotesValue(shown, value)) return { text: shown, usedRaw: false };
  return { text: `${shown} [${raw}]`, usedRaw: true };
}

/**
 * Format a 2D array as a markdown table.
 * Uses the first row as header. Cell text is escaped so pipes and line breaks
 * stay inside their cell.
 */
export function formatAsMarkdownTable(values: unknown[][]): string {
  if (!values || values.length === 0) return "(empty)";

  const rows = values.map((row) => row.map((v) => {
    const text = v === null || v === undefined ? ""
      : typeof v === "string" || typeof v === "number" || typeof v === "boolean" ? formatCellValue(v)
        : JSON.stringify(v);
    return escapeMarkdownCell(text);
  }));
  const numCols = Math.max(...rows.map((r) => r.length));

  // Pad rows to uniform width
  const padded = rows.map((r) => {
    while (r.length < numCols) r.push("");
    return r;
  });

  const header = padded[0];
  if (header === undefined) return "(empty)";

  const lines: string[] = [];
  lines.push(`| ${header.join(" | ")} |`);
  lines.push(`| ${header.map(() => "---").join(" | ")} |`);
  for (let i = 1; i < padded.length; i++) {
    const row = padded[i];
    if (row === undefined) continue;
    lines.push(`| ${row.join(" | ")} |`);
  }
  return lines.join("\n");
}

/**
 * Extract formula cells from a 2D formulas array.
 * Returns a list of "CellAddr: =FORMULA" strings.
 *
 * @param formulas - 2D array from Range.formulas
 * @param startAddress - top-left cell of the range (e.g. "A1")
 */
export function extractFormulas(formulas: unknown[][], startAddress: string): string[] {
  const result: string[] = [];
  const start = parseCell(startAddress);

  for (let r = 0; r < formulas.length; r++) {
    const row = formulas[r];
    if (row === undefined) continue;
    for (let c = 0; c < row.length; c++) {
      const f = row[c];
      if (typeof f === "string" && f.startsWith("=")) {
        const addr = `${colToLetter(start.col + c)}${start.row + r}`;
        result.push(`${addr}: ${f}`);
      }
    }
  }
  return result;
}

/**
 * Find error values in a 2D values array.
 * Returns cell addresses and error types.
 */
export function findErrors(
  values: unknown[][],
  startAddress: string,
): { address: string; error: string; formula?: string }[] {
  const errors: { address: string; error: string; formula?: string }[] = [];
  const start = parseCell(startAddress);

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    if (row === undefined) continue;
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (typeof v === "string" && v.startsWith("#")) {
        errors.push({
          address: `${colToLetter(start.col + c)}${start.row + r}`,
          error: v,
        });
      }
    }
  }
  return errors;
}

/**
 * Count non-empty cells in a 2D array.
 */
export function countNonEmpty(values: unknown[][]): number {
  let count = 0;
  for (const row of values) {
    for (const v of row) {
      if (v !== null && v !== undefined && v !== "") count++;
    }
  }
  return count;
}

/** True for Excel error values like #REF!, #VALUE!, #N/A, etc. */
export function isExcelError(value: unknown): value is string {
  return typeof value === "string" && /^#\w+!?$/.test(value);
}

/**
 * Truncate text to approximate token limit.
 * Rough heuristic: 1 token ≈ 4 characters.
 */
export function truncateForTokens(text: string, maxTokens: number): string {
  const maxChars = maxTokens * 4;
  if (text.length <= maxChars) return text;
  return text.substring(0, maxChars) + "\n\n… (truncated)";
}
