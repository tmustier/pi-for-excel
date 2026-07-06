/**
 * Formatting utilities for LLM-facing output.
 * Converts Excel data into compact, token-efficient text.
 */

import { colToLetter, parseCell } from "../excel/helpers.js";

/**
 * Format a 2D array as a markdown table.
 * Uses the first row as header.
 */
export function formatAsMarkdownTable(values: DynamicValue[][]): string {
  if (!values || values.length === 0) return "(empty)";

  const stringify = (v: DynamicValue): string => {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") {
      // Avoid scientific notation for readability
      if (Math.abs(v) < 1e15 && Math.abs(v) > 1e-6) return String(v);
      return v.toExponential(4);
    }
    if (typeof v === "string") return v;
    if (typeof v === "boolean") return String(v);
    return JSON.stringify(v);
  };

  const rows = values.map((row) => row.map(stringify));
  const numCols = Math.max(...rows.map((r) => r.length));

  // Pad rows to uniform width
  const padded = rows.map((r) => {
    while (r.length < numCols) r.push("");
    return r;
  });

  const lines: string[] = [];
  lines.push(`| ${padded[0].join(" | ")} |`);
  lines.push(`| ${padded[0].map(() => "---").join(" | ")} |`);
  for (let i = 1; i < padded.length; i++) {
    lines.push(`| ${padded[i].join(" | ")} |`);
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
export function extractFormulas(formulas: DynamicValue[][], startAddress: string): string[] {
  const result: string[] = [];
  const start = parseCell(startAddress);

  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      const f = formulas[r][c];
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
  values: DynamicValue[][],
  startAddress: string,
): { address: string; error: string; formula?: string }[] {
  const errors: { address: string; error: string; formula?: string }[] = [];
  const start = parseCell(startAddress);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = values[r][c];
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
export function countNonEmpty(values: DynamicValue[][]): number {
  let count = 0;
  for (const row of values) {
    for (const v of row) {
      if (v !== null && v !== undefined && v !== "") count++;
    }
  }
  return count;
}

/** True for Excel error values like #REF!, #VALUE!, #N/A, etc. */
export function isExcelError(value: DynamicValue): value is string {
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
