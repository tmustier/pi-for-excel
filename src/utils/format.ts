import { Type } from "typebox";

import { colToLetter, parseCell } from "../excel/helpers.js";

export type CellValue = string | number | boolean | null | undefined;

export const cellValueGridSchema = Type.Refine(
  Type.Unsafe<CellValue[][]>({}),
  (raw) => Array.isArray(raw) && raw.every((row) => Array.isArray(row) && row.every((cell) =>
    cell === null || typeof cell === "string" || typeof cell === "number" || typeof cell === "boolean")),
  () => "cell values must be a grid of strings, numbers, booleans or nulls",
);

export const textGridSchema = Type.Refine(
  Type.Unsafe<string[][]>({}),
  (raw) => Array.isArray(raw) && raw.every((row) => Array.isArray(row) && row.every((cell) => typeof cell === "string")),
  () => "cell text must be a grid of strings",
);

export function formatCellValue(v: CellValue): string {
  if (v === null || v === undefined) return "";
  return typeof v === "string" ? v : String(v);
}

export function escapeMarkdownCell(text: string): string {
  return text
    .replace(/\\/g, "\\\\")
    .replace(/\|/g, "\\|")
    .replace(/\r\n|\r|\n/g, "\\n");
}

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

export function formatAsMarkdownTable(values: unknown[][]): string {
  if (!values || values.length === 0) return "(empty)";

  const rows = values.map((row) => row.map((v) => {
    const text = v === null || v === undefined ? ""
      : typeof v === "string" || typeof v === "number" || typeof v === "boolean" ? formatCellValue(v)
        : JSON.stringify(v);
    return escapeMarkdownCell(text);
  }));
  const numCols = Math.max(...rows.map((r) => r.length));

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
      if (isExcelError(v)) {
        errors.push({
          address: `${colToLetter(start.col + c)}${start.row + r}`,
          error: v,
        });
      }
    }
  }
  return errors;
}

export function countNonEmpty(values: unknown[][]): number {
  let count = 0;
  for (const row of values) {
    for (const v of row) {
      if (v !== null && v !== undefined && v !== "") count++;
    }
  }
  return count;
}

const EXCEL_ERROR_VALUES = new Set([
  "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
  "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!",
  "#BUSY!", "#UNKNOWN!", "#EXTERNAL!", "#PYTHON!", "#TIMEOUT",
]);

export function isExcelError(value: unknown): value is string {
  return typeof value === "string" && EXCEL_ERROR_VALUES.has(value);
}

export function truncateForTokens(text: string, maxTokens: number): string {
  const maxChars = maxTokens * 4;
  if (text.length <= maxChars) return text;
  return text.substring(0, maxChars) + "\n\n… (truncated)";
}
