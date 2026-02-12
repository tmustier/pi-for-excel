/**
 * Workbook cell change diff helpers.
 */

import { cellAddress, parseCell, qualifiedAddress } from "../excel/helpers.js";

const DEFAULT_SAMPLE_LIMIT = 20;
const DEFAULT_PREVIEW_CHARS = 80;

export interface WorkbookCellChange {
  address: string;
  beforeValue: string;
  afterValue: string;
  beforeFormula?: string;
  afterFormula?: string;
}

export interface WorkbookCellChangeSummary {
  changedCount: number;
  truncated: boolean;
  sample: WorkbookCellChange[];
}

export interface BuildWorkbookCellChangeSummaryArgs {
  sheetName: string;
  startCell: string;
  beforeValues: unknown[][];
  beforeFormulas: unknown[][];
  afterValues: unknown[][];
  afterFormulas: unknown[][];
  sampleLimit?: number;
  previewChars?: number;
}

function clampPositiveInteger(value: number | undefined, fallback: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return fallback;
  const rounded = Math.floor(value);
  return rounded > 0 ? rounded : fallback;
}

function normalizeFormula(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;

  const trimmed = raw.trim();
  if (!trimmed.startsWith("=")) return undefined;
  return trimmed;
}

function serializeComparable(raw: unknown): string {
  if (raw === null || raw === undefined || raw === "") return "";

  if (typeof raw === "string") return raw;
  if (typeof raw === "number") return Number.isNaN(raw) ? "NaN" : String(raw);
  if (typeof raw === "boolean") return raw ? "true" : "false";
  if (typeof raw === "bigint") return String(raw);
  if (typeof raw === "symbol") return raw.description ?? "";
  if (typeof raw === "function") return "[function]";

  try {
    return JSON.stringify(raw);
  } catch {
    return "[unserializable]";
  }
}

function toPreview(raw: unknown, maxChars: number): string {
  const text = serializeComparable(raw);
  if (text.length <= maxChars) return text;
  return `${text.slice(0, maxChars - 1)}…`;
}

function cellAtOffset(startCell: string, rowOffset: number, colOffset: number): string {
  const start = parseCell(startCell);
  return cellAddress(start.col + colOffset, start.row + rowOffset);
}

function valueAt(grid: unknown[][], row: number, col: number): unknown {
  const rowValues = grid[row];
  if (!Array.isArray(rowValues)) return undefined;
  return rowValues[col];
}

function rowWidth(grid: unknown[][], row: number): number {
  const rowValues = grid[row];
  return Array.isArray(rowValues) ? rowValues.length : 0;
}

export function buildWorkbookCellChangeSummary(
  args: BuildWorkbookCellChangeSummaryArgs,
): WorkbookCellChangeSummary {
  const sampleLimit = clampPositiveInteger(args.sampleLimit, DEFAULT_SAMPLE_LIMIT);
  const previewChars = clampPositiveInteger(args.previewChars, DEFAULT_PREVIEW_CHARS);

  const rowCount = Math.max(
    args.beforeValues.length,
    args.afterValues.length,
    args.beforeFormulas.length,
    args.afterFormulas.length,
  );

  const sample: WorkbookCellChange[] = [];
  let changedCount = 0;

  for (let row = 0; row < rowCount; row += 1) {
    const colCount = Math.max(
      rowWidth(args.beforeValues, row),
      rowWidth(args.afterValues, row),
      rowWidth(args.beforeFormulas, row),
      rowWidth(args.afterFormulas, row),
    );

    for (let col = 0; col < colCount; col += 1) {
      const beforeValueRaw = valueAt(args.beforeValues, row, col);
      const afterValueRaw = valueAt(args.afterValues, row, col);

      const beforeFormula = normalizeFormula(valueAt(args.beforeFormulas, row, col));
      const afterFormula = normalizeFormula(valueAt(args.afterFormulas, row, col));

      const valueChanged = serializeComparable(beforeValueRaw) !== serializeComparable(afterValueRaw);
      const formulaChanged = beforeFormula !== afterFormula;

      if (!valueChanged && !formulaChanged) {
        continue;
      }

      changedCount += 1;

      if (sample.length >= sampleLimit) {
        continue;
      }

      const localAddress = cellAtOffset(args.startCell, row, col);
      sample.push({
        address: qualifiedAddress(args.sheetName, localAddress),
        beforeValue: toPreview(beforeValueRaw, previewChars),
        afterValue: toPreview(afterValueRaw, previewChars),
        beforeFormula,
        afterFormula,
      });
    }
  }

  return {
    changedCount,
    truncated: changedCount > sample.length,
    sample,
  };
}
