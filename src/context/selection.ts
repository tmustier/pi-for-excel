/**
 * Selection context — what the user has selected, sent with every request.
 *
 * The address and size come first and are always exact. The cell contents
 * are a bounded sample: the selection plus a few rows around it, clipped to
 * the sheet's used range, at most 20 columns, and for tall selections only
 * the first and last rows. A whole-column selection therefore costs the same
 * as a small one, and nothing the user selects can flood the model context.
 *
 * Failures are reported in the text (the agent learns that the selection
 * could not be read and why) and logged, never swallowed.
 */

import { excelRun, qualifiedAddress, colToLetter, ExcelCellEditModeError } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import { Value } from "typebox/value";

import {
  cellValueGridSchema,
  extractFormulas,
  findErrors,
  formatDisplayedCell,
  textGridSchema,
  type CellValue,
} from "../utils/format.js";

/** Rows of context above and below the selection. */
const CONTEXT_ROWS = 5;
const MAX_COLS = 20;
/** Cells in the sample; tall samples show fewer columns. */
const MAX_SAMPLE_CELLS = 240;
/** Tall windows show this many rows from the top and the bottom. */
const HEAD_ROWS = 15;
const TAIL_ROWS = 5;
const MAX_CELL_CHARS = 60;
const MIN_CELL_CHARS = 8;
/** Upper bound for the sample table, whatever the cell contents. */
const MAX_TABLE_CHARS = 4000;
const MAX_FORMULAS = 15;
const MAX_ERRORS = 10;

export interface SelectionContext {
  /** Qualified selection address, or null when it could not be read. */
  address: string | null;
  /** Agent-facing text. */
  text: string;
}

const HEADER = "Current selection (re-read for every request, not stored in the conversation):";

interface RowBlock {
  startRow: number;
  values: CellValue[][];
  formulas: CellValue[][];
  text: string[][];
  numberFormats: string[][];
}

function plural(count: number, noun: string): string {
  return `${count.toLocaleString("en-US")} ${noun}${count === 1 ? "" : "s"}`;
}

function formatGrid(blocks: readonly RowBlock[], startCol: number, colCount: number): string {
  const letters = Array.from({ length: colCount }, (_, i) => colToLetter(startCol + i));
  const lines = [`| | ${letters.join(" | ")} |`, `|---|${letters.map(() => "---").join("|")}|`];
  let chars = lines.join("\n").length;
  let previousEnd: number | null = null;
  // Share the budget across every sampled cell (leaving room for row numbers
  // and separators) so the last rows fit too.
  const rowCount = blocks.reduce((sum, block) => sum + block.values.length, 0);
  const perCell = Math.floor((MAX_TABLE_CHARS * 0.8) / Math.max(1, rowCount * colCount)) - 3;
  const maxCellChars = Math.max(MIN_CELL_CHARS, Math.min(MAX_CELL_CHARS, perCell));

  for (const block of blocks) {
    if (previousEnd !== null && block.startRow > previousEnd + 1) {
      lines.push(`| … | (${plural(block.startRow - previousEnd - 1, "row")} not shown) |`);
    }
    for (let r = 0; r < block.values.length; r++) {
      const cells = letters.map((_, c) => {
        const shown = formatDisplayedCell(block.values[r]?.[c], block.text[r]?.[c], block.numberFormats[r]?.[c]).text;
        const text = shown.replace(/\r?\n/g, " ").replace(/\\/g, "\\\\").replace(/\|/g, "\\|");
        return text.length > maxCellChars ? `${text.slice(0, maxCellChars - 1)}…` : text;
      });
      const line = `| ${block.startRow + r + 1} | ${cells.join(" | ")} |`;
      if (chars + line.length > MAX_TABLE_CHARS) {
        lines.push(`| … | (sample cut at ${MAX_TABLE_CHARS.toLocaleString("en-US")} characters) |`);
        return lines.join("\n");
      }
      lines.push(line);
      chars += line.length + 1;
    }
    previousEnd = block.startRow + block.values.length - 1;
  }
  return lines.join("\n");
}

async function readSelection(context: Excel.RequestContext): Promise<SelectionContext> {
  const sel = context.workbook.getSelectedRange();
  sel.load("address,rowIndex,columnIndex,rowCount,columnCount,worksheet/name");
  const used = sel.worksheet.getUsedRangeOrNullObject(true);
  used.load("address,rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();

  const sheetName = sel.worksheet.name;
  const address = qualifiedAddress(sheetName, sel.address);
  const lines = [
    HEADER,
    `Selection: ${address} — ${plural(sel.rowCount, "row")} × ${plural(sel.columnCount, "column")}`,
  ];

  if (used.isNullObject) {
    lines.push(`The sheet ${sheetName} has no values.`);
    return { address, text: lines.join("\n") };
  }

  const usedEndRow = used.rowIndex + used.rowCount - 1;
  const usedEndCol = used.columnIndex + used.columnCount - 1;
  const selEndRow = sel.rowIndex + sel.rowCount - 1;
  const selEndCol = sel.columnIndex + sel.columnCount - 1;
  lines.push(`Used range of ${sheetName}: ${qualifiedAddress(sheetName, used.address)}`);

  const startRow = Math.max(used.rowIndex, sel.rowIndex - CONTEXT_ROWS);
  const endRow = Math.min(usedEndRow, selEndRow + CONTEXT_ROWS);
  const windowRows = endRow - startRow + 1;
  const spans: Array<[number, number]> = windowRows <= HEAD_ROWS + TAIL_ROWS + 2
    ? [[startRow, windowRows]]
    : [[startRow, HEAD_ROWS], [endRow - TAIL_ROWS + 1, TAIL_ROWS]];
  const sampleRows = spans.reduce((sum, [, count]) => sum + count, 0);
  const maxCols = Math.max(1, Math.min(MAX_COLS, Math.floor(MAX_SAMPLE_CELLS / sampleRows)));
  // Keep the left-hand labels when the selection is near them; otherwise start
  // just left of the selection.
  const startCol = sel.columnIndex < used.columnIndex + maxCols
    ? used.columnIndex
    : Math.max(used.columnIndex, sel.columnIndex - 2);
  const endCol = Math.min(usedEndCol, startCol + maxCols - 1);

  const overlapsUsed = sel.rowIndex <= usedEndRow && selEndRow >= used.rowIndex
    && sel.columnIndex <= usedEndCol && selEndCol >= used.columnIndex;
  if (!overlapsUsed) lines.push("The selected cells are empty (outside the used range).");
  if (startRow > endRow || startCol > endCol) return { address, text: lines.join("\n") };

  const colCount = endCol - startCol + 1;

  const sheet = sel.worksheet;
  const ranges = spans.map(([row, count]) => {
    const range = sheet.getRangeByIndexes(row, startCol, count, colCount);
    range.load("values,formulas,text,numberFormat");
    return { row, range };
  });
  await context.sync();
  const blocks: RowBlock[] = ranges.map(({ row, range }) => ({
    startRow: row,
    values: Value.Parse(cellValueGridSchema, range.values),
    formulas: Value.Parse(cellValueGridSchema, range.formulas),
    text: Value.Parse(textGridSchema, range.text),
    numberFormats: Value.Parse(textGridSchema, range.numberFormat),
  }));

  const shownRows = spans.map(([row, count]) => `${row + 1}–${row + count}`).join(" and ");
  lines.push(
    `Sample: rows ${shownRows}, columns ${colToLetter(startCol)}–${colToLetter(endCol)}` +
    ` (the selection with up to ${CONTEXT_ROWS} rows around it, within the used range). Use read_range for anything else.`,
  );
  lines.push("", formatGrid(blocks, startCol, colCount));

  const formulas: string[] = [];
  const errors: string[] = [];
  for (const block of blocks) {
    const start = `${colToLetter(startCol)}${block.startRow + 1}`;
    formulas.push(...extractFormulas(block.formulas, start));
    errors.push(...findErrors(block.values, start).map((e) => `${e.address}=${e.error}`));
  }
  if (formulas.length > 0) {
    const more = formulas.length > MAX_FORMULAS ? `, … (+${formulas.length - MAX_FORMULAS} more)` : "";
    lines.push("", `Formulas in the sample: ${formulas.slice(0, MAX_FORMULAS).join("; ")}${more}`);
  }
  if (errors.length > 0) {
    const more = errors.length > MAX_ERRORS ? `, … (+${errors.length - MAX_ERRORS} more)` : "";
    lines.push("", `Errors in the sample: ${errors.slice(0, MAX_ERRORS).join(", ")}${more}`);
  }

  return { address, text: lines.join("\n") };
}

/**
 * Read the user's current selection. Never waits for the user to leave
 * cell-edit mode, and never throws: a failure becomes the text.
 */
export async function readSelectionContext(): Promise<SelectionContext> {
  try {
    return await excelRun(readSelection, { waitForCellEdit: false });
  } catch (error) {
    if (error instanceof ExcelCellEditModeError) {
      return {
        address: null,
        text: `${HEADER}\nUnavailable: Excel is in cell-edit mode (the user is typing in a cell). Workbook tools wait until the user finishes.`,
      };
    }
    console.warn("[selection-context] Could not read the selection:", getErrorMessage(error));
    return { address: null, text: `${HEADER}\nUnavailable: ${getErrorMessage(error)}` };
  }
}
