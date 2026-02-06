/**
 * Selection context — auto-read around the user's current selection.
 *
 * When the user sends a message, this reads the active cell/range
 * plus surrounding context (±5 rows, full column extent of the data region).
 *
 * Most add-ins only push the selection address — we auto-read the content
 * but doesn't auto-read the content. Our agent can answer "what's wrong
 * with this formula?" without needing a tool call first.
 */

import { excelRun, qualifiedAddress, colToLetter } from "../excel/helpers.js";
import { formatAsMarkdownTable, extractFormulas, findErrors } from "../utils/format.js";

/** How many rows of context to read above and below the selection */
const CONTEXT_ROWS = 5;
/** Maximum columns to include in context */
const MAX_CONTEXT_COLS = 20;

export interface SelectionContext {
  /** Fully qualified address of the selection */
  address: string;
  /** The formatted context text for injection */
  text: string;
}

/**
 * Read the user's current selection and surrounding context.
 * Returns null if Office.js is not available.
 */
export async function readSelectionContext(): Promise<SelectionContext | null> {
  try {
    return await excelRun(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load("address,values,formulas,rowIndex,columnIndex,rowCount,columnCount,worksheet/name");
      await context.sync();

      const sheetName = sel.worksheet.name;
      const selAddress = qualifiedAddress(sheetName, sel.address);

      // Determine the context window
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const used = sheet.getUsedRangeOrNullObject();
      used.load("rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();

      if (used.isNullObject) {
        return {
          address: selAddress,
          text: `**Selection:** ${selAddress} (sheet is empty)`,
        };
      }

      // Compute context window: ±CONTEXT_ROWS around selection, bounded by used range
      const usedStartRow = used.rowIndex;
      const usedEndRow = used.rowIndex + used.rowCount - 1;
      const usedStartCol = used.columnIndex;
      const usedEndCol = Math.min(
        used.columnIndex + used.columnCount - 1,
        used.columnIndex + MAX_CONTEXT_COLS - 1,
      );

      const contextStartRow = Math.max(usedStartRow, sel.rowIndex - CONTEXT_ROWS);
      const contextEndRow = Math.min(usedEndRow, sel.rowIndex + sel.rowCount - 1 + CONTEXT_ROWS);

      // Build context range address
      const startAddr = `${colToLetter(usedStartCol)}${contextStartRow + 1}`;
      const endAddr = `${colToLetter(usedEndCol)}${contextEndRow + 1}`;
      const contextRangeAddr = `${startAddr}:${endAddr}`;

      const contextRange = sheet.getRange(contextRangeAddr);
      contextRange.load("values,formulas,address");
      await context.sync();

      // Build the context text
      const lines: string[] = [];
      lines.push(`**Selection:** ${selAddress}`);
      lines.push(`**Context:** ${qualifiedAddress(sheetName, contextRange.address)}`);
      lines.push("");
      lines.push(formatAsMarkdownTable(contextRange.values));

      // Selection's own formulas
      const selLocalAddress = sel.address.includes("!") ? sel.address.split("!")[1] : sel.address;
      const selFormulas = extractFormulas(sel.formulas, selLocalAddress.split(":")[0]);
      if (selFormulas.length > 0) {
        lines.push("");
        lines.push(`**Selected formulas:** ${selFormulas.join(", ")}`);
      }

      // Errors in context
      const errors = findErrors(contextRange.values, startAddr);
      if (errors.length > 0) {
        lines.push("");
        lines.push(`⚠️ **Errors nearby:** ${errors.map((e) => `${e.address}=${e.error}`).join(", ")}`);
      }

      return { address: selAddress, text: lines.join("\n") };
    });
  } catch {
    return null;
  }
}
