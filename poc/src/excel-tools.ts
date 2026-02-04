/**
 * Excel Agent Tools — read_range & write_cells
 *
 * These are AgentTools that let the LLM read and write Excel data
 * via Office.js. They're registered with the ChatPanel's agent
 * through the toolsFactory callback.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";

// ============================================================================
// SCHEMAS
// ============================================================================

const readRangeSchema = Type.Object({
  range: Type.String({
    description:
      'Cell range in A1 notation. Can include sheet name, e.g. "A1:D10", "Sheet2!A1:B5". ' +
      'If no sheet is specified, uses the active sheet. ' +
      'Use named ranges like "SalesData" if they exist.',
  }),
});

const writeCellsSchema = Type.Object({
  start_cell: Type.String({
    description:
      'Top-left cell to start writing from, in A1 notation. Can include sheet name, e.g. "A1", "Sheet2!B3".',
  }),
  values: Type.Array(Type.Array(Type.Any(), { description: "Row of cell values" }), {
    description:
      "2D array of values to write. Each inner array is a row. " +
      'Use strings for text, numbers for numeric values, and strings starting with "=" for formulas. ' +
      "Example: [[\"Name\", \"Age\"], [\"Alice\", 30], [\"Bob\", 25]]",
  }),
});

type ReadRangeParams = Static<typeof readRangeSchema>;
type WriteCellsParams = Static<typeof writeCellsSchema>;

// ============================================================================
// HELPERS
// ============================================================================

/** Convert column number (0-indexed) to letter (0=A, 25=Z, 26=AA) */
function colToLetter(col: number): string {
  let letter = "";
  while (col >= 0) {
    letter = String.fromCharCode((col % 26) + 65) + letter;
    col = Math.floor(col / 26) - 1;
  }
  return letter;
}

/** Compute the range address from a start cell and a 2D values array */
function computeEndRange(startCell: string, values: any[][]): string {
  // Parse "Sheet1!A1" or "A1"
  let sheet = "";
  let cell = startCell;
  if (cell.includes("!")) {
    const parts = cell.split("!");
    sheet = parts[0] + "!";
    cell = parts[1];
  }

  // Parse column letters and row number from cell like "B3"
  const match = cell.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return startCell;

  const colStr = match[1].toUpperCase();
  const startRow = parseInt(match[2], 10);

  // Convert column letters to 0-indexed number
  let startCol = 0;
  for (let i = 0; i < colStr.length; i++) {
    startCol = startCol * 26 + (colStr.charCodeAt(i) - 64);
  }
  startCol--; // 0-indexed

  const numRows = values.length;
  const numCols = Math.max(...values.map((r) => r.length));

  const endCol = startCol + numCols - 1;
  const endRow = startRow + numRows - 1;

  return `${sheet}${colToLetter(startCol)}${startRow}:${colToLetter(endCol)}${endRow}`;
}

/** Format a 2D values array as a compact markdown table */
function formatAsTable(values: any[][], includeRowNumbers = true): string {
  if (!values || values.length === 0) return "(empty)";

  const rows = values.map((row, i) => {
    const cells = row.map((v: any) => {
      if (v === null || v === undefined || v === "") return "";
      if (typeof v === "number") return String(v);
      return String(v);
    });
    if (includeRowNumbers) {
      return `| ${i + 1} | ${cells.join(" | ")} |`;
    }
    return `| ${cells.join(" | ")} |`;
  });

  // Header separator
  const numCols = Math.max(...values.map((r) => r.length));
  const sep = includeRowNumbers
    ? `| --- | ${Array(numCols).fill("---").join(" | ")} |`
    : `| ${Array(numCols).fill("---").join(" | ")} |`;

  // Use first row as header
  return [rows[0], sep, ...rows.slice(1)].join("\n");
}

// ============================================================================
// TOOLS
// ============================================================================

export function createReadRangeTool(): AgentTool<typeof readRangeSchema> {
  return {
    name: "read_range",
    label: "Read Range",
    description:
      "Read cell values, formulas, and number formats from an Excel range. " +
      "Returns the data as a structured object with values (computed results), " +
      "formulas (raw formulas where present), and number formats. " +
      "Use this to understand the current state of the spreadsheet before making changes.",
    parameters: readRangeSchema,
    execute: async (
      _toolCallId: string,
      params: ReadRangeParams,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        // @ts-ignore — Excel is loaded via CDN
        const result = await Excel.run(async (context: any) => {
          // Parse sheet name if present
          let rangeRef = params.range;
          let sheet: any;

          if (rangeRef.includes("!")) {
            const [sheetName, addr] = rangeRef.split("!");
            // Strip quotes from sheet name (e.g., "'My Sheet'!A1" → "My Sheet")
            const cleanName = sheetName.replace(/^'|'$/g, "");
            sheet = context.workbook.worksheets.getItem(cleanName);
            rangeRef = addr;
          } else {
            sheet = context.workbook.worksheets.getActiveWorksheet();
          }

          const range = sheet.getRange(rangeRef);
          range.load("values,formulas,numberFormat,address,rowCount,columnCount");
          sheet.load("name");
          await context.sync();

          return {
            address: `${sheet.name}!${range.address}`,
            dimensions: `${range.rowCount}×${range.columnCount}`,
            values: range.values,
            formulas: range.formulas,
            numberFormats: range.numberFormat,
          };
        });

        // Build compact text response for the LLM
        const lines: string[] = [];
        lines.push(`**${result.address}** (${result.dimensions})`);
        lines.push("");

        // Values as markdown table
        lines.push("**Values:**");
        lines.push(formatAsTable(result.values));

        // Only include formulas if any cell has a formula (starts with =)
        const hasFormulas = result.formulas.some((row: any[]) =>
          row.some((f: any) => typeof f === "string" && f.startsWith("=")),
        );
        if (hasFormulas) {
          lines.push("");
          lines.push("**Formulas** (cells with formulas only):");
          const formulaCells: string[] = [];
          for (let r = 0; r < result.formulas.length; r++) {
            for (let c = 0; c < result.formulas[r].length; c++) {
              const f = result.formulas[r][c];
              if (typeof f === "string" && f.startsWith("=")) {
                // Reconstruct cell address
                const match = result.address.match(/!?([A-Z]+)(\d+)/i);
                if (match) {
                  let startCol = 0;
                  for (let i = 0; i < match[1].length; i++) {
                    startCol = startCol * 26 + (match[1].charCodeAt(i) - 64);
                  }
                  startCol--; // 0-indexed
                  const startRow = parseInt(match[2], 10);
                  const cellAddr = `${colToLetter(startCol + c)}${startRow + r}`;
                  formulaCells.push(`${cellAddr}: ${f}`);
                }
              }
            }
          }
          lines.push(formulaCells.join("\n"));
        }

        // Check for errors in values
        const errors = result.values
          .flat()
          .filter((v: any) => typeof v === "string" && v.startsWith("#"));
        if (errors.length > 0) {
          lines.push("");
          lines.push(`⚠️ **Errors detected:** ${[...new Set(errors)].join(", ")}`);
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error reading range "${params.range}": ${e.message}` }],
          details: undefined,
        };
      }
    },
  };
}

export function createWriteCellsTool(): AgentTool<typeof writeCellsSchema> {
  return {
    name: "write_cells",
    label: "Write Cells",
    description:
      "Write values and formulas to Excel cells. Provide a start cell and a 2D array of values. " +
      "The array will be written starting from the given cell, expanding right and down. " +
      'Use strings starting with "=" for formulas (e.g., "=SUM(A1:A10)"). ' +
      "After writing, the tool automatically reads back the results to verify correctness " +
      "and reports any formula errors (#DIV/0!, #VALUE!, #REF!, etc.).",
    parameters: writeCellsSchema,
    execute: async (
      _toolCallId: string,
      params: WriteCellsParams,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        // Validate values array
        if (!params.values || params.values.length === 0) {
          return {
            content: [{ type: "text", text: "Error: values array is empty" }],
            details: undefined,
          };
        }

        const rangeAddr = computeEndRange(params.start_cell, params.values);

        // @ts-ignore — Excel is loaded via CDN
        const result = await Excel.run(async (context: any) => {
          // Parse sheet name if present
          let cellRef = params.start_cell;
          let sheet: any;

          if (cellRef.includes("!")) {
            const [sheetName] = cellRef.split("!");
            const cleanName = sheetName.replace(/^'|'$/g, "");
            sheet = context.workbook.worksheets.getItem(cleanName);
          } else {
            sheet = context.workbook.worksheets.getActiveWorksheet();
          }

          // Compute the target range from start_cell + values dimensions
          const numRows = params.values.length;
          const numCols = Math.max(...params.values.map((r) => r.length));

          // Pad rows to uniform length
          const paddedValues = params.values.map((row) => {
            const padded = [...row];
            while (padded.length < numCols) padded.push("");
            return padded;
          });

          // Parse start cell to compute range
          let addr = cellRef;
          if (addr.includes("!")) addr = addr.split("!")[1];
          const match = addr.match(/^([A-Z]+)(\d+)$/i);
          if (!match) throw new Error(`Invalid cell address: ${cellRef}`);

          const colStr = match[1].toUpperCase();
          const startRow = parseInt(match[2], 10);
          let startCol = 0;
          for (let i = 0; i < colStr.length; i++) {
            startCol = startCol * 26 + (colStr.charCodeAt(i) - 64);
          }
          startCol--;

          const endCol = colToLetter(startCol + numCols - 1);
          const endRow = startRow + numRows - 1;
          const rangeAddress = `${colStr}${startRow}:${endCol}${endRow}`;

          const range = sheet.getRange(rangeAddress);
          range.values = paddedValues;
          range.format.autofitColumns();

          sheet.load("name");
          await context.sync();

          // Read back to verify
          const verify = sheet.getRange(rangeAddress);
          verify.load("values,formulas,address");
          await context.sync();

          return {
            address: `${sheet.name}!${verify.address}`,
            writtenValues: paddedValues,
            readBackValues: verify.values,
            readBackFormulas: verify.formulas,
          };
        });

        // Build response
        const lines: string[] = [];
        lines.push(`✅ Written to **${result.address}** (${params.values.length} rows × ${Math.max(...params.values.map((r: any[]) => r.length))} cols)`);

        // Check for formula errors in read-back values
        const errors: string[] = [];
        const errorCells: string[] = [];
        for (let r = 0; r < result.readBackValues.length; r++) {
          for (let c = 0; c < result.readBackValues[r].length; c++) {
            const v = result.readBackValues[r][c];
            if (typeof v === "string" && v.startsWith("#")) {
              errors.push(v);
              // Reconstruct cell address
              const addrMatch = result.address.match(/!?([A-Z]+)(\d+)/i);
              if (addrMatch) {
                let sc = 0;
                for (let i = 0; i < addrMatch[1].length; i++) {
                  sc = sc * 26 + (addrMatch[1].charCodeAt(i) - 64);
                }
                sc--;
                const cellAddr = `${colToLetter(sc + c)}${parseInt(addrMatch[2], 10) + r}`;
                const formula = result.readBackFormulas[r][c];
                errorCells.push(`${cellAddr}: ${v} (formula: ${formula})`);
              }
            }
          }
        }

        if (errors.length > 0) {
          lines.push("");
          lines.push(`⚠️ **${errors.length} formula error(s) detected:**`);
          for (const ec of errorCells) {
            lines.push(`- ${ec}`);
          }
          lines.push("");
          lines.push("Review the formulas and use write_cells again to fix them.");
        } else {
          // Show read-back values as confirmation
          lines.push("");
          lines.push("**Verified values:**");
          lines.push(formatAsTable(result.readBackValues));
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error writing cells: ${e.message}` }],
          details: undefined,
        };
      }
    },
  };
}
