/**
 * read_range — Read cell values, formulas, and optionally formatting.
 *
 * Three modes:
 * - "compact" (default): Markdown table of values. Token-efficient.
 * - "csv": Raw CSV values only. Analysis-friendly.
 * - "detailed": Full JSON with formulas, number formats. For debugging.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { excelRun, getRange, qualifiedAddress, parseCell, colToLetter } from "../excel/helpers.js";
import { formatAsMarkdownTable, extractFormulas, findErrors } from "../utils/format.js";
import { getErrorMessage } from "../utils/errors.js";
import { buildResolvedFormatLabels, getResolvedConventions, humanizeFormat } from "../conventions/index.js";
import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";
import type { ReadRangeCsvDetails } from "./tool-details.js";

const schema = Type.Object({
  range: Type.String({
    description:
      'Cell range in A1 notation, e.g. "A1:D10", "Sheet2!A1:B5". ' +
      "If no sheet is specified, uses the active sheet.",
  }),
  mode: Type.Optional(
    Type.Union([Type.Literal("compact"), Type.Literal("csv"), Type.Literal("detailed")], {
      description:
        '"compact" (default): markdown table of values. ' +
        '"csv": values only as CSV. ' +
        '"detailed": includes formulas and number formats.',
    }),
  ),
});

type Params = Static<typeof schema>;

interface CommentSummary {
  cell: string;
  content: string;
  author: string;
  resolved: boolean;
  replyCount: number;
}

interface ReadRangeResult {
  sheetName: string;
  address: string;
  rows: number;
  cols: number;
  values: unknown[][];
  formulas: unknown[][];
  numberFormats: unknown[][];
  comments: CommentSummary[];
}

export function createReadRangeTool(): AgentTool<typeof schema> {
  return {
    name: "read_range",
    label: "Read Range",
    description:
      "Read cell values from an Excel range. Returns a markdown table by default (compact mode). " +
      'Use mode "csv" for raw CSV values (analysis-friendly). ' +
      'Use mode "detailed" to also see formulas and number formats. ' +
      "Always read before modifying — never guess what's in the spreadsheet.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ReadRangeCsvDetails | undefined>> => {
      try {
        const mode = params.mode || "compact";

        const result = await excelRun<ReadRangeResult>(async (context) => {
          const { sheet, range } = getRange(context, params.range);
          range.load("values,formulas,numberFormat,address,rowCount,columnCount");
          sheet.load("name");

          // Pre-load comments collection for detailed mode
          const loadComments = mode === "detailed";
          const commentsCol = loadComments ? sheet.comments : undefined;
          if (commentsCol) {
            commentsCol.load("items");
          }

          await context.sync();

          // Collect comments within the range (detailed mode only)
          const comments: CommentSummary[] = [];
          if (commentsCol && commentsCol.items.length > 0) {
            const entries = commentsCol.items.map((comment) => {
              comment.load("content,authorName,resolved");
              const location = comment.getLocation();
              location.load("address");
              const replyCount = comment.replies.getCount();
              return { comment, location, replyCount };
            });
            await context.sync();

            const rangeAddr = range.address;
            for (const { comment, location, replyCount } of entries) {
              const locCell = location.address.includes("!")
                ? location.address.split("!")[1]
                : location.address;
              if (isCellInRange(locCell, rangeAddr)) {
                comments.push({
                  cell: locCell,
                  content: comment.content,
                  author: comment.authorName,
                  resolved: comment.resolved,
                  replyCount: replyCount.value,
                });
              }
            }
          }

          return {
            sheetName: sheet.name,
            address: range.address,
            rows: range.rowCount,
            cols: range.columnCount,
            values: range.values,
            formulas: range.formulas,
            numberFormats: range.numberFormat,
            comments,
          };
        });

        const fullAddress = qualifiedAddress(result.sheetName, result.address);
        // Extract just the cell part (without sheet!) for offset calculations
        const cellPart = result.address.includes("!") ? result.address.split("!")[1] : result.address;
        const startCell = cellPart.split(":")[0];

        if (mode === "compact") {
          return formatCompact(fullAddress, result, startCell);
        } else if (mode === "csv") {
          return formatCsvOutput(fullAddress, result, startCell);
        } else {
          let resolvedLabels: Map<string, string> | undefined;
          try {
            const conventions = await getResolvedConventions(getAppStorage().settings);
            resolvedLabels = buildResolvedFormatLabels(conventions);
          } catch {
            // Ignore conventions lookup failures in read-only output.
          }

          return formatDetailed(fullAddress, result, startCell, resolvedLabels);
        }
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error reading "${params.range}": ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

/** Check if a cell address falls within a range address (both without sheet prefix). */
function isCellInRange(cellAddr: string, rangeAddr: string): boolean {
  const clean = rangeAddr.includes("!") ? rangeAddr.split("!")[1] : rangeAddr;
  const parts = clean.includes(":") ? clean.split(":") : [clean, clean];
  const start = parseCell(parts[0]);
  const end = parseCell(parts[1]);
  const cell = parseCell(cellAddr);
  return (
    cell.col >= start.col &&
    cell.col <= end.col &&
    cell.row >= start.row &&
    cell.row <= end.row
  );
}

function hasAnyNonEmptyCell(values: unknown[][]): boolean {
  for (const row of values) {
    for (const v of row) {
      if (v !== null && v !== undefined && v !== "") return true;
    }
  }
  return false;
}

function formatAsExcelMarkdownTable(values: unknown[][], startCell: string): string {
  if (!values || values.length === 0) return "(empty)";

  const start = parseCell(startCell);
  const numCols = Math.max(...values.map((r) => r.length));

  const header: unknown[] = [""];
  for (let c = 0; c < numCols; c++) {
    header.push(colToLetter(start.col + c));
  }

  const rows: unknown[][] = [header];

  for (let r = 0; r < values.length; r++) {
    const row: unknown[] = [start.row + r, ...values[r]];
    while (row.length < numCols + 1) row.push("");
    rows.push(row);
  }

  return formatAsMarkdownTable(rows);
}

function formatCompact(
  address: string,
  result: ReadRangeResult,
  startCell: string,
): AgentToolResult<undefined> {
  const lines: string[] = [];

  // Collect these once; we also use them to detect a truly empty range.
  const formulas = extractFormulas(result.formulas, startCell);
  const errors = findErrors(result.values, startCell);
  const hasValues = hasAnyNonEmptyCell(result.values);

  lines.push(`**${address}** (${result.rows}×${result.cols})`);

  if (!hasValues && formulas.length === 0 && errors.length === 0) {
    lines.push("");
    lines.push("_All cells are empty._");
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  lines.push("");
  lines.push(formatAsExcelMarkdownTable(result.values, startCell));

  // Append formulas if any exist
  if (formulas.length > 0) {
    lines.push("");
    lines.push(`**Formulas:** ${formulas.join(", ")}`);
  }

  // Append errors if any
  if (errors.length > 0) {
    lines.push("");
    lines.push(`⚠️ **Errors:** ${errors.map((e) => `${e.address}=${e.error}`).join(", ")}`);
  }

  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}

function formatDetailed(
  address: string,
  result: ReadRangeResult,
  startCell: string,
  resolvedLabels?: Map<string, string>,
): AgentToolResult<undefined> {
  const lines: string[] = [];

  // Collect these once; we also use them to detect a truly empty range.
  const formulas = extractFormulas(result.formulas, startCell);
  const errors = findErrors(result.values, startCell);
  const hasValues = hasAnyNonEmptyCell(result.values);

  lines.push(`**${address}** (${result.rows}×${result.cols})`);

  if (!hasValues && formulas.length === 0 && errors.length === 0) {
    lines.push("");
    lines.push("_All cells are empty._");
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  lines.push("");

  // Values table
  lines.push("### Values");
  lines.push(formatAsExcelMarkdownTable(result.values, startCell));

  // All formulas
  if (formulas.length > 0) {
    lines.push("");
    lines.push("### Formulas");
    for (const f of formulas) {
      lines.push(`- ${f}`);
    }
  }

  // Number formats (deduplicated)
  const formatMap = new Map<string, string[]>();
  const start = parseCell(startCell);
  for (let r = 0; r < result.numberFormats.length; r++) {
    for (let c = 0; c < result.numberFormats[r].length; c++) {
      const fmt = result.numberFormats[r][c];
      if (typeof fmt === "string" && fmt !== "" && fmt !== "General") {
        const addr = `${colToLetter(start.col + c)}${start.row + r}`;
        const existing = formatMap.get(fmt) || [];
        existing.push(addr);
        formatMap.set(fmt, existing);
      }
    }
  }
  if (formatMap.size > 0) {
    lines.push("");
    lines.push("### Number Formats");
    for (const [fmt, cells] of formatMap) {
      const label = humanizeFormat(fmt, resolvedLabels);
      const display = label !== fmt ? `**${label}** (\`${fmt}\`)` : `\`${fmt}\``;
      lines.push(`- ${display} → ${cells.join(", ")}`);
    }
  }

  // Errors
  if (errors.length > 0) {
    lines.push("");
    lines.push("### ⚠️ Errors");
    for (const e of errors) {
      lines.push(`- ${e.address}: ${e.error}`);
    }
  }

  // Comments (only when present)
  if (result.comments.length > 0) {
    lines.push("");
    lines.push("### Comments");
    for (const c of result.comments) {
      const resolved = c.resolved ? " ✓" : "";
      const replies = c.replyCount > 0 ? ` (${c.replyCount} ${c.replyCount === 1 ? "reply" : "replies"})` : "";
      lines.push(`- **${c.cell}**: "${c.content}" — *${c.author}*${resolved}${replies}`);
    }
  }

  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}

/* ── CSV helpers (migrated from get-range-as-csv) ──────────────────── */

function toCsvField(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.includes(",") || value.includes('"') || value.includes("\n") || value.includes("\r") ? `"${value.replace(/"/g, '""')}"` : value;
  const str = typeof value === "number" || typeof value === "boolean" ? String(value) : JSON.stringify(value);
  if (/[",\n\r]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

function valuesToCsv(values: unknown[][]): string {
  if (!values || values.length === 0) return "";
  return values
    .map((row) => row.map((v) => toCsvField(v)).join(","))
    .join("\n");
}

function formatCsvOutput(
  address: string,
  result: ReadRangeResult,
  startCell: string,
): AgentToolResult<ReadRangeCsvDetails | undefined> {
  const lines: string[] = [];
  lines.push(`**${address}** (${result.rows}×${result.cols})`);
  lines.push("");

  const csv = valuesToCsv(result.values);
  if (!csv) {
    lines.push("(empty)");
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  lines.push("```csv");
  lines.push(csv);
  lines.push("```");

  const start = parseCell(startCell);
  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details: {
      kind: "read_range_csv",
      startCol: start.col,
      startRow: start.row,
      values: result.values,
      csv,
    },
  };
}
