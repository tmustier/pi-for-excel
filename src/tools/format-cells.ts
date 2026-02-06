/**
 * format_cells — Apply formatting to a range (separate from write_cells).
 *
 * Handles: font (bold, italic, color, size), fill color, number format,
 * borders, alignment, column width.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, getRange, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const DEFAULT_FONT_NAME = "Arial";
const DEFAULT_FONT_SIZE = 10;
// Excel columnWidth in Office.js uses points. Approx conversion for Arial 10:
// 1 character width ≈ 7.2 points (based on Excel UI measurement).
const POINTS_PER_CHAR_ARIAL_10 = 7.2;

const schema = Type.Object({
  range: Type.String({
    description: 'Range to format, e.g. "A1:D1", "Sheet2!B3:B20". Supports comma/semicolon-separated ranges on the same sheet (e.g. "A1:B2, D1:D2").',
  }),
  bold: Type.Optional(Type.Boolean({ description: "Set bold." })),
  italic: Type.Optional(Type.Boolean({ description: "Set italic." })),
  underline: Type.Optional(Type.Boolean({ description: "Set underline." })),
  font_color: Type.Optional(
    Type.String({ description: 'Font color as hex, e.g. "#0000FF" for blue.' }),
  ),
  font_size: Type.Optional(Type.Number({ description: "Font size in points." })),
  font_name: Type.Optional(Type.String({ description: 'Font name, e.g. "Arial", "Calibri".' })),
  fill_color: Type.Optional(
    Type.String({ description: 'Background fill color as hex, e.g. "#FFFF00" for yellow.' }),
  ),
  number_format: Type.Optional(
    Type.String({
      description:
        'Excel number format string, e.g. "#,##0.00", "0%", "$#,##0", "yyyy-mm-dd".',
    }),
  ),
  horizontal_alignment: Type.Optional(
    Type.String({
      description: '"Left", "Center", "Right", or "General".',
    }),
  ),
  vertical_alignment: Type.Optional(
    Type.String({
      description: '"Top", "Center", "Bottom".',
    }),
  ),
  wrap_text: Type.Optional(Type.Boolean({ description: "Enable text wrapping." })),
  column_width: Type.Optional(Type.Number({ description: "Set column width in Excel character-width units (assumes Arial 10). Converted to points internally." })),
  row_height: Type.Optional(Type.Number({ description: "Set row height in points." })),
  auto_fit: Type.Optional(
    Type.Boolean({ description: "Auto-fit column widths to content. Default: false." }),
  ),
  borders: Type.Optional(
    Type.Union(
      [
        Type.Literal("thin"),
        Type.Literal("medium"),
        Type.Literal("thick"),
        Type.Literal("none"),
      ],
      {
        description:
          'Border weight: "thin", "medium", "thick", or "none" to remove borders. Applied to all edges.',
      },
    ),
  ),
  merge: Type.Optional(
    Type.Boolean({ description: "Merge the range into a single cell." }),
  ),
});

type Params = Static<typeof schema>;

type HorizontalAlignment = "Left" | "Center" | "Right" | "General";
type VerticalAlignment = "Top" | "Center" | "Bottom";

function isHorizontalAlignment(value: string): value is HorizontalAlignment {
  return value === "Left" || value === "Center" || value === "Right" || value === "General";
}

function isVerticalAlignment(value: string): value is VerticalAlignment {
  return value === "Top" || value === "Center" || value === "Bottom";
}

export function createFormatCellsTool(): AgentTool<typeof schema> {
  return {
    name: "format_cells",
    label: "Format Cells",
    description:
      "Apply formatting to a range of cells (supports comma-separated ranges on one sheet). " +
      "Set font properties (bold, italic, color, size), fill color, number format, alignment, borders, " +
      "column width (Excel character units), and more. Does NOT modify cell values — use write_cells for that.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await excelRun(async (context) => {
          const resolved = resolveFormatTarget(context, params.range);
          resolved.sheet.load("name");
          resolved.target.load("address");

          const requestedColumnWidth = params.column_width;

          const needsAreas = resolved.isMultiRange && (params.number_format || params.merge !== undefined);
          if (!resolved.isMultiRange) {
            resolved.target.load("rowCount,columnCount");
          } else if (needsAreas) {
            resolved.target.areas.load("items/rowCount,items/columnCount");
          }

          await context.sync();

          const sheet = resolved.sheet;
          const target = resolved.target;
          const isMultiRange = resolved.isMultiRange;

          const applied: string[] = [];
          const warnings: string[] = [];
          const formatTarget = target.format;
          let columnWidthFormat: Excel.RangeFormat | null = null;

          // Font properties
          if (params.bold !== undefined) {
            formatTarget.font.bold = params.bold;
            applied.push(params.bold ? "bold" : "not bold");
          }
          if (params.italic !== undefined) {
            formatTarget.font.italic = params.italic;
            applied.push(params.italic ? "italic" : "not italic");
          }
          if (params.underline !== undefined) {
            formatTarget.font.underline = params.underline ? "Single" : "None";
            applied.push(params.underline ? "underline" : "no underline");
          }
          if (params.font_color) {
            formatTarget.font.color = params.font_color;
            applied.push(`font color ${params.font_color}`);
          }
          if (params.font_size) {
            formatTarget.font.size = params.font_size;
            applied.push(`${params.font_size}pt`);
          }
          if (params.font_name) {
            formatTarget.font.name = params.font_name;
            applied.push(`font ${params.font_name}`);
          }

          // Fill
          if (params.fill_color) {
            formatTarget.fill.color = params.fill_color;
            applied.push(`fill ${params.fill_color}`);
          }

          // Number format
          if (params.number_format) {
            const numberFormat = params.number_format;
            if (!resolved.isMultiRange) {
              const range = resolved.target;
              const formatMatrix = Array.from({ length: range.rowCount }, () =>
                Array.from({ length: range.columnCount }, () => numberFormat),
              );
              range.numberFormat = formatMatrix;
            } else {
              const areas = resolved.target;
              for (const area of areas.areas.items) {
                const formatMatrix = Array.from({ length: area.rowCount }, () =>
                  Array.from({ length: area.columnCount }, () => numberFormat),
                );
                area.numberFormat = formatMatrix;
              }
            }
            applied.push(`format "${numberFormat}"`);
          }

          // Alignment
          if (params.horizontal_alignment) {
            if (!isHorizontalAlignment(params.horizontal_alignment)) {
              throw new Error(
                `Invalid horizontal_alignment "${params.horizontal_alignment}". Use Left, Center, Right, or General.`,
              );
            }
            formatTarget.horizontalAlignment = params.horizontal_alignment;
            applied.push(`align ${params.horizontal_alignment.toLowerCase()}`);
          }
          if (params.vertical_alignment) {
            if (!isVerticalAlignment(params.vertical_alignment)) {
              throw new Error(
                `Invalid vertical_alignment "${params.vertical_alignment}". Use Top, Center, or Bottom.`,
              );
            }
            formatTarget.verticalAlignment = params.vertical_alignment;
            applied.push(`v-align ${params.vertical_alignment.toLowerCase()}`);
          }
          if (params.wrap_text !== undefined) {
            formatTarget.wrapText = params.wrap_text;
            applied.push(params.wrap_text ? "wrap" : "no wrap");
          }

          // Dimensions
          if (params.column_width !== undefined) {
            const columnTarget = target.getEntireColumn();

            if (params.font_name && params.font_name !== DEFAULT_FONT_NAME) {
              warnings.push(
                `Column width assumes ${DEFAULT_FONT_NAME} ${DEFAULT_FONT_SIZE}; using ${params.font_name} may differ.`
              );
            }
            if (params.font_size && params.font_size !== DEFAULT_FONT_SIZE) {
              warnings.push(
                `Column width assumes ${DEFAULT_FONT_NAME} ${DEFAULT_FONT_SIZE}; using ${params.font_size}pt may differ.`
              );
            }

            columnTarget.format.columnWidth = params.column_width * POINTS_PER_CHAR_ARIAL_10;
            columnTarget.format.load("columnWidth");
            columnWidthFormat = columnTarget.format;
            applied.push(`col width ${params.column_width}`);
          }
          if (params.row_height !== undefined) {
            const rowTarget = target.getEntireRow();
            rowTarget.format.rowHeight = params.row_height;
            applied.push(`row height ${params.row_height}`);
          }
          if (params.auto_fit) {
            formatTarget.autofitColumns();
            formatTarget.autofitRows();
            applied.push("auto-fit");
          }

          // Borders
          if (params.borders) {
            const borderValue = params.borders;

            const borderIndexes = [
              "EdgeTop",
              "EdgeBottom",
              "EdgeLeft",
              "EdgeRight",
              "InsideHorizontal",
              "InsideVertical",
            ] as const;

            for (const border of borderIndexes) {
              const borderItem = formatTarget.borders.getItem(border);
              if (borderValue === "none") {
                borderItem.style = "None";
              } else {
                borderItem.style = "Continuous";
                const borderWeight: "Thin" | "Medium" | "Thick" =
                  borderValue === "thin"
                    ? "Thin"
                    : borderValue === "medium"
                      ? "Medium"
                      : "Thick";
                borderItem.weight = borderWeight;
              }
            }
            applied.push(`${params.borders} borders`);
          }

          // Merge
          if (params.merge !== undefined) {
            if (resolved.isMultiRange) {
              const areas = resolved.target;
              for (const area of areas.areas.items) {
                if (params.merge) {
                  area.merge();
                } else {
                  area.unmerge();
                }
              }
              applied.push(params.merge ? "merged" : "unmerged");
            } else if (params.merge) {
              const range = resolved.target;
              range.merge();
              applied.push("merged");
            } else {
              const range = resolved.target;
              range.unmerge();
              applied.push("unmerged");
            }
          }

          await context.sync();

          if (columnWidthFormat && typeof requestedColumnWidth === "number") {
            const actualPoints = columnWidthFormat.columnWidth;
            if (typeof actualPoints === "number") {
              const actualChars = actualPoints / POINTS_PER_CHAR_ARIAL_10;
              const delta = Math.abs(actualChars - requestedColumnWidth);
              if (delta > 0.1) {
                warnings.push(
                  `Requested column width ${requestedColumnWidth}, Excel applied ${actualChars.toFixed(2)}.`
                );
              }
            } else {
              warnings.push("Column widths are not uniform; Excel returned no single width value.");
            }
          }

          return { sheetName: sheet.name, address: target.address, applied, warnings, isMultiRange };
        });

        const fullAddr = result.isMultiRange
          ? result.address
          : qualifiedAddress(result.sheetName, result.address);
        const warningText = result.warnings.length
          ? `\n\n⚠️ ${result.warnings.join("\n")}`
          : "";
        return {
          content: [
            {
              type: "text",
              text: `✅ Formatted **${fullAddr}**: ${result.applied.join(", ")}.${warningText}`,
            },
          ],
          details: undefined,
        };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error formatting: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

function splitRangeList(range: string): string[] {
  return range
    .split(/[;,]/)
    .map((part) => part.trim())
    .filter(Boolean);
}

type FormatResolution =
  | { sheet: Excel.Worksheet; target: Excel.Range; isMultiRange: false }
  | { sheet: Excel.Worksheet; target: Excel.RangeAreas; isMultiRange: true };

function resolveFormatTarget(context: Excel.RequestContext, ref: string): FormatResolution {
  const parts = splitRangeList(ref);
  if (parts.length <= 1) {
    const { sheet, range } = getRange(context, ref);
    return { sheet, target: range, isMultiRange: false };
  }

  let sheetName: string | undefined;
  const addresses: string[] = [];

  for (const part of parts) {
    const parsed = parseRangeRef(part);
    if (parsed.sheet) {
      if (sheetName && sheetName !== parsed.sheet) {
        throw new Error("Multi-range formatting must target a single sheet.");
      }
      sheetName = parsed.sheet;
    }
    addresses.push(parsed.address);
  }

  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();
  const target = sheet.getRanges(addresses.join(","));
  return { sheet, target, isMultiRange: true };
}
