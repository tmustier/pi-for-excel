/**
 * format_cells — Apply formatting to a range (separate from write_cells).
 *
 * Handles: font (bold, italic, color, size), fill color, number format,
 * borders, alignment, column width.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, getRange, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";

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
  column_width: Type.Optional(Type.Number({ description: "Set column width in Excel character-width units (same as Excel UI)." })),
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

export function createFormatCellsTool(): AgentTool<typeof schema> {
  return {
    name: "format_cells",
    label: "Format Cells",
    description:
      "Apply formatting to a range of cells (supports comma-separated ranges on one sheet). " +
      "Set font properties (bold, italic, color, size), fill color, number format, alignment, borders, " +
      "column width, and more. Does NOT modify cell values — use write_cells for that.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await excelRun(async (context: any) => {
          const { sheet, target, isMultiRange } = resolveFormatTarget(context, params.range);
          sheet.load("name");
          target.load("address");

          const needsAreas = isMultiRange && (params.number_format || params.merge !== undefined);
          if (!isMultiRange) {
            target.load("rowCount,columnCount");
          } else if (needsAreas) {
            target.areas.load("items/rowCount,items/columnCount");
          }

          await context.sync();

          const applied: string[] = [];
          const warnings: string[] = [];
          const formatTarget = target.format;
          let columnWidthFormat: any | null = null;

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
            if (!isMultiRange) {
              const formatMatrix = Array.from({ length: target.rowCount }, () =>
                Array.from({ length: target.columnCount }, () => params.number_format),
              );
              target.numberFormat = formatMatrix as any;
            } else {
              for (const area of target.areas.items) {
                const formatMatrix = Array.from({ length: area.rowCount }, () =>
                  Array.from({ length: area.columnCount }, () => params.number_format),
                );
                area.numberFormat = formatMatrix as any;
              }
            }
            applied.push(`format "${params.number_format}"`);
          }

          // Alignment
          if (params.horizontal_alignment) {
            formatTarget.horizontalAlignment = params.horizontal_alignment;
            applied.push(`align ${params.horizontal_alignment.toLowerCase()}`);
          }
          if (params.vertical_alignment) {
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
            columnTarget.format.columnWidth = params.column_width;
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
            const borderValue = params.borders.toLowerCase();
            if (!['none', 'thin', 'medium', 'thick'].includes(borderValue)) {
              throw new Error(`Invalid borders value "${params.borders}". Use thin, medium, thick, or none.`);
            }
            const borderWeight =
              borderValue === "thin"
                ? "Thin"
                : borderValue === "medium"
                  ? "Medium"
                  : borderValue === "thick"
                    ? "Thick"
                    : null;

            const borders = [
              "EdgeTop",
              "EdgeBottom",
              "EdgeLeft",
              "EdgeRight",
              "InsideHorizontal",
              "InsideVertical",
            ];
            for (const border of borders) {
              const borderItem = formatTarget.borders.getItem(border);
              if (borderValue === "none") {
                borderItem.style = "None";
              } else {
                borderItem.style = "Continuous";
                borderItem.weight = borderWeight;
              }
            }
            applied.push(`${params.borders} borders`);
          }

          // Merge
          if (params.merge !== undefined) {
            if (isMultiRange) {
              for (const area of target.areas.items) {
                if (params.merge) {
                  area.merge();
                } else {
                  area.unmerge();
                }
              }
              applied.push(params.merge ? "merged" : "unmerged");
            } else if (params.merge) {
              target.merge();
              applied.push("merged");
            } else {
              target.unmerge();
              applied.push("unmerged");
            }
          }

          await context.sync();

          if (columnWidthFormat) {
            const actual = columnWidthFormat.columnWidth;
            if (typeof actual === "number") {
              const delta = Math.abs(actual - params.column_width!);
              if (delta > 0.1) {
                warnings.push(
                  `Requested column width ${params.column_width}, Excel applied ${actual.toFixed(2)}.`
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
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error formatting: ${e.message}` }],
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

function resolveFormatTarget(context: any, ref: string): {
  sheet: any;
  target: any;
  isMultiRange: boolean;
} {
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
