/**
 * format_cells — Apply formatting to a range (separate from write_cells).
 *
 * Handles: font (bold, italic, color, size), fill color, number format,
 * borders, alignment, column width.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";

const schema = Type.Object({
  range: Type.String({
    description: 'Range to format, e.g. "A1:D1", "Sheet2!B3:B20".',
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
  column_width: Type.Optional(Type.Number({ description: "Set column width in points." })),
  row_height: Type.Optional(Type.Number({ description: "Set row height in points." })),
  auto_fit: Type.Optional(
    Type.Boolean({ description: "Auto-fit column widths to content. Default: false." }),
  ),
  borders: Type.Optional(
    Type.String({
      description:
        'Border style: "thin", "medium", "thick", or "none" to remove borders. Applied to all edges.',
    }),
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
      "Apply formatting to a range of cells. Set font properties (bold, italic, color, size), " +
      "fill color, number format, alignment, borders, column width, and more. " +
      "Does NOT modify cell values — use write_cells for that.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await excelRun(async (context: any) => {
          const { sheet, range } = getRange(context, params.range);
          sheet.load("name");
          range.load("address,rowCount,columnCount");
          await context.sync();

          const applied: string[] = [];

          // Font properties
          if (params.bold !== undefined) {
            range.format.font.bold = params.bold;
            applied.push(params.bold ? "bold" : "not bold");
          }
          if (params.italic !== undefined) {
            range.format.font.italic = params.italic;
            applied.push(params.italic ? "italic" : "not italic");
          }
          if (params.underline !== undefined) {
            range.format.font.underline = params.underline ? "Single" : "None";
            applied.push(params.underline ? "underline" : "no underline");
          }
          if (params.font_color) {
            range.format.font.color = params.font_color;
            applied.push(`font color ${params.font_color}`);
          }
          if (params.font_size) {
            range.format.font.size = params.font_size;
            applied.push(`${params.font_size}pt`);
          }
          if (params.font_name) {
            range.format.font.name = params.font_name;
            applied.push(`font ${params.font_name}`);
          }

          // Fill
          if (params.fill_color) {
            range.format.fill.color = params.fill_color;
            applied.push(`fill ${params.fill_color}`);
          }

          // Number format
          if (params.number_format) {
            const formatMatrix = Array.from({ length: range.rowCount }, () =>
              Array.from({ length: range.columnCount }, () => params.number_format),
            );
            range.numberFormat = formatMatrix as any;
            applied.push(`format "${params.number_format}"`);
          }

          // Alignment
          if (params.horizontal_alignment) {
            range.format.horizontalAlignment = params.horizontal_alignment;
            applied.push(`align ${params.horizontal_alignment.toLowerCase()}`);
          }
          if (params.vertical_alignment) {
            range.format.verticalAlignment = params.vertical_alignment;
            applied.push(`v-align ${params.vertical_alignment.toLowerCase()}`);
          }
          if (params.wrap_text !== undefined) {
            range.format.wrapText = params.wrap_text;
            applied.push(params.wrap_text ? "wrap" : "no wrap");
          }

          // Dimensions
          if (params.column_width) {
            range.format.columnWidth = params.column_width;
            applied.push(`col width ${params.column_width}`);
          }
          if (params.row_height) {
            range.format.rowHeight = params.row_height;
            applied.push(`row height ${params.row_height}`);
          }
          if (params.auto_fit) {
            range.format.autofitColumns();
            range.format.autofitRows();
            applied.push("auto-fit");
          }

          // Borders
          if (params.borders) {
            const style =
              params.borders === "none"
                ? "None"
                : params.borders === "thin"
                  ? "Thin"
                  : params.borders === "medium"
                    ? "Medium"
                    : params.borders === "thick"
                      ? "Thick"
                      : "Thin";

            const borders = [
              "EdgeTop",
              "EdgeBottom",
              "EdgeLeft",
              "EdgeRight",
              "InsideHorizontal",
              "InsideVertical",
            ];
            for (const border of borders) {
              range.format.borders.getItem(border).style = style;
            }
            applied.push(`${params.borders} borders`);
          }

          // Merge
          if (params.merge !== undefined) {
            if (params.merge) {
              range.merge();
              applied.push("merged");
            } else {
              range.unmerge();
              applied.push("unmerged");
            }
          }

          await context.sync();
          return { sheetName: sheet.name, address: range.address, applied };
        });

        const fullAddr = qualifiedAddress(result.sheetName, result.address);
        return {
          content: [
            {
              type: "text",
              text: `✅ Formatted **${fullAddr}**: ${result.applied.join(", ")}.`,
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
