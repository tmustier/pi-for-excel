/**
 * format_cells — Apply formatting to a range (separate from write_cells).
 *
 * Handles: font (bold, italic, color, size), fill color, number format,
 * borders, alignment, column width.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { FormatCellsDetails } from "./tool-details.js";
import { excelRun, getRange, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { dispatchWorkbookSnapshotCreated } from "../workbook/recovery-events.js";
import { getWorkbookRecoveryLog, MAX_RECOVERY_CELLS } from "../workbook/recovery-log.js";
import { captureFormatCellsState, type RecoveryFormatSelection } from "../workbook/recovery-states.js";
import { getErrorMessage } from "../utils/errors.js";
import { resolveStyles } from "../conventions/index.js";
import {
  CHECKPOINT_SKIPPED_NOTE,
  CHECKPOINT_SKIPPED_REASON,
} from "./recovery-metadata.js";
import { finalizeMutationOperation, finalizeMutationRecoveryStep } from "./mutation/finalize.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";
import type { BorderWeight } from "../conventions/index.js";
import {
  buildBorderInstructions,
  normalizeBorderParams,
  type BorderEdgeIndex,
  type NormalizedBorderParams,
} from "./format-cells-borders.js";
import { getResolvedConventions } from "../conventions/store.js";
import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

const DEFAULT_FONT_NAME = "Arial";
const DEFAULT_FONT_SIZE = 10;
// Excel columnWidth in Office.js uses points. Approx conversion for Arial 10:
// 1 character width ≈ 7.2 points (based on Excel UI measurement).
const POINTS_PER_CHAR_ARIAL_10 = 7.2;

const schema = Type.Object({
  range: Type.String({
    description: 'Range to format, e.g. "A1:D1", "Sheet2!B3:B20". Supports comma/semicolon-separated ranges on the same sheet (e.g. "A1:B2, D1:D2").',
  }),
  style: Type.Optional(
    Type.Union([Type.String(), Type.Array(Type.String())], {
      description:
        'Named style(s) to apply. Compose as array (left-to-right). ' +
        'Format: "number", "integer", "currency", "percent", "ratio", "text". ' +
        'Structural: "header", "total-row", "subtotal", "input", "blank-section". ' +
        'Example: ["currency", "total-row"] = currency format + bold + top border.',
    }),
  ),
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
        'Preset name ("number", "integer", "currency", "percent", "ratio", "text") ' +
        'or raw Excel format string. Overrides style\'s number format.',
    }),
  ),
  number_format_dp: Type.Optional(
    Type.Number({
      description: "Override decimal places for a number format preset.",
    }),
  ),
  currency_symbol: Type.Optional(
    Type.String({
      description: 'Override currency symbol, e.g. "£", "€". Only with currency preset.',
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
          'Border weight for ALL edges (shorthand). Individual edge params override this.',
      },
    ),
  ),
  border_top: Type.Optional(Type.Union(
    [Type.Literal("thin"), Type.Literal("medium"), Type.Literal("thick"), Type.Literal("none")],
    { description: "Top border weight." },
  )),
  border_bottom: Type.Optional(Type.Union(
    [Type.Literal("thin"), Type.Literal("medium"), Type.Literal("thick"), Type.Literal("none")],
    { description: "Bottom border weight." },
  )),
  border_left: Type.Optional(Type.Union(
    [Type.Literal("thin"), Type.Literal("medium"), Type.Literal("thick"), Type.Literal("none")],
    { description: "Left border weight." },
  )),
  border_right: Type.Optional(Type.Union(
    [Type.Literal("thin"), Type.Literal("medium"), Type.Literal("thick"), Type.Literal("none")],
    { description: "Right border weight." },
  )),
  border_color: Type.Optional(
    Type.String({ description: 'Hex color for borders (e.g. "#000000"). Applies to all borders set in this call. Default: automatic (black).' }),
  ),
  merge: Type.Optional(
    Type.Boolean({ description: "Merge the range into a single cell." }),
  ),
});

type Params = Static<typeof schema>;

type HorizontalAlignment = "Left" | "Center" | "Right" | "General";
type VerticalAlignment = "Top" | "Center" | "Bottom";
type BorderVisibleWeight = Exclude<BorderWeight, "none">;

function isHorizontalAlignment(value: string): value is HorizontalAlignment {
  return value === "Left" || value === "Center" || value === "Right" || value === "General";
}

function isVerticalAlignment(value: string): value is VerticalAlignment {
  return value === "Top" || value === "Center" || value === "Bottom";
}

interface ResolvedFormatPropertiesForCheckpoint {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fontSize?: number;
  fontName?: string;
  fillColor?: string;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
  borderTop?: BorderWeight;
  borderBottom?: BorderWeight;
  borderLeft?: BorderWeight;
  borderRight?: BorderWeight;
}

interface FormatCheckpointPlan {
  selection: RecoveryFormatSelection;
  unsupportedReason: string | null;
}

function buildFormatCheckpointPlan(
  params: Params,
  borderParams: NormalizedBorderParams,
  props: ResolvedFormatPropertiesForCheckpoint,
  hasNumberFormat: boolean,
): FormatCheckpointPlan {
  const selection: RecoveryFormatSelection = {
    numberFormat: hasNumberFormat || undefined,
    fillColor: props.fillColor !== undefined || undefined,
    fontColor: props.fontColor !== undefined || undefined,
    bold: props.bold !== undefined || undefined,
    italic: props.italic !== undefined || undefined,
    underlineStyle: props.underline !== undefined || undefined,
    fontName: props.fontName !== undefined || undefined,
    fontSize: props.fontSize !== undefined || undefined,
    horizontalAlignment: props.horizontalAlignment !== undefined || undefined,
    verticalAlignment: props.verticalAlignment !== undefined || undefined,
    wrapText: props.wrapText !== undefined || undefined,
    columnWidth: params.column_width !== undefined || params.auto_fit === true || undefined,
    rowHeight: params.row_height !== undefined || params.auto_fit === true || undefined,
    mergedAreas: params.merge !== undefined || undefined,
  };

  const hasShorthand = borderParams.shorthand !== undefined;
  const hasParamEdges =
    borderParams.top !== undefined ||
    borderParams.bottom !== undefined ||
    borderParams.left !== undefined ||
    borderParams.right !== undefined;
  const hasStyleEdges =
    props.borderTop !== undefined ||
    props.borderBottom !== undefined ||
    props.borderLeft !== undefined ||
    props.borderRight !== undefined;

  if (hasShorthand && !hasParamEdges && !hasStyleEdges) {
    selection.borderTop = true;
    selection.borderBottom = true;
    selection.borderLeft = true;
    selection.borderRight = true;
    selection.borderInsideHorizontal = true;
    selection.borderInsideVertical = true;
  } else {
    selection.borderTop = borderParams.top !== undefined || props.borderTop !== undefined || undefined;
    selection.borderBottom = borderParams.bottom !== undefined || props.borderBottom !== undefined || undefined;
    selection.borderLeft = borderParams.left !== undefined || props.borderLeft !== undefined || undefined;
    selection.borderRight = borderParams.right !== undefined || props.borderRight !== undefined || undefined;
  }

  const hasSelectedProperty = Object.values(selection).some((value) => value === true);
  if (!hasSelectedProperty) {
    return {
      selection,
      unsupportedReason: "No restorable format properties were changed.",
    };
  }

  return {
    selection,
    unsupportedReason: null,
  };
}

const mutationFinalizeDependencies: MutationFinalizeDependencies = {
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

export function createFormatCellsTool(): AgentTool<typeof schema, FormatCellsDetails> {
  return {
    name: "format_cells",
    label: "Format Cells",
    description:
      "Apply formatting to a range of cells (supports comma-separated ranges on one sheet). " +
      "Use named styles for common patterns: style: \"currency\" or style: [\"currency\", \"total-row\"]. " +
      "Individual params (bold, fill_color, etc.) override style properties. " +
      "Does NOT modify cell values — use write_cells for that.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<FormatCellsDetails>> => {
      try {
        // ── Load stored conventions ──────────────────────────────────
        const storage = getAppStorage();
        const conventionConfig = await getResolvedConventions(storage.settings);

        const normalizedBorders = normalizeBorderParams(params);

        // ── Resolve styles + overrides into flat properties ──────────
        const styleResult = resolveStyles(params.style, {
          numberFormat: params.number_format,
          numberFormatDp: params.number_format_dp,
          currencySymbol: params.currency_symbol,
          bold: params.bold,
          italic: params.italic,
          underline: params.underline,
          fontColor: params.font_color,
          fontSize: params.font_size,
          fontName: params.font_name,
          fillColor: params.fill_color,
          horizontalAlignment: params.horizontal_alignment as "Left" | "Center" | "Right" | "General" | undefined,
          verticalAlignment: params.vertical_alignment as "Top" | "Center" | "Bottom" | undefined,
          wrapText: params.wrap_text,
          borderTop: normalizedBorders.top,
          borderBottom: normalizedBorders.bottom,
          borderLeft: normalizedBorders.left,
          borderRight: normalizedBorders.right,
        }, conventionConfig);
        const props = styleResult.properties;

        const checkpointPlan = buildFormatCheckpointPlan(
          params,
          normalizedBorders,
          props,
          styleResult.excelNumberFormat !== undefined,
        );

        let checkpointCapture: {
          supported: boolean;
          state?: Awaited<ReturnType<typeof captureFormatCellsState>>["state"];
          reason?: string;
        };

        if (checkpointPlan.unsupportedReason) {
          checkpointCapture = {
            supported: false,
            reason: checkpointPlan.unsupportedReason,
          };
        } else {
          try {
            checkpointCapture = await captureFormatCellsState(
              params.range,
              checkpointPlan.selection,
              { maxCellCount: MAX_RECOVERY_CELLS },
            );
          } catch (captureError: unknown) {
            checkpointCapture = {
              supported: false,
              reason: `Format backup capture failed: ${getErrorMessage(captureError)}`,
            };
          }
        }

        const result = await excelRun(async (context) => {
          const resolved = resolveFormatTarget(context, params.range);
          resolved.sheet.load("name");
          resolved.target.load("address");

          const requestedColumnWidth = params.column_width;

          if (!resolved.isMultiRange) {
            resolved.target.load("rowCount,columnCount");
          } else {
            resolved.target.areas.load("items/rowCount,items/columnCount");
          }

          await context.sync();

          const sheet = resolved.sheet;
          const target = resolved.target;
          const isMultiRange = resolved.isMultiRange;

          const cellCount = isMultiRange
            ? resolved.target.areas.items.reduce((total, area) => total + (area.rowCount * area.columnCount), 0)
            : resolved.target.rowCount * resolved.target.columnCount;

          const applied: string[] = [];
          const warnings: string[] = [...styleResult.warnings];
          const formatTarget = target.format;
          let columnWidthFormat: Excel.RangeFormat | null = null;

          // Report which styles were applied
          if (params.style) {
            const names = Array.isArray(params.style) ? params.style : [params.style];
            applied.push(`style ${names.join(" + ")}`);
          }

          // Font properties (from resolved style + overrides)
          if (props.bold !== undefined) {
            formatTarget.font.bold = props.bold;
            if (!params.style) applied.push(props.bold ? "bold" : "not bold");
          }
          if (props.italic !== undefined) {
            formatTarget.font.italic = props.italic;
            if (!params.style) applied.push(props.italic ? "italic" : "not italic");
          }
          if (props.underline !== undefined) {
            formatTarget.font.underline = props.underline ? "Single" : "None";
            if (!params.style) applied.push(props.underline ? "underline" : "no underline");
          }
          if (props.fontColor) {
            formatTarget.font.color = props.fontColor;
            if (!params.style) applied.push(`font color ${props.fontColor}`);
          }
          if (props.fontSize) {
            formatTarget.font.size = props.fontSize;
            if (!params.style) applied.push(`${props.fontSize}pt`);
          }
          if (props.fontName) {
            formatTarget.font.name = props.fontName;
            if (!params.style) applied.push(`font ${props.fontName}`);
          }

          // Fill (from resolved style + overrides)
          if (props.fillColor) {
            formatTarget.fill.color = props.fillColor;
            if (!params.style) applied.push(`fill ${props.fillColor}`);
          }

          // Number format (from resolved style or raw)
          if (styleResult.excelNumberFormat) {
            const numberFormat = styleResult.excelNumberFormat;
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
            if (!params.style) applied.push(`format "${numberFormat}"`);
          }

          // Alignment (from resolved style + overrides)
          if (props.horizontalAlignment) {
            const hAlign = props.horizontalAlignment;
            if (!isHorizontalAlignment(hAlign)) {
              throw new Error(
                `Invalid horizontal_alignment "${String(hAlign)}". Use Left, Center, Right, or General.`,
              );
            }
            formatTarget.horizontalAlignment = hAlign;
            if (!params.style) applied.push(`align ${hAlign.toLowerCase()}`);
          }
          if (props.verticalAlignment) {
            const vAlign = props.verticalAlignment;
            if (!isVerticalAlignment(vAlign)) {
              throw new Error(
                `Invalid vertical_alignment "${String(vAlign)}". Use Top, Center, or Bottom.`,
              );
            }
            formatTarget.verticalAlignment = vAlign;
            if (!params.style) applied.push(`v-align ${vAlign.toLowerCase()}`);
          }
          if (props.wrapText !== undefined) {
            formatTarget.wrapText = props.wrapText;
            if (!params.style) applied.push(props.wrapText ? "wrap" : "no wrap");
          }

          // Dimensions (not part of styles — always from direct params)
          if (params.column_width !== undefined) {
            const columnTarget = target.getEntireColumn();

            if (props.fontName && props.fontName !== DEFAULT_FONT_NAME) {
              warnings.push(
                `Column width assumes ${DEFAULT_FONT_NAME} ${DEFAULT_FONT_SIZE}; using ${props.fontName} may differ.`
              );
            }
            if (props.fontSize && props.fontSize !== DEFAULT_FONT_SIZE) {
              warnings.push(
                `Column width assumes ${DEFAULT_FONT_NAME} ${DEFAULT_FONT_SIZE}; using ${props.fontSize}pt may differ.`
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

          // Borders — resolve from: individual edge params > style edges > `borders` shorthand
          applyBorders(formatTarget, normalizedBorders, props, params.border_color, applied);

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

          return { sheetName: sheet.name, address: target.address, applied, warnings, isMultiRange, cellCount };
        });

        const fullAddr = result.isMultiRange
          ? result.address
          : qualifiedAddress(result.sheetName, result.address);
        const warningText = result.warnings.length
          ? `\n\n⚠️ ${result.warnings.join("\n")}`
          : "";

        const toolResult: AgentToolResult<FormatCellsDetails> = {
          content: [
            {
              type: "text",
              text: `Formatted **${fullAddr}**: ${result.applied.join(", ")}.${warningText}`,
            },
          ],
          details: {
            kind: "format_cells",
            address: fullAddr,
            warningsCount: result.warnings.length,
          },
        };

        const recoveryUnavailableReason = checkpointCapture.supported && checkpointCapture.state
          ? CHECKPOINT_SKIPPED_REASON
          : (checkpointCapture.reason ?? CHECKPOINT_SKIPPED_REASON);

        await finalizeMutationRecoveryStep({
          result: toolResult,
          appendRecoverySnapshot: () => {
            if (!checkpointCapture.supported || !checkpointCapture.state) {
              return Promise.resolve(null);
            }

            return getWorkbookRecoveryLog().appendFormatCells({
              toolName: "format_cells",
              toolCallId,
              address: fullAddr,
              changedCount: result.cellCount,
              formatRangeState: checkpointCapture.state,
            });
          },
          appendResultNote: appendMutationResultNote,
          unavailableReason: recoveryUnavailableReason,
          unavailableNote: CHECKPOINT_SKIPPED_NOTE,
          dispatchSnapshotCreated: (checkpoint) => {
            dispatchWorkbookSnapshotCreated({
              snapshotId: checkpoint.id,
              toolName: checkpoint.toolName,
              address: checkpoint.address,
              changedCount: checkpoint.changedCount,
            });
          },
        });

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "format_cells",
            toolCallId,
            blocked: false,
            outputAddress: fullAddr,
            changedCount: result.cellCount,
            changes: [],
            summary: `formatted ${result.cellCount} cell(s)` +
              (result.warnings.length > 0 ? ` with ${result.warnings.length} warning(s)` : ""),
          },
        });

        return toolResult;
      } catch (e: unknown) {
        const message = getErrorMessage(e);

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "format_cells",
            toolCallId,
            blocked: true,
            outputAddress: params.range,
            changedCount: 0,
            changes: [],
            summary: `error: ${message}`,
          },
        });

        return {
          content: [{ type: "text", text: `Error formatting: ${message}` }],
          details: { kind: "format_cells", address: params.range },
        };
      }
    },
  };
}

// ── Border application ───────────────────────────────────────────────

/** Map a border weight string to the Office.js enum value. */
function toBorderWeight(weight: BorderVisibleWeight): "Thin" | "Medium" | "Thick" {
  if (weight === "thin") return "Thin";
  if (weight === "medium") return "Medium";
  return "Thick";
}

/** Apply a single border edge. */
function applyEdge(
  formatTarget: Excel.RangeFormat,
  edge: BorderEdgeIndex,
  weight: BorderWeight,
  color?: string,
): void {
  const borderItem = formatTarget.borders.getItem(edge);
  if (weight === "none") {
    borderItem.style = "None";
    return;
  }

  borderItem.style = "Continuous";
  borderItem.weight = toBorderWeight(weight);
  if (color) {
    borderItem.color = color;
  }
}

/**
 * Resolve and apply borders. Priority:
 *   1. Individual edge params (border_top, etc.) — highest
 *   2. Style-resolved edges (from named styles)
 *   3. `borders` shorthand (all edges + inside) — lowest
 */
function applyBorders(
  formatTarget: Excel.RangeFormat,
  borderParams: NormalizedBorderParams,
  props: { borderTop?: BorderWeight; borderBottom?: BorderWeight; borderLeft?: BorderWeight; borderRight?: BorderWeight },
  color: string | undefined,
  applied: string[],
): void {
  const instructions = buildBorderInstructions(borderParams, props, color);
  if (instructions === null) {
    return;
  }

  for (const { edge, weight } of instructions.operations) {
    applyEdge(formatTarget, edge, weight, color);
  }

  applied.push(instructions.appliedText);
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
