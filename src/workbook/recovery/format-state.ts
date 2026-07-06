/** Format-state capture/apply for workbook recovery snapshots. */

import { excelRun, getRange, parseRangeRef } from "../../excel/helpers.js";
import { qualifyAddressWithSheet, splitRangeList } from "./address.js";
import {
  cloneRecoveryFormatRangeState,
  cloneRecoveryFormatSelection,
  cloneStringGrid,
} from "./clone.js";
import {
  estimateFormatCaptureCellCount,
  hasSelectedFormatProperty,
} from "./format-selection.js";
import type {
  RecoveryFormatAreaState,
  RecoveryFormatCaptureResult,
  RecoveryFormatRangeState,
  RecoveryFormatSelection,
} from "./types.js";

import {
  BORDER_KEY_TO_EDGE,
  RECOVERY_BORDER_KEYS,
  applyBorderState,
  captureBorderState,
  isRecoveryHorizontalAlignment,
  isRecoveryUnderlineStyle,
  isRecoveryVerticalAlignment,
  normalizeOptionalBoolean,
  normalizeOptionalNumber,
  normalizeOptionalString,
  type RecoveryBorderKey,
} from "./format-state-normalization.js";
import {
  collectMergedAreaAddresses,
  dedupeRecoveryAddresses,
  validateStringGrid,
} from "./format-state-utils.js";

interface ResolvedFormatCaptureTarget {
  sheetName: string;
  areas: Excel.Range[];
}

async function resolveFormatCaptureTarget(
  context: Excel.RequestContext,
  ref: string,
): Promise<ResolvedFormatCaptureTarget> {
  const parts = splitRangeList(ref);

  if (parts.length <= 1) {
    const { sheet, range } = getRange(context, ref);
    sheet.load("name");
    range.load("address,rowCount,columnCount");
    await context.sync();
    return {
      sheetName: sheet.name,
      areas: [range],
    };
  }

  let sheetNameFromRef: string | undefined;
  const areaAddresses: string[] = [];

  for (const part of parts) {
    const parsed = parseRangeRef(part);
    if (parsed.sheet) {
      if (sheetNameFromRef && parsed.sheet !== sheetNameFromRef) {
        throw new Error("Format checkpoint capture supports a single sheet per mutation.");
      }
      sheetNameFromRef = parsed.sheet;
    }

    areaAddresses.push(parsed.address);
  }

  const sheet = sheetNameFromRef
    ? context.workbook.worksheets.getItem(sheetNameFromRef)
    : context.workbook.worksheets.getActiveWorksheet();
  const areasTarget = sheet.getRanges(areaAddresses.join(","));

  sheet.load("name");
  areasTarget.areas.load("items/address,items/rowCount,items/columnCount");
  await context.sync();

  return {
    sheetName: sheet.name,
    areas: [...areasTarget.areas.items],
  };
}

interface PreparedFormatAreaCapture {
  range: Excel.Range;
  address: string;
  rowCount: number;
  columnCount: number;
  columnFormats: Excel.RangeFormat[];
  rowFormats: Excel.RangeFormat[];
  mergedAreas?: Excel.RangeAreas;
  mergedAreaAddresses: string[];
  borders: Partial<Record<RecoveryBorderKey, Excel.RangeBorder>>;
}

async function captureFormatRangeStateWithSelection(
  context: Excel.RequestContext,
  target: ResolvedFormatCaptureTarget,
  selection: RecoveryFormatSelection,
  maxCellCount?: number,
): Promise<RecoveryFormatCaptureResult> {
  if (!hasSelectedFormatProperty(selection)) {
    return {
      supported: false,
      reason: "No restorable format properties were selected.",
    };
  }

  const captureCellCount = estimateFormatCaptureCellCount(target.areas, selection);

  if (typeof maxCellCount === "number" && Number.isFinite(maxCellCount) && captureCellCount > maxCellCount) {
    return {
      supported: false,
      reason: `Format checkpoint capture skipped: snapshot size exceeds ${maxCellCount.toLocaleString()} units.`,
    };
  }

  const preparedAreas: PreparedFormatAreaCapture[] = [];

  const needsFontLoad =
    selection.fontColor === true ||
    selection.bold === true ||
    selection.italic === true ||
    selection.underlineStyle === true ||
    selection.fontName === true ||
    selection.fontSize === true;

  const needsBorderLoad =
    selection.borderTop === true ||
    selection.borderBottom === true ||
    selection.borderLeft === true ||
    selection.borderRight === true ||
    selection.borderInsideHorizontal === true ||
    selection.borderInsideVertical === true;

  for (const area of target.areas) {
    const prepared: PreparedFormatAreaCapture = {
      range: area,
      address: qualifyAddressWithSheet(target.sheetName, area.address),
      rowCount: area.rowCount,
      columnCount: area.columnCount,
      columnFormats: [],
      rowFormats: [],
      mergedAreaAddresses: [],
      borders: {},
    };

    if (selection.numberFormat === true) {
      area.load("numberFormat");
    }

    if (selection.fillColor === true) {
      area.format.fill.load("color");
    }

    if (needsFontLoad) {
      area.format.font.load("color,bold,italic,underline,name,size");
    }

    if (selection.horizontalAlignment === true || selection.verticalAlignment === true || selection.wrapText === true) {
      area.format.load("horizontalAlignment,verticalAlignment,wrapText");
    }

    if (selection.columnWidth === true) {
      for (let columnIndex = 0; columnIndex < area.columnCount; columnIndex += 1) {
        const columnFormat = area.getColumn(columnIndex).format;
        columnFormat.load("columnWidth");
        prepared.columnFormats.push(columnFormat);
      }
    }

    if (selection.rowHeight === true) {
      for (let rowIndex = 0; rowIndex < area.rowCount; rowIndex += 1) {
        const rowFormat = area.getRow(rowIndex).format;
        rowFormat.load("rowHeight");
        prepared.rowFormats.push(rowFormat);
      }
    }

    if (selection.mergedAreas === true) {
      const mergedAreas = area.getMergedAreasOrNullObject();
      mergedAreas.load("isNullObject");
      prepared.mergedAreas = mergedAreas;
    }

    if (needsBorderLoad) {
      for (const borderKey of RECOVERY_BORDER_KEYS) {
        if (selection[borderKey] !== true) continue;

        const border = area.format.borders.getItem(BORDER_KEY_TO_EDGE[borderKey]);
        border.load("style,weight,color");
        prepared.borders[borderKey] = border;
      }
    }

    preparedAreas.push(prepared);
  }

  await context.sync();

  if (selection.mergedAreas === true) {
    for (const prepared of preparedAreas) {
      const mergedAreas = prepared.mergedAreas;
      if (!mergedAreas || mergedAreas.isNullObject) {
        prepared.mergedAreaAddresses = [];
        continue;
      }

      mergedAreas.areas.load("items/address");
    }

    await context.sync();

    for (const prepared of preparedAreas) {
      const mergedAreas = prepared.mergedAreas;
      if (!mergedAreas || mergedAreas.isNullObject) {
        prepared.mergedAreaAddresses = [];
        continue;
      }

      prepared.mergedAreaAddresses = dedupeRecoveryAddresses(
        mergedAreas.areas.items.map((areaRange) =>
          qualifyAddressWithSheet(target.sheetName, areaRange.address),
        ),
      );
    }
  }

  const areaStates: RecoveryFormatAreaState[] = [];

  for (const prepared of preparedAreas) {
    const areaState: RecoveryFormatAreaState = {
      address: prepared.address,
      rowCount: prepared.rowCount,
      columnCount: prepared.columnCount,
    };

    if (selection.numberFormat === true) {
      const matrix = validateStringGrid(prepared.range.numberFormat, prepared.rowCount, prepared.columnCount);
      if (!matrix) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: number format matrix is invalid.",
        };
      }

      areaState.numberFormat = matrix;
    }

    if (selection.fillColor === true) {
      const fillColor = normalizeOptionalString(prepared.range.format.fill.color);
      if (fillColor === undefined) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: fill color is not restorable.",
        };
      }

      areaState.fillColor = fillColor;
    }

    if (selection.fontColor === true) {
      const fontColor = normalizeOptionalString(prepared.range.format.font.color);
      if (fontColor === undefined) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: font color is not restorable.",
        };
      }

      areaState.fontColor = fontColor;
    }

    if (selection.bold === true) {
      const bold = normalizeOptionalBoolean(prepared.range.format.font.bold);
      if (bold === undefined) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: bold state is mixed or unsupported.",
        };
      }

      areaState.bold = bold;
    }

    if (selection.italic === true) {
      const italic = normalizeOptionalBoolean(prepared.range.format.font.italic);
      if (italic === undefined) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: italic state is mixed or unsupported.",
        };
      }

      areaState.italic = italic;
    }

    if (selection.underlineStyle === true) {
      const underline = prepared.range.format.font.underline;
      if (!isRecoveryUnderlineStyle(underline)) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: underline style is unsupported.",
        };
      }

      areaState.underlineStyle = underline;
    }

    if (selection.fontName === true) {
      const fontName = normalizeOptionalString(prepared.range.format.font.name);
      if (fontName === undefined) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: font name is not restorable.",
        };
      }

      areaState.fontName = fontName;
    }

    if (selection.fontSize === true) {
      const fontSize = normalizeOptionalNumber(prepared.range.format.font.size);
      if (fontSize === undefined) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: font size is mixed or unsupported.",
        };
      }

      areaState.fontSize = fontSize;
    }

    if (selection.horizontalAlignment === true) {
      const horizontalAlignment = prepared.range.format.horizontalAlignment;
      if (!isRecoveryHorizontalAlignment(horizontalAlignment)) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: horizontal alignment is unsupported.",
        };
      }

      areaState.horizontalAlignment = horizontalAlignment;
    }

    if (selection.verticalAlignment === true) {
      const verticalAlignment = prepared.range.format.verticalAlignment;
      if (!isRecoveryVerticalAlignment(verticalAlignment)) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: vertical alignment is unsupported.",
        };
      }

      areaState.verticalAlignment = verticalAlignment;
    }

    if (selection.wrapText === true) {
      const wrapText = normalizeOptionalBoolean(prepared.range.format.wrapText);
      if (wrapText === undefined) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: wrap-text state is mixed or unsupported.",
        };
      }

      areaState.wrapText = wrapText;
    }

    if (selection.columnWidth === true) {
      const columnWidths: number[] = [];
      for (const columnFormat of prepared.columnFormats) {
        const width = normalizeOptionalNumber(columnFormat.columnWidth);
        if (width === undefined) {
          return {
            supported: false,
            reason: "Format checkpoint capture failed: column width is mixed or unsupported.",
          };
        }

        columnWidths.push(width);
      }

      if (columnWidths.length !== prepared.columnCount) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: column-width count mismatch.",
        };
      }

      areaState.columnWidths = columnWidths;
    }

    if (selection.rowHeight === true) {
      const rowHeights: number[] = [];
      for (const rowFormat of prepared.rowFormats) {
        const height = normalizeOptionalNumber(rowFormat.rowHeight);
        if (height === undefined) {
          return {
            supported: false,
            reason: "Format checkpoint capture failed: row height is mixed or unsupported.",
          };
        }

        rowHeights.push(height);
      }

      if (rowHeights.length !== prepared.rowCount) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: row-height count mismatch.",
        };
      }

      areaState.rowHeights = rowHeights;
    }

    if (selection.mergedAreas === true) {
      areaState.mergedAreas = [...prepared.mergedAreaAddresses];
    }

    for (const borderKey of RECOVERY_BORDER_KEYS) {
      if (selection[borderKey] !== true) continue;

      const border = prepared.borders[borderKey];
      if (!border) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: border state is unavailable.",
        };
      }

      const borderState = captureBorderState(border);
      if (!borderState) {
        return {
          supported: false,
          reason: "Format checkpoint capture failed: border state is unsupported.",
        };
      }

      areaState[borderKey] = borderState;
    }

    areaStates.push(areaState);
  }

  return {
    supported: true,
    state: {
      selection: cloneRecoveryFormatSelection(selection),
      areas: areaStates,
      cellCount: captureCellCount,
    },
  };
}

function applyFormatRangeStateToArea(range: Excel.Range, state: RecoveryFormatAreaState): void {
  if (state.numberFormat) {
    range.numberFormat = cloneStringGrid(state.numberFormat);
  }

  if (typeof state.fillColor === "string") {
    range.format.fill.color = state.fillColor;
  }

  if (typeof state.fontColor === "string") {
    range.format.font.color = state.fontColor;
  }

  if (typeof state.bold === "boolean") {
    range.format.font.bold = state.bold;
  }

  if (typeof state.italic === "boolean") {
    range.format.font.italic = state.italic;
  }

  if (typeof state.underlineStyle === "string") {
    if (!isRecoveryUnderlineStyle(state.underlineStyle)) {
      throw new Error("Format checkpoint is invalid: underline style is unsupported.");
    }
    range.format.font.underline = state.underlineStyle;
  }

  if (typeof state.fontName === "string") {
    range.format.font.name = state.fontName;
  }

  if (typeof state.fontSize === "number") {
    range.format.font.size = state.fontSize;
  }

  if (typeof state.horizontalAlignment === "string") {
    if (!isRecoveryHorizontalAlignment(state.horizontalAlignment)) {
      throw new Error("Format checkpoint is invalid: horizontal alignment is unsupported.");
    }
    range.format.horizontalAlignment = state.horizontalAlignment;
  }

  if (typeof state.verticalAlignment === "string") {
    if (!isRecoveryVerticalAlignment(state.verticalAlignment)) {
      throw new Error("Format checkpoint is invalid: vertical alignment is unsupported.");
    }
    range.format.verticalAlignment = state.verticalAlignment;
  }

  if (typeof state.wrapText === "boolean") {
    range.format.wrapText = state.wrapText;
  }

  if (Array.isArray(state.columnWidths)) {
    for (const [columnIndex, width] of state.columnWidths.entries()) {
      range.getColumn(columnIndex).format.columnWidth = width;
    }
  }

  if (Array.isArray(state.rowHeights)) {
    for (const [rowIndex, height] of state.rowHeights.entries()) {
      range.getRow(rowIndex).format.rowHeight = height;
    }
  }

  for (const borderKey of RECOVERY_BORDER_KEYS) {
    const borderState = state[borderKey];
    if (!borderState) continue;

    const border = range.format.borders.getItem(BORDER_KEY_TO_EDGE[borderKey]);
    applyBorderState(border, borderState);
  }
}

function captureFormatRangeStateUnsupported(reason: string): RecoveryFormatCaptureResult {
  return {
    supported: false,
    reason,
  };
}

export interface CaptureFormatCellsStateOptions {
  maxCellCount?: number;
}

export async function captureFormatCellsState(
  address: string,
  selection: RecoveryFormatSelection,
  options: CaptureFormatCellsStateOptions = {},
): Promise<RecoveryFormatCaptureResult> {
  if (!hasSelectedFormatProperty(selection)) {
    return captureFormatRangeStateUnsupported("No restorable format properties were selected.");
  }

  return excelRun<RecoveryFormatCaptureResult>(async (context) => {
    const target = await resolveFormatCaptureTarget(context, address);
    return captureFormatRangeStateWithSelection(context, target, selection, options.maxCellCount);
  });
}

export async function applyFormatCellsState(
  address: string,
  targetState: RecoveryFormatRangeState,
): Promise<RecoveryFormatRangeState> {
  const previousStateResult = await captureFormatCellsState(address, targetState.selection);
  if (!previousStateResult.supported || !previousStateResult.state) {
    throw new Error(previousStateResult.reason ?? "Format checkpoint cannot be restored safely.");
  }
  const previousState = previousStateResult.state;

  return excelRun<RecoveryFormatRangeState>(async (context) => {
    const loadedAreas = targetState.areas.map((areaState) => {
      const { range } = getRange(context, areaState.address);
      range.load("rowCount,columnCount");
      return { areaState, range };
    });

    await context.sync();

    const restoreMergedAreas = targetState.selection.mergedAreas === true;
    const currentMergedAddresses = restoreMergedAreas
      ? collectMergedAreaAddresses(previousState)
      : [];
    const targetMergedAddresses = restoreMergedAreas
      ? collectMergedAreaAddresses(targetState)
      : [];

    for (const loaded of loadedAreas) {
      const { areaState, range } = loaded;

      const requiresExactShape =
        typeof areaState.numberFormat !== "undefined" ||
        typeof areaState.columnWidths !== "undefined" ||
        typeof areaState.rowHeights !== "undefined";

      if (requiresExactShape) {
        if (range.rowCount !== areaState.rowCount || range.columnCount !== areaState.columnCount) {
          throw new Error("Format checkpoint range shape changed and cannot be restored safely.");
        }
      }

      if (Array.isArray(areaState.columnWidths) && areaState.columnWidths.length !== areaState.columnCount) {
        throw new Error("Format checkpoint is invalid: column-width data does not match range shape.");
      }

      if (Array.isArray(areaState.rowHeights) && areaState.rowHeights.length !== areaState.rowCount) {
        throw new Error("Format checkpoint is invalid: row-height data does not match range shape.");
      }

      applyFormatRangeStateToArea(range, areaState);
    }

    if (restoreMergedAreas) {
      for (const mergedAddress of currentMergedAddresses) {
        const { range } = getRange(context, mergedAddress);
        range.unmerge();
      }

      for (const mergedAddress of targetMergedAddresses) {
        const { range } = getRange(context, mergedAddress);
        range.merge();
      }
    }

    await context.sync();
    return cloneRecoveryFormatRangeState(previousState);
  });
}
