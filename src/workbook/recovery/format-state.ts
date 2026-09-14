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

type CapturedFormatArea =
  | { supported: true; state: RecoveryFormatAreaState }
  | { supported: false; reason: string };

function prepareFormatAreaCapture(
  area: Excel.Range,
  sheetName: string,
  selection: RecoveryFormatSelection,
): PreparedFormatAreaCapture {
  const prepared: PreparedFormatAreaCapture = {
    range: area,
    address: qualifyAddressWithSheet(sheetName, area.address),
    rowCount: area.rowCount,
    columnCount: area.columnCount,
    columnFormats: [],
    rowFormats: [],
    mergedAreaAddresses: [],
    borders: {},
  };
  const loadFont = selection.fontColor === true ||
    selection.bold === true ||
    selection.italic === true ||
    selection.underlineStyle === true ||
    selection.fontName === true ||
    selection.fontSize === true;
  const loadBorders = selection.borderTop === true ||
    selection.borderBottom === true ||
    selection.borderLeft === true ||
    selection.borderRight === true ||
    selection.borderInsideHorizontal === true ||
    selection.borderInsideVertical === true;

  if (selection.numberFormat === true) area.load("numberFormat");
  if (selection.fillColor === true) area.format.fill.load("color");
  if (loadFont) area.format.font.load("color,bold,italic,underline,name,size");
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

  if (loadBorders) {
    for (const borderKey of RECOVERY_BORDER_KEYS) {
      if (selection[borderKey] !== true) continue;
      const border = area.format.borders.getItem(BORDER_KEY_TO_EDGE[borderKey]);
      border.load("style,weight,color");
      prepared.borders[borderKey] = border;
    }
  }

  return prepared;
}

async function loadMergedAreaAddresses(
  context: Excel.RequestContext,
  preparedAreas: PreparedFormatAreaCapture[],
  sheetName: string,
  selection: RecoveryFormatSelection,
): Promise<void> {
  if (selection.mergedAreas !== true) return;

  for (const prepared of preparedAreas) {
    const mergedAreas = prepared.mergedAreas;
    if (!mergedAreas || mergedAreas.isNullObject) continue;
    mergedAreas.areas.load("items/address");
  }

  await context.sync();

  for (const prepared of preparedAreas) {
    const mergedAreas = prepared.mergedAreas;
    if (!mergedAreas || mergedAreas.isNullObject) continue;
    prepared.mergedAreaAddresses = dedupeRecoveryAddresses(
      mergedAreas.areas.items.map((areaRange) => qualifyAddressWithSheet(sheetName, areaRange.address)),
    );
  }
}

function captureScalarFormatState(
  prepared: PreparedFormatAreaCapture,
  selection: RecoveryFormatSelection,
  areaState: RecoveryFormatAreaState,
): string | null {
  if (selection.numberFormat === true) {
    const matrix = validateStringGrid(prepared.range.numberFormat, prepared.rowCount, prepared.columnCount);
    if (!matrix) return "Format checkpoint capture failed: number format matrix is invalid.";
    areaState.numberFormat = matrix;
  }

  if (selection.fillColor === true) {
    const value = normalizeOptionalString(prepared.range.format.fill.color);
    if (value === undefined) return "Format checkpoint capture failed: fill color is not restorable.";
    areaState.fillColor = value;
  }

  if (selection.fontColor === true) {
    const value = normalizeOptionalString(prepared.range.format.font.color);
    if (value === undefined) return "Format checkpoint capture failed: font color is not restorable.";
    areaState.fontColor = value;
  }

  if (selection.bold === true) {
    const value = normalizeOptionalBoolean(prepared.range.format.font.bold);
    if (value === undefined) return "Format checkpoint capture failed: bold state is mixed or unsupported.";
    areaState.bold = value;
  }

  if (selection.italic === true) {
    const value = normalizeOptionalBoolean(prepared.range.format.font.italic);
    if (value === undefined) return "Format checkpoint capture failed: italic state is mixed or unsupported.";
    areaState.italic = value;
  }

  if (selection.underlineStyle === true) {
    const value = prepared.range.format.font.underline;
    if (!isRecoveryUnderlineStyle(value)) return "Format checkpoint capture failed: underline style is unsupported.";
    areaState.underlineStyle = value;
  }

  if (selection.fontName === true) {
    const value = normalizeOptionalString(prepared.range.format.font.name);
    if (value === undefined) return "Format checkpoint capture failed: font name is not restorable.";
    areaState.fontName = value;
  }

  if (selection.fontSize === true) {
    const value = normalizeOptionalNumber(prepared.range.format.font.size);
    if (value === undefined) return "Format checkpoint capture failed: font size is mixed or unsupported.";
    areaState.fontSize = value;
  }

  return null;
}

function captureLayoutFormatState(
  prepared: PreparedFormatAreaCapture,
  selection: RecoveryFormatSelection,
  areaState: RecoveryFormatAreaState,
): string | null {
  if (selection.horizontalAlignment === true) {
    const value = prepared.range.format.horizontalAlignment;
    if (!isRecoveryHorizontalAlignment(value)) {
      return "Format checkpoint capture failed: horizontal alignment is unsupported.";
    }
    areaState.horizontalAlignment = value;
  }

  if (selection.verticalAlignment === true) {
    const value = prepared.range.format.verticalAlignment;
    if (!isRecoveryVerticalAlignment(value)) {
      return "Format checkpoint capture failed: vertical alignment is unsupported.";
    }
    areaState.verticalAlignment = value;
  }

  if (selection.wrapText === true) {
    const value = normalizeOptionalBoolean(prepared.range.format.wrapText);
    if (value === undefined) return "Format checkpoint capture failed: wrap-text state is mixed or unsupported.";
    areaState.wrapText = value;
  }

  if (selection.mergedAreas === true) {
    areaState.mergedAreas = [...prepared.mergedAreaAddresses];
  }

  return null;
}

function captureDimensionFormatState(
  prepared: PreparedFormatAreaCapture,
  selection: RecoveryFormatSelection,
  areaState: RecoveryFormatAreaState,
): string | null {
  if (selection.columnWidth === true) {
    const columnWidths: number[] = [];
    for (const columnFormat of prepared.columnFormats) {
      const width = normalizeOptionalNumber(columnFormat.columnWidth);
      if (width === undefined) {
        return "Format checkpoint capture failed: column width is mixed or unsupported.";
      }
      columnWidths.push(width);
    }
    areaState.columnWidths = columnWidths;
  }

  if (selection.rowHeight === true) {
    const rowHeights: number[] = [];
    for (const rowFormat of prepared.rowFormats) {
      const height = normalizeOptionalNumber(rowFormat.rowHeight);
      if (height === undefined) {
        return "Format checkpoint capture failed: row height is mixed or unsupported.";
      }
      rowHeights.push(height);
    }
    areaState.rowHeights = rowHeights;
  }

  return null;
}

function captureBorderFormatState(
  prepared: PreparedFormatAreaCapture,
  selection: RecoveryFormatSelection,
  areaState: RecoveryFormatAreaState,
): string | null {
  for (const borderKey of RECOVERY_BORDER_KEYS) {
    if (selection[borderKey] !== true) continue;
    const border = prepared.borders[borderKey];
    if (!border) return "Format checkpoint capture failed: border state is unavailable.";
    const borderState = captureBorderState(border);
    if (!borderState) return "Format checkpoint capture failed: border state is unsupported.";
    areaState[borderKey] = borderState;
  }
  return null;
}

function capturePreparedFormatArea(
  prepared: PreparedFormatAreaCapture,
  selection: RecoveryFormatSelection,
): CapturedFormatArea {
  const state: RecoveryFormatAreaState = {
    address: prepared.address,
    rowCount: prepared.rowCount,
    columnCount: prepared.columnCount,
  };

  const reason = captureScalarFormatState(prepared, selection, state) ??
    captureLayoutFormatState(prepared, selection, state) ??
    captureDimensionFormatState(prepared, selection, state) ??
    captureBorderFormatState(prepared, selection, state);

  return reason ? { supported: false, reason } : { supported: true, state };
}

async function captureFormatRangeStateWithSelection(
  context: Excel.RequestContext,
  target: ResolvedFormatCaptureTarget,
  selection: RecoveryFormatSelection,
  maxCellCount?: number,
): Promise<RecoveryFormatCaptureResult> {
  const captureCellCount = estimateFormatCaptureCellCount(target.areas, selection);

  if (typeof maxCellCount === "number" && Number.isFinite(maxCellCount) && captureCellCount > maxCellCount) {
    return {
      supported: false,
      reason: `Format checkpoint capture skipped: snapshot size exceeds ${maxCellCount.toLocaleString()} units.`,
    };
  }

  const preparedAreas = target.areas.map((area) =>
    prepareFormatAreaCapture(area, target.sheetName, selection));

  await context.sync();
  await loadMergedAreaAddresses(context, preparedAreas, target.sheetName, selection);

  const areaStates: RecoveryFormatAreaState[] = [];
  for (const prepared of preparedAreas) {
    const captured = capturePreparedFormatArea(prepared, selection);
    if (!captured.supported) return captured;
    areaStates.push(captured.state);
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
  if (state.numberFormat !== undefined) {
    range.numberFormat = cloneStringGrid(state.numberFormat);
  }

  if (state.fillColor !== undefined) {
    range.format.fill.color = state.fillColor;
  }

  if (state.fontColor !== undefined) {
    range.format.font.color = state.fontColor;
  }

  if (state.bold !== undefined) {
    range.format.font.bold = state.bold;
  }

  if (state.italic !== undefined) {
    range.format.font.italic = state.italic;
  }

  if (state.underlineStyle !== undefined) {
    if (!isRecoveryUnderlineStyle(state.underlineStyle)) {
      throw new Error("Format checkpoint is invalid: underline style is unsupported.");
    }
    range.format.font.underline = state.underlineStyle;
  }

  if (state.fontName !== undefined) {
    range.format.font.name = state.fontName;
  }

  if (state.fontSize !== undefined) {
    range.format.font.size = state.fontSize;
  }

  if (state.horizontalAlignment !== undefined) {
    if (!isRecoveryHorizontalAlignment(state.horizontalAlignment)) {
      throw new Error("Format checkpoint is invalid: horizontal alignment is unsupported.");
    }
    range.format.horizontalAlignment = state.horizontalAlignment;
  }

  if (state.verticalAlignment !== undefined) {
    if (!isRecoveryVerticalAlignment(state.verticalAlignment)) {
      throw new Error("Format checkpoint is invalid: vertical alignment is unsupported.");
    }
    range.format.verticalAlignment = state.verticalAlignment;
  }

  if (state.wrapText !== undefined) {
    range.format.wrapText = state.wrapText;
  }

  if (state.columnWidths !== undefined) {
    for (const [columnIndex, width] of state.columnWidths.entries()) {
      range.getColumn(columnIndex).format.columnWidth = width;
    }
  }

  if (state.rowHeights !== undefined) {
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

export interface CaptureFormatCellsStateOptions {
  maxCellCount?: number;
}

export async function captureFormatCellsState(
  address: string,
  selection: RecoveryFormatSelection,
  options: CaptureFormatCellsStateOptions = {},
): Promise<RecoveryFormatCaptureResult> {
  if (!hasSelectedFormatProperty(selection)) {
    return {
      supported: false,
      reason: "No restorable format properties were selected.",
    };
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
        areaState.numberFormat !== undefined ||
        areaState.columnWidths !== undefined ||
        areaState.rowHeights !== undefined;

      if (requiresExactShape) {
        if (range.rowCount !== areaState.rowCount || range.columnCount !== areaState.columnCount) {
          throw new Error("Format checkpoint range shape changed and cannot be restored safely.");
        }
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
