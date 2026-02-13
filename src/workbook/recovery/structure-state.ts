/** Structure-state capture/apply for workbook recovery snapshots. */

import { excelRun } from "../../excel/helpers.js";
import { localAddressPart } from "./address.js";
import { cloneRecoveryModifyStructureState } from "./clone.js";
import { MAX_RECOVERY_CELLS } from "./constants.js";
import type {
  RecoveryModifyStructureState,
  RecoverySheetVisibility,
  RecoveryStructureValueRangeState,
} from "./types.js";

interface SyncContext {
  sync(): Promise<unknown>;
}

interface LoadableNullObject {
  isNullObject: boolean;
  load(propertyNames: string | string[]): void;
}

interface UsedRangeSnapshot extends LoadableNullObject {
  address: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: unknown[][];
}

interface UsedRangeSource {
  getUsedRangeOrNullObject(valuesOnly?: boolean): UsedRangeSnapshot;
}

export function isRecoverySheetVisibility(value: unknown): value is RecoverySheetVisibility {
  return value === "Visible" || value === "Hidden" || value === "VeryHidden";
}

export type CaptureModifyStructureStateArgs =
  | {
    kind: "sheet_name" | "sheet_visibility" | "sheet_absent";
    sheetRef: string;
  }
  | {
    kind: "rows_absent" | "columns_absent";
    sheetRef: string;
    position: number;
    count: number;
  };

export type StructureValueDataCaptureResult =
  | {
    status: "empty";
  }
  | {
    status: "captured";
    dataRange: RecoveryStructureValueRangeState;
  }
  | {
    status: "too_large";
    cellCount: number;
  };

function normalizePositiveInteger(value: number): number | null {
  if (!Number.isFinite(value)) {
    return null;
  }

  const normalized = Math.floor(value);
  if (normalized <= 0) {
    return null;
  }

  return normalized;
}

export function columnNumberToLetter(position: number): string {
  let col = position - 1;
  let letter = "";

  while (col >= 0) {
    letter = String.fromCharCode((col % 26) + 65) + letter;
    col = Math.floor(col / 26) - 1;
  }

  return letter;
}

function cloneUnknownGrid(grid: unknown[][]): unknown[][] {
  return grid.map((row) => {
    if (!Array.isArray(row)) {
      return [];
    }

    return [...row];
  });
}

function rowLength(grid: unknown[][], row: number): number {
  const rowValues = grid[row];
  return Array.isArray(rowValues) ? rowValues.length : 0;
}

function gridStats(values: unknown[][], formulas: unknown[][]): {
  rows: number;
  cols: number;
} {
  const rows = Math.max(values.length, formulas.length);
  let cols = 0;

  for (let row = 0; row < rows; row += 1) {
    cols = Math.max(cols, rowLength(values, row), rowLength(formulas, row));
  }

  return {
    rows,
    cols,
  };
}

function valueAt(grid: unknown[][], row: number, col: number): unknown {
  const rowValues = grid[row];
  if (!Array.isArray(rowValues)) {
    return "";
  }

  return col < rowValues.length ? rowValues[col] : "";
}

function normalizeFormula(raw: unknown): string | undefined {
  if (typeof raw !== "string") {
    return undefined;
  }

  const trimmed = raw.trim();
  if (!trimmed.startsWith("=")) {
    return undefined;
  }

  return trimmed;
}

function toRestoreValues(values: unknown[][], formulas: unknown[][]): unknown[][] {
  const { rows, cols } = gridStats(values, formulas);
  const restored: unknown[][] = [];

  for (let row = 0; row < rows; row += 1) {
    const outRow: unknown[] = [];

    for (let col = 0; col < cols; col += 1) {
      const formula = normalizeFormula(valueAt(formulas, row, col));
      outRow.push(formula ?? valueAt(values, row, col));
    }

    restored.push(outRow);
  }

  return restored;
}

function isStructureValueRangeStateShapeValid(dataRange: RecoveryStructureValueRangeState): boolean {
  if (typeof dataRange.address !== "string") {
    return false;
  }

  if (!Number.isInteger(dataRange.rowCount) || dataRange.rowCount <= 0) {
    return false;
  }

  if (!Number.isInteger(dataRange.columnCount) || dataRange.columnCount <= 0) {
    return false;
  }

  const stats = gridStats(dataRange.values, dataRange.formulas);
  return stats.rows === dataRange.rowCount && stats.cols === dataRange.columnCount;
}

export function estimateModifyStructureCellCount(state: RecoveryModifyStructureState): number {
  const dataRange = state.kind === "sheet_present" || state.kind === "rows_present" || state.kind === "columns_present"
    ? state.dataRange
    : undefined;

  if (!dataRange) {
    return 1;
  }

  const estimated = dataRange.rowCount * dataRange.columnCount;
  if (!Number.isFinite(estimated) || estimated <= 0) {
    return 1;
  }

  return estimated;
}

async function captureUsedRangeSnapshot(
  context: SyncContext,
  source: UsedRangeSource,
  maxCellCount: number,
): Promise<StructureValueDataCaptureResult> {
  const usedRange = source.getUsedRangeOrNullObject(true);
  usedRange.load(["isNullObject", "address", "rowCount", "columnCount"]);
  await context.sync();

  if (usedRange.isNullObject) {
    return { status: "empty" };
  }

  const cellCount = usedRange.rowCount * usedRange.columnCount;
  if (cellCount > maxCellCount) {
    return {
      status: "too_large",
      cellCount,
    };
  }

  usedRange.load(["values", "formulas"]);
  await context.sync();

  const values = cloneUnknownGrid(usedRange.values);
  const formulas = cloneUnknownGrid(usedRange.formulas);
  const dataRange: RecoveryStructureValueRangeState = {
    address: localAddressPart(usedRange.address),
    rowCount: usedRange.rowCount,
    columnCount: usedRange.columnCount,
    values,
    formulas,
  };

  if (!isStructureValueRangeStateShapeValid(dataRange)) {
    return { status: "empty" };
  }

  return {
    status: "captured",
    dataRange,
  };
}

export async function captureValueDataRange(
  context: SyncContext,
  targetRange: UsedRangeSource,
  maxCellCount = MAX_RECOVERY_CELLS,
): Promise<StructureValueDataCaptureResult> {
  return captureUsedRangeSnapshot(context, targetRange, maxCellCount);
}

export async function captureSheetValueDataRange(
  context: SyncContext,
  sheet: UsedRangeSource,
  maxCellCount = MAX_RECOVERY_CELLS,
): Promise<StructureValueDataCaptureResult> {
  return captureUsedRangeSnapshot(context, sheet, maxCellCount);
}

export async function hasValueDataInSheet(
  context: SyncContext,
  sheet: UsedRangeSource,
): Promise<boolean> {
  const usedRange = sheet.getUsedRangeOrNullObject(true);
  usedRange.load("isNullObject");
  await context.sync();
  return !usedRange.isNullObject;
}

export async function hasValueDataInRange(
  context: SyncContext,
  targetRange: UsedRangeSource,
): Promise<boolean> {
  const usedRange = targetRange.getUsedRangeOrNullObject(true);
  usedRange.load("isNullObject");
  await context.sync();

  return !usedRange.isNullObject;
}

async function loadSheetById(
  context: Excel.RequestContext,
  sheetId: string,
): Promise<Excel.Worksheet | null> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetId);
  sheet.load("isNullObject,id,name,visibility,position");
  await context.sync();

  if (sheet.isNullObject) {
    return null;
  }

  return sheet;
}

async function loadSheetByIdOrName(
  context: Excel.RequestContext,
  sheetId: string,
  sheetName: string,
): Promise<Excel.Worksheet | null> {
  const byId = await loadSheetById(context, sheetId);
  if (byId) {
    return byId;
  }

  const byName = context.workbook.worksheets.getItemOrNullObject(sheetName);
  byName.load("isNullObject,id,name,visibility,position");
  await context.sync();

  if (byName.isNullObject) {
    return null;
  }

  return byName;
}

async function restoreStructureValueRange(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  dataRange: RecoveryStructureValueRangeState,
): Promise<void> {
  if (!isStructureValueRangeStateShapeValid(dataRange)) {
    throw new Error("Structure checkpoint is invalid: captured data range is inconsistent.");
  }

  const range = sheet.getRange(dataRange.address);
  range.load(["rowCount", "columnCount"]);
  await context.sync();

  if (range.rowCount !== dataRange.rowCount || range.columnCount !== dataRange.columnCount) {
    throw new Error("Structure checkpoint is invalid: captured data range shape does not match target range.");
  }

  range.values = toRestoreValues(dataRange.values, dataRange.formulas);
  await context.sync();
}

export async function captureModifyStructureState(
  args: CaptureModifyStructureStateArgs,
): Promise<RecoveryModifyStructureState | null> {
  return excelRun<RecoveryModifyStructureState | null>(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(args.sheetRef);
    sheet.load("isNullObject,id,name,visibility");
    await context.sync();

    if (sheet.isNullObject) {
      return null;
    }

    if (args.kind === "sheet_name") {
      return {
        kind: "sheet_name",
        sheetId: sheet.id,
        name: sheet.name,
      };
    }

    if (args.kind === "sheet_visibility") {
      const visibility = sheet.visibility;
      if (!isRecoverySheetVisibility(visibility)) {
        return null;
      }

      return {
        kind: "sheet_visibility",
        sheetId: sheet.id,
        visibility,
      };
    }

    if (args.kind === "sheet_absent") {
      return {
        kind: "sheet_absent",
        sheetId: sheet.id,
        sheetName: sheet.name,
      };
    }

    if (args.kind !== "rows_absent" && args.kind !== "columns_absent") {
      return null;
    }

    const position = normalizePositiveInteger(args.position);
    const count = normalizePositiveInteger(args.count);
    if (position === null || count === null) {
      return null;
    }

    return {
      kind: args.kind,
      sheetId: sheet.id,
      sheetName: sheet.name,
      position,
      count,
    };
  });
}

export async function applyModifyStructureState(
  targetState: RecoveryModifyStructureState,
): Promise<RecoveryModifyStructureState> {
  return excelRun<RecoveryModifyStructureState>(async (context) => {
    if (targetState.kind === "sheet_name") {
      const sheet = await loadSheetById(context, targetState.sheetId);
      if (!sheet) {
        throw new Error("Sheet referenced by structure checkpoint no longer exists.");
      }

      const currentState: RecoveryModifyStructureState = {
        kind: "sheet_name",
        sheetId: sheet.id,
        name: sheet.name,
      };

      sheet.name = targetState.name;
      await context.sync();
      return currentState;
    }

    if (targetState.kind === "sheet_visibility") {
      const sheet = await loadSheetById(context, targetState.sheetId);
      if (!sheet) {
        throw new Error("Sheet referenced by structure checkpoint no longer exists.");
      }

      const currentVisibility = sheet.visibility;
      if (!isRecoverySheetVisibility(currentVisibility)) {
        throw new Error("Sheet visibility is unsupported for structure checkpoint restore.");
      }

      const currentState: RecoveryModifyStructureState = {
        kind: "sheet_visibility",
        sheetId: sheet.id,
        visibility: currentVisibility,
      };

      sheet.visibility = targetState.visibility;
      await context.sync();
      return currentState;
    }

    if (targetState.kind === "sheet_absent") {
      const sheet = targetState.allowDataDelete
        ? await loadSheetById(context, targetState.sheetId)
        : await loadSheetByIdOrName(context, targetState.sheetId, targetState.sheetName);
      if (!sheet) {
        return cloneRecoveryModifyStructureState(targetState);
      }

      const currentVisibility = sheet.visibility;
      if (!isRecoverySheetVisibility(currentVisibility)) {
        throw new Error("Sheet visibility is unsupported for structure checkpoint restore.");
      }

      const hasValueData = await hasValueDataInSheet(context, sheet);
      if (hasValueData && !targetState.allowDataDelete) {
        throw new Error(
          "Structure checkpoint restore is blocked: target sheet contains data and cannot be deleted safely.",
        );
      }

      let currentDataRange: RecoveryStructureValueRangeState | undefined;
      if (hasValueData) {
        const currentDataCapture = await captureSheetValueDataRange(context, sheet);
        if (currentDataCapture.status === "too_large") {
          throw new Error("Structure checkpoint restore failed: target sheet data exceeds recovery size limits.");
        }

        if (currentDataCapture.status !== "captured") {
          throw new Error("Structure checkpoint restore failed: could not capture current sheet data before delete.");
        }

        currentDataRange = currentDataCapture.dataRange;
      }

      const currentState: RecoveryModifyStructureState = {
        kind: "sheet_present",
        sheetId: sheet.id,
        sheetName: sheet.name,
        position: sheet.position,
        visibility: currentVisibility,
        ...(currentDataRange ? { dataRange: currentDataRange } : {}),
      };

      sheet.delete();
      await context.sync();
      return currentState;
    }

    if (targetState.kind === "sheet_present") {
      const existing = await loadSheetByIdOrName(context, targetState.sheetId, targetState.sheetName);

      if (!existing) {
        const created = context.workbook.worksheets.add(targetState.sheetName);
        created.position = targetState.position;
        created.visibility = targetState.visibility;
        created.load(["id", "name"]);
        await context.sync();

        if (targetState.dataRange) {
          await restoreStructureValueRange(context, created, targetState.dataRange);
        }

        const currentState: RecoveryModifyStructureState = {
          kind: "sheet_absent",
          sheetId: created.id,
          sheetName: created.name,
          ...(targetState.dataRange ? { allowDataDelete: true } : {}),
        };

        return currentState;
      }

      if (targetState.dataRange) {
        throw new Error(
          "Structure checkpoint restore is blocked: target sheet already exists and cannot be overwritten safely.",
        );
      }

      const currentVisibility = existing.visibility;
      if (!isRecoverySheetVisibility(currentVisibility)) {
        throw new Error("Sheet visibility is unsupported for structure checkpoint restore.");
      }

      const currentState: RecoveryModifyStructureState = {
        kind: "sheet_present",
        sheetId: existing.id,
        sheetName: existing.name,
        position: existing.position,
        visibility: currentVisibility,
      };

      existing.name = targetState.sheetName;
      existing.position = targetState.position;
      existing.visibility = targetState.visibility;
      await context.sync();
      return currentState;
    }

    if (targetState.kind === "rows_absent" || targetState.kind === "rows_present") {
      const position = normalizePositiveInteger(targetState.position);
      const count = normalizePositiveInteger(targetState.count);
      if (position === null || count === null) {
        throw new Error("Structure checkpoint is invalid: row position/count is invalid.");
      }

      const sheet = await loadSheetById(context, targetState.sheetId);
      if (!sheet) {
        throw new Error("Sheet referenced by row checkpoint no longer exists.");
      }

      const endRow = position + count - 1;
      const range = sheet.getRange(`${position}:${endRow}`);

      if (targetState.kind === "rows_absent") {
        const hasValueData = await hasValueDataInRange(context, range);
        if (hasValueData && !targetState.allowDataDelete) {
          throw new Error(
            "Structure checkpoint restore is blocked: target rows contain data and cannot be deleted safely.",
          );
        }

        let currentDataRange: RecoveryStructureValueRangeState | undefined;
        if (hasValueData) {
          const currentDataCapture = await captureValueDataRange(context, range);
          if (currentDataCapture.status === "too_large") {
            throw new Error("Structure checkpoint restore failed: target row data exceeds recovery size limits.");
          }

          if (currentDataCapture.status !== "captured") {
            throw new Error("Structure checkpoint restore failed: could not capture current row data before delete.");
          }

          currentDataRange = currentDataCapture.dataRange;
        }

        const currentState: RecoveryModifyStructureState = {
          kind: "rows_present",
          sheetId: sheet.id,
          sheetName: sheet.name,
          position,
          count,
          ...(currentDataRange ? { dataRange: currentDataRange } : {}),
        };

        range.delete("Up");
        await context.sync();
        return currentState;
      }

      const currentState: RecoveryModifyStructureState = {
        kind: "rows_absent",
        sheetId: sheet.id,
        sheetName: sheet.name,
        position,
        count,
        ...(targetState.dataRange ? { allowDataDelete: true } : {}),
      };

      range.insert("Down");
      await context.sync();

      if (targetState.dataRange) {
        await restoreStructureValueRange(context, sheet, targetState.dataRange);
      }

      return currentState;
    }

    const position = normalizePositiveInteger(targetState.position);
    const count = normalizePositiveInteger(targetState.count);
    if (position === null || count === null) {
      throw new Error("Structure checkpoint is invalid: column position/count is invalid.");
    }

    const sheet = await loadSheetById(context, targetState.sheetId);
    if (!sheet) {
      throw new Error("Sheet referenced by column checkpoint no longer exists.");
    }

    const startLetter = columnNumberToLetter(position);
    const endLetter = columnNumberToLetter(position + count - 1);

    if (targetState.kind === "columns_absent") {
      const range = sheet.getRange(`${startLetter}:${endLetter}`);
      const hasValueData = await hasValueDataInRange(context, range);

      if (hasValueData && !targetState.allowDataDelete) {
        throw new Error(
          "Structure checkpoint restore is blocked: target columns contain data and cannot be deleted safely.",
        );
      }

      let currentDataRange: RecoveryStructureValueRangeState | undefined;
      if (hasValueData) {
        const currentDataCapture = await captureValueDataRange(context, range);
        if (currentDataCapture.status === "too_large") {
          throw new Error("Structure checkpoint restore failed: target column data exceeds recovery size limits.");
        }

        if (currentDataCapture.status !== "captured") {
          throw new Error("Structure checkpoint restore failed: could not capture current column data before delete.");
        }

        currentDataRange = currentDataCapture.dataRange;
      }

      const currentState: RecoveryModifyStructureState = {
        kind: "columns_present",
        sheetId: sheet.id,
        sheetName: sheet.name,
        position,
        count,
        ...(currentDataRange ? { dataRange: currentDataRange } : {}),
      };

      range.delete("Left");
      await context.sync();
      return currentState;
    }

    const currentState: RecoveryModifyStructureState = {
      kind: "columns_absent",
      sheetId: sheet.id,
      sheetName: sheet.name,
      position,
      count,
      ...(targetState.dataRange ? { allowDataDelete: true } : {}),
    };

    const range = sheet.getRange(`${startLetter}:${startLetter}`);
    for (let index = 0; index < count; index += 1) {
      range.insert("Right");
    }

    await context.sync();

    if (targetState.dataRange) {
      await restoreStructureValueRange(context, sheet, targetState.dataRange);
    }

    return currentState;
  });
}
