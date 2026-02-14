/** Structure-state capture helpers for workbook recovery snapshots. */

import { excelRun } from "../../excel/helpers.js";
import { localAddressPart } from "./address.js";
import { MAX_RECOVERY_CELLS } from "./constants.js";
import { cloneGrid } from "./grid.js";
import {
  columnNumberToLetter,
  estimateModifyStructureCellCount,
  isRecoverySheetVisibility,
  isStructureValueRangeStateShapeValid,
  normalizePositiveInteger,
  type CaptureModifyStructureStateArgs,
  type StructureValueDataCaptureResult,
  type SyncContext,
  type UsedRangeSource,
} from "./structure-common.js";
import type { RecoveryModifyStructureState, RecoveryStructureValueRangeState } from "./types.js";

export {
  columnNumberToLetter,
  estimateModifyStructureCellCount,
  isRecoverySheetVisibility,
};
export type {
  CaptureModifyStructureStateArgs,
  StructureValueDataCaptureResult,
};

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

  const values = cloneGrid(usedRange.values);
  const formulas = cloneGrid(usedRange.formulas);
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
