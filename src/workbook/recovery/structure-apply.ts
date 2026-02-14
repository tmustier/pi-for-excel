/** Structure-state restore/apply helpers for workbook recovery snapshots. */

import { excelRun } from "../../excel/helpers.js";
import { cloneRecoveryModifyStructureState } from "./clone.js";
import { toRestoreValues } from "./grid.js";
import {
  columnNumberToLetter,
  isRecoverySheetVisibility,
  isStructureValueRangeStateShapeValid,
  normalizePositiveInteger,
} from "./structure-common.js";
import {
  captureSheetValueDataRange,
  captureValueDataRange,
  hasValueDataInRange,
  hasValueDataInSheet,
} from "./structure-capture.js";
import type {
  RecoveryModifyStructureState,
  RecoveryStructureValueRangeState,
} from "./types.js";

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
