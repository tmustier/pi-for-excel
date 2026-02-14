/** Shared helpers and types for structure-state capture/apply. */

import { gridStats } from "./grid.js";
import type {
  RecoveryModifyStructureState,
  RecoverySheetVisibility,
  RecoveryStructureValueRangeState,
} from "./types.js";

export interface SyncContext {
  sync(): Promise<unknown>;
}

export interface LoadableNullObject {
  isNullObject: boolean;
  load(propertyNames: string | string[]): void;
}

export interface UsedRangeSnapshot extends LoadableNullObject {
  address: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: unknown[][];
}

export interface UsedRangeSource {
  getUsedRangeOrNullObject(valuesOnly?: boolean): UsedRangeSnapshot;
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

export function isRecoverySheetVisibility(value: unknown): value is RecoverySheetVisibility {
  return value === "Visible" || value === "Hidden" || value === "VeryHidden";
}

export function normalizePositiveInteger(value: number): number | null {
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

export function isStructureValueRangeStateShapeValid(dataRange: RecoveryStructureValueRangeState): boolean {
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
