/**
 * Codec helpers for persisted workbook recovery snapshots.
 */

import { isRecord } from "../../utils/type-guards.js";
import {
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  cloneRecoveryFormatRangeState,
  cloneRecoveryModifyStructureState,
  isRecoveryConditionalFormatRule,
  type RecoveryCommentThreadState,
  type RecoveryConditionalFormatRule,
  type RecoveryFormatBorderState,
  type RecoveryFormatRangeState,
  type RecoveryModifyStructureState,
  type RecoveryStructureValueRangeState,
} from "../recovery-states.js";
import { cloneGrid, gridStats } from "./grid.js";
import { estimateModifyStructureCellCount } from "./structure-state.js";
import type {
  WorkbookRecoverySnapshot,
  WorkbookRecoverySnapshotKind,
  WorkbookRecoveryToolName,
} from "../recovery-log.js";

export interface ParsePersistedSnapshotsOptions {
  maxEntries: number;
}

export interface PersistedWorkbookRecoveryPayload {
  version: 1;
  snapshots: WorkbookRecoverySnapshot[];
}

function defaultCreateId(): string {
  const randomUuid = globalThis.crypto?.randomUUID;
  if (typeof randomUuid === "function") {
    return randomUuid.call(globalThis.crypto);
  }

  const randomChunk = Math.floor(Math.random() * 1_000_000)
    .toString(36)
    .padStart(4, "0");

  return `checkpoint_${Date.now().toString(36)}_${randomChunk}`;
}

function isWorkbookRecoveryToolName(value: unknown): value is WorkbookRecoveryToolName {
  return (
    value === "write_cells" ||
    value === "fill_formula" ||
    value === "python_transform_range" ||
    value === "format_cells" ||
    value === "conditional_format" ||
    value === "comments" ||
    value === "modify_structure" ||
    value === "restore_snapshot"
  );
}

function isGrid(value: unknown): value is unknown[][] {
  return Array.isArray(value) && value.every((row) => Array.isArray(row));
}

function parseWorkbookRecoverySnapshotKind(value: unknown): WorkbookRecoverySnapshotKind {
  return value === "conditional_format_rules" ||
      value === "comment_thread" ||
      value === "format_cells_state" ||
      value === "modify_structure_state" ||
      value === "range_values"
    ? value
    : "range_values";
}

function isRecoveryFormatSelection(value: unknown): value is RecoveryFormatRangeState["selection"] {
  if (!isRecord(value)) return false;

  const keys: Array<keyof RecoveryFormatRangeState["selection"]> = [
    "numberFormat",
    "fillColor",
    "fontColor",
    "bold",
    "italic",
    "underlineStyle",
    "fontName",
    "fontSize",
    "horizontalAlignment",
    "verticalAlignment",
    "wrapText",
    "columnWidth",
    "rowHeight",
    "mergedAreas",
    "borderTop",
    "borderBottom",
    "borderLeft",
    "borderRight",
    "borderInsideHorizontal",
    "borderInsideVertical",
  ];

  for (const key of keys) {
    const candidate = value[key];
    if (candidate !== undefined && typeof candidate !== "boolean") {
      return false;
    }
  }

  return true;
}

function isStringGrid(value: unknown): value is string[][] {
  return Array.isArray(value) &&
    value.every((row) => Array.isArray(row) && row.every((cell) => typeof cell === "string"));
}

function isNumberList(value: unknown): value is number[] {
  return Array.isArray(value) && value.every((item) => typeof item === "number" && Number.isFinite(item));
}

function isStringList(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === "string");
}

function isRecoveryFormatBorderState(value: unknown): value is RecoveryFormatBorderState {
  if (!isRecord(value)) return false;

  return (
    typeof value.style === "string" &&
    (value.weight === undefined || typeof value.weight === "string") &&
    (value.color === undefined || typeof value.color === "string")
  );
}

function isRecoveryFormatAreaState(value: unknown): value is RecoveryFormatRangeState["areas"][number] {
  if (!isRecord(value)) return false;
  if (typeof value.address !== "string") return false;
  if (typeof value.rowCount !== "number") return false;
  if (typeof value.columnCount !== "number") return false;

  if (value.numberFormat !== undefined && !isStringGrid(value.numberFormat)) return false;
  if (value.fillColor !== undefined && typeof value.fillColor !== "string") return false;
  if (value.fontColor !== undefined && typeof value.fontColor !== "string") return false;
  if (value.bold !== undefined && typeof value.bold !== "boolean") return false;
  if (value.italic !== undefined && typeof value.italic !== "boolean") return false;
  if (value.underlineStyle !== undefined && typeof value.underlineStyle !== "string") return false;
  if (value.fontName !== undefined && typeof value.fontName !== "string") return false;
  if (value.fontSize !== undefined && typeof value.fontSize !== "number") return false;
  if (value.horizontalAlignment !== undefined && typeof value.horizontalAlignment !== "string") return false;
  if (value.verticalAlignment !== undefined && typeof value.verticalAlignment !== "string") return false;
  if (value.wrapText !== undefined && typeof value.wrapText !== "boolean") return false;
  if (value.columnWidths !== undefined && !isNumberList(value.columnWidths)) return false;
  if (value.rowHeights !== undefined && !isNumberList(value.rowHeights)) return false;
  if (value.mergedAreas !== undefined && !isStringList(value.mergedAreas)) return false;

  if (value.borderTop !== undefined && !isRecoveryFormatBorderState(value.borderTop)) return false;
  if (value.borderBottom !== undefined && !isRecoveryFormatBorderState(value.borderBottom)) return false;
  if (value.borderLeft !== undefined && !isRecoveryFormatBorderState(value.borderLeft)) return false;
  if (value.borderRight !== undefined && !isRecoveryFormatBorderState(value.borderRight)) return false;
  if (value.borderInsideHorizontal !== undefined && !isRecoveryFormatBorderState(value.borderInsideHorizontal)) {
    return false;
  }
  if (value.borderInsideVertical !== undefined && !isRecoveryFormatBorderState(value.borderInsideVertical)) {
    return false;
  }

  if (Array.isArray(value.columnWidths) && value.columnWidths.length !== value.columnCount) {
    return false;
  }

  if (Array.isArray(value.rowHeights) && value.rowHeights.length !== value.rowCount) {
    return false;
  }

  return true;
}

function isRecoveryFormatRangeState(value: unknown): value is RecoveryFormatRangeState {
  if (!isRecord(value)) return false;
  if (!isRecoveryFormatSelection(value.selection)) return false;
  if (!Array.isArray(value.areas) || !value.areas.every((area) => isRecoveryFormatAreaState(area))) return false;
  if (typeof value.cellCount !== "number") return false;

  return true;
}

function isRecoverySheetVisibility(value: unknown): value is "Visible" | "Hidden" | "VeryHidden" {
  return value === "Visible" || value === "Hidden" || value === "VeryHidden";
}

function isPositiveInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value > 0;
}

function isRecoveryStructureValueRangeState(value: unknown): value is RecoveryStructureValueRangeState {
  if (!isRecord(value)) {
    return false;
  }

  if (typeof value.address !== "string") {
    return false;
  }

  if (!isPositiveInteger(value.rowCount) || !isPositiveInteger(value.columnCount)) {
    return false;
  }

  if (!isGrid(value.values) || !isGrid(value.formulas)) {
    return false;
  }

  const stats = gridStats(value.values, value.formulas);
  return stats.rows === value.rowCount && stats.cols === value.columnCount;
}

function isRecoveryModifyStructureState(value: unknown): value is RecoveryModifyStructureState {
  if (!isRecord(value)) return false;

  if (value.kind === "sheet_name") {
    return typeof value.sheetId === "string" && typeof value.name === "string";
  }

  if (value.kind === "sheet_visibility") {
    return typeof value.sheetId === "string" && isRecoverySheetVisibility(value.visibility);
  }

  if (value.kind === "sheet_absent") {
    return (
      typeof value.sheetId === "string" &&
      typeof value.sheetName === "string" &&
      (value.allowDataDelete === undefined || typeof value.allowDataDelete === "boolean")
    );
  }

  if (value.kind === "sheet_present") {
    return (
      typeof value.sheetId === "string" &&
      typeof value.sheetName === "string" &&
      typeof value.position === "number" &&
      Number.isInteger(value.position) &&
      value.position >= 0 &&
      isRecoverySheetVisibility(value.visibility) &&
      (value.dataRange === undefined || isRecoveryStructureValueRangeState(value.dataRange))
    );
  }

  if (value.kind === "rows_absent") {
    return (
      typeof value.sheetId === "string" &&
      typeof value.sheetName === "string" &&
      isPositiveInteger(value.position) &&
      isPositiveInteger(value.count) &&
      (value.allowDataDelete === undefined || typeof value.allowDataDelete === "boolean")
    );
  }

  if (value.kind === "rows_present") {
    return (
      typeof value.sheetId === "string" &&
      typeof value.sheetName === "string" &&
      isPositiveInteger(value.position) &&
      isPositiveInteger(value.count) &&
      (value.dataRange === undefined || isRecoveryStructureValueRangeState(value.dataRange))
    );
  }

  if (value.kind === "columns_absent") {
    return (
      typeof value.sheetId === "string" &&
      typeof value.sheetName === "string" &&
      isPositiveInteger(value.position) &&
      isPositiveInteger(value.count) &&
      (value.allowDataDelete === undefined || typeof value.allowDataDelete === "boolean")
    );
  }

  if (value.kind === "columns_present") {
    return (
      typeof value.sheetId === "string" &&
      typeof value.sheetName === "string" &&
      isPositiveInteger(value.position) &&
      isPositiveInteger(value.count) &&
      (value.dataRange === undefined || isRecoveryStructureValueRangeState(value.dataRange))
    );
  }

  return false;
}

function isRecoveryCommentThreadState(value: unknown): value is RecoveryCommentThreadState {
  if (!isRecord(value)) return false;
  if (typeof value.exists !== "boolean") return false;
  if (typeof value.content !== "string") return false;
  if (typeof value.resolved !== "boolean") return false;
  if (!Array.isArray(value.replies)) return false;

  return value.replies.every((reply) => typeof reply === "string");
}

function parseWorkbookRecoverySnapshot(value: unknown): WorkbookRecoverySnapshot | null {
  if (!isRecord(value)) return null;

  if (!isWorkbookRecoveryToolName(value.toolName)) return null;
  if (typeof value.toolCallId !== "string") return null;
  if (typeof value.address !== "string") return null;

  const snapshotKind = parseWorkbookRecoverySnapshotKind(value.snapshotKind);

  const beforeValues = isGrid(value.beforeValues)
    ? cloneGrid(value.beforeValues)
    : [];
  const beforeFormulas = isGrid(value.beforeFormulas)
    ? cloneGrid(value.beforeFormulas)
    : [];

  if (snapshotKind === "range_values" && (!isGrid(value.beforeValues) || !isGrid(value.beforeFormulas))) {
    return null;
  }

  let conditionalFormatRules: RecoveryConditionalFormatRule[] = [];
  if (Array.isArray(value.conditionalFormatRules)) {
    conditionalFormatRules = [];
    for (const rule of value.conditionalFormatRules) {
      if (!isRecoveryConditionalFormatRule(rule)) {
        return null;
      }

      conditionalFormatRules.push(rule);
    }
  }

  const formatRangeState = isRecoveryFormatRangeState(value.formatRangeState)
    ? cloneRecoveryFormatRangeState(value.formatRangeState)
    : undefined;

  const modifyStructureState = isRecoveryModifyStructureState(value.modifyStructureState)
    ? cloneRecoveryModifyStructureState(value.modifyStructureState)
    : undefined;

  const commentThreadState = isRecoveryCommentThreadState(value.commentThreadState)
    ? cloneRecoveryCommentThreadState(value.commentThreadState)
    : undefined;

  if (snapshotKind === "format_cells_state" && !formatRangeState) {
    return null;
  }

  if (snapshotKind === "modify_structure_state" && !modifyStructureState) {
    return null;
  }

  if (snapshotKind === "conditional_format_rules" && !Array.isArray(value.conditionalFormatRules)) {
    return null;
  }

  if (snapshotKind === "comment_thread" && !commentThreadState) {
    return null;
  }

  const id = typeof value.id === "string" ? value.id : defaultCreateId();
  const at = typeof value.at === "number" ? value.at : Date.now();

  const cellCountFromGrid = gridStats(beforeValues, beforeFormulas).cellCount;
  const fallbackCellCount = snapshotKind === "range_values"
    ? cellCountFromGrid
    : snapshotKind === "format_cells_state"
      ? (formatRangeState?.cellCount ?? 0)
      : snapshotKind === "modify_structure_state"
        ? (modifyStructureState ? estimateModifyStructureCellCount(modifyStructureState) : 1)
        : snapshotKind === "conditional_format_rules"
          ? conditionalFormatRules.length
          : 1;

  const cellCount = typeof value.cellCount === "number"
    ? Math.max(0, value.cellCount)
    : fallbackCellCount;

  const changedCount = typeof value.changedCount === "number"
    ? Math.max(0, value.changedCount)
    : cellCount;

  const snapshot: WorkbookRecoverySnapshot = {
    id,
    at,
    toolName: value.toolName,
    toolCallId: value.toolCallId,
    address: value.address,
    changedCount,
    cellCount,
    beforeValues,
    beforeFormulas,
    snapshotKind,
    workbookId: typeof value.workbookId === "string" ? value.workbookId : undefined,
    workbookLabel: typeof value.workbookLabel === "string" ? value.workbookLabel : undefined,
    restoredFromSnapshotId: typeof value.restoredFromSnapshotId === "string" ? value.restoredFromSnapshotId : undefined,
  };

  if (snapshotKind === "format_cells_state" && formatRangeState) {
    snapshot.formatRangeState = cloneRecoveryFormatRangeState(formatRangeState);
  }

  if (snapshotKind === "modify_structure_state" && modifyStructureState) {
    snapshot.modifyStructureState = cloneRecoveryModifyStructureState(modifyStructureState);
  }

  if (snapshotKind === "conditional_format_rules") {
    snapshot.conditionalFormatRules = cloneRecoveryConditionalFormatRules(conditionalFormatRules);
  }

  if (snapshotKind === "comment_thread") {
    snapshot.commentThreadState = commentThreadState;
  }

  return snapshot;
}

export function parsePersistedSnapshots(
  payload: unknown,
  options: ParsePersistedSnapshotsOptions,
): WorkbookRecoverySnapshot[] {
  if (!isRecord(payload)) return [];

  const snapshotsRaw = payload.snapshots;
  if (!Array.isArray(snapshotsRaw)) return [];

  const snapshots: WorkbookRecoverySnapshot[] = [];
  for (const item of snapshotsRaw) {
    const parsed = parseWorkbookRecoverySnapshot(item);
    if (parsed) {
      snapshots.push(parsed);
    }
  }

  const maxEntries = Number.isFinite(options.maxEntries)
    ? Math.max(0, Math.floor(options.maxEntries))
    : snapshots.length;

  return snapshots
    .sort((a, b) => b.at - a.at)
    .slice(0, maxEntries);
}

export function createPersistedWorkbookRecoveryPayload(
  snapshots: WorkbookRecoverySnapshot[],
): PersistedWorkbookRecoveryPayload {
  return {
    version: 1,
    snapshots,
  };
}
