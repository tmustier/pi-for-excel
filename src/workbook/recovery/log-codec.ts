/**
 * Codec helpers for persisted workbook recovery snapshots.
 *
 * A persisted snapshot is a schema keyed by `snapshotKind`; each kind requires
 * its own state and nothing else. An entry that fails its schema is dropped
 * rather than repaired: the writer is this module, so a mismatch means version
 * skew (a kind this build does not know) or corruption, and neither is a
 * checkpoint we can restore.
 */

import { type Static, Type } from "typebox";
import { Value } from "typebox/value";

import { StringEnum } from "../../tools/string-enum.js";
import { cloneRecoveryConditionalFormatRules, cloneRecoveryState } from "./clone.js";
import { cloneGrid, gridStats } from "./grid.js";
import {
  RecoveryChartStateSchema,
  RecoveryCommentThreadStateSchema,
  RecoveryConditionalFormatRuleSchema,
  RecoveryFormatRangeStateSchema,
  RecoveryGridSchema,
  RecoveryModifyStructureStateSchema,
} from "./schemas.js";
import { estimateModifyStructureCellCount } from "./structure-state.js";
import type { WorkbookRecoverySnapshot } from "../recovery-log.js";

export interface ParsePersistedSnapshotsOptions {
  maxEntries: number;
}

const WorkbookRecoveryToolNameSchema = StringEnum([
  "write_cells",
  "fill_formula",
  "python_transform_range",
  "format_cells",
  "conditional_format",
  "comments",
  "charts",
  "modify_structure",
  "restore_snapshot",
]);

const persistedSnapshotBase = {
  toolName: WorkbookRecoveryToolNameSchema,
  toolCallId: Type.String(),
  address: Type.String(),
  id: Type.Optional(Type.String()),
  at: Type.Optional(Type.Number()),
  cellCount: Type.Optional(Type.Number({ minimum: 0 })),
  changedCount: Type.Optional(Type.Number({ minimum: 0 })),
  workbookId: Type.Optional(Type.String()),
  workbookLabel: Type.Optional(Type.String()),
  restoredFromSnapshotId: Type.Optional(Type.String()),
};

/**
 * Grids are optional on every kind but `range_values`. Chart-state snapshots
 * are stored without them so that codecs predating `chart_state` fail their
 * range grid check and drop the entry, instead of misreading it as an empty
 * range backup after a version downgrade.
 */
const optionalGrids = {
  beforeValues: Type.Optional(RecoveryGridSchema),
  beforeFormulas: Type.Optional(RecoveryGridSchema),
};

/** `snapshotKind` was added after the first range snapshots shipped; absent means `range_values`. */
export const PersistedWorkbookRecoverySnapshotSchema = Type.Union([
  Type.Object({
    ...persistedSnapshotBase,
    snapshotKind: Type.Optional(Type.Literal("range_values")),
    beforeValues: RecoveryGridSchema,
    beforeFormulas: RecoveryGridSchema,
  }),
  Type.Object({
    ...persistedSnapshotBase,
    ...optionalGrids,
    snapshotKind: Type.Literal("format_cells_state"),
    formatRangeState: RecoveryFormatRangeStateSchema,
  }),
  Type.Object({
    ...persistedSnapshotBase,
    ...optionalGrids,
    snapshotKind: Type.Literal("modify_structure_state"),
    modifyStructureState: RecoveryModifyStructureStateSchema,
  }),
  Type.Object({
    ...persistedSnapshotBase,
    ...optionalGrids,
    snapshotKind: Type.Literal("conditional_format_rules"),
    conditionalFormatRules: Type.Array(RecoveryConditionalFormatRuleSchema),
  }),
  Type.Object({
    ...persistedSnapshotBase,
    ...optionalGrids,
    snapshotKind: Type.Literal("comment_thread"),
    commentThreadState: RecoveryCommentThreadStateSchema,
  }),
  Type.Object({
    ...persistedSnapshotBase,
    ...optionalGrids,
    snapshotKind: Type.Literal("chart_state"),
    chartState: RecoveryChartStateSchema,
  }),
]);

export type PersistedWorkbookRecoverySnapshot = Static<typeof PersistedWorkbookRecoverySnapshotSchema>;

export const PersistedWorkbookRecoveryPayloadSchema = Type.Object({
  version: Type.Literal(1),
  snapshots: Type.Array(Type.Unknown()),
});

export interface PersistedWorkbookRecoveryPayload {
  version: 1;
  snapshots: PersistedWorkbookRecoverySnapshot[];
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

/**
 * Builds the in-memory snapshot from a checked persisted entry. Grids take the
 * plain row-copy path; the kind states go through the schema clone, which also
 * drops properties an older build may have written. Neither the persisted
 * entry nor anything it references is kept.
 */
function toSnapshot(persisted: PersistedWorkbookRecoverySnapshot): WorkbookRecoverySnapshot {
  const beforeValues = cloneGrid(persisted.beforeValues ?? []);
  const beforeFormulas = cloneGrid(persisted.beforeFormulas ?? []);

  const snapshot: WorkbookRecoverySnapshot = {
    id: persisted.id ?? defaultCreateId(),
    at: persisted.at ?? Date.now(),
    toolName: persisted.toolName,
    toolCallId: persisted.toolCallId,
    address: persisted.address,
    changedCount: 0,
    cellCount: 0,
    beforeValues,
    beforeFormulas,
    snapshotKind: persisted.snapshotKind ?? "range_values",
  };

  if (persisted.workbookId !== undefined) snapshot.workbookId = persisted.workbookId;
  if (persisted.workbookLabel !== undefined) snapshot.workbookLabel = persisted.workbookLabel;
  if (persisted.restoredFromSnapshotId !== undefined) {
    snapshot.restoredFromSnapshotId = persisted.restoredFromSnapshotId;
  }

  let fallbackCellCount: number;
  switch (persisted.snapshotKind) {
    case "format_cells_state":
      snapshot.formatRangeState = cloneRecoveryState(RecoveryFormatRangeStateSchema, persisted.formatRangeState);
      fallbackCellCount = persisted.formatRangeState.cellCount;
      break;
    case "modify_structure_state":
      snapshot.modifyStructureState = cloneRecoveryState(RecoveryModifyStructureStateSchema, persisted.modifyStructureState);
      fallbackCellCount = estimateModifyStructureCellCount(persisted.modifyStructureState);
      break;
    case "conditional_format_rules":
      snapshot.conditionalFormatRules = cloneRecoveryConditionalFormatRules(persisted.conditionalFormatRules);
      fallbackCellCount = persisted.conditionalFormatRules.length;
      break;
    case "comment_thread":
      snapshot.commentThreadState = cloneRecoveryState(RecoveryCommentThreadStateSchema, persisted.commentThreadState);
      fallbackCellCount = 1;
      break;
    case "chart_state":
      snapshot.chartState = cloneRecoveryState(RecoveryChartStateSchema, persisted.chartState);
      fallbackCellCount = 1;
      break;
    default:
      // "range_values", or a legacy entry written before snapshotKind existed.
      fallbackCellCount = gridStats(beforeValues, beforeFormulas).cellCount;
  }

  snapshot.cellCount = persisted.cellCount ?? fallbackCellCount;
  snapshot.changedCount = persisted.changedCount ?? snapshot.cellCount;

  return snapshot;
}

export function parsePersistedSnapshots(
  payload: unknown,
  options: ParsePersistedSnapshotsOptions,
): WorkbookRecoverySnapshot[] {
  if (!Value.Check(PersistedWorkbookRecoveryPayloadSchema, payload)) return [];

  const snapshots: WorkbookRecoverySnapshot[] = [];
  for (const item of payload.snapshots) {
    if (Value.Check(PersistedWorkbookRecoverySnapshotSchema, item)) {
      snapshots.push(toSnapshot(item));
    }
  }

  const maxEntries = Number.isFinite(options.maxEntries)
    ? Math.max(0, Math.floor(options.maxEntries))
    : snapshots.length;

  return snapshots
    .sort((a, b) => b.at - a.at)
    .slice(0, maxEntries);
}

function toPersistedSnapshot(snapshot: WorkbookRecoverySnapshot): PersistedWorkbookRecoverySnapshot {
  const { snapshotKind: kind, beforeValues, beforeFormulas, ...rest } = snapshot;
  const snapshotKind = kind ?? "range_values";

  switch (snapshotKind) {
    case "range_values":
      return { ...rest, snapshotKind, beforeValues, beforeFormulas };
    case "format_cells_state":
      return { ...rest, snapshotKind, beforeValues, beforeFormulas, formatRangeState: requireState(snapshot.formatRangeState, snapshotKind) };
    case "modify_structure_state":
      return {
        ...rest,
        snapshotKind,
        beforeValues,
        beforeFormulas,
        modifyStructureState: requireState(snapshot.modifyStructureState, snapshotKind),
      };
    case "conditional_format_rules":
      return {
        ...rest,
        snapshotKind,
        beforeValues,
        beforeFormulas,
        conditionalFormatRules: requireState(snapshot.conditionalFormatRules, snapshotKind),
      };
    case "comment_thread":
      return { ...rest, snapshotKind, beforeValues, beforeFormulas, commentThreadState: requireState(snapshot.commentThreadState, snapshotKind) };
    case "chart_state":
      return { ...rest, snapshotKind, chartState: requireState(snapshot.chartState, snapshotKind) };
  }
}

function requireState<T>(state: T | undefined, snapshotKind: string): T {
  if (state === undefined) {
    throw new Error(`Recovery snapshot of kind "${snapshotKind}" has no state to persist.`);
  }
  return state;
}

export function createPersistedWorkbookRecoveryPayload(
  snapshots: WorkbookRecoverySnapshot[],
): PersistedWorkbookRecoveryPayload {
  return {
    version: 1,
    snapshots: snapshots.map(toPersistedSnapshot),
  };
}
