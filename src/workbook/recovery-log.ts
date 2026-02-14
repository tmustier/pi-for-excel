/**
 * Workbook recovery snapshots.
 *
 * Stores lightweight, local-only backups for workbook cell edits so users
 * can revert Pi's changes without pre-execution approval prompts.
 */

import { excelRun, getRange } from "../excel/helpers.js";
import { formatWorkbookLabel, getWorkbookContext, type WorkbookContext } from "./context.js";
import {
  createPersistedWorkbookRecoveryPayload,
  parsePersistedSnapshots,
} from "./recovery/log-codec.js";
import {
  restoreWorkbookRecoverySnapshot,
} from "./recovery/log-restore.js";
import {
  defaultGetSettingsStore,
  readPersistedWorkbookRecoveryPayload,
  type SettingsStoreLike,
  writePersistedWorkbookRecoveryPayload,
} from "./recovery/log-store.js";
import {
  applyCommentThreadState,
  applyConditionalFormatState,
  applyFormatCellsState,
  applyModifyStructureState,
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  cloneRecoveryFormatRangeState,
  cloneRecoveryModifyStructureState,
  type RecoveryCommentThreadState,
  type RecoveryConditionalFormatCaptureResult,
  type RecoveryConditionalFormatRule,
  type RecoveryFormatRangeState,
  type RecoveryModifyStructureState,
} from "./recovery-states.js";
import {
  cloneGrid,
  gridStats,
  normalizeFormula,
  rowLength,
  toRestoreValues,
  valueAt,
} from "./recovery/grid.js";
import { estimateModifyStructureCellCount } from "./recovery/structure-state.js";
import { MAX_RECOVERY_CELLS, MAX_RECOVERY_ENTRIES } from "./recovery/constants.js";

export { MAX_RECOVERY_CELLS };

export type WorkbookRecoveryToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "format_cells"
  | "conditional_format"
  | "comments"
  | "modify_structure"
  | "restore_snapshot";

export type WorkbookRecoverySnapshotKind =
  | "range_values"
  | "format_cells_state"
  | "modify_structure_state"
  | "conditional_format_rules"
  | "comment_thread";

export interface WorkbookRecoverySnapshot {
  id: string;
  at: number;
  toolName: WorkbookRecoveryToolName;
  toolCallId: string;
  address: string;
  changedCount: number;
  cellCount: number;
  beforeValues: unknown[][];
  beforeFormulas: unknown[][];
  snapshotKind?: WorkbookRecoverySnapshotKind;
  formatRangeState?: RecoveryFormatRangeState;
  modifyStructureState?: RecoveryModifyStructureState;
  conditionalFormatRules?: RecoveryConditionalFormatRule[];
  commentThreadState?: RecoveryCommentThreadState;
  workbookId?: string;
  workbookLabel?: string;
  restoredFromSnapshotId?: string;
}

export interface AppendWorkbookRecoverySnapshotArgs {
  toolName: WorkbookRecoveryToolName;
  toolCallId: string;
  address: string;
  changedCount?: number;
  beforeValues: unknown[][];
  beforeFormulas: unknown[][];
  restoredFromSnapshotId?: string;
}

export interface AppendFormatCellsRecoverySnapshotArgs {
  toolName: WorkbookRecoveryToolName;
  toolCallId: string;
  address: string;
  changedCount?: number;
  formatRangeState: RecoveryFormatRangeState;
  restoredFromSnapshotId?: string;
}

export interface AppendModifyStructureRecoverySnapshotArgs {
  toolName: WorkbookRecoveryToolName;
  toolCallId: string;
  address: string;
  changedCount?: number;
  modifyStructureState: RecoveryModifyStructureState;
  restoredFromSnapshotId?: string;
}

export interface AppendConditionalFormatRecoverySnapshotArgs {
  toolName: WorkbookRecoveryToolName;
  toolCallId: string;
  address: string;
  changedCount?: number;
  cellCount: number;
  conditionalFormatRules: RecoveryConditionalFormatRule[];
  restoredFromSnapshotId?: string;
}

export interface AppendCommentThreadRecoverySnapshotArgs {
  toolName: WorkbookRecoveryToolName;
  toolCallId: string;
  address: string;
  changedCount?: number;
  commentThreadState: RecoveryCommentThreadState;
  restoredFromSnapshotId?: string;
}

export interface RestoreWorkbookRecoverySnapshotResult {
  restoredSnapshotId: string;
  inverseSnapshotId: string | null;
  address: string;
  changedCount: number;
}

interface WorkbookRangeState {
  values: unknown[][];
  formulas: unknown[][];
}

interface WorkbookRecoveryLogDependencies {
  getSettingsStore: () => Promise<SettingsStoreLike | null>;
  getWorkbookContext: () => Promise<WorkbookContext>;
  now: () => number;
  createId: () => string;
  applySnapshot: (address: string, values: unknown[][]) => Promise<WorkbookRangeState>;
  applyFormatCellsSnapshot: (
    address: string,
    state: RecoveryFormatRangeState,
  ) => Promise<RecoveryFormatRangeState>;
  applyModifyStructureSnapshot: (
    address: string,
    state: RecoveryModifyStructureState,
  ) => Promise<RecoveryModifyStructureState>;
  applyConditionalFormatSnapshot: (
    address: string,
    rules: RecoveryConditionalFormatRule[],
  ) => Promise<RecoveryConditionalFormatCaptureResult>;
  applyCommentThreadSnapshot: (
    address: string,
    state: RecoveryCommentThreadState,
  ) => Promise<RecoveryCommentThreadState>;
}

function defaultNow(): number {
  return Date.now();
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

async function defaultApplySnapshot(address: string, values: unknown[][]): Promise<WorkbookRangeState> {
  return excelRun<WorkbookRangeState>(async (context) => {
    const { range } = getRange(context, address);
    range.load("values,formulas");
    await context.sync();

    const beforeValues = cloneGrid(range.values);
    const beforeFormulas = cloneGrid(range.formulas);

    range.values = values;
    await context.sync();

    return {
      values: beforeValues,
      formulas: beforeFormulas,
    };
  });
}

async function defaultApplyFormatCellsSnapshot(
  address: string,
  state: RecoveryFormatRangeState,
): Promise<RecoveryFormatRangeState> {
  return applyFormatCellsState(address, state);
}

async function defaultApplyModifyStructureSnapshot(
  _address: string,
  state: RecoveryModifyStructureState,
): Promise<RecoveryModifyStructureState> {
  return applyModifyStructureState(state);
}

async function defaultApplyConditionalFormatSnapshot(
  address: string,
  rules: RecoveryConditionalFormatRule[],
): Promise<RecoveryConditionalFormatCaptureResult> {
  return applyConditionalFormatState(address, rules);
}

async function defaultApplyCommentThreadSnapshot(
  address: string,
  state: RecoveryCommentThreadState,
): Promise<RecoveryCommentThreadState> {
  return applyCommentThreadState(address, state);
}

function serializeComparable(raw: unknown): string {
  if (raw === null || raw === undefined || raw === "") return "";

  if (typeof raw === "string") return raw;
  if (typeof raw === "number") return Number.isNaN(raw) ? "NaN" : String(raw);
  if (typeof raw === "boolean") return raw ? "true" : "false";
  if (typeof raw === "bigint") return String(raw);
  if (typeof raw === "symbol") return raw.description ?? "";
  if (typeof raw === "function") return "[function]";

  try {
    return JSON.stringify(raw);
  } catch {
    return "[unserializable]";
  }
}

function countChangedCells(args: {
  beforeValues: unknown[][];
  beforeFormulas: unknown[][];
  afterValues: unknown[][];
  afterFormulas: unknown[][];
}): number {
  const rowCount = Math.max(
    args.beforeValues.length,
    args.beforeFormulas.length,
    args.afterValues.length,
    args.afterFormulas.length,
  );

  let changedCount = 0;

  for (let row = 0; row < rowCount; row += 1) {
    const colCount = Math.max(
      rowLength(args.beforeValues, row),
      rowLength(args.beforeFormulas, row),
      rowLength(args.afterValues, row),
      rowLength(args.afterFormulas, row),
    );

    for (let col = 0; col < colCount; col += 1) {
      const beforeValue = serializeComparable(valueAt(args.beforeValues, row, col));
      const afterValue = serializeComparable(valueAt(args.afterValues, row, col));
      const beforeFormula = normalizeFormula(valueAt(args.beforeFormulas, row, col));
      const afterFormula = normalizeFormula(valueAt(args.afterFormulas, row, col));

      if (beforeValue !== afterValue || beforeFormula !== afterFormula) {
        changedCount += 1;
      }
    }
  }

  return changedCount;
}

function clampLimit(limit: number): number {
  if (!Number.isFinite(limit)) return 20;

  const rounded = Math.floor(limit);
  if (rounded <= 0) return 0;
  if (rounded > MAX_RECOVERY_ENTRIES) return MAX_RECOVERY_ENTRIES;
  return rounded;
}

function matchesWorkbook(snapshot: WorkbookRecoverySnapshot, workbookId: string): boolean {
  return snapshot.workbookId === workbookId;
}

export class WorkbookRecoveryLog {
  private readonly dependencies: WorkbookRecoveryLogDependencies;
  private loaded = false;
  private snapshots: WorkbookRecoverySnapshot[] = [];

  constructor(dependencies: Partial<WorkbookRecoveryLogDependencies> = {}) {
    this.dependencies = {
      getSettingsStore: dependencies.getSettingsStore ?? defaultGetSettingsStore,
      getWorkbookContext: dependencies.getWorkbookContext ?? getWorkbookContext,
      now: dependencies.now ?? defaultNow,
      createId: dependencies.createId ?? defaultCreateId,
      applySnapshot: dependencies.applySnapshot ?? defaultApplySnapshot,
      applyFormatCellsSnapshot:
        dependencies.applyFormatCellsSnapshot ?? defaultApplyFormatCellsSnapshot,
      applyModifyStructureSnapshot:
        dependencies.applyModifyStructureSnapshot ?? defaultApplyModifyStructureSnapshot,
      applyConditionalFormatSnapshot:
        dependencies.applyConditionalFormatSnapshot ?? defaultApplyConditionalFormatSnapshot,
      applyCommentThreadSnapshot:
        dependencies.applyCommentThreadSnapshot ?? defaultApplyCommentThreadSnapshot,
    };
  }

  private async ensureLoaded(): Promise<void> {
    if (this.loaded) return;
    this.loaded = true;

    const settings = await this.dependencies.getSettingsStore();
    const payload = await readPersistedWorkbookRecoveryPayload(settings);
    this.snapshots = parsePersistedSnapshots(payload, {
      maxEntries: MAX_RECOVERY_ENTRIES,
    });
  }

  private async persist(): Promise<void> {
    const settings = await this.dependencies.getSettingsStore();
    const payload = createPersistedWorkbookRecoveryPayload(this.snapshots);
    await writePersistedWorkbookRecoveryPayload(settings, payload);
  }

  private async resolveWorkbookIdentity(
    workbookContextOverride?: WorkbookContext,
  ): Promise<{ workbookContext: WorkbookContext; workbookId: string; workbookLabel?: string } | null> {
    try {
      const workbookContext = workbookContextOverride ?? await this.dependencies.getWorkbookContext();
      if (!workbookContext.workbookId) {
        return null;
      }

      return {
        workbookContext,
        workbookId: workbookContext.workbookId,
        workbookLabel: formatWorkbookLabel(workbookContext),
      };
    } catch {
      return null;
    }
  }

  private async appendSnapshot(
    snapshot: WorkbookRecoverySnapshot,
  ): Promise<WorkbookRecoverySnapshot> {
    this.snapshots = [snapshot, ...this.snapshots].slice(0, MAX_RECOVERY_ENTRIES);
    await this.persist();
    return snapshot;
  }

  private async appendRangeWithContext(
    args: AppendWorkbookRecoverySnapshotArgs,
    workbookContextOverride?: WorkbookContext,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const values = cloneGrid(args.beforeValues);
    const formulas = cloneGrid(args.beforeFormulas);

    const stats = gridStats(values, formulas);
    if (stats.cellCount <= 0) return null;
    if (stats.cellCount > MAX_RECOVERY_CELLS) return null;

    const changedCount = typeof args.changedCount === "number"
      ? Math.max(0, Math.floor(args.changedCount))
      : stats.cellCount;

    const workbookIdentity = await this.resolveWorkbookIdentity(workbookContextOverride);
    if (!workbookIdentity) return null;

    return this.appendSnapshot({
      id: this.dependencies.createId(),
      at: this.dependencies.now(),
      toolName: args.toolName,
      toolCallId: args.toolCallId,
      address: args.address,
      changedCount,
      cellCount: stats.cellCount,
      beforeValues: values,
      beforeFormulas: formulas,
      snapshotKind: "range_values",
      workbookId: workbookIdentity.workbookId,
      workbookLabel: workbookIdentity.workbookLabel,
      restoredFromSnapshotId: args.restoredFromSnapshotId,
    });
  }

  private async appendFormatCellsWithContext(
    args: AppendFormatCellsRecoverySnapshotArgs,
    workbookContextOverride?: WorkbookContext,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const formatRangeState = cloneRecoveryFormatRangeState(args.formatRangeState);
    if (formatRangeState.cellCount <= 0) return null;
    if (formatRangeState.cellCount > MAX_RECOVERY_CELLS) return null;

    const workbookIdentity = await this.resolveWorkbookIdentity(workbookContextOverride);
    if (!workbookIdentity) return null;

    const changedCount = typeof args.changedCount === "number"
      ? Math.max(0, Math.floor(args.changedCount))
      : formatRangeState.cellCount;

    return this.appendSnapshot({
      id: this.dependencies.createId(),
      at: this.dependencies.now(),
      toolName: args.toolName,
      toolCallId: args.toolCallId,
      address: args.address,
      changedCount,
      cellCount: formatRangeState.cellCount,
      beforeValues: [],
      beforeFormulas: [],
      snapshotKind: "format_cells_state",
      formatRangeState,
      workbookId: workbookIdentity.workbookId,
      workbookLabel: workbookIdentity.workbookLabel,
      restoredFromSnapshotId: args.restoredFromSnapshotId,
    });
  }

  private async appendModifyStructureWithContext(
    args: AppendModifyStructureRecoverySnapshotArgs,
    workbookContextOverride?: WorkbookContext,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const modifyStructureState = cloneRecoveryModifyStructureState(args.modifyStructureState);
    const cellCount = estimateModifyStructureCellCount(modifyStructureState);
    if (cellCount <= 0) return null;
    if (cellCount > MAX_RECOVERY_CELLS) return null;

    const workbookIdentity = await this.resolveWorkbookIdentity(workbookContextOverride);
    if (!workbookIdentity) return null;

    const changedCount = typeof args.changedCount === "number"
      ? Math.max(0, Math.floor(args.changedCount))
      : cellCount;

    return this.appendSnapshot({
      id: this.dependencies.createId(),
      at: this.dependencies.now(),
      toolName: args.toolName,
      toolCallId: args.toolCallId,
      address: args.address,
      changedCount,
      cellCount,
      beforeValues: [],
      beforeFormulas: [],
      snapshotKind: "modify_structure_state",
      modifyStructureState,
      workbookId: workbookIdentity.workbookId,
      workbookLabel: workbookIdentity.workbookLabel,
      restoredFromSnapshotId: args.restoredFromSnapshotId,
    });
  }

  private async appendConditionalFormatWithContext(
    args: AppendConditionalFormatRecoverySnapshotArgs,
    workbookContextOverride?: WorkbookContext,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const rules = cloneRecoveryConditionalFormatRules(args.conditionalFormatRules);
    const normalizedCellCount = Math.max(0, Math.floor(args.cellCount));
    if (normalizedCellCount <= 0) return null;

    const workbookIdentity = await this.resolveWorkbookIdentity(workbookContextOverride);
    if (!workbookIdentity) return null;

    const changedCount = typeof args.changedCount === "number"
      ? Math.max(0, Math.floor(args.changedCount))
      : normalizedCellCount;

    return this.appendSnapshot({
      id: this.dependencies.createId(),
      at: this.dependencies.now(),
      toolName: args.toolName,
      toolCallId: args.toolCallId,
      address: args.address,
      changedCount,
      cellCount: normalizedCellCount,
      beforeValues: [],
      beforeFormulas: [],
      snapshotKind: "conditional_format_rules",
      conditionalFormatRules: rules,
      workbookId: workbookIdentity.workbookId,
      workbookLabel: workbookIdentity.workbookLabel,
      restoredFromSnapshotId: args.restoredFromSnapshotId,
    });
  }

  private async appendCommentThreadWithContext(
    args: AppendCommentThreadRecoverySnapshotArgs,
    workbookContextOverride?: WorkbookContext,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const workbookIdentity = await this.resolveWorkbookIdentity(workbookContextOverride);
    if (!workbookIdentity) return null;

    const changedCount = typeof args.changedCount === "number"
      ? Math.max(0, Math.floor(args.changedCount))
      : 1;

    return this.appendSnapshot({
      id: this.dependencies.createId(),
      at: this.dependencies.now(),
      toolName: args.toolName,
      toolCallId: args.toolCallId,
      address: args.address,
      changedCount,
      cellCount: 1,
      beforeValues: [],
      beforeFormulas: [],
      snapshotKind: "comment_thread",
      commentThreadState: cloneRecoveryCommentThreadState(args.commentThreadState),
      workbookId: workbookIdentity.workbookId,
      workbookLabel: workbookIdentity.workbookLabel,
      restoredFromSnapshotId: args.restoredFromSnapshotId,
    });
  }

  async append(args: AppendWorkbookRecoverySnapshotArgs): Promise<WorkbookRecoverySnapshot | null> {
    await this.ensureLoaded();
    return this.appendRangeWithContext(args);
  }

  async appendFormatCells(
    args: AppendFormatCellsRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    await this.ensureLoaded();
    return this.appendFormatCellsWithContext(args);
  }

  async appendModifyStructure(
    args: AppendModifyStructureRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    await this.ensureLoaded();
    return this.appendModifyStructureWithContext(args);
  }

  async appendConditionalFormat(
    args: AppendConditionalFormatRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    await this.ensureLoaded();
    return this.appendConditionalFormatWithContext(args);
  }

  async appendCommentThread(
    args: AppendCommentThreadRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    await this.ensureLoaded();
    return this.appendCommentThreadWithContext(args);
  }

  async list(opts: { limit?: number; workbookId?: string | null } = {}): Promise<WorkbookRecoverySnapshot[]> {
    await this.ensureLoaded();

    const limit = clampLimit(opts.limit ?? 20);
    const workbookId = opts.workbookId;

    if (workbookId === undefined) {
      return this.snapshots.slice(0, limit);
    }

    if (!workbookId) {
      return [];
    }

    const filtered = this.snapshots.filter((snapshot) => matchesWorkbook(snapshot, workbookId));
    return filtered.slice(0, limit);
  }

  async listForCurrentWorkbook(limit = 20): Promise<WorkbookRecoverySnapshot[]> {
    const workbookContext = await this.dependencies.getWorkbookContext();
    const workbookId = workbookContext.workbookId;
    if (!workbookId) return [];

    return this.list({ limit, workbookId });
  }

  async delete(snapshotId: string): Promise<boolean> {
    await this.ensureLoaded();

    const workbookContext = await this.dependencies.getWorkbookContext();
    const workbookId = workbookContext.workbookId;
    if (!workbookId) return false;

    const previousLength = this.snapshots.length;
    this.snapshots = this.snapshots.filter(
      (snapshot) => !(snapshot.id === snapshotId && matchesWorkbook(snapshot, workbookId)),
    );

    if (this.snapshots.length === previousLength) {
      return false;
    }

    await this.persist();
    return true;
  }

  async clearForCurrentWorkbook(): Promise<number> {
    await this.ensureLoaded();

    const workbookContext = await this.dependencies.getWorkbookContext();
    const workbookId = workbookContext.workbookId;
    if (!workbookId) return 0;

    const previousLength = this.snapshots.length;
    this.snapshots = this.snapshots.filter((snapshot) => !matchesWorkbook(snapshot, workbookId));

    const removed = previousLength - this.snapshots.length;
    if (removed > 0) {
      await this.persist();
    }

    return removed;
  }

  async restore(snapshotId: string): Promise<RestoreWorkbookRecoverySnapshotResult> {
    await this.ensureLoaded();

    const snapshot = this.snapshots.find((item) => item.id === snapshotId);
    if (!snapshot) {
      throw new Error("Snapshot not found.");
    }

    const workbookContext = await this.dependencies.getWorkbookContext();

    return restoreWorkbookRecoverySnapshot({
      snapshot,
      workbookContext,
      dependencies: {
        applySnapshot: this.dependencies.applySnapshot,
        applyFormatCellsSnapshot: this.dependencies.applyFormatCellsSnapshot,
        applyModifyStructureSnapshot: this.dependencies.applyModifyStructureSnapshot,
        applyConditionalFormatSnapshot: this.dependencies.applyConditionalFormatSnapshot,
        applyCommentThreadSnapshot: this.dependencies.applyCommentThreadSnapshot,
        appendRangeSnapshot: (args, context) => this.appendRangeWithContext(args, context),
        appendFormatCellsSnapshot: (args, context) => this.appendFormatCellsWithContext(args, context),
        appendModifyStructureSnapshot: (args, context) => this.appendModifyStructureWithContext(args, context),
        appendConditionalFormatSnapshot: (args, context) => this.appendConditionalFormatWithContext(args, context),
        appendCommentThreadSnapshot: (args, context) => this.appendCommentThreadWithContext(args, context),
        toRestoreValues,
        countChangedCells,
      },
    });
  }

  async restoreLatestForCurrentWorkbook(): Promise<RestoreWorkbookRecoverySnapshotResult> {
    const snapshots = await this.listForCurrentWorkbook(1);
    const latest = snapshots[0];
    if (!latest) {
      throw new Error("No backups found for this workbook.");
    }

    return this.restore(latest.id);
  }
}

let singleton: WorkbookRecoveryLog | null = null;

export function getWorkbookRecoveryLog(): WorkbookRecoveryLog {
  if (!singleton) {
    singleton = new WorkbookRecoveryLog();
  }

  return singleton;
}
