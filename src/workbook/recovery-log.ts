/**
 * Workbook recovery snapshots.
 *
 * Stores lightweight, local-only backups for workbook cell edits so users
 * can revert Pi's changes without pre-execution approval prompts.
 */

import { colToLetter, excelRun, getRange, parseCell, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import { getWorkbookContext, type WorkbookContext } from "./context.js";
import type { DocumentInstanceIdentity } from "../host/document-instance.js";
import {
  createRecoveryScopeResolver,
  defaultGetDocumentInstance,
  type RecoveryScopeResolver,
  type RecoveryWorkbookScope,
} from "./recovery-scope.js";
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
  readRetentionLimit,
  RECOVERY_SETTING_KEY,
} from "./recovery/log-store.js";
import {
  applyChartState,
  applyCommentThreadState,
  applyConditionalFormatState,
  applyFormatCellsState,
  applyModifyStructureState,
  cloneRecoveryChartState,
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  cloneRecoveryFormatRangeState,
  cloneRecoveryModifyStructureState,
  type RecoveryChartApplyResult,
  type RecoveryChartState,
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
import { getErrorMessage } from "../utils/errors.js";
import { estimateModifyStructureCellCount } from "./recovery/structure-state.js";
import {
  MAX_RECOVERY_CELLS,
  MAX_RECOVERY_ENTRIES,
  MAX_STORED_RECOVERY_ENTRIES,
} from "./recovery/constants.js";
import type { SettingsUpdater } from "../storage/local/settings-store.js";

export { MAX_RECOVERY_CELLS };

export type WorkbookRecoveryToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "format_cells"
  | "conditional_format"
  | "comments"
  | "charts"
  | "modify_structure"
  | "restore_snapshot";

export type WorkbookRecoverySnapshotKind =
  | "range_values"
  | "format_cells_state"
  | "modify_structure_state"
  | "conditional_format_rules"
  | "comment_thread"
  | "chart_state";

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
  /** Range backups: the formulas/values Pi's change left (absent on older backups). */
  afterFormulas?: unknown[][];
  snapshotKind?: WorkbookRecoverySnapshotKind;
  formatRangeState?: RecoveryFormatRangeState;
  modifyStructureState?: RecoveryModifyStructureState;
  conditionalFormatRules?: RecoveryConditionalFormatRule[];
  commentThreadState?: RecoveryCommentThreadState;
  chartState?: RecoveryChartState;
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
  /** What the change left in the range, so a later restore can spot edits made since. */
  afterFormulas?: unknown[][];
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

export interface AppendChartRecoverySnapshotArgs {
  toolName: WorkbookRecoveryToolName;
  toolCallId: string;
  address: string;
  changedCount?: number;
  chartState: RecoveryChartState;
  restoredFromSnapshotId?: string;
}

export interface RestoreWorkbookRecoverySnapshotResult {
  restoredSnapshotId: string;
  inverseSnapshotId: string | null;
  /** Set when the restore succeeded but its inverse backup could not be saved. */
  inverseSnapshotError?: string;
  address: string;
  changedCount: number;
}

interface WorkbookRangeState {
  values: unknown[][];
  formulas: unknown[][];
}

interface WorkbookRecoveryLogDependencies {
  settings: SettingsUpdater | null;
  getSettingsStore: () => Promise<SettingsUpdater | null>;
  getWorkbookContext: () => Promise<WorkbookContext>;
  getDocumentInstance: () => DocumentInstanceIdentity | null;
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
  applyChartSnapshot: (
    address: string,
    state: RecoveryChartState,
  ) => Promise<RecoveryChartApplyResult>;
  /** Current formulas (values for constants) of a range. */
  readRangeFormulas: (address: string) => Promise<unknown[][]>;
}

/** What restoring a backup would do, for a warning before it overwrites anything. */
export interface WorkbookRecoveryRestorePreview {
  id: string;
  /** When the backup was taken (ms since epoch). */
  at: number;
  address: string;
  /**
   * Cells changed since Pi's change, which the restore would overwrite.
   * `null` when none changed, or when the backup cannot tell (non-range
   * backups, and range backups saved before this was recorded).
   */
  laterEdits: { count: number; sample: Array<{ address: string; content: string }> } | null;
}

const LATER_EDITS_SAMPLE_SIZE = 3;

async function defaultReadRangeFormulas(address: string): Promise<unknown[][]> {
  return excelRun(async (context) => {
    const { range } = getRange(context, address);
    range.load("formulas");
    await context.sync();
    return cloneGrid(range.formulas);
  });
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

async function defaultApplyChartSnapshot(
  address: string,
  state: RecoveryChartState,
): Promise<RecoveryChartApplyResult> {
  return applyChartState(address, state);
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

/**
 * Keeps the newest `perWorkbookLimit` backups of each workbook, then the
 * newest `MAX_STORED_RECOVERY_ENTRIES` overall. Saving the workbook does not
 * expire backups: on autosave hosts (Excel on the web, AutoSave) a save
 * happens every few seconds, which would leave nothing to revert.
 */
function applyRetention(
  snapshots: WorkbookRecoverySnapshot[],
  perWorkbookLimit: number,
): WorkbookRecoverySnapshot[] {
  const keptPerWorkbook = new Map<string, number>();
  return [...snapshots]
    .sort((a, b) => b.at - a.at)
    .filter((snapshot) => {
      const key = snapshot.workbookId ?? "";
      const kept = keptPerWorkbook.get(key) ?? 0;
      if (kept >= perWorkbookLimit) return false;
      keptPerWorkbook.set(key, kept + 1);
      return true;
    })
    .slice(0, MAX_STORED_RECOVERY_ENTRIES);
}

export class WorkbookRecoveryLog {
  private readonly dependencies: WorkbookRecoveryLogDependencies;
  private readonly scope: RecoveryScopeResolver;

  constructor(dependencies: Partial<WorkbookRecoveryLogDependencies> = {}) {
    const getSettingsStore = dependencies.settings !== undefined
      ? () => Promise.resolve(dependencies.settings ?? null)
      : dependencies.getSettingsStore ?? defaultGetSettingsStore;

    this.dependencies = {
      settings: dependencies.settings ?? null,
      getSettingsStore,
      getWorkbookContext: dependencies.getWorkbookContext ?? getWorkbookContext,
      getDocumentInstance: dependencies.getDocumentInstance ?? defaultGetDocumentInstance,
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
      applyChartSnapshot:
        dependencies.applyChartSnapshot ?? defaultApplyChartSnapshot,
      readRangeFormulas: dependencies.readRangeFormulas ?? defaultReadRangeFormulas,
    };
    this.scope = createRecoveryScopeResolver({
      getWorkbookContext: this.dependencies.getWorkbookContext,
      getDocumentInstance: this.dependencies.getDocumentInstance,
    });
  }

  /**
   * The stored list, read fresh: other taskpanes on this computer append to
   * and delete from the same list, so nothing is cached between calls.
   */
  private async readSnapshots(): Promise<WorkbookRecoverySnapshot[]> {
    const settings = await this.dependencies.getSettingsStore();
    const payload = await readPersistedWorkbookRecoveryPayload(settings);
    return parsePersistedSnapshots(payload, { maxEntries: MAX_STORED_RECOVERY_ENTRIES });
  }

  /**
   * Applies `change` to the list as stored at write time, atomically, so a
   * concurrent write from another taskpane is never overwritten. Returns
   * `null` when there is no storage. Rejects when the write does not commit.
   */
  private async updateSnapshots<T>(
    change: (snapshots: WorkbookRecoverySnapshot[]) => { snapshots: WorkbookRecoverySnapshot[]; result: T },
  ): Promise<T | null> {
    const settings = await this.dependencies.getSettingsStore();
    if (!settings) return null;

    const outcome: { result: T | null } = { result: null };
    await settings.update(
      RECOVERY_SETTING_KEY,
      (raw) => parsePersistedSnapshots(raw, { maxEntries: MAX_STORED_RECOVERY_ENTRIES }),
      (current) => {
        const next = change(current);
        outcome.result = next.result;
        return createPersistedWorkbookRecoveryPayload(next.snapshots);
      },
    );
    return outcome.result;
  }

  private async readRetentionLimit(): Promise<number> {
    return readRetentionLimit(await this.dependencies.getSettingsStore());
  }

  /**
   * Scope for a new checkpoint. A restore passes the scope it already resolved so
   * the inverse checkpoint lands under the same identity. Host failures mean
   * "no identity": the mutation must still succeed, so nothing is thrown here.
   */
  private async resolveAppendScope(
    scopeOverride?: RecoveryWorkbookScope,
  ): Promise<RecoveryWorkbookScope | null> {
    if (scopeOverride) return scopeOverride;
    try {
      return await this.scope.resolveForAppend();
    } catch {
      return null;
    }
  }

  /**
   * Stores a new backup. Returns `null` when there is no storage; throws
   * "Backup could not be saved: …" when the write fails, so callers can tell
   * the user this change cannot be reverted rather than implying it can.
   */
  private async appendSnapshot(
    snapshot: WorkbookRecoverySnapshot,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const limit = await this.readRetentionLimit();
    try {
      const stored = await this.updateSnapshots((snapshots) => ({
        snapshots: applyRetention([snapshot, ...snapshots], limit),
        result: snapshot,
      }));
      return stored;
    } catch (error) {
      throw new Error(`Backup could not be saved: ${getErrorMessage(error)}`);
    }
  }

  private async appendRangeWithContext(
    args: AppendWorkbookRecoverySnapshotArgs,
    scopeOverride?: RecoveryWorkbookScope,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const values = cloneGrid(args.beforeValues);
    const formulas = cloneGrid(args.beforeFormulas);

    const stats = gridStats(values, formulas);
    if (stats.cellCount <= 0) return null;
    if (stats.cellCount > MAX_RECOVERY_CELLS) return null;

    const changedCount = typeof args.changedCount === "number"
      ? Math.max(0, Math.floor(args.changedCount))
      : stats.cellCount;

    const scope = await this.resolveAppendScope(scopeOverride);
    if (!scope) return null;

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
      ...(args.afterFormulas ? { afterFormulas: cloneGrid(args.afterFormulas) } : {}),
      snapshotKind: "range_values",
      workbookId: scope.workbookId,
      workbookLabel: scope.workbookLabel,
      ...(args.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: args.restoredFromSnapshotId } : {}),
    });
  }

  private async appendFormatCellsWithContext(
    args: AppendFormatCellsRecoverySnapshotArgs,
    scopeOverride?: RecoveryWorkbookScope,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const formatRangeState = cloneRecoveryFormatRangeState(args.formatRangeState);
    if (formatRangeState.cellCount <= 0) return null;
    if (formatRangeState.cellCount > MAX_RECOVERY_CELLS) return null;

    const scope = await this.resolveAppendScope(scopeOverride);
    if (!scope) return null;

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
      workbookId: scope.workbookId,
      workbookLabel: scope.workbookLabel,
      ...(args.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: args.restoredFromSnapshotId } : {}),
    });
  }

  private async appendModifyStructureWithContext(
    args: AppendModifyStructureRecoverySnapshotArgs,
    scopeOverride?: RecoveryWorkbookScope,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const modifyStructureState = cloneRecoveryModifyStructureState(args.modifyStructureState);
    const cellCount = estimateModifyStructureCellCount(modifyStructureState);
    if (cellCount <= 0) return null;
    if (cellCount > MAX_RECOVERY_CELLS) return null;

    const scope = await this.resolveAppendScope(scopeOverride);
    if (!scope) return null;

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
      workbookId: scope.workbookId,
      workbookLabel: scope.workbookLabel,
      ...(args.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: args.restoredFromSnapshotId } : {}),
    });
  }

  private async appendConditionalFormatWithContext(
    args: AppendConditionalFormatRecoverySnapshotArgs,
    scopeOverride?: RecoveryWorkbookScope,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const rules = cloneRecoveryConditionalFormatRules(args.conditionalFormatRules);
    const normalizedCellCount = Math.max(0, Math.floor(args.cellCount));
    if (normalizedCellCount <= 0) return null;

    const scope = await this.resolveAppendScope(scopeOverride);
    if (!scope) return null;

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
      workbookId: scope.workbookId,
      workbookLabel: scope.workbookLabel,
      ...(args.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: args.restoredFromSnapshotId } : {}),
    });
  }

  private async appendCommentThreadWithContext(
    args: AppendCommentThreadRecoverySnapshotArgs,
    scopeOverride?: RecoveryWorkbookScope,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const scope = await this.resolveAppendScope(scopeOverride);
    if (!scope) return null;

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
      workbookId: scope.workbookId,
      workbookLabel: scope.workbookLabel,
      ...(args.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: args.restoredFromSnapshotId } : {}),
    });
  }

  private async appendChartWithContext(
    args: AppendChartRecoverySnapshotArgs,
    scopeOverride?: RecoveryWorkbookScope,
  ): Promise<WorkbookRecoverySnapshot | null> {
    const scope = await this.resolveAppendScope(scopeOverride);
    if (!scope) return null;

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
      snapshotKind: "chart_state",
      chartState: cloneRecoveryChartState(args.chartState),
      workbookId: scope.workbookId,
      workbookLabel: scope.workbookLabel,
      ...(args.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: args.restoredFromSnapshotId } : {}),
    });
  }

  async append(args: AppendWorkbookRecoverySnapshotArgs): Promise<WorkbookRecoverySnapshot | null> {
    return this.appendRangeWithContext(args);
  }

  async appendFormatCells(
    args: AppendFormatCellsRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    return this.appendFormatCellsWithContext(args);
  }

  async appendModifyStructure(
    args: AppendModifyStructureRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    return this.appendModifyStructureWithContext(args);
  }

  async appendConditionalFormat(
    args: AppendConditionalFormatRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    return this.appendConditionalFormatWithContext(args);
  }

  async appendCommentThread(
    args: AppendCommentThreadRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    return this.appendCommentThreadWithContext(args);
  }

  async appendChart(
    args: AppendChartRecoverySnapshotArgs,
  ): Promise<WorkbookRecoverySnapshot | null> {
    return this.appendChartWithContext(args);
  }

  async list(opts: { limit?: number; workbookId?: string | null } = {}): Promise<WorkbookRecoverySnapshot[]> {
    const snapshots = await this.readSnapshots();

    const limit = clampLimit(opts.limit ?? 20);
    const workbookId = opts.workbookId;

    if (workbookId === undefined) {
      return snapshots.slice(0, limit);
    }

    if (!workbookId) {
      return [];
    }

    const filtered = snapshots.filter((snapshot) => matchesWorkbook(snapshot, workbookId));
    return filtered.slice(0, limit);
  }

  async listForCurrentWorkbook(limit = 20): Promise<WorkbookRecoverySnapshot[]> {
    const scope = await this.scope.resolveForRead();
    if (!scope) return [];

    return this.list({ limit, workbookId: scope.workbookId });
  }

  async delete(snapshotId: string): Promise<boolean> {
    const scope = await this.scope.resolveForRead();
    if (!scope) return false;
    const workbookId = scope.workbookId;

    const deleted = await this.updateSnapshots((snapshots) => {
      const kept = snapshots.filter(
        (snapshot) => !(snapshot.id === snapshotId && matchesWorkbook(snapshot, workbookId)),
      );
      return { snapshots: kept, result: kept.length < snapshots.length };
    });
    return deleted ?? false;
  }

  async clearForCurrentWorkbook(): Promise<number> {
    const scope = await this.scope.resolveForRead();
    if (!scope) return 0;

    const removed = await this.updateSnapshots((snapshots) => {
      const kept = snapshots.filter((snapshot) => !matchesWorkbook(snapshot, scope.workbookId));
      return { snapshots: kept, result: snapshots.length - kept.length };
    });
    return removed ?? 0;
  }

  /**
   * Moves backups from one identity to another. A never-saved workbook's
   * backups are filed under its document token; when its first save gives it
   * a path identity, they follow it there instead of being stranded.
   */
  async rescopeWorkbook(args: { fromWorkbookId: string; toWorkbookId: string; workbookLabel: string }): Promise<number> {
    const limit = await this.readRetentionLimit();
    const moved = await this.updateSnapshots((snapshots) => {
      let count = 0;
      const next = snapshots.map((snapshot) => {
        if (!matchesWorkbook(snapshot, args.fromWorkbookId)) return snapshot;
        count += 1;
        return { ...snapshot, workbookId: args.toWorkbookId, workbookLabel: args.workbookLabel };
      });
      return { snapshots: count > 0 ? applyRetention(next, limit) : snapshots, result: count };
    });
    return moved ?? 0;
  }

  /** Applies the "Keep at most" limit now, for example right after the user lowers it. */
  async applyRetentionLimit(): Promise<void> {
    const limit = await this.readRetentionLimit();
    await this.updateSnapshots((snapshots) => ({ snapshots: applyRetention(snapshots, limit), result: null }));
  }

  async restore(snapshotId: string): Promise<RestoreWorkbookRecoverySnapshotResult> {
    const snapshot = (await this.readSnapshots()).find((item) => item.id === snapshotId);
    if (!snapshot) {
      throw new Error("Snapshot not found.");
    }

    const scope = await this.scope.resolveForRead();
    if (!scope) {
      throw new Error("Current workbook identity is unavailable; cannot safely restore this snapshot.");
    }

    // The workbook has changed by the time the inverse backup is stored, so a
    // failed save must not turn a completed restore into a reported failure.
    const inverse: { error: string | null } = { error: null };
    const keepRestoreOnAppendFailure = <TArgs>(
      append: (args: TArgs, scope: RecoveryWorkbookScope) => Promise<WorkbookRecoverySnapshot | null>,
    ) => async (args: TArgs, appendScope: RecoveryWorkbookScope): Promise<WorkbookRecoverySnapshot | null> => {
      try {
        return await append(args, appendScope);
      } catch (error) {
        inverse.error = getErrorMessage(error);
        return null;
      }
    };

    const result = await restoreWorkbookRecoverySnapshot({
      snapshot,
      scope,
      dependencies: {
        applySnapshot: this.dependencies.applySnapshot,
        applyFormatCellsSnapshot: this.dependencies.applyFormatCellsSnapshot,
        applyModifyStructureSnapshot: this.dependencies.applyModifyStructureSnapshot,
        applyConditionalFormatSnapshot: this.dependencies.applyConditionalFormatSnapshot,
        applyCommentThreadSnapshot: this.dependencies.applyCommentThreadSnapshot,
        appendRangeSnapshot: keepRestoreOnAppendFailure((args: AppendWorkbookRecoverySnapshotArgs, context) => this.appendRangeWithContext(args, context)),
        appendFormatCellsSnapshot: keepRestoreOnAppendFailure((args: AppendFormatCellsRecoverySnapshotArgs, context) => this.appendFormatCellsWithContext(args, context)),
        appendModifyStructureSnapshot: keepRestoreOnAppendFailure((args: AppendModifyStructureRecoverySnapshotArgs, context) => this.appendModifyStructureWithContext(args, context)),
        appendConditionalFormatSnapshot: keepRestoreOnAppendFailure((args: AppendConditionalFormatRecoverySnapshotArgs, context) => this.appendConditionalFormatWithContext(args, context)),
        appendCommentThreadSnapshot: keepRestoreOnAppendFailure((args: AppendCommentThreadRecoverySnapshotArgs, context) => this.appendCommentThreadWithContext(args, context)),
        appendChartSnapshot: keepRestoreOnAppendFailure((args: AppendChartRecoverySnapshotArgs, context) => this.appendChartWithContext(args, context)),
        applyChartSnapshot: this.dependencies.applyChartSnapshot,
        toRestoreValues,
        countChangedCells,
      },
    });

    if (!result.inverseSnapshotId) {
      // Nothing to roll back to (a created chart cannot be re-created, or the
      // rollback could not be stored). The backup is used up: re-applying it
      // does nothing, and left in place it would stay "latest" and hide every
      // older backup from /revert.
      try {
        await this.updateSnapshots((snapshots) => ({
          snapshots: snapshots.filter((item) => item.id !== snapshot.id),
          result: null,
        }));
      } catch (error) {
        console.warn("[pi] Used backup could not be removed.", error);
      }
    }

    return inverse.error === null ? result : { ...result, inverseSnapshotError: inverse.error };
  }

  /**
   * Reads, without changing anything, what a restore of `snapshotId` would
   * overwrite: cells edited since Pi's change. Backups now last until the
   * "Keep at most" limit removes them, often days, so the user may have
   * worked in the range since.
   */
  async previewRestore(snapshotId: string): Promise<WorkbookRecoveryRestorePreview> {
    const snapshot = (await this.readSnapshots()).find((item) => item.id === snapshotId);
    if (!snapshot) {
      throw new Error("Snapshot not found.");
    }
    // Same refusals as restore, so a preview never reads another workbook.
    const scope = await this.scope.resolveForRead();
    if (!scope) {
      throw new Error("Current workbook identity is unavailable; cannot safely restore this snapshot.");
    }
    if (!matchesWorkbook(snapshot, scope.workbookId)) {
      throw new Error("Snapshot belongs to a different workbook.");
    }

    const preview: WorkbookRecoveryRestorePreview = {
      id: snapshot.id,
      at: snapshot.at,
      address: snapshot.address,
      laterEdits: null,
    };
    const after = snapshot.afterFormulas;
    if ((snapshot.snapshotKind ?? "range_values") !== "range_values" || !after) return preview;

    const current = await this.dependencies.readRangeFormulas(snapshot.address);
    const { sheet, address } = parseRangeRef(snapshot.address);
    const start = parseCell(address.split(":")[0] ?? address);
    const sample: Array<{ address: string; content: string }> = [];
    let count = 0;
    after.forEach((row, r) => {
      row.forEach((expected, c) => {
        const now = current[r]?.[c];
        if (JSON.stringify(now ?? "") === JSON.stringify(expected ?? "")) return;
        const content = typeof now === "string" || typeof now === "number" || typeof now === "boolean" ? String(now) : "";
        count += 1;
        if (sample.length >= LATER_EDITS_SAMPLE_SIZE) return;
        const cell = `${colToLetter(start.col + c)}${start.row + r}`;
        sample.push({ address: sheet ? qualifiedAddress(sheet, cell) : cell, content: content === "" ? "(empty)" : content });
      });
    });
    if (count > 0) preview.laterEdits = { count, sample };
    return preview;
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

export function getWorkbookRecoveryLog(
  defaultLog?: WorkbookRecoveryLog,
): WorkbookRecoveryLog {
  if (!singleton) {
    singleton = defaultLog ?? new WorkbookRecoveryLog();
  }

  return singleton;
}
