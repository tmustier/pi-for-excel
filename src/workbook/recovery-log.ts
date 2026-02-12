/**
 * Workbook recovery snapshots.
 *
 * Stores lightweight, local-only checkpoints for workbook cell edits so users
 * can revert mistakes without pre-execution approval prompts.
 */

import { excelRun, getRange } from "../excel/helpers.js";
import { formatWorkbookLabel, getWorkbookContext, type WorkbookContext } from "./context.js";
import { isRecord } from "../utils/type-guards.js";
import {
  applyCommentThreadState,
  applyConditionalFormatState,
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  type RecoveryCommentThreadState,
  type RecoveryConditionalFormatCaptureResult,
  type RecoveryConditionalFormatRule,
} from "./recovery-states.js";

const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";
const MAX_RECOVERY_ENTRIES = 120;
export const MAX_RECOVERY_CELLS = 20_000;

export type WorkbookRecoveryToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "conditional_format"
  | "comments"
  | "restore_snapshot";

export type WorkbookRecoverySnapshotKind =
  | "range_values"
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

interface SettingsStoreLike {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
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
  applyConditionalFormatSnapshot: (
    address: string,
    rules: RecoveryConditionalFormatRule[],
  ) => Promise<RecoveryConditionalFormatCaptureResult>;
  applyCommentThreadSnapshot: (
    address: string,
    state: RecoveryCommentThreadState,
  ) => Promise<RecoveryCommentThreadState>;
}

interface PersistedWorkbookRecoveryPayload {
  version: 1;
  snapshots: WorkbookRecoverySnapshot[];
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

function isSettingsStoreLike(value: unknown): value is SettingsStoreLike {
  if (!isRecord(value)) return false;

  return (
    typeof value.get === "function" &&
    typeof value.set === "function"
  );
}

async function defaultGetSettingsStore(): Promise<SettingsStoreLike | null> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const appStorage = storageModule.getAppStorage();
    const settings = isRecord(appStorage) ? appStorage.settings : null;
    return isSettingsStoreLike(settings) ? settings : null;
  } catch {
    return null;
  }
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

function isWorkbookRecoveryToolName(value: unknown): value is WorkbookRecoveryToolName {
  return (
    value === "write_cells" ||
    value === "fill_formula" ||
    value === "python_transform_range" ||
    value === "conditional_format" ||
    value === "comments" ||
    value === "restore_snapshot"
  );
}

function isGrid(value: unknown): value is unknown[][] {
  return Array.isArray(value) && value.every((row) => Array.isArray(row));
}

function rowLength(grid: unknown[][], row: number): number {
  const rowValues = grid[row];
  return Array.isArray(rowValues) ? rowValues.length : 0;
}

function valueAt(grid: unknown[][], row: number, col: number): unknown {
  const rowValues = grid[row];
  if (!Array.isArray(rowValues)) return "";
  return col < rowValues.length ? rowValues[col] : "";
}

function cloneGrid(grid: unknown[][]): unknown[][] {
  return grid.map((row) => {
    if (!Array.isArray(row)) {
      return [];
    }

    return [...row];
  });
}

function gridStats(values: unknown[][], formulas: unknown[][]): {
  rows: number;
  cols: number;
  cellCount: number;
} {
  const rows = Math.max(values.length, formulas.length);
  let cols = 0;

  for (let row = 0; row < rows; row += 1) {
    cols = Math.max(cols, rowLength(values, row), rowLength(formulas, row));
  }

  return {
    rows,
    cols,
    cellCount: rows * cols,
  };
}

function normalizeFormula(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;

  const trimmed = raw.trim();
  if (!trimmed.startsWith("=")) return undefined;
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

function parseWorkbookRecoverySnapshotKind(value: unknown): WorkbookRecoverySnapshotKind {
  return value === "conditional_format_rules" || value === "comment_thread" || value === "range_values"
    ? value
    : "range_values";
}

function isRecoveryConditionalFormatRule(value: unknown): value is RecoveryConditionalFormatRule {
  if (!isRecord(value)) return false;

  const type = value.type;
  if (type !== "custom" && type !== "cell_value") return false;

  const operator = value.operator;
  const validOperator = operator === undefined || (
    operator === "Between" ||
    operator === "NotBetween" ||
    operator === "EqualTo" ||
    operator === "NotEqualTo" ||
    operator === "GreaterThan" ||
    operator === "LessThan" ||
    operator === "GreaterThanOrEqual" ||
    operator === "LessThanOrEqual"
  );

  if (!validOperator) return false;

  return (
    (value.stopIfTrue === undefined || typeof value.stopIfTrue === "boolean") &&
    (value.formula === undefined || typeof value.formula === "string") &&
    (value.formula1 === undefined || typeof value.formula1 === "string") &&
    (value.formula2 === undefined || typeof value.formula2 === "string") &&
    (value.fillColor === undefined || typeof value.fillColor === "string") &&
    (value.fontColor === undefined || typeof value.fontColor === "string") &&
    (value.bold === undefined || typeof value.bold === "boolean") &&
    (value.italic === undefined || typeof value.italic === "boolean") &&
    (value.underline === undefined || typeof value.underline === "boolean") &&
    (value.appliesToAddress === undefined || typeof value.appliesToAddress === "string")
  );
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

  const commentThreadState = isRecoveryCommentThreadState(value.commentThreadState)
    ? cloneRecoveryCommentThreadState(value.commentThreadState)
    : undefined;

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

  if (snapshotKind === "conditional_format_rules") {
    snapshot.conditionalFormatRules = cloneRecoveryConditionalFormatRules(conditionalFormatRules);
  }

  if (snapshotKind === "comment_thread") {
    snapshot.commentThreadState = commentThreadState;
  }

  return snapshot;
}

function parsePersistedSnapshots(payload: unknown): WorkbookRecoverySnapshot[] {
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

  return snapshots
    .sort((a, b) => b.at - a.at)
    .slice(0, MAX_RECOVERY_ENTRIES);
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
    if (!settings) return;

    try {
      const payload = await settings.get<unknown>(RECOVERY_SETTING_KEY);
      this.snapshots = parsePersistedSnapshots(payload);
    } catch {
      this.snapshots = [];
    }
  }

  private async persist(): Promise<void> {
    const settings = await this.dependencies.getSettingsStore();
    if (!settings) return;

    const payload: PersistedWorkbookRecoveryPayload = {
      version: 1,
      snapshots: this.snapshots,
    };

    try {
      await settings.set(RECOVERY_SETTING_KEY, payload);
    } catch {
      // ignore persistence failures
    }
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
    if (!snapshot.workbookId) {
      throw new Error("Snapshot is missing workbook identity and cannot be restored safely.");
    }

    if (!workbookContext.workbookId) {
      throw new Error("Current workbook identity is unavailable; cannot safely restore this snapshot.");
    }

    if (snapshot.workbookId !== workbookContext.workbookId) {
      throw new Error("Snapshot belongs to a different workbook.");
    }

    const snapshotKind = snapshot.snapshotKind ?? "range_values";

    if (snapshotKind === "conditional_format_rules") {
      const rules = snapshot.conditionalFormatRules ?? [];
      const currentState = await this.dependencies.applyConditionalFormatSnapshot(snapshot.address, rules);

      if (!currentState.supported) {
        throw new Error(currentState.reason ?? "Conditional format checkpoint cannot be restored safely.");
      }

      const inverseSnapshot = await this.appendConditionalFormatWithContext(
        {
          toolName: "restore_snapshot",
          toolCallId: `restore:${snapshot.id}`,
          address: snapshot.address,
          changedCount: snapshot.changedCount,
          cellCount: snapshot.cellCount,
          conditionalFormatRules: currentState.rules,
          restoredFromSnapshotId: snapshot.id,
        },
        workbookContext,
      );

      return {
        restoredSnapshotId: snapshot.id,
        inverseSnapshotId: inverseSnapshot?.id ?? null,
        address: snapshot.address,
        changedCount: snapshot.changedCount,
      };
    }

    if (snapshotKind === "comment_thread") {
      const targetState = snapshot.commentThreadState;
      if (!targetState) {
        throw new Error("Comment checkpoint data is missing.");
      }

      const currentState = await this.dependencies.applyCommentThreadSnapshot(snapshot.address, targetState);
      const inverseSnapshot = await this.appendCommentThreadWithContext(
        {
          toolName: "restore_snapshot",
          toolCallId: `restore:${snapshot.id}`,
          address: snapshot.address,
          changedCount: snapshot.changedCount,
          commentThreadState: currentState,
          restoredFromSnapshotId: snapshot.id,
        },
        workbookContext,
      );

      return {
        restoredSnapshotId: snapshot.id,
        inverseSnapshotId: inverseSnapshot?.id ?? null,
        address: snapshot.address,
        changedCount: snapshot.changedCount,
      };
    }

    const restoreValues = toRestoreValues(snapshot.beforeValues, snapshot.beforeFormulas);
    const currentState = await this.dependencies.applySnapshot(snapshot.address, restoreValues);

    const inverseChangedCount = countChangedCells({
      beforeValues: currentState.values,
      beforeFormulas: currentState.formulas,
      afterValues: snapshot.beforeValues,
      afterFormulas: snapshot.beforeFormulas,
    });

    const inverseSnapshot = await this.appendRangeWithContext(
      {
        toolName: "restore_snapshot",
        toolCallId: `restore:${snapshot.id}`,
        address: snapshot.address,
        changedCount: inverseChangedCount,
        beforeValues: currentState.values,
        beforeFormulas: currentState.formulas,
        restoredFromSnapshotId: snapshot.id,
      },
      workbookContext,
    );

    return {
      restoredSnapshotId: snapshot.id,
      inverseSnapshotId: inverseSnapshot?.id ?? null,
      address: snapshot.address,
      changedCount: inverseChangedCount,
    };
  }

  async restoreLatestForCurrentWorkbook(): Promise<RestoreWorkbookRecoverySnapshotResult> {
    const snapshots = await this.listForCurrentWorkbook(1);
    const latest = snapshots[0];
    if (!latest) {
      throw new Error("No recovery checkpoints found for this workbook.");
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
