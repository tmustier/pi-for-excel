/**
 * Workbook recovery snapshots.
 *
 * Stores lightweight, local-only checkpoints for workbook cell edits so users
 * can revert mistakes without pre-execution approval prompts.
 */

import { excelRun, getRange } from "../excel/helpers.js";
import { formatWorkbookLabel, getWorkbookContext, type WorkbookContext } from "./context.js";
import { isRecord } from "../utils/type-guards.js";

const RECOVERY_SETTING_KEY = "workbook.recovery-snapshots.v1";
const MAX_RECOVERY_ENTRIES = 120;
export const MAX_RECOVERY_CELLS = 20_000;

export type WorkbookRecoveryToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "restore_snapshot";

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

function isWorkbookRecoveryToolName(value: unknown): value is WorkbookRecoveryToolName {
  return (
    value === "write_cells" ||
    value === "fill_formula" ||
    value === "python_transform_range" ||
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

function parseWorkbookRecoverySnapshot(value: unknown): WorkbookRecoverySnapshot | null {
  if (!isRecord(value)) return null;

  if (!isWorkbookRecoveryToolName(value.toolName)) return null;
  if (typeof value.toolCallId !== "string") return null;
  if (typeof value.address !== "string") return null;
  if (!isGrid(value.beforeValues)) return null;
  if (!isGrid(value.beforeFormulas)) return null;

  const id = typeof value.id === "string" ? value.id : defaultCreateId();
  const at = typeof value.at === "number" ? value.at : Date.now();
  const cellCount = typeof value.cellCount === "number" ? Math.max(0, value.cellCount) : 0;

  return {
    id,
    at,
    toolName: value.toolName,
    toolCallId: value.toolCallId,
    address: value.address,
    changedCount: typeof value.changedCount === "number" ? Math.max(0, value.changedCount) : 0,
    cellCount,
    beforeValues: cloneGrid(value.beforeValues),
    beforeFormulas: cloneGrid(value.beforeFormulas),
    workbookId: typeof value.workbookId === "string" ? value.workbookId : undefined,
    workbookLabel: typeof value.workbookLabel === "string" ? value.workbookLabel : undefined,
    restoredFromSnapshotId: typeof value.restoredFromSnapshotId === "string" ? value.restoredFromSnapshotId : undefined,
  };
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

  private async appendWithContext(
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

    let workbookId: string | undefined;
    let workbookLabel: string | undefined;

    try {
      const workbookContext = workbookContextOverride ?? await this.dependencies.getWorkbookContext();
      if (!workbookContext.workbookId) {
        return null;
      }

      workbookId = workbookContext.workbookId;
      workbookLabel = formatWorkbookLabel(workbookContext);
    } catch {
      return null;
    }

    const snapshot: WorkbookRecoverySnapshot = {
      id: this.dependencies.createId(),
      at: this.dependencies.now(),
      toolName: args.toolName,
      toolCallId: args.toolCallId,
      address: args.address,
      changedCount,
      cellCount: stats.cellCount,
      beforeValues: values,
      beforeFormulas: formulas,
      workbookId,
      workbookLabel,
      restoredFromSnapshotId: args.restoredFromSnapshotId,
    };

    this.snapshots = [snapshot, ...this.snapshots].slice(0, MAX_RECOVERY_ENTRIES);
    await this.persist();

    return snapshot;
  }

  async append(args: AppendWorkbookRecoverySnapshotArgs): Promise<WorkbookRecoverySnapshot | null> {
    await this.ensureLoaded();
    return this.appendWithContext(args);
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

    const restoreValues = toRestoreValues(snapshot.beforeValues, snapshot.beforeFormulas);
    const currentState = await this.dependencies.applySnapshot(snapshot.address, restoreValues);

    const inverseChangedCount = countChangedCells({
      beforeValues: currentState.values,
      beforeFormulas: currentState.formulas,
      afterValues: snapshot.beforeValues,
      afterFormulas: snapshot.beforeFormulas,
    });

    const inverseSnapshot = await this.appendWithContext(
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
