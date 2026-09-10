import { excelRun } from "../excel/helpers.js";
import { createRecoveryScopeResolver } from "./recovery-scope.js";

interface WorkbookSaveBoundaryMonitorDependencies {
  /** Identity checkpoints are currently stored under, or `null` when there is none to clear. */
  resolveWorkbookId: () => Promise<string | null>;
  readWorkbookDirtyState: () => Promise<boolean | null>;
  clearBackupsForCurrentWorkbook: () => Promise<number>;
}

const DEFAULT_POLL_INTERVAL_MS = 4_000;

const defaultScopeResolver = createRecoveryScopeResolver();

async function defaultResolveWorkbookId(): Promise<string | null> {
  const scope = await defaultScopeResolver.resolveForRead();
  return scope?.workbookId ?? null;
}

async function defaultReadWorkbookDirtyState(): Promise<boolean | null> {
  try {
    return await excelRun(async (context) => {
      const workbook = context.workbook;
      workbook.load("isDirty");
      await context.sync();
      return workbook.isDirty;
    });
  } catch {
    // API set may be unavailable on some hosts/builds.
    return null;
  }
}

export class WorkbookSaveBoundaryMonitor {
  private readonly dependencies: WorkbookSaveBoundaryMonitorDependencies;
  private lastDirtyByWorkbookId = new Map<string, boolean>();

  constructor(dependencies: Partial<WorkbookSaveBoundaryMonitorDependencies> = {}) {
    this.dependencies = {
      resolveWorkbookId: dependencies.resolveWorkbookId ?? defaultResolveWorkbookId,
      readWorkbookDirtyState: dependencies.readWorkbookDirtyState ?? defaultReadWorkbookDirtyState,
      clearBackupsForCurrentWorkbook: dependencies.clearBackupsForCurrentWorkbook ?? (() => Promise.resolve(0)),
    };
  }

  async checkOnce(): Promise<void> {
    const workbookId = await this.dependencies.resolveWorkbookId();
    if (!workbookId) return;

    const isDirty = await this.dependencies.readWorkbookDirtyState();
    if (isDirty === null) return;

    const previous = this.lastDirtyByWorkbookId.get(workbookId);
    this.lastDirtyByWorkbookId.set(workbookId, isDirty);

    const savedTransitionObserved = previous === true && isDirty === false;
    const firstObservationIsAlreadySaved = previous === undefined && isDirty === false;

    if (savedTransitionObserved || firstObservationIsAlreadySaved) {
      await this.dependencies.clearBackupsForCurrentWorkbook();
    }
  }
}

export function startWorkbookSaveBoundaryPolling(args: {
  monitor: WorkbookSaveBoundaryMonitor;
  intervalMs?: number;
}): () => void {
  const intervalMs = args.intervalMs ?? DEFAULT_POLL_INTERVAL_MS;

  let stopped = false;
  const tick = () => {
    if (stopped) return;
    void args.monitor.checkOnce().catch(() => {
      // Ignore monitor errors — backup operations remain available manually.
    });
  };

  tick();
  const interval = setInterval(tick, intervalMs);

  return () => {
    if (stopped) return;
    stopped = true;
    clearInterval(interval);
  };
}
