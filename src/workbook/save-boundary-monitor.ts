import { excelRun } from "../excel/helpers.js";
import { createRecoveryScopeResolver, UNSAVED_WORKBOOK_ID_PREFIX } from "./recovery-scope.js";

interface WorkbookSaveBoundaryMonitorDependencies {
  /** Identity checkpoints are currently stored under, or `null` when there is none to clear. */
  resolveWorkbookId: () => Promise<string | null>;
  readWorkbookDirtyState: () => Promise<boolean | null>;
  clearBackupsForWorkbook: (workbookId: string) => Promise<number>;
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

/**
 * Clears checkpoints when the workbook crosses a save boundary (dirty → saved).
 *
 * A never-saved workbook's checkpoints are scoped to a document token. Its
 * first save gives the document a path identity, and that save also covers
 * everything captured under the token, so those checkpoints are cleared too;
 * otherwise they would linger, invisible, under an identity nothing resolves
 * to any more. The rule is limited to token → path transitions: a token only
 * exists in the document-bound Office host, where one taskpane serves one
 * document, so the transition can only be that document being saved. Hosts
 * whose taskpane follows the active workbook (WPS) never see a token, and a
 * plain identity change there means the user switched workbooks.
 */
export class WorkbookSaveBoundaryMonitor {
  private readonly dependencies: WorkbookSaveBoundaryMonitorDependencies;
  private lastDirtyByWorkbookId = new Map<string, boolean>();
  private lastObservedWorkbookId: string | null = null;

  constructor(dependencies: Partial<WorkbookSaveBoundaryMonitorDependencies> = {}) {
    this.dependencies = {
      resolveWorkbookId: dependencies.resolveWorkbookId ?? defaultResolveWorkbookId,
      readWorkbookDirtyState: dependencies.readWorkbookDirtyState ?? defaultReadWorkbookDirtyState,
      clearBackupsForWorkbook: dependencies.clearBackupsForWorkbook ?? (() => Promise.resolve(0)),
    };
  }

  async checkOnce(): Promise<void> {
    const workbookId = await this.dependencies.resolveWorkbookId();
    if (!workbookId) return;

    const isDirty = await this.dependencies.readWorkbookDirtyState();
    if (isDirty === null) return;

    const previousIdentity = this.lastObservedWorkbookId;
    this.lastObservedWorkbookId = workbookId;

    const previous = this.lastDirtyByWorkbookId.get(workbookId);
    this.lastDirtyByWorkbookId.set(workbookId, isDirty);

    const savedTransitionObserved = previous === true && isDirty === false;
    const firstObservationIsAlreadySaved = previous === undefined && isDirty === false;

    if (savedTransitionObserved || firstObservationIsAlreadySaved) {
      await this.dependencies.clearBackupsForWorkbook(workbookId);
    }

    const firstSaveOfNewWorkbook =
      previousIdentity !== null
      && previousIdentity !== workbookId
      && previousIdentity.startsWith(UNSAVED_WORKBOOK_ID_PREFIX)
      && !workbookId.startsWith(UNSAVED_WORKBOOK_ID_PREFIX);
    if (firstSaveOfNewWorkbook) {
      this.lastDirtyByWorkbookId.delete(previousIdentity);
      await this.dependencies.clearBackupsForWorkbook(previousIdentity);
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
