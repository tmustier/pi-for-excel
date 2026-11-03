import {
  createRecoveryScopeResolver,
  UNSAVED_WORKBOOK_ID_PREFIX,
  type RecoveryWorkbookScope,
} from "./recovery-scope.js";

interface BackupIdentityMonitorDependencies {
  resolveScope: () => Promise<RecoveryWorkbookScope | null>;
  rescopeBackups: (args: { fromWorkbookId: string; toWorkbookId: string; workbookLabel: string }) => Promise<number>;
}

const DEFAULT_POLL_INTERVAL_MS = 4_000;

const defaultScopeResolver = createRecoveryScopeResolver();

export class BackupIdentityMonitor {
  private readonly dependencies: BackupIdentityMonitorDependencies;
  private lastObservedWorkbookId: string | null = null;
  private checkInFlight: Promise<void> | null = null;

  constructor(dependencies: Pick<BackupIdentityMonitorDependencies, "rescopeBackups"> & Partial<BackupIdentityMonitorDependencies>) {
    this.dependencies = {
      resolveScope: dependencies.resolveScope ?? (() => defaultScopeResolver.resolveForRead()),
      rescopeBackups: dependencies.rescopeBackups,
    };
  }

  checkOnce(signal?: AbortSignal): Promise<void> {
    if (signal?.aborted) return Promise.resolve();
    if (this.checkInFlight) return this.checkInFlight;
    this.checkInFlight = this.check(signal).finally(() => {
      this.checkInFlight = null;
    });
    return this.checkInFlight;
  }

  private async check(signal?: AbortSignal): Promise<void> {
    const scope = await this.dependencies.resolveScope();
    if (!scope || signal?.aborted) return;

    const previous = this.lastObservedWorkbookId;

    const firstSaveOfNewWorkbook =
      previous !== null
      && previous !== scope.workbookId
      && previous.startsWith(UNSAVED_WORKBOOK_ID_PREFIX)
      && !scope.unsaved;
    if (firstSaveOfNewWorkbook) {
      await this.dependencies.rescopeBackups({
        fromWorkbookId: previous,
        toWorkbookId: scope.workbookId,
        workbookLabel: scope.workbookLabel,
      });
    }
    if (!signal?.aborted) this.lastObservedWorkbookId = scope.workbookId;
  }
}

export function startBackupIdentityPolling(args: {
  monitor: BackupIdentityMonitor;
  intervalMs?: number;
}): () => void {
  const intervalMs = args.intervalMs ?? DEFAULT_POLL_INTERVAL_MS;

  const controller = new AbortController();
  const tick = () => {
    if (controller.signal.aborted) return;
    void args.monitor.checkOnce(controller.signal).catch((error) => {
      console.warn("[pi] Backup identity check failed:", error);
    });
  };

  tick();
  const interval = setInterval(tick, intervalMs);

  return () => {
    if (controller.signal.aborted) return;
    controller.abort();
    clearInterval(interval);
  };
}
