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

  constructor(dependencies: Pick<BackupIdentityMonitorDependencies, "rescopeBackups"> & Partial<BackupIdentityMonitorDependencies>) {
    this.dependencies = {
      resolveScope: dependencies.resolveScope ?? (() => defaultScopeResolver.resolveForRead()),
      rescopeBackups: dependencies.rescopeBackups,
    };
  }

  async checkOnce(): Promise<void> {
    const scope = await this.dependencies.resolveScope();
    if (!scope) return;

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
    this.lastObservedWorkbookId = scope.workbookId;
  }
}

export function startBackupIdentityPolling(args: {
  monitor: BackupIdentityMonitor;
  intervalMs?: number;
}): () => void {
  const intervalMs = args.intervalMs ?? DEFAULT_POLL_INTERVAL_MS;

  let stopped = false;
  const tick = () => {
    if (stopped) return;
    void args.monitor.checkOnce().catch((error) => {
      console.warn("[pi] Backup identity check failed:", error);
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
