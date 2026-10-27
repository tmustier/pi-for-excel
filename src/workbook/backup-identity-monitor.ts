import {
  createRecoveryScopeResolver,
  UNSAVED_WORKBOOK_ID_PREFIX,
  type RecoveryWorkbookScope,
} from "./recovery-scope.js";

interface BackupIdentityMonitorDependencies {
  /** Identity backups are stored under now, or `null` when there is none. */
  resolveScope: () => Promise<RecoveryWorkbookScope | null>;
  rescopeBackups: (args: { fromWorkbookId: string; toWorkbookId: string; workbookLabel: string }) => Promise<number>;
}

const DEFAULT_POLL_INTERVAL_MS = 4_000;

const defaultScopeResolver = createRecoveryScopeResolver();

/**
 * Keeps a never-saved workbook's backups with it through its first save.
 *
 * Saving does not expire backups (see `applyRetention` in recovery-log.ts):
 * Excel on the web and desktop AutoSave save every few seconds and report
 * the workbook as never dirty, so a "cleared on save" rule left nothing to
 * revert there. What saving can change is the identity backups are filed
 * under: a never-saved workbook's backups are scoped to a document token,
 * and its first save gives it a path identity. This monitor moves them to
 * the new identity at that transition. The rule is limited to token → path
 * transitions: a token only exists in the document-bound Office host, where
 * one taskpane serves one document, so the transition can only be that
 * document being saved. Hosts whose taskpane follows the active workbook
 * (WPS) never see a token, and a plain identity change there means the user
 * switched workbooks.
 */
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
    this.lastObservedWorkbookId = scope.workbookId;

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
