/**
 * Recovery backups overlay.
 */

import { formatRelativeDate } from "./overlay-relative-date.js";
import {
  applyRecoveryFilters,
  buildToolFilterOptions,
  DEFAULT_FILTER_STATE,
  type RecoveryFilterState,
  type RecoverySortOrder,
  type RecoveryToolFilter,
} from "./recovery-filtering.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { RECOVERY_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";

export type RecoveryCheckpointToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "format_cells"
  | "conditional_format"
  | "comments"
  | "modify_structure"
  | "restore_snapshot";

export interface RecoveryCheckpointSummary {
  id: string;
  at: number;
  toolName: RecoveryCheckpointToolName;
  address: string;
  changedCount: number;
  restoredFromSnapshotId?: string;
}

function formatRecoveryToolLabel(toolName: RecoveryCheckpointToolName): string {
  switch (toolName) {
    case "write_cells":
      return "Write";
    case "fill_formula":
      return "Fill formula";
    case "python_transform_range":
      return "Python transform";
    case "format_cells":
      return "Format cells";
    case "conditional_format":
      return "Conditional format";
    case "comments":
      return "Comments";
    case "modify_structure":
      return "Modify structure";
    case "restore_snapshot":
      return "Restore";
    default:
      return toolName;
  }
}

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return "0 B";
  }

  if (bytes < 1024) {
    return `${Math.floor(bytes)} B`;
  }

  const units = ["KB", "MB", "GB"];
  let value = bytes / 1024;
  let unitIndex = 0;

  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[unitIndex]}`;
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

function exportCheckpointsAsJson(
  checkpoints: RecoveryCheckpointSummary[],
  workbookLabel: string,
): void {
  const payload = {
    exported: new Date().toISOString(),
    workbook: workbookLabel,
    count: checkpoints.length,
    checkpoints,
  };

  const json = JSON.stringify(payload, null, 2);
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  const opened = window.open(url, "_blank");
  if (!opened) {
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `pi-backups-${new Date().toISOString().slice(0, 10)}.json`;
    anchor.rel = "noopener";
    anchor.hidden = true;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }

  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

// ---------------------------------------------------------------------------
// Retention
// ---------------------------------------------------------------------------

export interface RetentionConfig {
  maxSnapshots: number;
}

export interface ManualFullBackupSummary {
  id: string;
  sizeBytes: number;
}

// ---------------------------------------------------------------------------
// Overlay
// ---------------------------------------------------------------------------

export async function showRecoveryDialog(opts: {
  workbookLabel: string;
  loadCheckpoints: () => Promise<RecoveryCheckpointSummary[]>;
  onRestore: (snapshotId: string) => Promise<void>;
  onDelete: (snapshotId: string) => Promise<boolean>;
  onClear: () => Promise<number>;
  onCreateManualFullBackup?: () => Promise<ManualFullBackupSummary>;
  getRetentionConfig?: () => Promise<RetentionConfig>;
  setRetentionConfig?: (config: RetentionConfig) => Promise<void>;
}): Promise<void> {
  if (closeOverlayById(RECOVERY_OVERLAY_ID)) {
    return;
  }

  const dialog = createOverlayDialog({
    overlayId: RECOVERY_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m pi-recovery-dialog",
  });

  const { header } = createOverlayHeader({
    onClose: dialog.close,
    closeLabel: "Close backups",
    title: "Backups (Beta)",
    subtitle: "Saved before Pi edits, in between saves. Entries are sheet-specific in this workbook.",
  });

  const workbookTag = document.createElement("p");
  workbookTag.className = "pi-overlay-workbook-tag";
  workbookTag.textContent = `Workbook: ${opts.workbookLabel}`;

  const saveBoundaryHint = document.createElement("p");
  saveBoundaryHint.className = "pi-overlay-hint";
  saveBoundaryHint.textContent = "Backups reset after you save this workbook.";

  // -- Search + filters --

  const searchRow = document.createElement("div");
  searchRow.className = "pi-recovery-search-row";

  const searchInput = document.createElement("input");
  searchInput.type = "text";
  searchInput.placeholder = "Search by id, tool, or range…";
  searchInput.className = "pi-recovery-search pi-overlay-inline-control";

  const toolFilterSelect = document.createElement("select");
  toolFilterSelect.className = "pi-recovery-filter-select pi-overlay-inline-control";

  const sortButton = document.createElement("button");
  sortButton.type = "button";
  sortButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-recovery-sort-btn";
  sortButton.textContent = "↓ Newest";

  searchRow.append(searchInput, toolFilterSelect, sortButton);

  // -- Toolbar --

  const toolbar = document.createElement("div");
  toolbar.className = "pi-recovery-toolbar";

  const fullBackupButton = document.createElement("button");
  fullBackupButton.type = "button";
  fullBackupButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
  fullBackupButton.textContent = "Full backup";
  fullBackupButton.hidden = opts.onCreateManualFullBackup === undefined;

  const refreshButton = document.createElement("button");
  refreshButton.type = "button";
  refreshButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
  refreshButton.textContent = "Refresh";

  const exportButton = document.createElement("button");
  exportButton.type = "button";
  exportButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
  exportButton.textContent = "Export";

  const clearButton = document.createElement("button");
  clearButton.type = "button";
  clearButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
  clearButton.textContent = "Clear all";

  const statusText = document.createElement("span");
  statusText.className = "pi-recovery-status";

  toolbar.append(fullBackupButton, refreshButton, exportButton, clearButton, statusText);

  // -- Retention --

  const retentionRow = document.createElement("div");
  retentionRow.className = "pi-recovery-retention";

  const retentionLabel = document.createElement("label");
  retentionLabel.className = "pi-recovery-retention__label";
  retentionLabel.textContent = "Keep at most";

  const retentionInput = document.createElement("input");
  retentionInput.type = "number";
  retentionInput.min = "5";
  retentionInput.max = "120";
  retentionInput.className = "pi-recovery-retention__input pi-overlay-inline-control";

  const retentionSuffix = document.createElement("span");
  retentionSuffix.className = "pi-recovery-retention__suffix";
  retentionSuffix.textContent = "backups";

  const retentionSave = document.createElement("button");
  retentionSave.type = "button";
  retentionSave.className = "pi-overlay-btn pi-overlay-btn--ghost";
  retentionSave.textContent = "Save";

  retentionRow.append(retentionLabel, retentionInput, retentionSuffix, retentionSave);

  const hasRetention = opts.getRetentionConfig !== undefined && opts.setRetentionConfig !== undefined;
  retentionRow.hidden = !hasRetention;

  // -- List --

  const list = document.createElement("div");
  list.className = "pi-recovery-list";

  // -- Assemble --

  dialog.card.append(
    header, workbookTag, saveBoundaryHint,
    searchRow, toolbar, retentionRow, list,
  );

  // -- State --

  let allCheckpoints: RecoveryCheckpointSummary[] = [];
  let busy = false;
  const filterState: RecoveryFilterState = { ...DEFAULT_FILTER_STATE };

  const formatChangedLabel = (changedCount: number): string =>
    `${changedCount.toLocaleString()} change${changedCount === 1 ? "" : "s"}`;

  const shortId = (id: string): string => (id.length > 12 ? id.slice(0, 12) : id);

  const setBusy = (next: boolean): void => {
    busy = next;
    fullBackupButton.disabled = next || opts.onCreateManualFullBackup === undefined;
    refreshButton.disabled = next;
    exportButton.disabled = next || allCheckpoints.length === 0;
    clearButton.disabled = next || allCheckpoints.length === 0;
    searchInput.disabled = next;
    toolFilterSelect.disabled = next;
    sortButton.disabled = next;
    retentionSave.disabled = next;

    for (const button of list.querySelectorAll<HTMLButtonElement>("button")) {
      button.disabled = next;
    }
  };

  const syncFilterControls = (): void => {
    const options = buildToolFilterOptions(allCheckpoints);
    toolFilterSelect.replaceChildren();
    for (const opt of options) {
      const el = document.createElement("option");
      el.value = opt.value;
      el.textContent = `${opt.label} (${opt.count})`;
      el.selected = opt.value === filterState.toolFilter;
      toolFilterSelect.appendChild(el);
    }

    sortButton.textContent = filterState.sortOrder === "newest" ? "↓ Newest" : "↑ Oldest";
  };

  const renderList = (): void => {
    const filtered = applyRecoveryFilters(allCheckpoints, filterState);
    syncFilterControls();

    list.replaceChildren();

    if (allCheckpoints.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-overlay-empty";
      empty.textContent = "No backups for this workbook yet.";
      list.appendChild(empty);
      statusText.textContent = "No backups";
      clearButton.disabled = true;
      exportButton.disabled = true;
      return;
    }

    if (filtered.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-overlay-empty";
      empty.textContent = "No backups match the current filters.";
      list.appendChild(empty);
      statusText.textContent = `0 of ${allCheckpoints.length} shown`;
      return;
    }

    for (const checkpoint of filtered) {
      const item = document.createElement("div");
      item.className = "pi-overlay-surface pi-recovery-item";

      const itemHeader = document.createElement("div");
      itemHeader.className = "pi-recovery-item__header";

      const titleEl = document.createElement("span");
      titleEl.className = "pi-recovery-item__title";
      titleEl.textContent = `${formatRecoveryToolLabel(checkpoint.toolName)} · ${checkpoint.address}`;

      const timeEl = document.createElement("span");
      timeEl.className = "pi-recovery-item__time";
      timeEl.textContent = formatRelativeDate(new Date(checkpoint.at).toISOString());

      itemHeader.append(titleEl, timeEl);

      const meta = document.createElement("div");
      meta.className = "pi-recovery-item__meta";
      meta.textContent = `${formatChangedLabel(checkpoint.changedCount)} · #${shortId(checkpoint.id)}`;

      const actions = document.createElement("div");
      actions.className = "pi-recovery-item__actions";

      const restoreButton = document.createElement("button");
      restoreButton.type = "button";
      restoreButton.className = "pi-overlay-btn pi-overlay-btn--primary";
      restoreButton.textContent = "Restore";

      const deleteButton = document.createElement("button");
      deleteButton.type = "button";
      deleteButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
      deleteButton.textContent = "Delete";

      restoreButton.addEventListener("click", () => {
        if (busy) return;

        void (async () => {
          setBusy(true);
          statusText.textContent = "Restoring…";

          try {
            await opts.onRestore(checkpoint.id);
            allCheckpoints = await opts.loadCheckpoints();
            renderList();
          } catch (error: unknown) {
            const message = error instanceof Error ? error.message : "Unknown error";
            showToast(`Restore failed: ${message}`);
            statusText.textContent = "Restore failed";
          } finally {
            setBusy(false);
          }
        })();
      });

      deleteButton.addEventListener("click", () => {
        if (busy) return;

        const proceed = window.confirm("Delete this backup?");
        if (!proceed) return;

        void (async () => {
          setBusy(true);
          statusText.textContent = "Deleting…";

          try {
            const deleted = await opts.onDelete(checkpoint.id);
            if (!deleted) {
              showToast("Backup not found");
            }

            allCheckpoints = await opts.loadCheckpoints();
            renderList();
          } catch (error: unknown) {
            const message = error instanceof Error ? error.message : "Unknown error";
            showToast(`Delete failed: ${message}`);
            statusText.textContent = "Delete failed";
          } finally {
            setBusy(false);
          }
        })();
      });

      actions.append(restoreButton, deleteButton);
      item.append(itemHeader, meta, actions);

      if (checkpoint.restoredFromSnapshotId) {
        const restoredMeta = document.createElement("div");
        restoredMeta.className = "pi-recovery-item__restored";
        restoredMeta.textContent = `Restored from #${shortId(checkpoint.restoredFromSnapshotId)}`;
        item.appendChild(restoredMeta);
      }

      list.appendChild(item);
    }

    if (filtered.length < allCheckpoints.length) {
      statusText.textContent = `${filtered.length} of ${allCheckpoints.length} shown`;
    } else {
      statusText.textContent = `${allCheckpoints.length} backup${allCheckpoints.length === 1 ? "" : "s"}`;
    }

    clearButton.disabled = busy || allCheckpoints.length === 0;
    exportButton.disabled = busy || allCheckpoints.length === 0;
  };

  const reload = async (): Promise<void> => {
    allCheckpoints = await opts.loadCheckpoints();
    renderList();
  };

  // -- Event listeners --

  let searchTimer: ReturnType<typeof setTimeout> | null = null;

  searchInput.addEventListener("input", () => {
    if (searchTimer !== null) clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      filterState.search = searchInput.value;
      renderList();
    }, 200);
  });

  toolFilterSelect.addEventListener("change", () => {
    filterState.toolFilter = toolFilterSelect.value as RecoveryToolFilter;
    renderList();
  });

  sortButton.addEventListener("click", () => {
    const next: RecoverySortOrder = filterState.sortOrder === "newest" ? "oldest" : "newest";
    filterState.sortOrder = next;
    renderList();
  });

  fullBackupButton.addEventListener("click", () => {
    if (busy) return;

    const createManualFullBackup = opts.onCreateManualFullBackup;
    if (!createManualFullBackup) return;

    void (async () => {
      setBusy(true);
      statusText.textContent = "Capturing full backup…";
      try {
        const backup = await createManualFullBackup();
        showToast(
          `Full backup downloaded: #${shortId(backup.id)} (${formatBytes(backup.sizeBytes)}). Open it in Excel to restore.`,
        );
        renderList();
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(`Full backup failed: ${message}`);
        statusText.textContent = "Full backup failed";
      } finally {
        setBusy(false);
      }
    })();
  });

  exportButton.addEventListener("click", () => {
    if (busy || allCheckpoints.length === 0) return;
    exportCheckpointsAsJson(allCheckpoints, opts.workbookLabel);
    showToast(`Exported ${allCheckpoints.length} backup${allCheckpoints.length === 1 ? "" : "s"}`);
  });

  refreshButton.addEventListener("click", () => {
    if (busy) return;

    void (async () => {
      setBusy(true);
      statusText.textContent = "Refreshing…";
      try {
        await reload();
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(`Refresh failed: ${message}`);
        statusText.textContent = "Refresh failed";
      } finally {
        setBusy(false);
      }
    })();
  });

  clearButton.addEventListener("click", () => {
    if (busy || allCheckpoints.length === 0) return;

    const proceed = window.confirm(`Delete all ${allCheckpoints.length} backups for this workbook?`);
    if (!proceed) return;

    void (async () => {
      setBusy(true);
      statusText.textContent = "Clearing…";
      try {
        const removed = await opts.onClear();
        showToast(`Cleared ${removed} backup${removed === 1 ? "" : "s"}`);
        await reload();
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(`Clear failed: ${message}`);
        statusText.textContent = "Clear failed";
      } finally {
        setBusy(false);
      }
    })();
  });

  retentionSave.addEventListener("click", () => {
    if (busy) return;

    const setConfig = opts.setRetentionConfig;
    if (!setConfig) return;

    const value = parseInt(retentionInput.value, 10);
    if (!Number.isFinite(value) || value < 5 || value > 120) {
      showToast("Retention limit must be between 5 and 120");
      return;
    }

    void (async () => {
      setBusy(true);
      try {
        await setConfig({ maxSnapshots: value });
        showToast(`Retention set to ${value} backups`);
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(`Failed to save retention: ${message}`);
      } finally {
        setBusy(false);
      }
    })();
  });

  // -- Cleanup --

  dialog.addCleanup(() => {
    if (searchTimer !== null) clearTimeout(searchTimer);
  });

  // -- Mount + initial load --

  dialog.mount();

  setBusy(true);
  statusText.textContent = "Loading…";
  try {
    if (opts.getRetentionConfig) {
      try {
        const config = await opts.getRetentionConfig();
        retentionInput.value = String(config.maxSnapshots);
      } catch {
        retentionInput.value = "120";
      }
    }

    await reload();
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : "Unknown error";
    showToast(`Failed to load backups: ${message}`);
    statusText.textContent = "Load failed";
  } finally {
    setBusy(false);
  }
}
