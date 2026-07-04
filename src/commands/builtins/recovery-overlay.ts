/**
 * Recovery backups overlay.
 *
 * Progressive-disclosure empty state: when there are no backups, only the
 * title, subtitle, warning callout, and empty message are shown. Search,
 * filter, toolbar, and retention controls appear once backups exist.
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
import { requestConfirmationDialog } from "../../ui/confirm-dialog.js";
import { RECOVERY_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import {
  createCallout,
  createEmptyInline,
  createButton,
  createActionsRow,
} from "../../ui/extensions-hub-components.js";
import { lucide, AlertTriangle, Package, Search } from "../../ui/lucide-icons.js";
import {
  MAX_RECOVERY_ENTRIES,
  MIN_RETENTION_LIMIT,
} from "../../workbook/recovery/constants.js";

export type RecoveryCheckpointToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "format_cells"
  | "conditional_format"
  | "comments"
  | "charts"
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
    case "charts":
      return "Charts";
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
    title: "Backups",
    subtitle: "Snapshots saved before Pi changes your data",
  });

  // -- Warning callout --

  const warningCallout = createCallout(
    "warn",
    lucide(AlertTriangle),
    "Backups clear when you save this workbook in Excel.",
    { compact: true },
  );

  // -- Search + filters (hidden when empty) --

  const searchRow = document.createElement("div");
  searchRow.className = "pi-recovery-search-row pi-overlay-inline-row pi-overlay-inline-row--compact pi-overlay-inline-row--wrap";

  const searchInput = document.createElement("input");
  searchInput.type = "text";
  searchInput.placeholder = "Search backups…";
  searchInput.className = "pi-recovery-search pi-overlay-inline-control";

  const toolFilterSelect = document.createElement("select");
  toolFilterSelect.className = "pi-recovery-filter-select pi-overlay-inline-control";

  const sortButton = document.createElement("button");
  sortButton.type = "button";
  sortButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-recovery-sort-btn";
  sortButton.textContent = "↓ Newest";

  searchRow.append(searchInput, toolFilterSelect, sortButton);

  // -- Toolbar (hidden when empty) --

  const toolbar = document.createElement("div");
  toolbar.className = "pi-overlay-toolbar";

  const toolbarActions = document.createElement("div");
  toolbarActions.className = "pi-overlay-toolbar-actions";

  const downloadBackupBtn = createButton("Download backup", {
    primary: true,
    compact: true,
    onClick: () => {
      if (busy) return;
      const createManualFullBackup = opts.onCreateManualFullBackup;
      if (!createManualFullBackup) return;
      void (async () => {
        setBusy(true);
        statusText.textContent = "Capturing…";
        try {
          const backup = await createManualFullBackup();
          showToast(`Backup downloaded: #${shortId(backup.id)} (${formatBytes(backup.sizeBytes)})`);
          renderList();
        } catch (error: unknown) {
          const message = error instanceof Error ? error.message : "Unknown error";
          showToast(`Backup failed: ${message}`);
          statusText.textContent = "Backup failed";
        } finally {
          setBusy(false);
        }
      })();
    },
  });
  downloadBackupBtn.hidden = opts.onCreateManualFullBackup === undefined;

  const refreshButton = createButton("Refresh", {
    compact: true,
    onClick: () => {
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
    },
  });

  const clearButton = createButton("Clear all", {
    danger: true,
    compact: true,
    onClick: () => {
      if (busy || allCheckpoints.length === 0) return;
      void (async () => {
        const proceed = await requestConfirmationDialog({
          title: "Delete all backups for this workbook?",
          message: `This will delete ${allCheckpoints.length} backup${allCheckpoints.length === 1 ? "" : "s"}.`,
          confirmLabel: "Delete all",
          cancelLabel: "Cancel",
          confirmButtonTone: "danger",
          restoreFocusOnClose: false,
        });
        if (!proceed || busy) return;
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
    },
  });

  toolbarActions.append(downloadBackupBtn, refreshButton, clearButton);

  const statusText = document.createElement("span");
  statusText.className = "pi-overlay-toolbar-status";

  toolbar.append(toolbarActions, statusText);

  // -- Retention (hidden by default, collapsible) --

  const hasRetention = opts.getRetentionConfig !== undefined && opts.setRetentionConfig !== undefined;
  const retentionDetails = document.createElement("details");
  retentionDetails.className = "pi-recovery-retention-details";
  retentionDetails.hidden = !hasRetention;

  const retentionSummary = document.createElement("summary");
  retentionSummary.className = "pi-recovery-retention-summary";
  retentionSummary.textContent = "Retention settings";
  retentionDetails.appendChild(retentionSummary);

  const retentionRow = document.createElement("div");
  retentionRow.className = "pi-recovery-retention pi-overlay-inline-row pi-overlay-inline-row--compact";

  const retentionLabel = document.createElement("label");
  retentionLabel.className = "pi-recovery-retention__label";
  retentionLabel.textContent = "Keep at most";

  const retentionInput = document.createElement("input");
  retentionInput.type = "number";
  retentionInput.min = String(MIN_RETENTION_LIMIT);
  retentionInput.max = String(MAX_RECOVERY_ENTRIES);
  retentionInput.className = "pi-recovery-retention__input pi-overlay-inline-control";

  const retentionSuffix = document.createElement("span");
  retentionSuffix.className = "pi-recovery-retention__suffix";
  retentionSuffix.textContent = "backups";

  const retentionSave = createButton("Save", {
    compact: true,
    onClick: () => {
      if (busy) return;
      const setConfig = opts.setRetentionConfig;
      if (!setConfig) return;
      const value = parseInt(retentionInput.value, 10);
      if (!Number.isFinite(value) || value < MIN_RETENTION_LIMIT || value > MAX_RECOVERY_ENTRIES) {
        showToast(`Retention limit must be between ${MIN_RETENTION_LIMIT} and ${MAX_RECOVERY_ENTRIES}`);
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
    },
  });

  retentionRow.append(retentionLabel, retentionInput, retentionSuffix, retentionSave);
  retentionDetails.appendChild(retentionRow);

  // -- List --

  const list = document.createElement("div");
  list.className = "pi-recovery-list";

  // -- Assemble --

  dialog.card.append(header, warningCallout, searchRow, toolbar, retentionDetails, list);

  // -- State --

  let allCheckpoints: RecoveryCheckpointSummary[] = [];
  let busy = false;
  const filterState: RecoveryFilterState = { ...DEFAULT_FILTER_STATE };

  const formatChangedLabel = (changedCount: number): string =>
    `${changedCount.toLocaleString()} change${changedCount === 1 ? "" : "s"}`;

  const shortId = (id: string): string => (id.length > 12 ? id.slice(0, 12) : id);

  /** Show/hide progressive-disclosure controls based on checkpoint count. */
  const syncVisibility = (): void => {
    const hasBackups = allCheckpoints.length > 0;
    searchRow.hidden = !hasBackups;
    toolbar.hidden = !hasBackups;
    retentionDetails.hidden = !hasRetention || !hasBackups;
  };

  const setBusy = (next: boolean): void => {
    busy = next;
    downloadBackupBtn.disabled = next || opts.onCreateManualFullBackup === undefined;
    refreshButton.disabled = next;
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
    syncVisibility();
    syncFilterControls();

    list.replaceChildren();

    if (allCheckpoints.length === 0) {
      const empty = createEmptyInline(
        lucide(Package),
        "No backups yet\nPi will save snapshots here before making changes to your data.",
      );
      list.appendChild(empty);
      statusText.textContent = "";
      return;
    }

    if (filtered.length === 0) {
      const empty = createEmptyInline(lucide(Search), "No backups match the current filters.");
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

      const restoreButton = createButton("Restore", {
        primary: true,
        compact: true,
        onClick: () => {
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
        },
      });

      const deleteButton = createButton("Delete", {
        danger: true,
        compact: true,
        onClick: () => {
          if (busy) return;
          void (async () => {
            const proceed = await requestConfirmationDialog({
              title: "Delete this backup?",
              message: `Backup: ${checkpoint.address} (#${shortId(checkpoint.id)})`,
              confirmLabel: "Delete",
              cancelLabel: "Cancel",
              confirmButtonTone: "danger",
              restoreFocusOnClose: false,
            });
            if (!proceed || busy) return;
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
        },
      });

      const actions = createActionsRow(restoreButton, deleteButton);
      actions.classList.add("pi-overlay-actions--inline");
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
        retentionInput.value = String(MAX_RECOVERY_ENTRIES);
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
