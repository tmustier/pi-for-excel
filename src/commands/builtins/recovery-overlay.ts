/**
 * Recovery backups overlay.
 */

import { formatRelativeDate } from "./overlay-relative-date.js";
import {
  closeOverlayById,
  createOverlayCloseButton,
  createOverlayDialog,
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

export async function showRecoveryDialog(opts: {
  workbookLabel: string;
  loadCheckpoints: () => Promise<RecoveryCheckpointSummary[]>;
  onRestore: (snapshotId: string) => Promise<void>;
  onDelete: (snapshotId: string) => Promise<boolean>;
  onClear: () => Promise<number>;
}): Promise<void> {
  if (closeOverlayById(RECOVERY_OVERLAY_ID)) {
    return;
  }

  const dialog = createOverlayDialog({
    overlayId: RECOVERY_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-recovery-dialog",
  });

  const closeOverlay = dialog.close;

  const header = document.createElement("div");
  header.className = "pi-overlay-header";

  const titleWrap = document.createElement("div");
  titleWrap.className = "pi-overlay-title-wrap";

  const title = document.createElement("h2");
  title.className = "pi-overlay-title";
  title.textContent = "Backups (Beta)";

  const subtitle = document.createElement("p");
  subtitle.className = "pi-overlay-subtitle";
  subtitle.textContent = "Saved before Pi edits, in between saves. Entries are sheet-specific in this workbook.";

  const closeButton = createOverlayCloseButton({
    onClose: closeOverlay,
    label: "Close backups",
  });

  titleWrap.append(title, subtitle);
  header.append(titleWrap, closeButton);

  const workbookTag = document.createElement("p");
  workbookTag.className = "pi-overlay-workbook-tag";
  workbookTag.textContent = `Workbook: ${opts.workbookLabel}`;

  const saveBoundaryHint = document.createElement("p");
  saveBoundaryHint.className = "pi-overlay-hint";
  saveBoundaryHint.textContent = "Backups reset after you save this workbook.";

  const toolbar = document.createElement("div");
  toolbar.className = "pi-recovery-toolbar";

  const refreshButton = document.createElement("button");
  refreshButton.type = "button";
  refreshButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
  refreshButton.textContent = "Refresh";

  const clearButton = document.createElement("button");
  clearButton.type = "button";
  clearButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
  clearButton.textContent = "Clear all";

  const statusText = document.createElement("span");
  statusText.className = "pi-recovery-status";

  toolbar.append(refreshButton, clearButton, statusText);

  const list = document.createElement("div");
  list.className = "pi-recovery-list";

  dialog.card.append(header, workbookTag, saveBoundaryHint, toolbar, list);

  let checkpoints: RecoveryCheckpointSummary[] = [];
  let busy = false;

  const formatChangedLabel = (changedCount: number): string =>
    `${changedCount.toLocaleString()} change${changedCount === 1 ? "" : "s"}`;

  const shortId = (id: string): string => (id.length > 12 ? id.slice(0, 12) : id);

  const setBusy = (next: boolean): void => {
    busy = next;
    refreshButton.disabled = next;
    clearButton.disabled = next || checkpoints.length === 0;

    for (const button of list.querySelectorAll<HTMLButtonElement>("button")) {
      button.disabled = next;
    }
  };

  const renderList = (): void => {
    list.replaceChildren();

    if (checkpoints.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-overlay-empty";
      empty.textContent = "No backups for this workbook yet.";
      list.appendChild(empty);
      statusText.textContent = "No backups";
      clearButton.disabled = true;
      return;
    }

    for (const checkpoint of checkpoints) {
      const item = document.createElement("div");
      item.className = "pi-overlay-surface pi-recovery-item";

      const header = document.createElement("div");
      header.className = "pi-recovery-item__header";

      const titleEl = document.createElement("span");
      titleEl.className = "pi-recovery-item__title";
      titleEl.textContent = `${formatRecoveryToolLabel(checkpoint.toolName)} · ${checkpoint.address}`;

      const timeEl = document.createElement("span");
      timeEl.className = "pi-recovery-item__time";
      timeEl.textContent = formatRelativeDate(new Date(checkpoint.at).toISOString());

      header.append(titleEl, timeEl);

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
            checkpoints = await opts.loadCheckpoints();
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

            checkpoints = await opts.loadCheckpoints();
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
      item.append(header, meta, actions);

      if (checkpoint.restoredFromSnapshotId) {
        const restoredMeta = document.createElement("div");
        restoredMeta.className = "pi-recovery-item__restored";
        restoredMeta.textContent = `Restored from #${shortId(checkpoint.restoredFromSnapshotId)}`;
        item.appendChild(restoredMeta);
      }

      list.appendChild(item);
    }

    statusText.textContent = `${checkpoints.length} backup${checkpoints.length === 1 ? "" : "s"}`;
    clearButton.disabled = busy || checkpoints.length === 0;
  };

  const reload = async (): Promise<void> => {
    checkpoints = await opts.loadCheckpoints();
    renderList();
  };

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
    if (busy || checkpoints.length === 0) return;

    const proceed = window.confirm(`Delete all ${checkpoints.length} backups for this workbook?`);
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

  dialog.mount();

  setBusy(true);
  statusText.textContent = "Loading…";
  try {
    await reload();
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : "Unknown error";
    showToast(`Failed to load backups: ${message}`);
    statusText.textContent = "Load failed";
  } finally {
    setBusy(false);
  }
}
