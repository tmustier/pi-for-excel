/**
 * Builtin command overlays (provider picker, resume, shortcuts).
 */

import type { SessionData, SessionMetadata } from "@mariozechner/pi-web-ui/dist/storage/types.js";
import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  getCrossWorkbookResumeConfirmMessage,
  getResumeTargetLabel,
  type ResumeDialogTarget,
} from "./resume-target.js";
import { requestChatInputFocus } from "../../ui/input-focus.js";
import { showToast } from "../../ui/toast.js";
import { installOverlayEscapeClose } from "../../ui/overlay-escape.js";
import { formatWorkbookLabel, getWorkbookContext } from "../../workbook/context.js";
import {
  getSessionWorkbookId,
  partitionSessionIdsByWorkbook,
} from "../../workbook/session-association.js";

export { showInstructionsDialog } from "./instructions-overlay.js";

const overlayClosers = new WeakMap<HTMLElement, () => void>();

function formatRelativeDate(iso: string): string {
  const d = new Date(iso);
  const now = new Date();
  const diff = now.getTime() - d.getTime();
  if (diff < 60000) return "just now";
  if (diff < 3600000) return `${Math.round(diff / 60000)}m ago`;
  if (diff < 86400000) return `${Math.round(diff / 3600000)}h ago`;
  if (diff < 604800000) return `${Math.round(diff / 86400000)}d ago`;
  return d.toLocaleDateString();
}

export type RecoveryCheckpointToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "conditional_format"
  | "comments"
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
    case "conditional_format":
      return "Conditional format";
    case "comments":
      return "Comments";
    case "restore_snapshot":
      return "Restore";
    default:
      return toolName;
  }
}

export async function showProviderPicker(): Promise<void> {
  const existing = document.getElementById("pi-login-overlay");
  if (existing) {
    const closeExisting = overlayClosers.get(existing);
    if (closeExisting) {
      closeExisting();
    } else {
      existing.remove();
    }

    return;
  }

  const { ALL_PROVIDERS, buildProviderRow } = await import("../../ui/provider-login.js");
  const storage = getAppStorage();
  const configuredKeys = await storage.providerKeys.list();
  const configuredSet = new Set(configuredKeys);

  const overlay = document.createElement("div");
  overlay.id = "pi-login-overlay";
  overlay.className = "pi-welcome-overlay";

  const card = document.createElement("div");
  card.className = "pi-welcome-card pi-overlay-card pi-provider-picker-card";

  const title = document.createElement("h2");
  title.className = "pi-overlay-title";
  title.textContent = "Providers";

  const subtitle = document.createElement("p");
  subtitle.className = "pi-overlay-subtitle";
  subtitle.textContent = "Connect providers to use their models.";

  const list = document.createElement("div");
  list.className = "pi-welcome-providers pi-provider-picker-list";

  card.append(title, subtitle, list);
  overlay.appendChild(card);

  const expandedRef: { current: HTMLElement | null } = { current: null };

  for (const provider of ALL_PROVIDERS) {
    const isActive = configuredSet.has(provider.id);
    const row = buildProviderRow(provider, {
      isActive,
      expandedRef,
      onConnected: (_row: HTMLElement, _id: string, label: string) => {
        document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        showToast(`${label} connected`);
      },
      onDisconnected: (_row: HTMLElement, _id: string, label: string) => {
        document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        showToast(`${label} disconnected`);
      },
    });
    list.appendChild(row);
  }

  let closed = false;
  const closeOverlay = () => {
    if (closed) {
      return;
    }

    closed = true;
    overlayClosers.delete(overlay);
    cleanupEscape();
    overlay.remove();
    requestChatInputFocus();
  };
  const cleanupEscape = installOverlayEscapeClose(overlay, closeOverlay);
  overlayClosers.set(overlay, closeOverlay);

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeOverlay();
  });

  document.body.appendChild(overlay);
}

function buildResumeListItem(session: SessionMetadata): HTMLButtonElement {
  const btn = document.createElement("button");
  btn.className = "pi-welcome-provider pi-resume-item";
  btn.dataset.id = session.id;

  const title = document.createElement("span");
  title.className = "pi-resume-item__title";
  title.textContent = session.title || "Untitled";

  const meta = document.createElement("span");
  meta.className = "pi-resume-item__meta";
  meta.textContent = `${session.messageCount || 0} messages · ${formatRelativeDate(session.lastModified)}`;

  btn.append(title, meta);
  return btn;
}

function buildWorkbookFilterRow(opts: {
  workbookLabel: string;
  checked: boolean;
  onToggle: (checked: boolean) => void;
}): HTMLElement {
  const row = document.createElement("label");
  row.className = "pi-resume-workbook-filter";

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = opts.checked;

  const labelText = document.createElement("span");
  labelText.textContent = "Show sessions from all workbooks";

  const workbookHint = document.createElement("span");
  workbookHint.className = "pi-resume-workbook-filter__hint";
  workbookHint.textContent = opts.workbookLabel;

  checkbox.addEventListener("change", () => {
    opts.onToggle(checkbox.checked);
  });

  row.append(checkbox, labelText, workbookHint);
  return row;
}

export async function showResumeDialog(opts: {
  defaultTarget?: ResumeDialogTarget;
  onOpenInNewTab: (sessionData: SessionData) => Promise<void>;
  onReplaceCurrent: (sessionData: SessionData) => Promise<void>;
}): Promise<void> {
  const storage = getAppStorage();
  const allSessions = await storage.sessions.getAllMetadata();

  if (allSessions.length === 0) {
    showToast("No previous sessions");
    return;
  }

  const existing = document.getElementById("pi-resume-overlay");
  if (existing) {
    const closeExisting = overlayClosers.get(existing);
    if (closeExisting) {
      closeExisting();
    } else {
      existing.remove();
    }

    return;
  }

  const workbookCtx = await getWorkbookContext();
  const workbookId = workbookCtx.workbookId;
  const workbookLabel = formatWorkbookLabel(workbookCtx);
  const metadataById = new Map(allSessions.map((s) => [s.id, s]));

  let defaultSessionIds = allSessions.map((s) => s.id);
  if (workbookId) {
    const partition = await partitionSessionIdsByWorkbook(
      storage.settings,
      allSessions.map((s) => s.id),
      workbookId,
    );
    defaultSessionIds = [...partition.matchingSessionIds, ...partition.unlinkedSessionIds];
  }

  let showAllWorkbooks = workbookId === null;
  let selectedTarget: ResumeDialogTarget = opts.defaultTarget ?? "new_tab";

  const overlay = document.createElement("div");
  overlay.id = "pi-resume-overlay";
  overlay.className = "pi-welcome-overlay";

  const card = document.createElement("div");
  card.className = "pi-welcome-card pi-overlay-card pi-resume-dialog";

  const title = document.createElement("h2");
  title.className = "pi-overlay-title pi-resume-dialog__title";
  title.textContent = "Resume Session";

  card.appendChild(title);

  const targetControls = document.createElement("div");
  targetControls.className = "pi-resume-target-controls";

  const openInNewTabButton = document.createElement("button");
  openInNewTabButton.type = "button";
  openInNewTabButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-resume-target-btn";
  openInNewTabButton.textContent = "Open in new tab";

  const replaceCurrentButton = document.createElement("button");
  replaceCurrentButton.type = "button";
  replaceCurrentButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-resume-target-btn";
  replaceCurrentButton.textContent = "Replace current";

  const targetHint = document.createElement("div");
  targetHint.className = "pi-resume-target-hint";

  const syncTargetButtons = () => {
    const isNewTab = selectedTarget === "new_tab";

    openInNewTabButton.classList.toggle("is-active", isNewTab);
    replaceCurrentButton.classList.toggle("is-active", !isNewTab);

    openInNewTabButton.setAttribute("aria-pressed", String(isNewTab));
    replaceCurrentButton.setAttribute("aria-pressed", String(!isNewTab));
    targetHint.textContent = `Default action: ${getResumeTargetLabel(selectedTarget)}`;
  };

  openInNewTabButton.addEventListener("click", () => {
    selectedTarget = "new_tab";
    syncTargetButtons();
  });

  replaceCurrentButton.addEventListener("click", () => {
    selectedTarget = "replace_current";
    syncTargetButtons();
  });

  targetControls.append(openInNewTabButton, replaceCurrentButton);
  card.append(targetControls, targetHint);
  syncTargetButtons();

  const list = document.createElement("div");
  list.className = "pi-resume-list";

  if (workbookId) {
    card.appendChild(
      buildWorkbookFilterRow({
        workbookLabel,
        checked: showAllWorkbooks,
        onToggle(checked) {
          showAllWorkbooks = checked;
          renderList();
        },
      }),
    );
  }

  card.appendChild(list);
  overlay.appendChild(card);

  function getVisibleSessions(): SessionMetadata[] {
    if (showAllWorkbooks || workbookId === null) {
      return allSessions;
    }

    const visible: SessionMetadata[] = [];
    for (const sessionId of defaultSessionIds) {
      const metadata = metadataById.get(sessionId);
      if (metadata) visible.push(metadata);
    }
    return visible;
  }

  function renderList(): void {
    const sessions = getVisibleSessions().slice(0, 30);

    list.replaceChildren();

    if (sessions.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-overlay-empty pi-resume-list-empty";
      empty.textContent = "No sessions available for this workbook.";
      list.appendChild(empty);
      return;
    }

    for (const session of sessions) {
      list.appendChild(buildResumeListItem(session));
    }
  }

  renderList();

  let closed = false;
  const closeOverlay = () => {
    if (closed) {
      return;
    }

    closed = true;
    overlayClosers.delete(overlay);
    cleanupEscape();
    overlay.remove();
    requestChatInputFocus();
  };
  const cleanupEscape = installOverlayEscapeClose(overlay, closeOverlay);
  overlayClosers.set(overlay, closeOverlay);

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) {
      closeOverlay();
      return;
    }

    const target = e.target;
    if (!(target instanceof HTMLElement)) return;

    const item = target.closest<HTMLElement>(".pi-resume-item");
    if (!item) return;

    const id = item.dataset.id;
    if (!id) return;

    void (async () => {
      const targetMode = selectedTarget;

      if (workbookId) {
        const linkedWorkbookId = await getSessionWorkbookId(storage.settings, id);
        if (linkedWorkbookId && linkedWorkbookId !== workbookId) {
          const proceed = window.confirm(getCrossWorkbookResumeConfirmMessage(targetMode));
          if (!proceed) return;
        }
      }

      const sessionData = await storage.sessions.loadSession(id);
      if (!sessionData) {
        showToast("Session not found");
        closeOverlay();
        return;
      }

      if (targetMode === "replace_current") {
        await opts.onReplaceCurrent(sessionData);
      } else {
        await opts.onOpenInNewTab(sessionData);
      }

      closeOverlay();
      const resumedMode = targetMode === "replace_current" ? "current tab" : "new tab";
      showToast(`Resumed in ${resumedMode}: ${sessionData.title || "Untitled"}`);
    })();
  });

  document.body.appendChild(overlay);
}

export async function showRecoveryDialog(opts: {
  workbookLabel: string;
  loadCheckpoints: () => Promise<RecoveryCheckpointSummary[]>;
  onRestore: (snapshotId: string) => Promise<void>;
  onDelete: (snapshotId: string) => Promise<boolean>;
  onClear: () => Promise<number>;
}): Promise<void> {
  const existing = document.getElementById("pi-recovery-overlay");
  if (existing) {
    const closeExisting = overlayClosers.get(existing);
    if (closeExisting) {
      closeExisting();
    } else {
      existing.remove();
    }

    return;
  }

  const overlay = document.createElement("div");
  overlay.id = "pi-recovery-overlay";
  overlay.className = "pi-welcome-overlay";

  const card = document.createElement("div");
  card.className = "pi-welcome-card pi-overlay-card pi-recovery-dialog";

  const title = document.createElement("h2");
  title.className = "pi-overlay-title";
  title.textContent = "Recovery Checkpoints";

  const subtitle = document.createElement("p");
  subtitle.className = "pi-overlay-subtitle";
  subtitle.textContent = "Revert worksheet edits with one click. Restores also create rollback checkpoints.";

  const workbookTag = document.createElement("p");
  workbookTag.className = "pi-overlay-workbook-tag";
  workbookTag.textContent = opts.workbookLabel;

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

  const footer = document.createElement("div");
  footer.className = "pi-overlay-actions";

  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "pi-overlay-btn pi-overlay-btn--ghost";
  closeButton.textContent = "Close";

  footer.append(closeButton);
  card.append(title, subtitle, workbookTag, toolbar, list, footer);
  overlay.appendChild(card);

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
      empty.textContent = "No checkpoints for this workbook yet.";
      list.appendChild(empty);
      statusText.textContent = "No checkpoints";
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

        const proceed = window.confirm("Delete this checkpoint?");
        if (!proceed) return;

        void (async () => {
          setBusy(true);
          statusText.textContent = "Deleting…";

          try {
            const deleted = await opts.onDelete(checkpoint.id);
            if (!deleted) {
              showToast("Checkpoint not found");
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

    statusText.textContent = `${checkpoints.length} checkpoint${checkpoints.length === 1 ? "" : "s"}`;
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

    const proceed = window.confirm(`Delete all ${checkpoints.length} checkpoints for this workbook?`);
    if (!proceed) return;

    void (async () => {
      setBusy(true);
      statusText.textContent = "Clearing…";
      try {
        const removed = await opts.onClear();
        showToast(`Cleared ${removed} checkpoint${removed === 1 ? "" : "s"}`);
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

  let closed = false;
  const closeOverlay = () => {
    if (closed) {
      return;
    }

    closed = true;
    overlayClosers.delete(overlay);
    cleanupEscape();
    overlay.remove();
    requestChatInputFocus();
  };

  const cleanupEscape = installOverlayEscapeClose(overlay, closeOverlay);
  overlayClosers.set(overlay, closeOverlay);

  closeButton.addEventListener("click", closeOverlay);

  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      closeOverlay();
    }
  });

  document.body.appendChild(overlay);

  setBusy(true);
  statusText.textContent = "Loading…";
  try {
    await reload();
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : "Unknown error";
    showToast(`Failed to load checkpoints: ${message}`);
    statusText.textContent = "Load failed";
  } finally {
    setBusy(false);
  }
}

export function showShortcutsDialog(): void {
  const shortcuts = [
    ["Enter", "Send message"],
    ["Shift+Tab", "Cycle thinking level"],
    ["Esc", "Exit input focus, dismiss overlays, or abort streaming"],
    ["Enter (streaming)", "Steer — redirect agent"],
    ["⌥Enter", "Queue follow-up message"],
    ["/", "Open command menu"],
    ["↑↓", "Navigate command menu"],
    ["←/→", "Switch chats (after Esc exits input focus, wraps)"],
    ["⌘⇧[/⌘⇧]", "Switch chats (fallback on macOS hosts)"],
    ["Ctrl+PageUp/PageDown", "Switch chats (fallback on Windows hosts)"],
    ["⌘/Ctrl+⇧T", "Reopen last closed tab"],
    ["F2", "Focus chat input"],
    ["F6", "Focus: Sheet ↔ Sidebar"],
    ["⇧F6", "Focus: reverse direction"],
  ];

  const existing = document.getElementById("pi-shortcuts-overlay");
  if (existing) {
    const closeExisting = overlayClosers.get(existing);
    if (closeExisting) {
      closeExisting();
    } else {
      existing.remove();
    }

    return;
  }

  const overlay = document.createElement("div");
  overlay.id = "pi-shortcuts-overlay";
  overlay.className = "pi-welcome-overlay";

  const card = document.createElement("div");
  card.className = "pi-welcome-card pi-overlay-card pi-shortcuts-dialog";

  const title = document.createElement("h2");
  title.className = "pi-overlay-title";
  title.textContent = "Keyboard Shortcuts";

  const list = document.createElement("div");
  list.className = "pi-shortcuts-list";

  for (const [key, desc] of shortcuts) {
    const row = document.createElement("div");
    row.className = "pi-shortcuts-row";

    const keyEl = document.createElement("kbd");
    keyEl.className = "pi-shortcuts-key";
    keyEl.textContent = key;

    const descEl = document.createElement("span");
    descEl.className = "pi-shortcuts-desc";
    descEl.textContent = desc;

    row.append(keyEl, descEl);
    list.appendChild(row);
  }

  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--full";
  closeButton.textContent = "Close";

  card.append(title, list, closeButton);
  overlay.appendChild(card);

  let closed = false;
  const closeOverlay = () => {
    if (closed) {
      return;
    }

    closed = true;
    overlayClosers.delete(overlay);
    cleanupEscape();
    overlay.remove();
    requestChatInputFocus();
  };
  const cleanupEscape = installOverlayEscapeClose(overlay, closeOverlay);
  overlayClosers.set(overlay, closeOverlay);

  closeButton?.addEventListener("click", () => {
    closeOverlay();
  });

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) {
      closeOverlay();
    }
  });

  document.body.appendChild(overlay);
}
