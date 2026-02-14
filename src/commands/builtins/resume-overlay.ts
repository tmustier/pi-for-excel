/**
 * Resume-session overlay.
 */

import type { SessionData, SessionMetadata } from "@mariozechner/pi-web-ui/dist/storage/types.js";
import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  getCrossWorkbookResumeConfirmMessage,
  getResumeTargetLabel,
  type ResumeDialogTarget,
} from "./resume-target.js";
import { formatRelativeDate } from "./overlay-relative-date.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { RESUME_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import { formatWorkbookLabel, getWorkbookContext } from "../../workbook/context.js";
import {
  getSessionWorkbookId,
  partitionSessionIdsByWorkbook,
} from "../../workbook/session-association.js";

function buildResumeListItem(session: SessionMetadata): HTMLButtonElement {
  const button = document.createElement("button");
  button.className = "pi-welcome-provider pi-resume-item";
  button.dataset.id = session.id;

  const title = document.createElement("span");
  title.className = "pi-resume-item__title";
  title.textContent = session.title || "Untitled";

  const meta = document.createElement("span");
  meta.className = "pi-resume-item__meta";
  meta.textContent = `${session.messageCount || 0} messages · ${formatRelativeDate(session.lastModified)}`;

  button.append(title, meta);
  return button;
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

  if (closeOverlayById(RESUME_OVERLAY_ID)) {
    return;
  }

  const workbookContext = await getWorkbookContext();
  const workbookId = workbookContext.workbookId;
  const workbookLabel = formatWorkbookLabel(workbookContext);
  const metadataById = new Map(allSessions.map((session) => [session.id, session]));

  let defaultSessionIds = allSessions.map((session) => session.id);
  if (workbookId) {
    const partition = await partitionSessionIdsByWorkbook(
      storage.settings,
      allSessions.map((session) => session.id),
      workbookId,
    );
    defaultSessionIds = [...partition.matchingSessionIds, ...partition.unlinkedSessionIds];
  }

  let showAllWorkbooks = workbookId === null;
  let selectedTarget: ResumeDialogTarget = opts.defaultTarget ?? "new_tab";

  const dialog = createOverlayDialog({
    overlayId: RESUME_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-resume-dialog",
  });

  const closeOverlay = dialog.close;

  const { header } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close resume sessions",
    title: "Resume Session",
    titleClassName: "pi-resume-dialog__title",
  });

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

  const list = document.createElement("div");
  list.className = "pi-resume-list";

  dialog.card.append(header, targetControls, targetHint);
  syncTargetButtons();

  if (workbookId) {
    dialog.card.appendChild(
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

  dialog.card.appendChild(list);

  const getVisibleSessions = (): SessionMetadata[] => {
    if (showAllWorkbooks || workbookId === null) {
      return allSessions;
    }

    const visible: SessionMetadata[] = [];
    for (const sessionId of defaultSessionIds) {
      const metadata = metadataById.get(sessionId);
      if (metadata) visible.push(metadata);
    }
    return visible;
  };

  const renderList = (): void => {
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
  };

  renderList();

  dialog.overlay.addEventListener("click", (event) => {
    const target = event.target;
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

  dialog.mount();
}
