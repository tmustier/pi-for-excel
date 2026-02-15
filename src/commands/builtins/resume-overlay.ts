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
import { requestConfirmationDialog } from "../../ui/confirm-dialog.js";
import { RESUME_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import { formatWorkbookLabel, getWorkbookContext } from "../../workbook/context.js";
import {
  getSessionWorkbookId,
  partitionSessionIdsByWorkbook,
} from "../../workbook/session-association.js";

const RESUME_ITEM_KIND_SESSION = "session";
const RESUME_ITEM_KIND_RECENTLY_CLOSED = "recently_closed";

type ResumeItemKind = typeof RESUME_ITEM_KIND_SESSION | typeof RESUME_ITEM_KIND_RECENTLY_CLOSED;

export interface ResumeRecentlyClosedItem {
  id: string;
  sessionId: string;
  title: string;
  closedAt: string;
  workbookId: string | null;
}

function buildResumeListItem(session: SessionMetadata): HTMLButtonElement {
  const button = document.createElement("button");
  button.className = "pi-welcome-provider pi-resume-item";
  button.dataset.id = session.id;
  button.dataset.resumeKind = RESUME_ITEM_KIND_SESSION;

  const title = document.createElement("span");
  title.className = "pi-resume-item__title";
  title.textContent = session.title || "Untitled";

  const meta = document.createElement("span");
  meta.className = "pi-resume-item__meta";
  meta.textContent = `${session.messageCount || 0} messages · ${formatRelativeDate(session.lastModified)}`;

  button.append(title, meta);
  return button;
}

function buildRecentlyClosedListItem(item: ResumeRecentlyClosedItem): HTMLButtonElement {
  const button = document.createElement("button");
  button.className = "pi-welcome-provider pi-resume-item pi-resume-item--recent";
  button.dataset.id = item.sessionId;
  button.dataset.recentId = item.id;
  button.dataset.resumeKind = RESUME_ITEM_KIND_RECENTLY_CLOSED;

  const title = document.createElement("span");
  title.className = "pi-resume-item__title";
  title.textContent = item.title || "Untitled";

  const meta = document.createElement("span");
  meta.className = "pi-resume-item__meta";
  meta.textContent = `Closed ${formatRelativeDate(item.closedAt)} · Reopens in new tab`;

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
  getRecentlyClosedItems?: () => readonly ResumeRecentlyClosedItem[];
  onReopenRecentlyClosed?: (item: ResumeRecentlyClosedItem) => Promise<boolean>;
}): Promise<void> {
  const storage = getAppStorage();
  const allSessions = await storage.sessions.getAllMetadata();

  const getRecentlyClosedItems = (): ResumeRecentlyClosedItem[] => {
    if (!opts.getRecentlyClosedItems) {
      return [];
    }

    return [...opts.getRecentlyClosedItems()];
  };

  if (allSessions.length === 0 && getRecentlyClosedItems().length === 0) {
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
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m pi-resume-dialog",
  });

  const closeOverlay = dialog.close;

  const { header } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close resume sessions",
    title: "Resume Session",
    subtitle: "Pick a previous session and choose whether to open it in a new tab or replace the current tab.",
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

  const body = document.createElement("div");
  body.className = "pi-overlay-body pi-resume-body";

  const recentSection = document.createElement("section");
  recentSection.className = "pi-overlay-section pi-resume-section";
  recentSection.dataset.resumeSection = "recently-closed";

  const recentTitle = document.createElement("h3");
  recentTitle.className = "pi-overlay-section-title";
  recentTitle.textContent = "Recently closed";

  const recentHint = document.createElement("p");
  recentHint.className = "pi-overlay-hint pi-resume-section-hint";
  recentHint.textContent = "Reopen a recently closed tab directly in a new tab.";

  const recentList = document.createElement("div");
  recentList.className = "pi-resume-list pi-resume-list--recent";

  recentSection.append(recentTitle, recentHint, recentList);

  const savedSection = document.createElement("section");
  savedSection.className = "pi-overlay-section pi-resume-section";
  savedSection.dataset.resumeSection = "saved";

  const savedTitle = document.createElement("h3");
  savedTitle.className = "pi-overlay-section-title";
  savedTitle.textContent = "Saved sessions";

  const list = document.createElement("div");
  list.className = "pi-resume-list";

  savedSection.append(savedTitle, list);
  body.append(recentSection, savedSection);

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

  dialog.card.append(header, targetControls, targetHint);
  syncTargetButtons();

  if (workbookId) {
    dialog.card.appendChild(
      buildWorkbookFilterRow({
        workbookLabel,
        checked: showAllWorkbooks,
        onToggle(checked) {
          showAllWorkbooks = checked;
          renderLists();
        },
      }),
    );
  }

  dialog.card.appendChild(body);

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

  const getVisibleRecentlyClosed = (): ResumeRecentlyClosedItem[] => {
    const items = getRecentlyClosedItems();
    if (showAllWorkbooks || workbookId === null) {
      return items;
    }

    return items.filter((item) => item.workbookId === workbookId || item.workbookId === null);
  };

  let recentlyClosedById = new Map<string, ResumeRecentlyClosedItem>();

  const reopenRecentlyClosedEntry = async (item: ResumeRecentlyClosedItem): Promise<boolean> => {
    if (opts.onReopenRecentlyClosed) {
      return opts.onReopenRecentlyClosed(item);
    }

    const sessionData = await storage.sessions.loadSession(item.sessionId);
    if (!sessionData) {
      showToast("Couldn't reopen session");
      return false;
    }

    await opts.onOpenInNewTab(sessionData);
    showToast(`Reopened: ${sessionData.title || "Untitled"}`);
    return true;
  };

  const renderLists = (): void => {
    const sessions = getVisibleSessions().slice(0, 30);
    const recentlyClosedItems = getVisibleRecentlyClosed().slice(0, 6);

    recentlyClosedById = new Map(recentlyClosedItems.map((item) => [item.id, item]));

    recentList.replaceChildren();
    list.replaceChildren();

    recentSection.hidden = recentlyClosedItems.length === 0;
    for (const item of recentlyClosedItems) {
      recentList.appendChild(buildRecentlyClosedListItem(item));
    }

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

  renderLists();

  dialog.overlay.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    const item = target.closest<HTMLElement>(".pi-resume-item");
    if (!item) return;

    const id = item.dataset.id;
    if (!id) return;

    const kindRaw = item.dataset.resumeKind;
    const kind: ResumeItemKind = kindRaw === RESUME_ITEM_KIND_RECENTLY_CLOSED
      ? RESUME_ITEM_KIND_RECENTLY_CLOSED
      : RESUME_ITEM_KIND_SESSION;

    void (async () => {
      if (kind === RESUME_ITEM_KIND_RECENTLY_CLOSED) {
        const recentId = item.dataset.recentId;
        if (!recentId) {
          showToast("Session is no longer in recently closed");
          renderLists();
          return;
        }

        const recentEntry = recentlyClosedById.get(recentId);
        if (!recentEntry) {
          showToast("Session is no longer in recently closed");
          renderLists();
          return;
        }

        const reopened = await reopenRecentlyClosedEntry(recentEntry);
        if (reopened) {
          closeOverlay();
          return;
        }

        renderLists();
        return;
      }

      const targetMode = selectedTarget;

      if (workbookId) {
        const linkedWorkbookId = await getSessionWorkbookId(storage.settings, id);
        if (linkedWorkbookId && linkedWorkbookId !== workbookId) {
          const proceed = await requestConfirmationDialog({
            title: "Resume session from another workbook?",
            message: getCrossWorkbookResumeConfirmMessage(targetMode),
            confirmLabel: "Resume anyway",
            cancelLabel: "Cancel",
            restoreFocusOnClose: false,
          });
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
