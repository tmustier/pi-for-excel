/**
 * Approval dialogs: the user's consent before Pi acts.
 *
 * Approvals are queued, not replaced. Parallel tool calls in one response,
 * or several tabs, can each ask at once; every request gets its own answer
 * and is shown one at a time in arrival order, with a count of the ones
 * waiting behind it. Nothing is declined unless the user declines it or
 * the turn that asked is stopped (its `signal` aborts), which withdraws the
 * request.
 *
 * Built on `createOverlayDialog` (the overlay primitive the `ux-system`
 * work standardizes); the queue only decides which request is on screen.
 */

import { render, type TemplateResult } from "lit";

import {
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
  type OverlayDialogController,
} from "./overlay-dialog.js";
import { TOOL_APPROVAL_OVERLAY_ID } from "./overlay-ids.js";
import { t } from "../language/index.js";

export interface ApprovalRequest {
  title: string;
  /** One line naming what will happen and where, e.g. "Write Sheet1!B2:C4". */
  headline?: string;
  /** Structured details (the tool card's humanized input). */
  details?: TemplateResult;
  /** Plain-text body; used when there are no structured details. */
  message?: string;
  /** What the change replaces, e.g. "Replaces 2 non-empty cells: …". Shown prominently. */
  overwrites?: string;
  /** Short note under the details, e.g. why this asks. */
  note?: string;
  /** Which tab asked, when it may not be the one on screen. */
  source?: string;
  confirmLabel: string;
  confirmTone?: "primary" | "danger";
  signal?: AbortSignal;
  /**
   * Stops the turn that asked. With it, the dialog offers "Stop Pi": the
   * taskpane (and its Stop button) is inert behind the dialog, so this is
   * how the user stops Pi while it waits for an answer.
   */
  onStop?: () => void;
}

interface QueuedApproval {
  request: ApprovalRequest;
  resolve: (approved: boolean) => void;
  onAbort: (() => void) | null;
}

const queue: QueuedApproval[] = [];
let current: { entry: QueuedApproval; dialog: OverlayDialogController; setWaiting: (count: number) => void } | null = null;

function settle(entry: QueuedApproval, approved: boolean): void {
  if (entry.onAbort && entry.request.signal) {
    entry.request.signal.removeEventListener("abort", entry.onAbort);
  }
  entry.onAbort = null;
  entry.resolve(approved);
}

/** Declines every request, shown or queued, and stops each turn that asked. */
function stopAll(): void {
  const entries = [...(current ? [current.entry] : []), ...queue.splice(0)];
  const stops = new Set<() => void>();
  for (const entry of entries) {
    if (entry.request.onStop) stops.add(entry.request.onStop);
    settle(entry, false);
  }
  current?.dialog.close();
  for (const stop of stops) stop();
}

function updateWaitingCount(): void {
  current?.setWaiting(queue.length);
}

function showNext(): void {
  if (current) return;
  const entry = queue.shift();
  if (!entry) return;

  const { request } = entry;
  const dialog = createOverlayDialog({
    overlayId: TOOL_APPROVAL_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--s pi-approval-card",
    restoreFocusOnClose: queue.length === 0,
  });

  let answered = false;
  const answer = (approved: boolean): void => {
    if (answered) return;
    answered = true;
    settle(entry, approved);
    dialog.close();
  };

  const { header } = createOverlayHeader({
    onClose: () => answer(false),
    closeLabel: t("confirm.cancel"),
    title: request.title,
  });

  const body = document.createElement("div");
  body.className = "pi-overlay-body pi-approval__body";

  if (request.source) {
    const source = document.createElement("p");
    source.className = "pi-approval__source";
    source.textContent = request.source;
    body.appendChild(source);
  }

  if (request.headline) {
    const headline = document.createElement("p");
    headline.className = "pi-approval__headline";
    headline.textContent = request.headline;
    body.appendChild(headline);
  }

  if (request.details) {
    const details = document.createElement("div");
    details.className = "pi-approval__details";
    render(request.details, details);
    body.appendChild(details);
  } else if (request.message) {
    const message = document.createElement("p");
    message.className = "pi-overlay-subtitle pi-confirm-dialog__message";
    message.textContent = request.message;
    body.appendChild(message);
  }

  if (request.overwrites) {
    const overwrites = document.createElement("p");
    overwrites.className = "pi-approval__overwrites";
    overwrites.textContent = request.overwrites;
    body.appendChild(overwrites);
  }

  if (request.note) {
    const note = document.createElement("p");
    note.className = "pi-approval__note";
    note.textContent = request.note;
    body.appendChild(note);
  }

  const waiting = document.createElement("p");
  waiting.className = "pi-approval__waiting";
  waiting.setAttribute("aria-live", "polite");
  body.appendChild(waiting);

  const actions = document.createElement("div");
  actions.className = "pi-overlay-actions";
  const cancelButton = createOverlayButton({ text: t("approval.decline") });
  const confirmButton = createOverlayButton({
    text: request.confirmLabel,
    className: request.confirmTone === "danger" ? "pi-overlay-btn--danger" : "pi-overlay-btn--primary",
  });
  cancelButton.addEventListener("click", () => answer(false));
  confirmButton.addEventListener("click", () => answer(true));
  if (request.onStop) {
    const stopButton = createOverlayButton({ text: t("approval.stop"), className: "pi-approval__stop" });
    stopButton.addEventListener("click", () => {
      answered = true;
      stopAll();
    });
    actions.append(stopButton);
  }
  actions.append(cancelButton, confirmButton);

  dialog.card.append(header, body, actions);
  dialog.addCleanup(() => {
    // Escape, backdrop and the header × all decline.
    if (!answered) {
      answered = true;
      settle(entry, false);
    }
    if (current?.dialog === dialog) current = null;
    queueMicrotask(showNext);
  });

  current = {
    entry,
    dialog,
    setWaiting: (count) => {
      waiting.textContent = count > 0 ? t("approval.waiting", { count }) : "";
      waiting.hidden = count === 0;
    },
  };
  updateWaitingCount();
  // The question must be answered before anything else in the taskpane is
  // used (APPROVE-11): everything behind it, including the composer, is inert
  // until it closes, so Tab cannot leave it and typing cannot reach the chat.
  // The ux-system overlay primitive is taking this over for every dialog.
  const behind = Array.from(document.body.children).filter(
    (element): element is HTMLElement => element instanceof HTMLElement && !element.inert,
  );
  for (const element of behind) element.inert = true;
  dialog.addCleanup(() => {
    for (const element of behind) element.inert = false;
  });
  dialog.mount();
  // Focus the dialog itself, not "Allow": an Enter typed for the composer as
  // the dialog opens must not approve a change the user has not read.
  dialog.card.tabIndex = -1;
  dialog.card.dataset.autofocus = "";
  dialog.card.focus();
}

/**
 * Asks the user and resolves `true` (allowed) or `false` (declined or
 * withdrawn). Rejects only when no UI can be shown.
 */
export function requestApproval(request: ApprovalRequest): Promise<boolean> {
  if (typeof document === "undefined" || !(document.body instanceof HTMLElement)) {
    return Promise.reject(new Error(t("confirm.unavailable")));
  }
  if (request.signal?.aborted) return Promise.resolve(false);

  return new Promise<boolean>((resolve) => {
    const entry: QueuedApproval = { request, resolve, onAbort: null };

    if (request.signal) {
      entry.onAbort = () => {
        if (current?.entry === entry) {
          settle(entry, false);
          current.dialog.close();
          return;
        }
        const index = queue.indexOf(entry);
        if (index >= 0) queue.splice(index, 1);
        settle(entry, false);
        updateWaitingCount();
      };
      request.signal.addEventListener("abort", entry.onAbort, { once: true });
    }

    queue.push(entry);
    if (current) {
      updateWaitingCount();
    } else {
      showNext();
    }
  });
}
