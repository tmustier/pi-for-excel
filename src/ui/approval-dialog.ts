import { render, type TemplateResult } from "lit";

import {
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { TOOL_APPROVAL_OVERLAY_ID } from "./overlay-ids.js";
import { t } from "../language/index.js";
import { announce } from "./announcer.js";

export interface ApprovalRequest {
  title: string;
  toolCallId?: string;
  headline?: string;
  details?: TemplateResult;
  message?: string;
  overwrites?: string;
  note?: string;
  source?: string;
  confirmLabel: string;
  confirmTone?: "primary" | "danger";
  signal?: AbortSignal;
  allowForTurn?: boolean;
  secondaryApproval?: { label: string; onApprove: () => void | Promise<void> };
  onStop?: () => void;
}

type ApprovalPhase =
  | { kind: "inline" | "queued" | "settled" }
  | { kind: "preparing"; location: "inline" | "dialog" };

interface ApprovalEntry {
  phase: ApprovalPhase;
  request: ApprovalRequest;
  resolve: (approved: boolean) => void;
  onAbort: (() => void) | null;
}

const queue: ApprovalEntry[] = [];
const allowedTurns = new WeakSet<AbortSignal>();
let current: { entry: ApprovalEntry; answer: (approved: boolean) => void; setWaiting: (count: number) => void } | null = null;

export interface InlineApproval {
  toolCallId: string;
  request: ApprovalRequest;
  answer(approved: boolean): void;
  allowForTurn(): void;
  prepareAndApprove(): Promise<void>;
  preparing: boolean;
}

export interface InlineApprovalHost {
  claims(toolCallId: string): boolean;
  onChange(): void;
}

const inlineHosts = new Set<InlineApprovalHost>();
const inline = new Map<InlineApproval, { entry: ApprovalEntry; release(): void; finish(approved: boolean): void }>();

export function getInlineApprovals(): readonly InlineApproval[] {
  return [...inline].filter(([, { entry }]) => entry.phase.kind === "inline"
    || (entry.phase.kind === "preparing" && entry.phase.location === "inline")).map(([approval]) => approval);
}

function notifyInlineHosts(): void {
  for (const host of inlineHosts) host.onChange();
}

function claimedInline(toolCallId: string): boolean {
  for (const host of inlineHosts) {
    if (host.claims(toolCallId)) return true;
  }
  return false;
}

export function registerInlineApprovalHost(host: InlineApprovalHost): { refresh(): void; unregister(): void } {
  const refresh = (): void => {
    for (const [approval, { release }] of inline) {
      if (!claimedInline(approval.toolCallId)) release();
    }
  };
  inlineHosts.add(host);
  return {
    refresh,
    unregister() {
      inlineHosts.delete(host);
      refresh();
    },
  };
}

function settle(entry: ApprovalEntry, approved: boolean): void {
  if (entry.phase.kind === "settled") return;
  entry.phase = { kind: "settled" };
  if (entry.onAbort && entry.request.signal) {
    entry.request.signal.removeEventListener("abort", entry.onAbort);
  }
  entry.onAbort = null;
  entry.resolve(approved);
}

function allowTurn(signal: AbortSignal): void {
  allowedTurns.add(signal);
  for (const approval of inline.keys()) {
    if (approval.request.allowForTurn && approval.request.signal === signal) approval.answer(true);
  }
}

function stopAll(): void {
  const entries = queue.splice(0);
  const stops = new Set<() => void>();
  for (const request of [
    ...entries.map((entry) => entry.request),
    ...(current ? [current.entry.request] : []),
    ...[...inline.keys()].map((approval) => approval.request),
  ]) {
    if (request.onStop) stops.add(request.onStop);
  }
  for (const { finish } of inline.values()) finish(false);
  for (const entry of entries) settle(entry, false);
  current?.answer(false);
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
  if (request.allowForTurn && request.signal && allowedTurns.has(request.signal) && !request.signal.aborted) {
    settle(entry, true);
    showNext();
    return;
  }
  const dialog = createOverlayDialog({
    overlayId: TOOL_APPROVAL_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--s pi-approval-card",
    restoreFocusOnClose: queue.length === 0,
  });

  const answer = (approved: boolean): void => {
    if (entry.phase.kind === "settled") return;
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
  if (request.allowForTurn && request.signal) {
    const signal = request.signal;
    const turnButton = createOverlayButton({ text: t("approval.allowForTurn") });
    turnButton.addEventListener("click", () => {
      if (signal.aborted) return;
      allowTurn(signal);
      answer(true);
    });
    actions.append(turnButton);
  }
  if (request.secondaryApproval) {
    const secondary = request.secondaryApproval;
    const secondaryButton = createOverlayButton({ text: secondary.label });
    secondaryButton.addEventListener("click", () => {
      if (request.signal?.aborted || secondaryButton.disabled) return;
      entry.phase = { kind: "preparing", location: "dialog" };
      secondaryButton.disabled = true;
      confirmButton.disabled = true;
      void Promise.resolve().then(() => {
        if (entry.phase.kind !== "settled" && !request.signal?.aborted) return secondary.onApprove();
      }).then(() => {
        if (entry.phase.kind !== "settled" && !request.signal?.aborted) answer(true);
      }, () => {
        if (entry.phase.kind === "settled" || request.signal?.aborted) return;
        entry.phase = { kind: "queued" };
        const error = document.createElement("p");
        error.textContent = t("approval.preparationFailed");
        error.setAttribute("role", "alert");
        body.append(error);
        secondaryButton.disabled = false;
        confirmButton.disabled = false;
      });
    });
    actions.append(secondaryButton);
  }
  if (request.onStop) {
    const stopButton = createOverlayButton({ text: t("approval.stop"), className: "pi-approval__stop" });
    stopButton.addEventListener("click", () => {
      stopAll();
    });
    actions.append(stopButton);
  }
  actions.append(cancelButton, confirmButton);

  dialog.card.append(header, body, actions);
  dialog.addCleanup(() => {
    settle(entry, false);
    if (current?.entry === entry) current = null;
    queueMicrotask(showNext);
  });

  current = {
    entry,
    answer,
    setWaiting: (count) => {
      waiting.textContent = count > 0 ? t("approval.waiting", { count }) : "";
      waiting.hidden = count === 0;
    },
  };
  updateWaitingCount();
  const behind = Array.from(document.body.children).filter(
    (element): element is HTMLElement =>
      element instanceof HTMLElement && !element.inert && element.getAttribute("aria-live") === null,
  );
  for (const element of behind) element.inert = true;
  dialog.addCleanup(() => {
    for (const element of behind) element.inert = false;
  });
  dialog.mount();
  dialog.card.tabIndex = -1;
  dialog.card.dataset.autofocus = "";
  dialog.card.focus();
  dialog.addCleanup(announce(t("announce.approval", { title: request.title })));
}

export function requestApproval(request: ApprovalRequest): Promise<boolean> {
  if (typeof document === "undefined" || !(document.body instanceof HTMLElement)) {
    return Promise.reject(new Error(t("confirm.unavailable")));
  }
  if (request.signal?.aborted) return Promise.resolve(false);
  if (request.allowForTurn && request.signal && allowedTurns.has(request.signal)) return Promise.resolve(true);

  return new Promise<boolean>((resolve) => {
    const entry: ApprovalEntry = { request, resolve, onAbort: null, phase: { kind: "inline" } };
    if (request.toolCallId !== undefined && claimedInline(request.toolCallId)) {
      const toolCallId = request.toolCallId;
      const finish = (approved: boolean): void => {
        if (entry.phase.kind === "settled") return;
        inline.delete(approval);
        settle(entry, approved);
        notifyInlineHosts();
      };
      const approval: InlineApproval = {
        toolCallId,
        request: entry.request,
        answer(approved) {
          if (entry.phase.kind === "inline" || (entry.phase.kind === "preparing" && entry.phase.location === "inline")) finish(approved);
        },
        allowForTurn() {
          const { signal } = entry.request;
          if ((entry.phase.kind !== "inline" && (entry.phase.kind !== "preparing" || entry.phase.location === "dialog")) || !entry.request.allowForTurn || !signal || signal.aborted) return;
          allowTurn(signal);
        },
        get preparing() { return entry.phase.kind === "preparing"; },
        async prepareAndApprove() {
          const secondary = entry.request.secondaryApproval;
          if (!secondary || entry.phase.kind !== "inline" || entry.request.signal?.aborted) return;
          entry.phase = { kind: "preparing", location: "inline" };
          notifyInlineHosts();
          try {
            await secondary.onApprove();
            if (!entry.request.signal?.aborted) finish(true);
          } catch (error) {
            if (entry.phase.kind === "preparing") {
              if (entry.phase.location === "dialog") {
                inline.delete(approval);
                enqueue(entry);
              } else entry.phase = { kind: "inline" };
            }
            throw error;
          } finally {
            notifyInlineHosts();
          }
        },
      };
      entry.onAbort = () => finish(false);
      entry.request.signal?.addEventListener("abort", entry.onAbort, { once: true });
      inline.set(approval, { entry, finish, release: () => {
        if (entry.phase.kind === "preparing") {
          entry.phase = { kind: "preparing", location: "dialog" };
          return;
        }
        inline.delete(approval);
        enqueue(entry);
      } });
      notifyInlineHosts();
      return;
    }
    enqueue(entry);
  });
}

function enqueue(entry: ApprovalEntry): void {
  const { request } = entry;
  if (entry.onAbort) request.signal?.removeEventListener("abort", entry.onAbort);
  entry.onAbort = null;
  if (request.signal?.aborted) {
    settle(entry, false);
    return;
  }
  entry.phase = { kind: "queued" };
  if (request.signal) {
    entry.onAbort = () => {
      if (current?.entry === entry) {
        current.answer(false);
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
}
