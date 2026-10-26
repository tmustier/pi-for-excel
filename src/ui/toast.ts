/**
 * Toasts: short confirmations of something the user just did, with an
 * optional action (Undo). See src/ui/README.md, "Notifications".
 *
 * - One region, drawn above everything (overlays included), just above the
 *   composer dock so it never covers the tabs or banners.
 * - Up to three stack, newest at the bottom; an identical message restarts
 *   instead of stacking.
 * - Hovering or focusing a toast pauses its timer. Error toasts have a close
 *   button and are announced assertively.
 */

import { t } from "../language/index.js";

export type ToastVariant = "info" | "error";

export interface ToastOptions {
  duration?: number;
  variant?: ToastVariant;
}

interface ActionToastOptions {
  /** Lets the caller dismiss this toast later with `dismissToast(id)`. */
  id?: string;
  message: string;
  actionLabel: string;
  onAction: () => void;
  duration?: number;
}

interface ToastEntry {
  id?: string;
  root: HTMLDivElement;
  message: string;
  hasAction: boolean;
  remainingMs: number;
  startedAt: number;
  timer: ReturnType<typeof setTimeout> | null;
}

const ERROR_TOAST_PATTERN = /\b(fail(?:ed|ure)?|error|invalid|denied|blocked|could\s*not|couldn't|couldn’t|can\s*not|can't|can’t|unable|timed\s*out)\b/iu;
const DEFAULT_INFO_DURATION_MS = 2500;
const DEFAULT_ERROR_DURATION_MS = 6000;
const ACTION_TOAST_DEFAULT_MS = 7000;
const MAX_VISIBLE_TOASTS = 3;

let region: HTMLDivElement | null = null;
const entries: ToastEntry[] = [];

function ensureRegion(): HTMLDivElement {
  if (region && region.isConnected) {
    return region;
  }

  region = document.createElement("div");
  region.id = "pi-toast";
  region.className = "pi-toast-region";
  region.setAttribute("aria-live", "polite");
  region.setAttribute("aria-relevant", "additions");
  document.body.appendChild(region);
  return region;
}

function removeEntry(entry: ToastEntry): void {
  const index = entries.indexOf(entry);
  if (index < 0) return;
  entries.splice(index, 1);
  if (entry.timer !== null) clearTimeout(entry.timer);
  entry.timer = null;
  entry.root.remove();
}

function startTimer(entry: ToastEntry): void {
  if (entry.timer !== null) clearTimeout(entry.timer);
  entry.startedAt = Date.now();
  entry.timer = setTimeout(() => removeEntry(entry), Math.max(0, entry.remainingMs));
}

function pauseTimer(entry: ToastEntry): void {
  if (entry.timer === null) return;
  clearTimeout(entry.timer);
  entry.timer = null;
  entry.remainingMs = Math.max(800, entry.remainingMs - (Date.now() - entry.startedAt));
}

function inferToastVariant(message: string): ToastVariant {
  return ERROR_TOAST_PATTERN.test(message) ? "error" : "info";
}

function normalizeToastOptions(
  message: string,
  durationOrOptions: number | ToastOptions | undefined,
): { duration: number; variant: ToastVariant } {
  if (typeof durationOrOptions === "number") {
    return { duration: durationOrOptions, variant: inferToastVariant(message) };
  }

  const variant = durationOrOptions?.variant ?? inferToastVariant(message);
  const duration = durationOrOptions?.duration
    ?? (variant === "error" ? DEFAULT_ERROR_DURATION_MS : DEFAULT_INFO_DURATION_MS);

  return { duration, variant };
}

function renderToast(opts: {
  id?: string;
  message: string;
  duration: number;
  variant: ToastVariant;
  action?: { label: string; onAction: () => void };
}): void {
  const host = ensureRegion();

  const duplicate = entries.find((entry) => entry.message === opts.message && entry.hasAction === Boolean(opts.action));
  if (duplicate && !opts.action) {
    duplicate.remainingMs = opts.duration;
    startTimer(duplicate);
    return;
  }

  const root = document.createElement("div");
  root.className = `pi-toast visible${opts.variant === "error" ? " pi-toast--error" : ""}${opts.action ? " pi-toast--action" : ""}`;
  root.setAttribute("role", opts.variant === "error" ? "alert" : "status");

  const text = document.createElement("span");
  text.className = "pi-toast__message";
  text.textContent = opts.message;
  root.appendChild(text);

  const entry: ToastEntry = {
    ...(opts.id !== undefined ? { id: opts.id } : {}),
    root,
    message: opts.message,
    hasAction: Boolean(opts.action),
    remainingMs: opts.duration,
    startedAt: Date.now(),
    timer: null,
  };

  if (opts.action) {
    const action = document.createElement("button");
    action.type = "button";
    action.className = "pi-toast__action";
    action.textContent = opts.action.label;
    const onAction = opts.action.onAction;
    action.addEventListener("click", () => {
      removeEntry(entry);
      onAction();
    });
    root.appendChild(action);
  }

  if (opts.variant === "error") {
    const dismiss = document.createElement("button");
    dismiss.type = "button";
    dismiss.className = "pi-toast__dismiss";
    dismiss.textContent = "×";
    dismiss.setAttribute("aria-label", t("toast.dismiss"));
    dismiss.addEventListener("click", () => removeEntry(entry));
    root.appendChild(dismiss);
  }

  root.addEventListener("pointerenter", () => pauseTimer(entry));
  root.addEventListener("pointerleave", () => startTimer(entry));
  root.addEventListener("focusin", () => pauseTimer(entry));
  root.addEventListener("focusout", () => startTimer(entry));

  entries.push(entry);
  host.appendChild(root);
  startTimer(entry);

  // Drop the oldest toast without an action first; Undo stays as long as it can.
  while (entries.length > MAX_VISIBLE_TOASTS) {
    const victim = entries.find((candidate) => !candidate.hasAction) ?? entries[0];
    if (!victim) break;
    removeEntry(victim);
  }

  return toastGeneration;
}

export function showToast(message: string, duration?: number): void;
export function showToast(message: string, options?: ToastOptions): void;
export function showToast(message: string, durationOrOptions?: number | ToastOptions): void {
  if (typeof document === "undefined") return;
  const normalized = normalizeToastOptions(message, durationOrOptions);
  renderToast({ message, duration: normalized.duration, variant: normalized.variant });
}

export function showActionToast(opts: ActionToastOptions): void {
  if (typeof document === "undefined") return;
  renderToast({
    ...(opts.id !== undefined ? { id: opts.id } : {}),
    message: opts.message,
    duration: opts.duration ?? ACTION_TOAST_DEFAULT_MS,
    variant: "info",
    action: { label: opts.actionLabel, onAction: opts.onAction },
  });

  const isShowing = (): boolean =>
    generation === toastGeneration && toastElements !== null && toastElements.root.classList.contains("visible");

  return {
    dismiss: () => {
      if (!isShowing() || !toastElements) return;
      clearHideTimer();
      hideToastElements(toastElements);
    },
    isShowing,
  };
}

export function isActionToastVisible(): boolean {
  return entries.some((entry) => entry.hasAction);
}

/** Remove the toast shown with `id` (e.g. its Undo became stale). */
export function dismissToast(id: string): void {
  for (const entry of entries.filter((candidate) => candidate.id === id)) {
    removeEntry(entry);
  }
}
