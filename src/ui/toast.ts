/**
 * Shared toast helper used across taskpane and commands.
 */


interface ToastElements {
  root: HTMLDivElement;
  message: HTMLSpanElement;
  action: HTMLButtonElement;
}

export type ToastVariant = "info" | "error";

export interface ToastOptions {
  duration?: number;
  variant?: ToastVariant;
}

interface ActionToastOptions {
  message: string;
  actionLabel: string;
  onAction: () => void;
  duration?: number;
}

interface ResolvedToastOptions {
  message: string;
  duration: number;
  variant: ToastVariant;
  action?: {
    label: string;
    onAction: () => void;
  };
}

const ERROR_TOAST_PATTERN = /\b(fail(?:ed|ure)?|error|invalid|denied|blocked|could\s*not|couldn't|can\s*not|can't|unable|timed\s*out)\b/iu;
const DEFAULT_INFO_DURATION_MS = 2000;
const DEFAULT_ERROR_DURATION_MS = 6000;
const ACTION_TOAST_DEFAULT_MS = 7000;

let toastElements: ToastElements | null = null;
let hideTimer: ReturnType<typeof setTimeout> | null = null;
/** Backup timer for action toasts — fires even if the primary is cancelled. */
let actionHideTimer: ReturnType<typeof setTimeout> | null = null;

function clearHideTimer(): void {
  if (hideTimer !== null) {
    clearTimeout(hideTimer);
    hideTimer = null;
  }
  if (actionHideTimer !== null) {
    clearTimeout(actionHideTimer);
    actionHideTimer = null;
  }
}

function ensureToastElements(): ToastElements {
  if (toastElements) {
    return toastElements;
  }

  const root = document.createElement("div");
  root.id = "pi-toast";
  root.className = "pi-toast";
  root.setAttribute("aria-atomic", "true");

  const content = document.createElement("div");
  content.className = "pi-toast__content";

  const message = document.createElement("span");
  message.className = "pi-toast__message";

  const action = document.createElement("button");
  action.type = "button";
  action.className = "pi-toast__action";
  action.hidden = true;

  content.append(message, action);
  root.appendChild(content);
  document.body.appendChild(root);

  toastElements = { root, message, action };
  return toastElements;
}

function scheduleHide(duration: number): void {
  clearHideTimer();
  hideTimer = setTimeout(() => {
    const elements = toastElements;
    if (!elements) return;
    elements.root.classList.remove("visible");
    elements.root.classList.remove("pi-toast--action");
    elements.root.classList.remove("pi-toast--error");
    elements.action.hidden = true;
    elements.action.onclick = null;
  }, Math.max(0, duration));
}

function inferToastVariant(message: string): ToastVariant {
  return ERROR_TOAST_PATTERN.test(message) ? "error" : "info";
}

function normalizeToastOptions(
  message: string,
  durationOrOptions: number | ToastOptions | undefined,
): { duration: number; variant: ToastVariant } {
  if (typeof durationOrOptions === "number") {
    return {
      duration: durationOrOptions,
      variant: inferToastVariant(message),
    };
  }

  const variant = durationOrOptions?.variant ?? inferToastVariant(message);
  const duration = durationOrOptions?.duration
    ?? (variant === "error" ? DEFAULT_ERROR_DURATION_MS : DEFAULT_INFO_DURATION_MS);

  return { duration, variant };
}

function applyToastVariant(root: HTMLDivElement, variant: ToastVariant): void {
  root.classList.toggle("pi-toast--error", variant === "error");

  if (variant === "error") {
    root.setAttribute("role", "alert");
    root.setAttribute("aria-live", "assertive");
    return;
  }

  root.setAttribute("role", "status");
  root.setAttribute("aria-live", "polite");
}

function renderToast(opts: ResolvedToastOptions): void {
  const elements = ensureToastElements();
  applyToastVariant(elements.root, opts.variant);
  elements.message.textContent = opts.message;

  if (opts.action) {
    elements.root.classList.add("pi-toast--action");
    elements.action.hidden = false;
    elements.action.textContent = opts.action.label;
    elements.action.onclick = () => {
      opts.action?.onAction();
      elements.root.classList.remove("visible");
      elements.root.classList.remove("pi-toast--action");
      elements.root.classList.remove("pi-toast--error");
      elements.action.hidden = true;
      elements.action.onclick = null;
      clearHideTimer();
    };
  } else {
    elements.root.classList.remove("pi-toast--action");
    elements.action.hidden = true;
    elements.action.onclick = null;
  }

  elements.root.classList.add("visible");
  scheduleHide(opts.duration);

  // For action toasts, schedule an independent backup hide that can't be
  // cancelled by a subsequent plain showToast() call.  Adds 1s margin so
  // the primary timer normally wins.
  if (opts.action) {
    if (actionHideTimer !== null) clearTimeout(actionHideTimer);
    actionHideTimer = setTimeout(() => {
      actionHideTimer = null;
      // Only hide if the action toast is still showing (not already
      // dismissed by the user or the primary timer).
      if (isActionToastVisible()) {
        const el = toastElements;
        if (!el) return;
        el.root.classList.remove("visible", "pi-toast--action", "pi-toast--error");
        el.action.hidden = true;
        el.action.onclick = null;
      }
    }, opts.duration + 1000);
  }
}

export function showToast(message: string, duration?: number): void;
export function showToast(message: string, options?: ToastOptions): void;
export function showToast(message: string, durationOrOptions?: number | ToastOptions): void {
  const normalized = normalizeToastOptions(message, durationOrOptions);

  // Don't let a plain info toast overwrite an active action toast — the
  // action toast has an undo button the user may still need, and replacing
  // it would also cancel its hide timer, leaving the toast stuck forever.
  // Error toasts are allowed through — they're higher priority.
  if (normalized.variant !== "error" && isActionToastVisible()) return;

  renderToast({
    message,
    duration: normalized.duration,
    variant: normalized.variant,
  });
}

export function showActionToast(opts: ActionToastOptions): void {
  renderToast({
    message: opts.message,
    duration: opts.duration ?? ACTION_TOAST_DEFAULT_MS,
    variant: "info",
    action: {
      label: opts.actionLabel,
      onAction: opts.onAction,
    },
  });
}

export function isActionToastVisible(): boolean {
  if (!toastElements) {
    return false;
  }

  return toastElements.root.classList.contains("visible")
    && toastElements.root.classList.contains("pi-toast--action")
    && !toastElements.action.hidden;
}
