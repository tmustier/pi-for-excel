/**
 * Shared helpers for fullscreen overlays.
 *
 * Consolidates:
 * - single-instance toggle behavior by overlay id
 * - Escape handling
 * - backdrop-click close
 * - focus restoration after close
 */

import { requestChatInputFocus } from "./input-focus.js";
import { installOverlayEscapeClose } from "./overlay-escape.js";

const overlayClosers = new WeakMap<HTMLElement, () => void>();

export function closeOverlayById(overlayId: string): boolean {
  const existing = document.getElementById(overlayId);
  if (!(existing instanceof HTMLElement)) {
    return false;
  }

  const closeExisting = overlayClosers.get(existing);
  if (closeExisting) {
    closeExisting();
  } else {
    existing.remove();
  }

  return true;
}

export interface OverlayDialogOptions {
  overlayId: string;
  cardClassName: string;
  closeOnBackdrop?: boolean;
  restoreFocusOnClose?: boolean;
  zIndex?: number;
}

export interface OverlayDialogController {
  overlay: HTMLDivElement;
  card: HTMLDivElement;
  close: () => void;
  mount: () => void;
  addCleanup: (cleanup: () => void) => void;
}

export interface OverlayDialogManager {
  ensure: () => OverlayDialogController;
  dismiss: () => void;
  getCurrent: () => OverlayDialogController | null;
}

export function createOverlayCloseButton(opts: {
  onClose: () => void;
  label?: string;
}): HTMLButtonElement {
  const button = document.createElement("button");
  const label = opts.label ?? "Close dialog";

  button.type = "button";
  button.className = "pi-overlay-close";
  button.textContent = "×";
  button.setAttribute("aria-label", label);
  button.title = label;

  button.addEventListener("click", () => {
    opts.onClose();
  });

  return button;
}

export interface OverlayHeaderOptions {
  title: string;
  subtitle?: string;
  onClose: () => void;
  closeLabel?: string;
  titleClassName?: string;
  subtitleClassName?: string;
}

export interface OverlayHeaderElements {
  header: HTMLDivElement;
  titleWrap: HTMLDivElement;
  title: HTMLHeadingElement;
  subtitle: HTMLParagraphElement | null;
  closeButton: HTMLButtonElement;
}

function mergeClassName(baseClassName: string, className?: string): string {
  return className && className.trim().length > 0
    ? `${baseClassName} ${className}`
    : baseClassName;
}

export function createOverlayHeader(options: OverlayHeaderOptions): OverlayHeaderElements {
  const header = document.createElement("div");
  header.className = "pi-overlay-header";

  const titleWrap = document.createElement("div");
  titleWrap.className = "pi-overlay-title-wrap";

  const title = document.createElement("h2");
  title.className = mergeClassName("pi-overlay-title", options.titleClassName);
  title.textContent = options.title;

  titleWrap.appendChild(title);

  let subtitle: HTMLParagraphElement | null = null;
  if (options.subtitle !== undefined) {
    subtitle = document.createElement("p");
    subtitle.className = mergeClassName("pi-overlay-subtitle", options.subtitleClassName);
    subtitle.textContent = options.subtitle;
    titleWrap.appendChild(subtitle);
  }

  const closeButton = createOverlayCloseButton({
    onClose: options.onClose,
    label: options.closeLabel,
  });

  header.append(titleWrap, closeButton);

  return {
    header,
    titleWrap,
    title,
    subtitle,
    closeButton,
  };
}

export function createOverlayDialog(options: OverlayDialogOptions): OverlayDialogController {
  const overlay = document.createElement("div");
  overlay.id = options.overlayId;
  overlay.className = "pi-welcome-overlay";
  overlay.setAttribute("role", "dialog");
  overlay.setAttribute("aria-modal", "true");
  if (options.zIndex !== undefined) {
    overlay.style.zIndex = String(options.zIndex);
  }

  const card = document.createElement("div");
  card.className = options.cardClassName;
  overlay.appendChild(card);

  const cleanups: Array<() => void> = [];
  let closed = false;

  const close = () => {
    if (closed) {
      return;
    }

    closed = true;
    overlayClosers.delete(overlay);

    for (let index = cleanups.length - 1; index >= 0; index -= 1) {
      try {
        cleanups[index]();
      } catch {
        // ignore cleanup errors
      }
    }

    overlay.remove();

    if (options.restoreFocusOnClose !== false) {
      requestChatInputFocus();
    }
  };

  const cleanupEscape = installOverlayEscapeClose(overlay, close);
  cleanups.push(cleanupEscape);

  if (options.closeOnBackdrop !== false) {
    const onBackdropClick = (event: MouseEvent) => {
      if (event.target === overlay) {
        close();
      }
    };

    overlay.addEventListener("click", onBackdropClick);
    cleanups.push(() => {
      overlay.removeEventListener("click", onBackdropClick);
    });
  }

  overlayClosers.set(overlay, close);

  return {
    overlay,
    card,
    close,
    mount: () => {
      document.body.appendChild(overlay);
    },
    addCleanup: (cleanup) => {
      if (closed) {
        cleanup();
        return;
      }

      cleanups.push(cleanup);
    },
  };
}

export function createOverlayDialogManager(options: OverlayDialogOptions): OverlayDialogManager {
  let current: OverlayDialogController | null = null;

  const ensure = (): OverlayDialogController => {
    if (current && current.overlay.isConnected) {
      return current;
    }

    closeOverlayById(options.overlayId);

    const dialog = createOverlayDialog(options);
    dialog.addCleanup(() => {
      if (current === dialog) {
        current = null;
      }
    });

    current = dialog;
    return dialog;
  };

  return {
    ensure,
    dismiss: () => {
      if (current) {
        current.close();
        return;
      }

      closeOverlayById(options.overlayId);
    },
    getCurrent: () => current,
  };
}
