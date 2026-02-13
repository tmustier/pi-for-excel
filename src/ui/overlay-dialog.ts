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
