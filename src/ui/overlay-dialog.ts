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

/**
 * Register a close handler for an overlay element that is not built through
 * `createOverlayDialog` (e.g. the settings shell), so `closeOverlayById`
 * routes through its cleanup path instead of removing the node directly.
 */
export function registerOverlayCloser(overlay: HTMLElement, close: () => void): void {
  overlayClosers.set(overlay, close);
}

/** Remove a previously registered overlay close handler. */
export function unregisterOverlayCloser(overlay: HTMLElement): void {
  overlayClosers.delete(overlay);
}

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
  /**
   * Focus always returns to the element that opened the overlay when it is still
   * on the page. When it is gone, `true` (default) falls back to the composer.
   */
  restoreFocusOnClose?: boolean;
  /**
   * Called on Escape before the overlay closes. Return true when the dialog
   * handled it (e.g. went back one level) so the overlay stays open.
   */
  onEscape?: () => boolean;
  zIndex?: number;
  /**
   * Called on Escape before the overlay closes. Return true when the dialog
   * handled it (e.g. went back one level) so the overlay stays open.
   */
  onEscape?: () => boolean;
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

export interface OverlayButtonOptions {
  text: string;
  className?: string;
  type?: "button" | "submit" | "reset";
}

export function createOverlayButton(options: OverlayButtonOptions): HTMLButtonElement {
  const button = document.createElement("button");

  button.type = options.type ?? "button";
  button.textContent = options.text;
  const hasVariant = options.className?.includes("pi-overlay-btn--") ?? false;
  const base = hasVariant ? "pi-overlay-btn" : "pi-overlay-btn pi-overlay-btn--ghost";
  button.className = mergeClassName(base, options.className);

  return button;
}

export interface OverlayInputOptions {
  placeholder: string;
  type?: "text" | "password";
  className?: string;
}

export function createOverlayInput(options: OverlayInputOptions): HTMLInputElement {
  const input = document.createElement("input");

  input.type = options.type ?? "text";
  input.placeholder = options.placeholder;
  input.className = mergeClassName("pi-overlay-input", options.className);

  return input;
}

export function createOverlaySectionTitle(text: string): HTMLHeadingElement {
  const title = document.createElement("h3");
  title.className = "pi-overlay-section-title";
  title.textContent = text;
  return title;
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
    ...(options.closeLabel !== undefined ? { label: options.closeLabel } : {}),
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

const FOCUSABLE_SELECTOR = [
  "a[href]",
  "button:not([disabled])",
  "input:not([disabled]):not([type='hidden'])",
  "select:not([disabled])",
  "textarea:not([disabled])",
  "[tabindex]:not([tabindex='-1'])",
  "[contenteditable='true']",
].join(",");

function isRendered(element: HTMLElement): boolean {
  return element.getClientRects().length > 0 && !element.closest("[hidden], [inert]");
}

/** Tabbable elements inside `root`, in DOM order, skipping hidden ones. */
export function getTabbableElements(root: ParentNode): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)).filter(isRendered);
}

/** True while any modal overlay (Settings, a dialog, the welcome screen…) is open. */
export function isModalOverlayOpen(): boolean {
  if (typeof document === "undefined") return false;
  return getTopmostModalOverlay() !== null;
}

function getTopmostModalOverlay(): HTMLElement | null {
  const modals = Array.from(document.body.children).filter(
    (element): element is HTMLElement => element instanceof HTMLElement
      && element.getAttribute("aria-modal") === "true"
      && !element.inert,
  );
  return modals[modals.length - 1] ?? null;
}

function pickInitialFocus(card: HTMLElement): HTMLElement {
  const explicit = card.querySelector<HTMLElement>("[data-autofocus], [autofocus]");
  if (explicit && isRendered(explicit)) return explicit;

  const field = getTabbableElements(card).find((element) =>
    element instanceof HTMLTextAreaElement
    || element instanceof HTMLSelectElement
    || (element instanceof HTMLInputElement && element.type !== "checkbox" && element.type !== "radio"),
  );
  if (field) return field;

  // Focus the card itself: screen readers announce the dialog, and the next Tab
  // reaches the first control without anything being activated by accident.
  if (!card.hasAttribute("tabindex")) card.tabIndex = -1;
  return card;
}

export interface OverlayFocusHandle {
  /** Move focus into the card. Call after the overlay is in the document. */
  focusInitial: () => void;
  /** Return focus to the element that had it when the overlay opened. */
  restore: (fallbackToComposer: boolean) => void;
  dispose: () => void;
}

/**
 * Modal focus contract shared by every overlay: remember the opener, move focus
 * in, keep Tab/Shift+Tab inside the card while this overlay is topmost, and
 * hand focus back to the opener on close (or the composer when it is gone).
 */
export function installOverlayFocus(overlay: HTMLElement, card: HTMLElement): OverlayFocusHandle {
  const activeElement = document.activeElement;
  const opener = activeElement instanceof HTMLElement && activeElement !== document.body ? activeElement : null;

  const isTopmost = () => overlay.isConnected && getTopmostModalOverlay() === overlay;

  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== "Tab" || event.defaultPrevented || !isTopmost()) return;
    const tabbables = getTabbableElements(card);
    const first = tabbables[0];
    const last = tabbables[tabbables.length - 1];
    // Office on the web handles Tab itself from its own keydown listener and
    // ignores defaultPrevented: with no next tabbable it exits the add-in
    // frame. Every Tab the trap handles must stop there too.
    const handled = (target: HTMLElement): void => {
      event.preventDefault();
      event.stopPropagation();
      target.focus();
    };

    if (!first || !last) {
      handled(card);
      return;
    }

    // The trap moves focus itself for every Tab, wrapping at the ends. From
    // the card (tabindex=-1) or anywhere outside the list, start at an end.
    const index = document.activeElement instanceof HTMLElement ? tabbables.indexOf(document.activeElement) : -1;
    if (index === -1) {
      handled(event.shiftKey ? last : first);
      return;
    }
    const step = event.shiftKey ? -1 : 1;
    handled(tabbables[(index + step + tabbables.length) % tabbables.length] ?? first);
  };

  // Focus that lands behind the backdrop (a composer refocusing itself, a click
  // that slips through) is pulled back into the card. Toasts stay reachable.
  const onFocusIn = (event: FocusEvent) => {
    if (!isTopmost()) return;
    const target = event.target;
    if (!(target instanceof Node) || overlay.contains(target)) return;
    if (target instanceof Element && target.closest(".pi-toast-region")) return;
    pickInitialFocus(card).focus({ preventScroll: true });
  };

  document.addEventListener("keydown", onKeyDown, true);
  document.addEventListener("focusin", onFocusIn, true);

  return {
    focusInitial: () => {
      const title = card.querySelector<HTMLElement>(".pi-overlay-title");
      if (title && !overlay.hasAttribute("aria-labelledby")) {
        if (!title.id) title.id = `${overlay.id || "pi-overlay"}-title`;
        overlay.setAttribute("aria-labelledby", title.id);
      }
      const active = document.activeElement;
      if (active instanceof Node && card.contains(active)) return;
      pickInitialFocus(card).focus({ preventScroll: true });
    },
    restore: (fallbackToComposer) => {
      if (opener && opener.isConnected && isRendered(opener)) {
        opener.focus({ preventScroll: true });
        if (document.activeElement === opener) return;
      }
      if (fallbackToComposer) requestChatInputFocus();
    },
    dispose: () => {
      document.removeEventListener("keydown", onKeyDown, true);
      document.removeEventListener("focusin", onFocusIn, true);
    },
  };
}

/**
 * Make the overlays already open inert while `overlay` is on top. WebKit exposes
 * only one aria-modal dialog to assistive technology, and it may pick an older one.
 * Returns the function that releases them.
 */
export function coverOpenOverlays(overlay: HTMLElement): () => void {
  const covered = Array.from(document.body.children).filter(
    (element): element is HTMLElement => element instanceof HTMLElement
      && element !== overlay
      && element.getAttribute("aria-modal") === "true"
      && !element.inert,
  );
  for (const element of covered) element.inert = true;
  return () => {
    for (const element of covered) element.inert = false;
  };
}

export function createOverlayDialog(options: OverlayDialogOptions): OverlayDialogController {
  const overlay = document.createElement("div");
  overlay.id = options.overlayId;
  overlay.className = "pi-overlay";
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
  let focus: OverlayFocusHandle | null = null;

  const close = () => {
    if (closed) {
      return;
    }

    closed = true;
    overlayClosers.delete(overlay);

    for (let index = cleanups.length - 1; index >= 0; index -= 1) {
      const cleanup = cleanups[index];
      if (!cleanup) {
        continue;
      }

      try {
        cleanup();
      } catch {
        // ignore cleanup errors
      }
    }

    overlay.remove();
    focus?.dispose();
    focus?.restore(options.restoreFocusOnClose !== false);
  };

  const cleanupEscape = installOverlayEscapeClose(overlay, () => {
    if (options.onEscape?.() === true) return;
    close();
  }, card);
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
      if (closed || overlay.isConnected) return;
      focus = installOverlayFocus(overlay, card);
      cleanups.push(coverOpenOverlays(overlay));
      document.body.appendChild(overlay);
      focus.focusInitial();
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
