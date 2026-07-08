/**
 * Settings shell — one overlay containing a navigable page stack.
 *
 * Replaces the previous constellation of sibling settings overlays (settings,
 * extensions hub, rules, backups, shortcuts) with a single surface:
 *
 *   header: [‹ back]  Page title  [×]
 *   body:   scrollable page content
 *   footer: optional sticky page footer (e.g. Save/Cancel)
 *
 * Navigation model:
 * - `navigate(id)` pushes a page; `back()` pops.
 * - Deep links rebuild the stack from each page's `parentId` chain so Back
 *   always walks up the hierarchy.
 * - Escape goes back one level (closes at the root). The × button and
 *   backdrop click always request a full close.
 * - Pages may install a `beforeLeave` guard (used for unsaved-changes
 *   confirmation); it gates back, navigate, and close.
 */

import { requestChatInputFocus } from "./input-focus.js";
import { installOverlayEscapeClose } from "./overlay-escape.js";
import { registerOverlayCloser, unregisterOverlayCloser } from "./overlay-dialog.js";

export interface SettingsPageContext {
  /** Scrollable content region for the page. */
  body: HTMLElement;
  /** Set (or clear) the sticky footer below the scroll region. */
  setFooter: (footer: HTMLElement | null) => void;
  /** Push a nested page. */
  navigate: (pageId: string) => void;
  /** Pop back one level (no-op at root). */
  back: () => void;
  /** Close the whole settings overlay. */
  close: () => void;
  /** Register cleanup run when the page is left or the overlay closes. */
  addCleanup: (cleanup: () => void) => void;
  /**
   * Install a guard consulted before leaving this page (back, navigate, or
   * close). Return false from the guard to stay on the page.
   */
  setBeforeLeave: (guard: (() => Promise<boolean>) | null) => void;
}

export interface SettingsShellPage {
  id: string;
  /** Parent page id used to rebuild the stack for deep links. */
  parentId?: string;
  title: () => string;
  subtitle?: () => string;
  render: (ctx: SettingsPageContext) => void | Promise<void>;
}

export interface SettingsShellController {
  /** Open the overlay at a page (defaults to the root page), or navigate if open. */
  open: (pageId?: string) => Promise<void>;
  /** Whether the overlay is currently mounted. */
  isOpen: () => boolean;
  /** Request a close (honors the active page's beforeLeave guard). */
  requestClose: () => Promise<void>;
}

export interface SettingsShellOptions {
  overlayId: string;
  rootId: string;
  /** Resolve a page definition by id. */
  getPage: (pageId: string) => SettingsShellPage | undefined;
  backLabel: () => string;
  closeLabel: () => string;
}

interface MountedShell {
  overlay: HTMLDivElement;
  card: HTMLDivElement;
  backButton: HTMLButtonElement;
  titleEl: HTMLHeadingElement;
  subtitleEl: HTMLParagraphElement;
  bodyEl: HTMLDivElement;
  footerEl: HTMLDivElement;
  destroy: () => void;
}

export function createSettingsShell(options: SettingsShellOptions): SettingsShellController {
  let mounted: MountedShell | null = null;
  let stack: string[] = [];
  let pageCleanups: Array<() => void> = [];
  let beforeLeave: (() => Promise<boolean>) | null = null;
  let transitionInFlight = false;
  let renderToken = 0;

  const runPageCleanups = (): void => {
    for (let index = pageCleanups.length - 1; index >= 0; index -= 1) {
      const cleanup = pageCleanups[index];
      if (!cleanup) continue;
      try {
        cleanup();
      } catch {
        // ignore cleanup errors
      }
    }
    pageCleanups = [];
    beforeLeave = null;
  };

  const destroyShell = (): void => {
    if (!mounted) return;
    renderToken += 1;
    runPageCleanups();
    unregisterOverlayCloser(mounted.overlay);
    mounted.destroy();
    mounted.overlay.remove();
    mounted = null;
    stack = [];
    requestChatInputFocus();
  };

  const confirmLeave = async (): Promise<boolean> => {
    if (!beforeLeave) return true;
    if (transitionInFlight) return false;
    transitionInFlight = true;
    try {
      return await beforeLeave();
    } finally {
      transitionInFlight = false;
    }
  };

  const requestClose = async (): Promise<void> => {
    if (!mounted) return;
    const canLeave = await confirmLeave();
    if (!canLeave) return;
    destroyShell();
  };

  const buildStackFor = (pageId: string): string[] => {
    const chain: string[] = [];
    let currentId: string | undefined = pageId;
    const seen = new Set<string>();

    while (currentId !== undefined && !seen.has(currentId)) {
      seen.add(currentId);
      chain.unshift(currentId);
      const page = options.getPage(currentId);
      currentId = page?.parentId;
    }

    if (chain[0] !== options.rootId) {
      chain.unshift(options.rootId);
    }

    return chain;
  };

  const renderCurrentPage = (direction: "forward" | "back" | "none"): void => {
    if (!mounted) return;
    const shell = mounted;

    const pageId = stack[stack.length - 1] ?? options.rootId;
    const page = options.getPage(pageId);
    if (!page) {
      destroyShell();
      return;
    }

    const currentRenderToken = renderToken + 1;
    renderToken = currentRenderToken;
    runPageCleanups();

    shell.backButton.hidden = stack.length <= 1;
    shell.titleEl.textContent = page.title();
    const subtitleText = page.subtitle?.() ?? "";
    shell.subtitleEl.textContent = subtitleText;
    shell.subtitleEl.hidden = subtitleText.length === 0;

    const pageBody = document.createElement("div");
    pageBody.className = "pi-set-shell__page";

    shell.bodyEl.replaceChildren(pageBody);
    shell.bodyEl.scrollTop = 0;
    shell.footerEl.replaceChildren();
    shell.footerEl.hidden = true;

    // Retrigger the slide-in animation for the incoming page.
    shell.bodyEl.classList.remove("pi-set-shell__body--forward", "pi-set-shell__body--back");
    if (direction !== "none") {
      // Force a reflow so removing+adding the class restarts the animation.
      void shell.bodyEl.offsetWidth;
      shell.bodyEl.classList.add(
        direction === "forward" ? "pi-set-shell__body--forward" : "pi-set-shell__body--back",
      );
    }

    const isCurrentRender = (): boolean => (
      mounted === shell
      && renderToken === currentRenderToken
      && stack[stack.length - 1] === pageId
      && pageBody.isConnected
    );

    const ctx: SettingsPageContext = {
      body: pageBody,
      setFooter: (footer) => {
        if (!isCurrentRender()) return;
        shell.footerEl.replaceChildren();
        if (footer) {
          shell.footerEl.appendChild(footer);
          shell.footerEl.hidden = false;
        } else {
          shell.footerEl.hidden = true;
        }
      },
      navigate: (nextId) => {
        if (isCurrentRender()) {
          void navigateTo(nextId);
        }
      },
      back: () => {
        if (isCurrentRender()) {
          void goBack();
        }
      },
      close: () => {
        if (isCurrentRender()) {
          void requestClose();
        }
      },
      addCleanup: (cleanup) => {
        if (isCurrentRender()) {
          pageCleanups.push(cleanup);
          return;
        }

        try {
          cleanup();
        } catch {
          // ignore cleanup errors from stale async renders
        }
      },
      setBeforeLeave: (guard) => {
        if (!isCurrentRender()) return;
        beforeLeave = guard;
      },
    };

    const rendered = page.render(ctx);
    if (rendered instanceof Promise) {
      rendered.catch(() => {
        // Pages surface their own errors; keep the shell alive.
      });
    }

    // Move focus to the page container for keyboard/screen-reader users.
    shell.bodyEl.focus({ preventScroll: true });
  };

  const navigateTo = async (pageId: string): Promise<void> => {
    if (!mounted) return;
    if (stack[stack.length - 1] === pageId) return;
    const canLeave = await confirmLeave();
    if (!canLeave) return;
    stack.push(pageId);
    renderCurrentPage("forward");
  };

  const goBack = async (): Promise<void> => {
    if (!mounted) return;
    if (stack.length <= 1) {
      await requestClose();
      return;
    }
    const canLeave = await confirmLeave();
    if (!canLeave) return;
    stack.pop();
    renderCurrentPage("back");
  };

  const mountShell = (): MountedShell => {
    const overlay = document.createElement("div");
    overlay.id = options.overlayId;
    overlay.className = "pi-welcome-overlay";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");

    const card = document.createElement("div");
    card.className = "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-set-shell";
    overlay.appendChild(card);

    const header = document.createElement("div");
    header.className = "pi-set-shell__header";

    const backButton = document.createElement("button");
    backButton.type = "button";
    backButton.className = "pi-set-shell__back";
    backButton.setAttribute("aria-label", options.backLabel());
    backButton.title = options.backLabel();
    const backGlyph = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    backGlyph.setAttribute("viewBox", "0 0 16 16");
    backGlyph.setAttribute("fill", "none");
    backGlyph.setAttribute("stroke", "currentColor");
    backGlyph.setAttribute("stroke-width", "1.75");
    backGlyph.setAttribute("stroke-linecap", "round");
    backGlyph.setAttribute("stroke-linejoin", "round");
    backGlyph.setAttribute("aria-hidden", "true");
    const backPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
    backPath.setAttribute("d", "M10 3 5 8l5 5");
    backGlyph.appendChild(backPath);
    backButton.appendChild(backGlyph);
    backButton.addEventListener("click", () => {
      void goBack();
    });

    const titleWrap = document.createElement("div");
    titleWrap.className = "pi-set-shell__title-wrap";

    const titleEl = document.createElement("h2");
    titleEl.className = "pi-overlay-title pi-set-shell__title";

    const subtitleEl = document.createElement("p");
    subtitleEl.className = "pi-overlay-subtitle pi-set-shell__subtitle";
    subtitleEl.hidden = true;

    titleWrap.append(titleEl, subtitleEl);

    const closeButton = document.createElement("button");
    closeButton.type = "button";
    closeButton.className = "pi-overlay-close";
    closeButton.textContent = "×";
    closeButton.setAttribute("aria-label", options.closeLabel());
    closeButton.title = options.closeLabel();
    closeButton.addEventListener("click", () => {
      void requestClose();
    });

    header.append(backButton, titleWrap, closeButton);

    const bodyEl = document.createElement("div");
    bodyEl.className = "pi-set-shell__body";
    bodyEl.tabIndex = -1;

    const footerEl = document.createElement("div");
    footerEl.className = "pi-set-shell__footer";
    footerEl.hidden = true;

    card.append(header, bodyEl, footerEl);

    const cleanupEscape = installOverlayEscapeClose(overlay, () => {
      void goBack();
    });

    const onBackdropClick = (event: MouseEvent): void => {
      if (event.target === overlay) {
        void requestClose();
      }
    };
    overlay.addEventListener("click", onBackdropClick);

    registerOverlayCloser(overlay, () => {
      void requestClose();
    });

    document.body.appendChild(overlay);

    return {
      overlay,
      card,
      backButton,
      titleEl,
      subtitleEl,
      bodyEl,
      footerEl,
      destroy: () => {
        cleanupEscape();
        overlay.removeEventListener("click", onBackdropClick);
      },
    };
  };

  return {
    open: async (pageId?: string) => {
      const targetId = pageId ?? options.rootId;

      if (mounted) {
        if (stack[stack.length - 1] === targetId) return;
        const canLeave = await confirmLeave();
        if (!canLeave) return;
        stack = buildStackFor(targetId);
        renderCurrentPage("forward");
        return;
      }

      mounted = mountShell();
      stack = buildStackFor(targetId);
      renderCurrentPage("none");
    },
    isOpen: () => mounted !== null,
    requestClose,
  };
}
