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
 * - The history starts at the page the shell was opened on (`open(id)`, e.g.
 *   /rules or /backups opens straight at that page). `navigate(id)` pushes.
 * - Back/close rule: Escape and a page's own Save/Cancel (`ctx.back()`)
 *   return to where the user came from, and close the shell on the page it
 *   was opened at. The ‹ button is "up": it pops the history, or on the entry
 *   page goes to the page's `parentId` (so Extensions → Connections can still
 *   reach Plugins and Skills). × and backdrop click always request a full close.
 * - Pages may install a `beforeLeave` guard (used for unsaved-changes
 *   confirmation); it gates back, navigate, and close.
 */

import { createOverlayDialog, registerOverlayCloser } from "./overlay-dialog.js";

export interface SettingsPageContext {
  /** Scrollable content region for the page. */
  body: HTMLElement;
  /** Set (or clear) the sticky footer below the scroll region. */
  setFooter: (footer: HTMLElement | null) => void;
  /** Push a nested page. */
  navigate: (pageId: string) => void;
  /** Return to the previous page, or close the shell on the page it was opened at. */
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
  /** Where the ‹ button leads when this page is the one the shell opened at. */
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
  /** Stack depth of the page the shell was opened at; Escape closes at or above it. */
  let entryDepth = 1;
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
    const shell = mounted;
    mounted = null;
    stack = [];
    shell.destroy();
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

    shell.backButton.hidden = stack.length <= 1 && page.parentId === undefined;
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

  const goUp = async (): Promise<void> => {
    if (!mounted) return;
    const parentId = stack.length <= 1 ? options.getPage(stack[0] ?? options.rootId)?.parentId : undefined;
    if (parentId === undefined) {
      await goBack();
      return;
    }
    const canLeave = await confirmLeave();
    if (!canLeave) return;
    stack = [parentId];
    renderCurrentPage("back");
  };

  const mountShell = (): MountedShell => {
    const dialog = createOverlayDialog({
      overlayId: options.overlayId,
      cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-set-shell",
      closeOnBackdrop: false,
      // goBack closes (through the unsaved-changes guard) on the entry page.
      onEscape: () => {
        void goBack();
        return true;
      },
    });
    const { overlay, card } = dialog;

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
      void goUp();
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

    const onBackdropClick = (event: MouseEvent): void => {
      if (event.target === overlay) {
        void requestClose();
      }
    };
    overlay.addEventListener("click", onBackdropClick);
    dialog.addCleanup(() => overlay.removeEventListener("click", onBackdropClick));

    // closeOverlayById (e.g. reopening Settings from a command) must go
    // through the unsaved-changes guard, not straight to dialog.close().
    registerOverlayCloser(overlay, () => {
      void requestClose();
    });

    dialog.mount();

    return {
      overlay,
      card,
      backButton,
      titleEl,
      subtitleEl,
      bodyEl,
      footerEl,
      destroy: dialog.close,
    };
  };

  return {
    open: async (pageId?: string) => {
      const targetId = pageId ?? options.rootId;

      if (mounted) {
        if (stack[stack.length - 1] === targetId) return;
        const canLeave = await confirmLeave();
        if (!canLeave) return;
        stack = [targetId];
        renderCurrentPage("forward");
        return;
      }

      mounted = mountShell();
      stack = [targetId];
      renderCurrentPage("none");
    },
    isOpen: () => mounted !== null,
    requestClose,
  };
}
