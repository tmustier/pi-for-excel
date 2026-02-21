/**
 * Escape-key guard helpers.
 *
 * When transient UI (menus, popovers, dialogs, overlays) is open,
 * Esc should dismiss that UI instead of aborting the active agent turn.
 */

const ESCAPE_OWNER_SELECTORS = [
  ".pi-welcome-overlay[data-claims-escape='true']",
  ".pi-utilities-menu",
  ".pi-status-popover",
  "#pi-command-menu",
  "#pi-widget-slot:not(:empty)",
  "#pi-widget-slot-below:not(:empty)",
  "agent-model-selector",
  "agent-settings-dialog",
  "api-key-prompt-dialog",
  "agent-api-key-dialog",
] as const;

function isElementTarget(value: EventTarget): value is Element {
  return typeof Element !== "undefined" && value instanceof Element;
}

function isNodeTarget(value: EventTarget): value is Node {
  return typeof Node !== "undefined" && value instanceof Node;
}

function isHTMLElementNode(value: Element): value is HTMLElement {
  return typeof HTMLElement !== "undefined" && value instanceof HTMLElement;
}

function isElementVisible(element: Element): boolean {
  if (isHTMLElementNode(element) && element.hidden) {
    return false;
  }

  if (typeof window === "undefined") {
    return true;
  }

  const style = window.getComputedStyle(element);
  if (style.display === "none") return false;
  if (style.visibility === "hidden") return false;

  return true;
}

function getTargetElement(target: EventTarget | null | undefined): Element | null {
  if (!target) return null;
  if (isElementTarget(target)) return target;
  if (isNodeTarget(target)) return target.parentElement;
  return null;
}

function hasVisibleMatches(selector: string): boolean {
  const matches = document.querySelectorAll(selector);
  for (const match of matches) {
    if (isElementVisible(match)) return true;
  }
  return false;
}

export function doesOverlayClaimEscape(target?: EventTarget | null): boolean {
  if (typeof document === "undefined") {
    return false;
  }

  const targetElement = getTargetElement(target ?? null);
  if (targetElement) {
    for (const selector of ESCAPE_OWNER_SELECTORS) {
      if (targetElement.closest(selector)) return true;
    }
  }

  for (const selector of ESCAPE_OWNER_SELECTORS) {
    if (hasVisibleMatches(selector)) return true;
  }

  return false;
}
