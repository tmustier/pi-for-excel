/**
 * Escape-key guard helpers.
 *
 * When transient UI (menus, popovers, dialogs, overlays) is open,
 * Esc should dismiss that UI instead of aborting the active agent turn.
 */

const OVERLAY_ESCAPE_OWNER_SELECTORS = [
  ".pi-welcome-overlay[data-claims-escape='true']",
  ".pi-utilities-menu",
  ".pi-status-popover",
  "#pi-command-menu",
  "agent-model-selector",
  "agent-settings-dialog",
  "api-key-prompt-dialog",
  "agent-api-key-dialog",
] as const;

const WIDGET_ESCAPE_OWNER_SELECTORS = [
  "#pi-widget-slot:not(:empty)",
  "#pi-widget-slot-below:not(:empty)",
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

function doesSelectorSetClaimEscape(target: EventTarget | null | undefined, selectors: readonly string[]): boolean {
  const targetElement = getTargetElement(target ?? null);
  if (targetElement) {
    for (const selector of selectors) {
      if (targetElement.closest(selector)) return true;
    }
  }

  for (const selector of selectors) {
    if (hasVisibleMatches(selector)) return true;
  }

  return false;
}

export function doesOverlayClaimEscape(target?: EventTarget | null): boolean {
  if (typeof document === "undefined") {
    return false;
  }

  return doesSelectorSetClaimEscape(target, OVERLAY_ESCAPE_OWNER_SELECTORS);
}

export function doesExtensionWidgetClaimEscape(target?: EventTarget | null): boolean {
  if (typeof document === "undefined") {
    return false;
  }

  return doesSelectorSetClaimEscape(target, WIDGET_ESCAPE_OWNER_SELECTORS);
}

/**
 * Used by streaming abort handlers where both overlays and extension widgets
 * should consume Escape before aborting the active turn.
 */
export function doesUiClaimStreamingEscape(target?: EventTarget | null): boolean {
  return doesOverlayClaimEscape(target) || doesExtensionWidgetClaimEscape(target);
}
