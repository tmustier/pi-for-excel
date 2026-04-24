/**
 * Helper to make overlay dialogs Esc-dismissible and mark them as
 * Escape owners so streaming Esc abort is suppressed while open.
 */

import { blurTextEntryTarget } from "../utils/text-entry.js";

function getTargetElement(target: EventTarget | null): Element | null {
  if (!target) return null;
  if (typeof Element !== "undefined" && target instanceof Element) return target;
  if (typeof Node !== "undefined" && target instanceof Node) return target.parentElement;
  return null;
}

function isTopmostEscapeClaim(overlay: HTMLElement): boolean {
  const claimingOverlays = Array.from(
    document.querySelectorAll<HTMLElement>("[data-claims-escape='true']"),
  ).filter((candidate) => candidate.isConnected);

  if (claimingOverlays.length === 0) {
    return true;
  }

  return claimingOverlays[claimingOverlays.length - 1] === overlay;
}

export function installOverlayEscapeClose(
  overlay: HTMLElement,
  closeOverlay: () => void,
): () => void {
  overlay.dataset.claimsEscape = "true";

  if (typeof document === "undefined") {
    return () => {
      delete overlay.dataset.claimsEscape;
    };
  }

  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== "Escape") {
      return;
    }

    if (!overlay.isConnected || !isTopmostEscapeClaim(overlay)) {
      return;
    }

    const targetElement = getTargetElement(event.target);
    if (targetElement && overlay.contains(targetElement) && blurTextEntryTarget(targetElement)) {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    closeOverlay();
  };

  document.addEventListener("keydown", onKeyDown, true);

  return () => {
    delete overlay.dataset.claimsEscape;
    document.removeEventListener("keydown", onKeyDown, true);
  };
}
