const MENU_ITEM_SELECTOR = "[role='menuitem'], [role='menuitemradio'], [role='menuitemcheckbox']";

function getMenuItems(menu: HTMLElement): HTMLElement[] {
  return Array.from(menu.querySelectorAll<HTMLElement>(MENU_ITEM_SELECTOR)).filter((item) =>
    !item.hasAttribute("disabled")
    && item.getAttribute("aria-disabled") !== "true"
    && item.getClientRects().length > 0,
  );
}

export function focusInitialMenuItem(menu: HTMLElement): void {
  const items = getMenuItems(menu);
  const checked = items.find((item) => item.getAttribute("aria-checked") === "true");
  const target = checked ?? items[0];
  if (target) {
    target.focus({ preventScroll: true });
  } else {
    if (!menu.hasAttribute("tabindex")) menu.tabIndex = -1;
    menu.focus({ preventScroll: true });
  }
}

export function handleMenuKeydown(
  event: KeyboardEvent,
  menu: HTMLElement,
  onClose: () => void,
): void {
  if (event.altKey || event.ctrlKey || event.metaKey) return;

  if (event.key === "Escape" || event.key === "Esc") {
    event.preventDefault();
    event.stopPropagation();
    onClose();
    return;
  }

  if (event.key === "Tab") {
    event.preventDefault();
    event.stopPropagation();
    onClose();
    return;
  }

  const items = getMenuItems(menu);
  if (items.length === 0) return;

  const active = document.activeElement;
  const index = active instanceof HTMLElement ? items.indexOf(active) : -1;

  let next: HTMLElement | undefined;
  switch (event.key) {
    case "ArrowDown":
    case "Down":
      next = items[(index + 1) % items.length];
      break;
    case "ArrowUp":
    case "Up":
      next = items[(index - 1 + items.length) % items.length];
      break;
    case "Home":
      next = items[0];
      break;
    case "End":
      next = items[items.length - 1];
      break;
    default:
      if (event.key.length === 1 && /\S/u.test(event.key) && !event.shiftKey) {
        const letter = event.key.toLocaleLowerCase();
        const ordered = [...items.slice(index + 1), ...items.slice(0, index + 1)];
        next = ordered.find((item) => (item.textContent ?? "").trim().toLocaleLowerCase().startsWith(letter));
      }
  }

  if (!next) return;
  event.preventDefault();
  event.stopPropagation();
  next.focus({ preventScroll: true });
}
