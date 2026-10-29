function toElement(target: EventTarget | null | undefined): Element | null {
  if (!target) return null;
  if (typeof Element !== "undefined" && target instanceof Element) return target;
  if (typeof Node !== "undefined" && target instanceof Node) return target.parentElement;
  return null;
}

function isNonTextInputType(type: string): boolean {
  const normalized = type.trim().toLowerCase();
  return (
    normalized === "button"
    || normalized === "checkbox"
    || normalized === "color"
    || normalized === "file"
    || normalized === "hidden"
    || normalized === "image"
    || normalized === "radio"
    || normalized === "range"
    || normalized === "reset"
    || normalized === "submit"
  );
}

function isTextEntryElement(element: HTMLElement): boolean {
  if (typeof HTMLTextAreaElement !== "undefined" && element instanceof HTMLTextAreaElement) {
    return true;
  }

  if (typeof HTMLSelectElement !== "undefined" && element instanceof HTMLSelectElement) {
    return true;
  }

  if (typeof HTMLInputElement !== "undefined" && element instanceof HTMLInputElement) {
    return !isNonTextInputType(element.type);
  }

  return element.isContentEditable;
}

export function getTextEntryElement(target: EventTarget | null | undefined): HTMLElement | null {
  const element = toElement(target);
  if (!element) return null;

  const candidate = element.closest<HTMLElement>(
    "textarea,input,select,[contenteditable='true'],[contenteditable='plaintext-only']",
  );
  if (!candidate) return null;

  return isTextEntryElement(candidate) ? candidate : null;
}

export function isTextEntryTarget(target: EventTarget | null | undefined): boolean {
  return getTextEntryElement(target) !== null;
}

export function blurTextEntryTarget(target: EventTarget | null | undefined): boolean {
  const element = getTextEntryElement(target);
  if (!element) return false;

  element.blur();
  return true;
}
