/**
 * Shared host-side extension widget surface manager.
 *
 * Supports deterministic ordering, owner-scoped cleanup, and
 * optional placement metadata for Widget API v2.
 */

export type ExtensionWidgetPlacement = "above-input" | "below-input";

export interface ExtensionWidgetSpec {
  ownerId: string;
  id: string;
  element: HTMLElement;
  title?: string;
  placement?: ExtensionWidgetPlacement;
  order?: number;
  collapsible?: boolean;
  collapsed?: boolean;
  minHeightPx?: number;
  maxHeightPx?: number;
}

interface NormalizedWidgetSpec {
  ownerId: string;
  id: string;
  element: HTMLElement;
  title: string;
  placement: ExtensionWidgetPlacement;
  order: number;
  collapsible: boolean;
  collapsed: boolean;
  minHeightPx: number | null;
  maxHeightPx: number | null;
  createdAt: number;
}

const WIDGET_SLOT_ID = "pi-widget-slot";
const BELOW_SLOT_ID = "pi-widget-slot-below";
const LEGACY_TITLELESS_CLASS = "pi-widget-slot";
const MIN_WIDGET_HEIGHT_PX = 72;
const MAX_WIDGET_HEIGHT_PX = 640;

const widgetsByKey = new Map<string, NormalizedWidgetSpec>();
let creationCounter = 0;

function toWidgetKey(ownerId: string, id: string): string {
  return `${ownerId}::${id}`;
}

function parseOptionalNumber(value: number | undefined): number | null {
  if (typeof value !== "number" || Number.isNaN(value) || !Number.isFinite(value)) {
    return null;
  }

  return value;
}

function clampHeight(value: number | null): number | null {
  if (value === null) {
    return null;
  }

  return Math.max(MIN_WIDGET_HEIGHT_PX, Math.min(MAX_WIDGET_HEIGHT_PX, value));
}

function normalizeWidgetSpec(input: ExtensionWidgetSpec, existingCreatedAt: number | null): NormalizedWidgetSpec {
  const ownerId = input.ownerId.trim();
  if (ownerId.length === 0) {
    throw new Error("Widget owner id cannot be empty.");
  }

  const id = input.id.trim();
  if (id.length === 0) {
    throw new Error("Widget id cannot be empty.");
  }

  const title = typeof input.title === "string" ? input.title.trim() : "";

  return {
    ownerId,
    id,
    element: input.element,
    title,
    placement: input.placement ?? "above-input",
    order: typeof input.order === "number" && Number.isFinite(input.order) ? input.order : 0,
    collapsible: input.collapsible === true,
    collapsed: input.collapsed === true,
    minHeightPx: clampHeight(parseOptionalNumber(input.minHeightPx)),
    maxHeightPx: clampHeight(parseOptionalNumber(input.maxHeightPx)),
    createdAt: existingCreatedAt ?? creationCounter++,
  };
}

function ensureAboveSlot(): HTMLElement | null {
  let slot = document.getElementById(WIDGET_SLOT_ID);
  if (slot) {
    return slot;
  }

  const inputArea = document.querySelector<HTMLElement>(".pi-input-area");
  if (!inputArea) {
    return null;
  }

  const parent = inputArea.parentElement;
  if (!parent) {
    return null;
  }

  slot = document.createElement("div");
  slot.id = WIDGET_SLOT_ID;
  slot.className = LEGACY_TITLELESS_CLASS;
  parent.insertBefore(slot, inputArea);
  return slot;
}

function ensureBelowSlot(): HTMLElement | null {
  let slot = document.getElementById(BELOW_SLOT_ID);
  if (slot) {
    return slot;
  }

  const inputArea = document.querySelector<HTMLElement>(".pi-input-area");
  if (!inputArea) {
    return null;
  }

  const parent = inputArea.parentElement;
  if (!parent) {
    return null;
  }

  slot = document.createElement("div");
  slot.id = BELOW_SLOT_ID;
  slot.className = LEGACY_TITLELESS_CLASS;
  parent.insertBefore(slot, inputArea.nextSibling);
  return slot;
}

function sortWidgets(widgets: readonly NormalizedWidgetSpec[]): NormalizedWidgetSpec[] {
  return [...widgets].sort((left, right) => {
    if (left.order !== right.order) {
      return left.order - right.order;
    }

    if (left.createdAt !== right.createdAt) {
      return left.createdAt - right.createdAt;
    }

    return left.id.localeCompare(right.id);
  });
}

function buildWidgetCard(widget: NormalizedWidgetSpec): HTMLElement {
  const card = document.createElement("div");
  card.className = "pi-overlay-surface";

  const showHeader = widget.title.length > 0 || widget.collapsible;
  let contentContainer: HTMLElement = card;

  if (showHeader) {
    const header = document.createElement("div");
    header.className = "pi-ext-widget-header";

    const title = document.createElement("span");
    title.className = "pi-ext-widget-title";
    title.textContent = widget.title.length > 0 ? widget.title : widget.id;
    header.appendChild(title);

    if (widget.collapsible) {
      const toggle = document.createElement("button");
      toggle.type = "button";
      toggle.className = "pi-overlay-btn pi-overlay-btn--ghost";
      toggle.textContent = widget.collapsed ? "Expand" : "Collapse";
      toggle.addEventListener("click", () => {
        widget.collapsed = !widget.collapsed;
        renderExtensionWidgets();
      });
      header.appendChild(toggle);
    }

    const body = document.createElement("div");
    body.className = "pi-ext-widget-body";

    card.append(header, body);
    contentContainer = body;
  }

  if (widget.minHeightPx !== null) {
    contentContainer.style.minHeight = `${widget.minHeightPx}px`;
  }

  if (widget.maxHeightPx !== null) {
    contentContainer.style.maxHeight = `${widget.maxHeightPx}px`;
    contentContainer.style.overflow = "auto";
  }

  contentContainer.style.display = widget.collapsed ? "none" : "block";
  contentContainer.replaceChildren(widget.element);

  return card;
}

function renderSlot(slot: HTMLElement | null, widgets: readonly NormalizedWidgetSpec[]): void {
  if (!slot) {
    return;
  }

  slot.replaceChildren(...widgets.map((widget) => buildWidgetCard(widget)));
  slot.style.display = widgets.length > 0 ? "block" : "none";
}

export function renderExtensionWidgets(): void {
  const allWidgets = Array.from(widgetsByKey.values());
  const aboveWidgets = sortWidgets(allWidgets.filter((widget) => widget.placement === "above-input"));
  const belowWidgets = sortWidgets(allWidgets.filter((widget) => widget.placement === "below-input"));

  renderSlot(ensureAboveSlot(), aboveWidgets);
  renderSlot(ensureBelowSlot(), belowWidgets);
}

export function upsertExtensionWidget(input: ExtensionWidgetSpec): void {
  const key = toWidgetKey(input.ownerId, input.id);
  const existing = widgetsByKey.get(key);
  const normalized = normalizeWidgetSpec(input, existing ? existing.createdAt : null);

  widgetsByKey.set(key, normalized);
  renderExtensionWidgets();
}

export function removeExtensionWidget(ownerId: string, widgetId: string): void {
  const key = toWidgetKey(ownerId.trim(), widgetId.trim());
  if (!widgetsByKey.delete(key)) {
    return;
  }

  renderExtensionWidgets();
}

export function clearExtensionWidgets(ownerId: string): void {
  const normalizedOwnerId = ownerId.trim();
  if (normalizedOwnerId.length === 0) {
    return;
  }

  let changed = false;
  for (const key of widgetsByKey.keys()) {
    if (!key.startsWith(`${normalizedOwnerId}::`)) {
      continue;
    }

    widgetsByKey.delete(key);
    changed = true;
  }

  if (!changed) {
    return;
  }

  renderExtensionWidgets();
}

export function clearAllExtensionWidgets(): void {
  if (widgetsByKey.size === 0) {
    return;
  }

  widgetsByKey.clear();
  renderExtensionWidgets();
}
