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
  minHeightPx?: number | null;
  maxHeightPx?: number | null;
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

export interface ExtensionWidgetHeightBounds {
  minHeightPx: number | null;
  maxHeightPx: number | null;
}

export interface ExtensionWidgetCollapseState {
  collapsible: boolean;
  collapsed: boolean;
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

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && !Number.isNaN(value) && Number.isFinite(value);
}

function normalizePlacement(
  placement: ExtensionWidgetPlacement | undefined,
  existing: NormalizedWidgetSpec | null,
): ExtensionWidgetPlacement {
  if (placement === "above-input" || placement === "below-input") {
    return placement;
  }

  return existing?.placement ?? "above-input";
}

function normalizeOrder(value: number | undefined, existing: NormalizedWidgetSpec | null): number {
  if (isFiniteNumber(value)) {
    return value;
  }

  return existing?.order ?? 0;
}

function clampHeight(value: number): number {
  return Math.round(Math.max(MIN_WIDGET_HEIGHT_PX, Math.min(MAX_WIDGET_HEIGHT_PX, value)));
}

function normalizeHeightOverride(value: number | null | undefined): number | null | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (value === null) {
    return null;
  }

  if (!isFiniteNumber(value)) {
    return undefined;
  }

  return clampHeight(value);
}

export function normalizeWidgetHeightBounds(
  minHeightPx: number | null,
  maxHeightPx: number | null,
): ExtensionWidgetHeightBounds {
  if (minHeightPx !== null && maxHeightPx !== null && maxHeightPx < minHeightPx) {
    return {
      minHeightPx,
      maxHeightPx: minHeightPx,
    };
  }

  return {
    minHeightPx,
    maxHeightPx,
  };
}

export function resolveWidgetHeightBoundsForUpsert(
  nextMinHeightPx: number | null | undefined,
  nextMaxHeightPx: number | null | undefined,
  existing: ExtensionWidgetHeightBounds | null,
): ExtensionWidgetHeightBounds {
  const minOverride = normalizeHeightOverride(nextMinHeightPx);
  const maxOverride = normalizeHeightOverride(nextMaxHeightPx);

  const resolvedMin = minOverride === undefined ? (existing?.minHeightPx ?? null) : minOverride;
  const resolvedMax = maxOverride === undefined ? (existing?.maxHeightPx ?? null) : maxOverride;

  return normalizeWidgetHeightBounds(resolvedMin, resolvedMax);
}

export function resolveWidgetCollapseState(
  nextCollapsible: boolean | undefined,
  nextCollapsed: boolean | undefined,
  existing: ExtensionWidgetCollapseState | null,
): ExtensionWidgetCollapseState {
  const collapsible = typeof nextCollapsible === "boolean"
    ? nextCollapsible
    : (existing?.collapsible ?? false);

  if (!collapsible) {
    return {
      collapsible: false,
      collapsed: false,
    };
  }

  const collapsed = typeof nextCollapsed === "boolean"
    ? nextCollapsed
    : (existing?.collapsed ?? false);

  return {
    collapsible,
    collapsed,
  };
}

function normalizeWidgetSpec(input: ExtensionWidgetSpec, existing: NormalizedWidgetSpec | null): NormalizedWidgetSpec {
  const ownerId = input.ownerId.trim();
  if (ownerId.length === 0) {
    throw new Error("Widget owner id cannot be empty.");
  }

  const id = input.id.trim();
  if (id.length === 0) {
    throw new Error("Widget id cannot be empty.");
  }

  const title = typeof input.title === "string"
    ? input.title.trim()
    : (existing?.title ?? "");

  const placement = normalizePlacement(input.placement, existing);
  const order = normalizeOrder(input.order, existing);

  const collapseState = resolveWidgetCollapseState(
    input.collapsible,
    input.collapsed,
    existing
      ? {
        collapsible: existing.collapsible,
        collapsed: existing.collapsed,
      }
      : null,
  );

  const heightBounds = resolveWidgetHeightBoundsForUpsert(
    input.minHeightPx,
    input.maxHeightPx,
    existing
      ? {
        minHeightPx: existing.minHeightPx,
        maxHeightPx: existing.maxHeightPx,
      }
      : null,
  );

  return {
    ownerId,
    id,
    element: input.element,
    title,
    placement,
    order,
    collapsible: collapseState.collapsible,
    collapsed: collapseState.collapsed,
    minHeightPx: heightBounds.minHeightPx,
    maxHeightPx: heightBounds.maxHeightPx,
    createdAt: existing?.createdAt ?? creationCounter++,
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

function createWidgetTitle(widget: NormalizedWidgetSpec): HTMLSpanElement {
  const title = document.createElement("span");
  title.className = "pi-ext-widget-title";
  title.textContent = widget.title.length > 0 ? widget.title : widget.id;
  return title;
}

function buildWidgetCard(widget: NormalizedWidgetSpec): HTMLElement {
  const card = document.createElement("section");
  card.className = "pi-overlay-surface pi-ext-widget-card";
  card.dataset.widgetOwnerId = widget.ownerId;
  card.dataset.widgetId = widget.id;
  card.dataset.widgetPlacement = widget.placement;

  const showHeader = widget.title.length > 0 || widget.collapsible;

  if (showHeader) {
    if (widget.collapsible) {
      const headerButton = document.createElement("button");
      headerButton.type = "button";
      headerButton.className = "pi-ext-widget-header pi-ext-widget-header--toggle";
      headerButton.setAttribute("aria-expanded", widget.collapsed ? "false" : "true");
      headerButton.setAttribute(
        "aria-label",
        `${widget.collapsed ? "Expand" : "Collapse"} widget ${widget.title.length > 0 ? widget.title : widget.id}`,
      );

      const title = createWidgetTitle(widget);

      const toggleState = document.createElement("span");
      toggleState.className = "pi-ext-widget-header__state";

      const icon = document.createElement("span");
      icon.className = "pi-ext-widget-header__icon";
      icon.textContent = widget.collapsed ? "▸" : "▾";

      const action = document.createElement("span");
      action.className = "pi-ext-widget-header__action";
      action.textContent = widget.collapsed ? "Expand" : "Collapse";

      toggleState.append(icon, action);
      headerButton.append(title, toggleState);
      headerButton.addEventListener("click", () => {
        widget.collapsed = !widget.collapsed;
        renderExtensionWidgets();
      });

      card.appendChild(headerButton);
    } else {
      const header = document.createElement("div");
      header.className = "pi-ext-widget-header";
      header.appendChild(createWidgetTitle(widget));
      card.appendChild(header);
    }
  }

  const body = document.createElement("div");
  body.className = "pi-ext-widget-body";

  if (widget.minHeightPx !== null) {
    body.style.minHeight = `${widget.minHeightPx}px`;
  }

  if (widget.maxHeightPx !== null) {
    body.style.maxHeight = `${widget.maxHeightPx}px`;
    body.style.overflowY = "auto";
    body.classList.add("pi-ext-widget-body--scrollable");
  }

  body.hidden = widget.collapsed;
  if (widget.collapsed) {
    card.classList.add("is-collapsed");
  }

  body.replaceChildren(widget.element);
  card.appendChild(body);

  return card;
}

function renderSlot(slot: HTMLElement | null, widgets: readonly NormalizedWidgetSpec[]): void {
  if (!slot) {
    return;
  }

  slot.replaceChildren(...widgets.map((widget) => buildWidgetCard(widget)));
  slot.style.display = widgets.length > 0 ? "flex" : "none";
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
  const existing = widgetsByKey.get(key) ?? null;
  const normalized = normalizeWidgetSpec(input, existing);

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
