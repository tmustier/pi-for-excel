import type { WidgetPlacement } from "../../commands/extension-api.js";
import { EXTENSION_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { createOverlayDialogManager } from "../../ui/overlay-dialog.js";
import { upsertExtensionWidget } from "../internal/widget-surface.js";
import {
  collectSandboxUiActionIds,
  renderSandboxUiTree,
  type SandboxUiNode,
} from "../sandbox-ui.js";

const SANDBOX_WIDGET_SLOT_ID = "pi-widget-slot";

const sandboxOverlayDialogManager = createOverlayDialogManager({
  overlayId: EXTENSION_OVERLAY_ID,
  cardClassName: "pi-welcome-card pi-overlay-card",
  zIndex: 260,
});

export function createTextOnlyUiNode(text: string): SandboxUiNode {
  return {
    kind: "element",
    tag: "pre",
    className: "pi-overlay-code",
    children: [
      {
        kind: "text",
        text,
      },
    ],
  };
}

export function showOverlayNode(
  node: SandboxUiNode,
  onAction: (actionId: string) => void,
): Set<string> {
  const dialog = sandboxOverlayDialogManager.ensure();

  const body = document.createElement("div");
  renderSandboxUiTree(body, node, onAction);

  dialog.card.replaceChildren(body);

  if (!dialog.overlay.isConnected) {
    dialog.mount();
  }

  return new Set(collectSandboxUiActionIds(node));
}

export function dismissOverlay(): void {
  sandboxOverlayDialogManager.dismiss();
}

function ensureWidgetSlot(): HTMLElement | null {
  let slot = document.getElementById(SANDBOX_WIDGET_SLOT_ID);
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
  slot.id = SANDBOX_WIDGET_SLOT_ID;
  slot.className = "pi-widget-slot";
  parent.insertBefore(slot, inputArea);
  return slot;
}

export function showWidgetNode(
  node: SandboxUiNode,
  onAction: (actionId: string) => void,
): Set<string> {
  const slot = ensureWidgetSlot();
  if (!slot) {
    return new Set<string>();
  }

  const card = document.createElement("div");
  card.className = "pi-overlay-surface";

  const body = document.createElement("div");
  renderSandboxUiTree(body, node, onAction);

  card.appendChild(body);
  slot.replaceChildren(card);
  slot.style.display = "flex";

  return new Set(collectSandboxUiActionIds(node));
}

export function dismissWidget(): void {
  const slot = document.getElementById(SANDBOX_WIDGET_SLOT_ID);
  if (!slot) {
    return;
  }

  slot.style.display = "none";
  slot.replaceChildren();
}

export interface SandboxWidgetUpsertOptions {
  ownerId: string;
  widgetId: string;
  node: SandboxUiNode;
  onAction: (actionId: string) => void;
  title?: string;
  placement?: WidgetPlacement;
  order?: number;
  collapsible?: boolean;
  collapsed?: boolean;
  minHeightPx?: number | null;
  maxHeightPx?: number | null;
}

export function upsertSandboxWidgetNode(options: SandboxWidgetUpsertOptions): Set<string> {
  const body = document.createElement("div");
  renderSandboxUiTree(body, options.node, options.onAction);

  upsertExtensionWidget({
    ownerId: options.ownerId,
    id: options.widgetId,
    element: body,
    title: options.title,
    placement: options.placement,
    order: options.order,
    collapsible: options.collapsible,
    collapsed: options.collapsed,
    minHeightPx: options.minHeightPx,
    maxHeightPx: options.maxHeightPx,
  });

  return new Set(collectSandboxUiActionIds(options.node));
}
