/**
 * Small DOM-only queue display for queued steering / follow-up messages.
 *
 * Stores queue state per runtime, and can attach/detach to the visible sidebar.
 *
 * Security note: avoid `innerHTML` so queued text can't inject markup.
 */

import type { Agent } from "@mariozechner/pi-agent-core";

import type { PiSidebar } from "../ui/pi-sidebar.js";
import { extractTextFromContent } from "../utils/content.js";

export type QueuedMessageType = "steer" | "follow-up";
export type QueuedActionType = "prompt" | "command";

export type QueuedMessageItem = { type: QueuedMessageType; text: string };
type QueuedActionItem = { type: QueuedActionType; label: string; text: string };

function getQueuedRestoreShortcutHint(): string {
  if (typeof navigator === "undefined") {
    return "Alt+↑";
  }

  const platform = navigator.platform ?? "";
  const userAgent = navigator.userAgent ?? "";
  if (platform.startsWith("Mac") || userAgent.includes("Macintosh")) {
    return "⌥↑";
  }

  return "Alt+↑";
}

function renderQueuedItem({ type, text }: QueuedMessageItem): HTMLElement {
  const itemEl = document.createElement("div");
  itemEl.className = "pi-queue__item";

  const label = type === "steer" ? "Steering" : "Follow-up";
  const cls = type === "steer" ? "pi-queue__label--steer" : "pi-queue__label--followup";

  const labelEl = document.createElement("span");
  labelEl.className = `pi-queue__label ${cls}`;
  labelEl.textContent = label;

  const truncated = text.length > 50 ? text.slice(0, 47) + "…" : text;
  const textEl = document.createElement("span");
  textEl.className = "pi-queue__text";
  textEl.textContent = truncated;

  itemEl.append(labelEl, textEl);
  return itemEl;
}

export type QueueDisplay = {
  add: (type: QueuedMessageType, text: string) => void;
  clear: () => void;
  drainQueuedMessages: () => QueuedMessageItem[];
  setActionQueue: (items: Array<{ type: QueuedActionType; label: string; text: string }>) => void;
  attach: (sidebar: PiSidebar) => void;
  detach: () => void;
};

export function createQueueDisplay(opts: {
  agent: Agent;
}): QueueDisplay {
  const { agent } = opts;

  const queued: QueuedMessageItem[] = [];
  let queuedActions: QueuedActionItem[] = [];
  let attachedSidebar: PiSidebar | null = null;

  function findContainer(sidebar: PiSidebar): HTMLElement | null {
    return sidebar.querySelector<HTMLElement>("#pi-queue-display");
  }

  function removeContainer(): void {
    if (!attachedSidebar) return;
    const existing = findContainer(attachedSidebar);
    existing?.remove();
  }

  function updateQueueDisplay(): void {
    if (!attachedSidebar) return;

    let container = findContainer(attachedSidebar);

    if (queued.length === 0 && queuedActions.length === 0) {
      container?.remove();
      return;
    }

    if (!container) {
      container = document.createElement("div");
      container.id = "pi-queue-display";
      container.className = "pi-queue";

      const inputArea = attachedSidebar.querySelector<HTMLElement>(".pi-input-area");
      if (inputArea && inputArea.parentElement) {
        inputArea.parentElement.insertBefore(container, inputArea);
      } else {
        attachedSidebar.appendChild(container);
      }
    }

    const fragment = document.createDocumentFragment();

    for (const item of queued) {
      fragment.appendChild(renderQueuedItem(item));
    }

    for (const action of queuedActions) {
      const itemEl = document.createElement("div");
      itemEl.className = "pi-queue__item";

      const labelEl = document.createElement("span");
      labelEl.className = "pi-queue__label pi-queue__label--action";
      labelEl.textContent = action.label;

      const truncated = action.text.length > 50 ? action.text.slice(0, 47) + "…" : action.text;
      const textEl = document.createElement("span");
      textEl.className = "pi-queue__text";
      textEl.textContent = truncated;

      itemEl.append(labelEl, textEl);
      fragment.appendChild(itemEl);
    }

    const hintEl = document.createElement("div");
    hintEl.className = "pi-queue__hint";
    hintEl.textContent = `↳ ${getQueuedRestoreShortcutHint()} to edit queued messages`;
    fragment.appendChild(hintEl);

    container.replaceChildren(fragment);
  }

  function add(type: QueuedMessageType, text: string): void {
    queued.push({ type, text });
    updateQueueDisplay();
  }

  function clear(): void {
    queued.length = 0;
    updateQueueDisplay();
  }

  function drainQueuedMessages(): QueuedMessageItem[] {
    if (queued.length === 0) {
      return [];
    }

    const drained = [...queued];
    queued.length = 0;
    updateQueueDisplay();
    return drained;
  }

  function setActionQueue(items: QueuedActionItem[]): void {
    queuedActions = items;
    updateQueueDisplay();
  }

  function attach(sidebar: PiSidebar): void {
    if (attachedSidebar === sidebar) {
      updateQueueDisplay();
      return;
    }

    removeContainer();
    attachedSidebar = sidebar;
    updateQueueDisplay();
  }

  function detach(): void {
    removeContainer();
    attachedSidebar = null;
  }

  agent.subscribe((ev) => {
    if (queued.length === 0) return;

    if (ev.type === "message_start" && ev.message.role === "user") {
      const msgText = extractTextFromContent(ev.message.content);
      const idx = queued.findIndex((q) => q.text === msgText);
      if (idx !== -1) {
        queued.splice(idx, 1);
        updateQueueDisplay();
      }
    }

    // Only clear steer/follow-up on agent end. Action queue is owned elsewhere.
    if (ev.type === "agent_end" && queued.length > 0) {
      clear();
    }
  });

  return { add, clear, drainQueuedMessages, setActionQueue, attach, detach };
}
