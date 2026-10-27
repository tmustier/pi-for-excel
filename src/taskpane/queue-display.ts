/**
 * The queue strip above the composer: steers and follow-ups waiting for the
 * agent, prompts and commands waiting in the action queue, and messages a
 * stopped or failed turn never delivered ("Not sent", kept in the tab state).
 *
 * Each row shows its full text (on hover and focus) and has its own Edit and
 * Remove; the action queue decides what those do (see `action-queue.ts`).
 * Stores queue state per runtime, and can attach/detach to the visible sidebar.
 *
 * Security note: avoid `innerHTML` so queued text can't inject markup.
 */

import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";

import type { PiSidebar } from "../ui/pi-sidebar.js";
import { t } from "../language/index.js";

export type QueuedMessageType = "steer" | "follow-up";
export type QueuedActionType = "prompt" | "command";

export type QueuedMessageItem = {
  type: QueuedMessageType;
  text: string;
  /** The message handed to the agent; its delivery removes exactly this row. */
  message: AgentMessage;
};
type QueuedActionItem = { type: QueuedActionType; label: string; text: string };

/** One row of the strip, in display order: held, then messages, then actions. */
export type QueueRow =
  | { kind: "held"; index: number }
  | { kind: "message"; item: QueuedMessageItem }
  | { kind: "action"; index: number };

export interface QueueRowHandlers {
  edit: (row: QueueRow) => void;
  remove: (row: QueueRow) => void;
}

interface HeldSource {
  readonly held: readonly string[];
  subscribe: (listener: () => void) => () => void;
}

export type QueueDisplay = {
  add: (type: QueuedMessageType, text: string, message: AgentMessage) => void;
  clear: () => void;
  drainQueuedMessages: () => QueuedMessageItem[];
  /** Removes one waiting steer/follow-up from the strip and from the agent's queue. */
  removeQueuedMessage: (item: QueuedMessageItem) => void;
  setActionQueue: (items: QueuedActionItem[]) => void;
  setRowHandlers: (handlers: QueueRowHandlers) => void;
  /** The row Alt+Up edits: the newest row that holds the user's text (not a command). */
  lastEditableRow: () => QueueRow | null;
  attach: (sidebar: PiSidebar) => void;
  detach: () => void;
};

export function createQueueDisplay(opts: {
  agent: Agent;
  /** Where "Not sent" rows live (the tab state). */
  held: HeldSource;
}): QueueDisplay {
  const { agent } = opts;

  const queued: QueuedMessageItem[] = [];
  let queuedActions: QueuedActionItem[] = [];
  let attachedSidebar: PiSidebar | null = null;
  let handlers: QueueRowHandlers | null = null;

  function removeContainer(): void {
    attachedSidebar?.querySelector("#pi-queue-display")?.remove();
  }

  /** Polite screen-reader announcements; kept outside the strip so it survives the strip emptying. */
  function announce(text: string): void {
    if (!attachedSidebar) return;
    let announcer = attachedSidebar.querySelector<HTMLElement>("#pi-queue-announcer");
    if (!announcer) {
      announcer = document.createElement("div");
      announcer.id = "pi-queue-announcer";
      announcer.className = "pi-queue__announcer";
      announcer.setAttribute("aria-live", "polite");
      announcer.setAttribute("role", "status");
      attachedSidebar.appendChild(announcer);
    }
    announcer.textContent = text;
  }

  function focusComposer(): void {
    const textarea = attachedSidebar?.querySelector<HTMLTextAreaElement>("pi-input textarea");
    if (!textarea) return;
    textarea.focus();
    textarea.setSelectionRange(textarea.value.length, textarea.value.length);
  }

  function renderRow(row: QueueRow, label: string, labelClass: string, text: string, editable: boolean): HTMLElement {
    const itemEl = document.createElement("div");
    itemEl.className = "pi-queue__item";
    itemEl.setAttribute("role", "listitem");

    const labelEl = document.createElement("span");
    labelEl.className = `pi-queue__label ${labelClass}`;
    labelEl.textContent = label;

    // The full text is in the row; CSS shortens it to one line until the
    // row is hovered or focused.
    const textEl = document.createElement("span");
    textEl.className = "pi-queue__text";
    textEl.textContent = text;
    textEl.title = text;

    const actionsEl = document.createElement("span");
    actionsEl.className = "pi-queue__actions";
    if (editable) {
      const editBtn = document.createElement("button");
      editBtn.type = "button";
      editBtn.className = "pi-queue__edit";
      editBtn.textContent = t("queue-display.edit");
      editBtn.setAttribute("aria-label", t("queue-display.editAria", { text }));
      editBtn.addEventListener("click", () => {
        handlers?.edit(row);
        requestAnimationFrame(focusComposer);
      });
      actionsEl.appendChild(editBtn);
    }
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "pi-queue__remove";
    removeBtn.textContent = "×";
    removeBtn.setAttribute("aria-label", t("queue-display.removeAria", { text: text || label }));
    removeBtn.title = t("queue-display.remove");
    removeBtn.addEventListener("click", () => handlers?.remove(row));
    actionsEl.appendChild(removeBtn);

    itemEl.append(labelEl, textEl, actionsEl);
    return itemEl;
  }

  function updateQueueDisplay(): void {
    if (!attachedSidebar) return;

    let container = attachedSidebar.querySelector<HTMLElement>("#pi-queue-display");
    const held = opts.held.held;

    if (queued.length === 0 && queuedActions.length === 0 && held.length === 0) {
      const hadFocus = container?.contains(document.activeElement) ?? false;
      container?.remove();
      if (hadFocus) focusComposer();
      return;
    }

    if (!container) {
      container = document.createElement("div");
      container.id = "pi-queue-display";
      container.className = "pi-queue";
      container.setAttribute("role", "list");
      container.setAttribute("aria-label", t("queue-display.aria"));

      const inputArea = attachedSidebar.querySelector<HTMLElement>(".pi-input-area");
      if (inputArea && inputArea.parentElement) {
        inputArea.parentElement.insertBefore(container, inputArea);
      } else {
        attachedSidebar.appendChild(container);
      }
    }

    // Keep keyboard focus on the same row position across a re-render.
    const focused = document.activeElement;
    const focusedButton = focused && container.contains(focused)
      ? { index: [...container.querySelectorAll("button")].findIndex((button) => button === focused) }
      : null;

    const fragment = document.createDocumentFragment();

    held.forEach((text, index) => {
      fragment.appendChild(renderRow({ kind: "held", index }, t("queue-display.heldLabel"), "pi-queue__label--held", text, true));
    });

    for (const item of queued) {
      const steer = item.type === "steer";
      fragment.appendChild(renderRow(
        { kind: "message", item },
        steer ? t("queue-display.steerLabel") : t("queue-display.followUpLabel"),
        steer ? "pi-queue__label--steer" : "pi-queue__label--followup",
        item.text,
        true,
      ));
    }

    queuedActions.forEach((action, index) => {
      fragment.appendChild(renderRow({ kind: "action", index }, action.label, "pi-queue__label--action", action.text, action.type === "prompt"));
    });

    // Alt+Up edits the last row holding the user's text; commands stay queued.
    if (lastEditableRow()) {
      const hintEl = document.createElement("div");
      hintEl.className = "pi-queue__hint";
      const shortcut = navigator.platform.startsWith("Mac") || navigator.userAgent.includes("Macintosh") ? "⌥↑" : "Alt+↑";
      hintEl.textContent = t("queue-display.shortcutHint", { shortcut });
      fragment.appendChild(hintEl);
    }

    container.replaceChildren(fragment);

    if (focusedButton && focusedButton.index >= 0) {
      const buttons = container.querySelectorAll("button");
      const target = buttons[Math.min(focusedButton.index, buttons.length - 1)];
      target?.focus();
    }
  }

  function lastEditableRow(): QueueRow | null {
    for (let index = queuedActions.length - 1; index >= 0; index -= 1) {
      if (queuedActions[index]?.type === "prompt") return { kind: "action", index };
    }
    const lastMessage = queued[queued.length - 1];
    if (lastMessage) return { kind: "message", item: lastMessage };
    const held = opts.held.held;
    return held.length > 0 ? { kind: "held", index: held.length - 1 } : null;
  }

  function add(type: QueuedMessageType, text: string, message: AgentMessage): void {
    queued.push({ type, text, message });
    updateQueueDisplay();
    announce(t("queue-display.announceQueued", { text }));
  }

  function clear(): void {
    queued.length = 0;
    updateQueueDisplay();
  }

  function drainQueuedMessages(): QueuedMessageItem[] {
    const drained = queued.splice(0);
    if (drained.length > 0) updateQueueDisplay();
    return drained;
  }

  function removeQueuedMessage(item: QueuedMessageItem): void {
    const idx = queued.indexOf(item);
    if (idx === -1) return;
    queued.splice(idx, 1);
    // The agent's queues can only be cleared as a whole; put back the rest
    // of this kind in order (synchronously, so nothing is delivered between).
    if (item.type === "steer") {
      agent.clearSteeringQueue();
      for (const rest of queued) if (rest.type === "steer") agent.steer(rest.message);
    } else {
      agent.clearFollowUpQueue();
      for (const rest of queued) if (rest.type === "follow-up") agent.followUp(rest.message);
    }
    updateQueueDisplay();
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

  let renderedHeld = opts.held.held;
  opts.held.subscribe(() => {
    if (opts.held.held === renderedHeld) return;
    renderedHeld = opts.held.held;
    updateQueueDisplay();
  });

  // A row leaves the strip when its own message is delivered. Rows still
  // waiting when a turn ends are settled by the action queue (delivered,
  // or kept as "Not sent" rows), never silently dropped here.
  agent.subscribe((ev) => {
    if (queued.length === 0) return;
    if (ev.type !== "message_start" || ev.message.role !== "user") return;

    const idx = queued.findIndex((q) => q.message === ev.message);
    if (idx !== -1) {
      queued.splice(idx, 1);
      updateQueueDisplay();
      const text = typeof ev.message.content === "string"
        ? ev.message.content
        : ev.message.content.map((part) => (part.type === "text" ? part.text : "")).join("");
      announce(t("queue-display.announceDelivered", { text }));
    }
  });

  return {
    add,
    clear,
    drainQueuedMessages,
    removeQueuedMessage,
    setActionQueue,
    setRowHandlers: (next) => { handlers = next; },
    lastEditableRow,
    attach,
    detach,
  };
}
