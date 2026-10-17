/**
 * Small DOM-only queue display for queued steering / follow-up messages.
 *
 * Projects the committed inbox, and can attach/detach to the visible sidebar.
 *
 * Security note: avoid `innerHTML` so queued text can't inject markup.
 */

import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { SubmissionId, UserInput } from "@earendil-works/pi-durable";
import type { ConversationPresenter } from "../durable/conversation-view.js";
import { editQueuedInput, takeQueuedInputs, type QueuedInput } from "../durable/inbox.js";

import type { PiSidebar } from "../ui/pi-sidebar.js";
import { extractTextFromContent } from "../utils/content.js";
import { t } from "../language/index.js";

export type QueuedMessageType = "steer" | "follow-up";

export type QueuedMessageItem = { type: QueuedMessageType; text: string };

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

  const label = type === "steer" ? t("queue-display.steerLabel") : t("queue-display.followUpLabel");
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
  takeQueuedInputs: () => Promise<QueuedInput[]>;
  editQueuedInput: (id: SubmissionId, content: UserInput) => Promise<boolean>;
  attach: (sidebar: PiSidebar) => void;
  detach: () => void;
  dispose: () => void;
};

export function createQueueDisplay(opts: {
  view: ConversationPresenter;
}): QueueDisplay {
  const { view } = opts;
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

    const queued = view.current.inbox.filter((item) => item.mode !== "write");
    if (queued.length === 0) {
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
      fragment.appendChild(renderQueuedItem({
        type: item.mode === "steer" ? "steer" : "follow-up",
        text: extractTextFromContent(item.content),
      }));
    }

    const hintEl = document.createElement("div");
    hintEl.className = "pi-queue__hint";
    hintEl.textContent = t("queue-display.shortcutHint", { shortcut: getQueuedRestoreShortcutHint() });
    fragment.appendChild(hintEl);

    container.replaceChildren(fragment);
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

  const unsubscribe = view.subscribe(updateQueueDisplay);
  return {
    attach, detach,
    dispose: () => { detach(); unsubscribe(); },
    takeQueuedInputs: () => takeQueuedInputs(view.conversation, context),
    editQueuedInput: (id, content) => editQueuedInput(view.conversation, id, content, context),
  };
}
