/**
 * Keyboard shortcuts + key-driven UX.
 *
 * Extracted from taskpane.ts to keep the entrypoint thin.
 */

import type { ModelThinkingLevel as ThinkingLevel } from "@earendil-works/pi-ai";
import { BACKGROUND_CONTEXT as durableContext } from "@earendil-works/chord/context";
import type { ConversationPresenter } from "../durable/conversation-view.js";
import type { QueueDisplay } from "./queue-display.js";
import { extractTextFromContent } from "../utils/content.js";
import { normalizePromptContent } from "../messages/image-input-limits.js";
import { executeSlashCommand } from "../commands/slash-command-execution.js";

import { getThinkingLevelsForModel } from "../models/thinking-levels.js";
import type { PiSidebar } from "../ui/pi-sidebar.js";
import { moveCursorToEnd } from "../ui/input-focus.js";
import { isActionToastVisible, showToast } from "../ui/toast.js";
import { t } from "../language/index.js";

import { doesUiClaimStreamingEscape } from "../utils/escape-guard.js";
import { blurTextEntryTarget, isTextEntryTarget } from "../utils/text-entry.js";
import {
  handleCommandMenuKey,
  hideCommandMenu,
  isCommandMenuVisible,
} from "../commands/command-menu.js";

import { flashThinkingLevel } from "./status-bar.js";
import { THINKING_LEVEL_COLORS } from "./thinking-display.js";
import { shouldHandleUndoCloseTabShortcut, textEntryHasContent } from "./keyboard-shortcuts/guards.js";
import {
  getAdjacentTabDirectionFromShortcut,
  isCloseActiveTabShortcut,
  isCreateTabShortcut,
  isFocusInputShortcut,
  isReopenLastClosedShortcut,
  isRestoreQueuedMessagesShortcut,
  isUndoCloseTabShortcut,
  shouldAbortFromEscape,
  shouldBlurEditorFromEscape,
} from "./keyboard-shortcuts/matchers.js";

export {
  getAdjacentTabDirectionFromShortcut,
  isCloseActiveTabShortcut,
  isCreateTabShortcut,
  isFocusInputShortcut,
  isReopenLastClosedShortcut,
  isRestoreQueuedMessagesShortcut,
  isUndoCloseTabShortcut,
  shouldAbortFromEscape,
  shouldBlurEditorFromEscape,
} from "./keyboard-shortcuts/matchers.js";
export { shouldHandleUndoCloseTabShortcut } from "./keyboard-shortcuts/guards.js";

interface KeydownContext {
  event: KeyboardEvent;
  agent: ConversationPresenter | null;
  textarea: HTMLTextAreaElement | undefined;
  keyTarget: Node | null;
  isInEditor: boolean;
  isStreaming: boolean;
}

type ShortcutHandler = (context: KeydownContext) => boolean;

function reportKeyboardError(): void {
  showToast(t("slash-command.toast.failed"));
}

/** Withdraw only inputs still queued at commit time; preserve inbox order. */
export async function restoreQueuedMessagesToEditor(
  sidebar: { getInput: () => { value: string } | undefined; getTextarea: () => HTMLTextAreaElement | undefined },
  queueDisplay: Pick<QueueDisplay, "takeQueuedInputs"> | null,
): Promise<number> {
  const queued = await queueDisplay?.takeQueuedInputs() ?? [];
  if (!queued.length) return 0;
  const input = sidebar.getInput();
  const textarea = sidebar.getTextarea();
  const text = [...queued.map((item) => extractTextFromContent(item.content)), input?.value ?? textarea?.value ?? ""]
    .filter((value) => value.trim().length > 0).join("\n\n");
  if (input) input.value = text;
  else if (textarea) textarea.value = text;
  textarea?.focus();
  textarea?.setSelectionRange(text.length, text.length);
  return queued.length;
}

function isInsideSessionTabs(target: EventTarget | null | undefined): boolean {
  if (!(target instanceof Element)) {
    return false;
  }

  return target.closest(".pi-session-tabs") !== null;
}

function setExcelToolCardsExpanded(expanded: boolean): void {
  const toolMessages = document.querySelectorAll("tool-message");

  for (const toolMessage of toolMessages) {
    const body = toolMessage.querySelector<HTMLElement>(".pi-tool-card__body");
    if (!body) continue;

    body.classList.toggle("pi-tool-card__body--collapsed", !expanded);

    const up = toolMessage.querySelector<HTMLElement>(".pi-tool-card__chevron-up");
    const down = toolMessage.querySelector<HTMLElement>(".pi-tool-card__chevron-collapsed");
    if (!up || !down) continue;

    if (expanded) {
      up.classList.remove("hidden");
      down.classList.add("hidden");
    } else {
      up.classList.add("hidden");
      down.classList.remove("hidden");
    }
  }
}

function collapseThinkingBlocks(): void {
  const blocks = document.querySelectorAll("thinking-block");
  for (const block of blocks) {
    // When expanded, ThinkingBlock renders a markdown-block for its body.
    const isExpanded = Boolean(block.querySelector("markdown-block"));
    if (!isExpanded) continue;

    const header = block.querySelector<HTMLElement>(".thinking-header");
    header?.click();
  }
}

function expandThinkingBlocks(): void {
  const blocks = document.querySelectorAll("thinking-block");
  for (const block of blocks) {
    const isExpanded = Boolean(block.querySelector("markdown-block"));
    if (isExpanded) continue;

    const header = block.querySelector<HTMLElement>(".thinking-header");
    header?.click();
  }
}

function buildKeydownContext(args: {
  event: KeyboardEvent;
  agent: ConversationPresenter | null;
  sidebar: PiSidebar;
}): KeydownContext {
  const { event, agent, sidebar } = args;
  const textarea = sidebar.getTextarea();
  const eventTarget = event.target instanceof Node ? event.target : null;
  const keyTarget = eventTarget ?? (document.activeElement instanceof Node ? document.activeElement : null);
  const isInSidebarInput = keyTarget instanceof Element
    ? keyTarget.closest(".pi-input-card") !== null
    : false;

  const isInEditor = Boolean(
    isInSidebarInput
    || (textarea && keyTarget && (keyTarget === textarea || textarea.contains(keyTarget))),
  );

  return {
    event,
    agent,
    textarea,
    keyTarget,
    isInEditor,
    isStreaming: agent?.current.busy ?? false,
  };
}

export function getThinkingLevels(agent: ConversationPresenter): ThinkingLevel[] {
  return getThinkingLevelsForModel(agent.current.model);
}

export async function cycleThinkingLevel(agent: ConversationPresenter): Promise<ThinkingLevel> {
  const levels = getThinkingLevels(agent);
  const current = agent.current.thinkingLevel;
  const idx = levels.indexOf(current);
  const next = levels[(idx >= 0 ? idx + 1 : 0) % levels.length] ?? "off";

  await agent.conversation.configure({ thinkingLevel: next }, durableContext);
  flashThinkingLevel(next, THINKING_LEVEL_COLORS[next]);

  return next;
}

export function installKeyboardShortcuts(opts: {
  getActiveConversation: () => ConversationPresenter | null;
  getActiveQueueDisplay: () => QueueDisplay | null;
  sidebar: PiSidebar;
  markUserAborted?: (view: ConversationPresenter) => void;
  onCreateTab?: () => void;
  onCloseActiveTab?: () => void;
  onReopenLastClosed?: () => void;
  canUndoCloseTab?: () => boolean;
  onSwitchAdjacentTab?: (direction: -1 | 1) => void;
}): () => void {
  const {
    getActiveConversation,
    getActiveQueueDisplay,
    sidebar,
    markUserAborted,
    onCreateTab,
    onCloseActiveTab,
    onReopenLastClosed,
    canUndoCloseTab,
    onSwitchAdjacentTab,
  } = opts;
  const submitting = new WeakSet<ConversationPresenter>();

  const shortcutHandlers: readonly ShortcutHandler[] = [
    (context) => {
      const { event } = context;
      if (!isFocusInputShortcut(event)) return false;

      const input = sidebar.getInput();
      if (!input) return false;

      event.preventDefault();
      event.stopPropagation();
      input.focus();

      const activeTextarea = sidebar.getTextarea();
      if (activeTextarea) {
        moveCursorToEnd(activeTextarea);
      }

      return true;
    },
    (context) => {
      const { event } = context;
      if (event.key !== "Escape" || !isCommandMenuVisible()) return false;

      event.preventDefault();
      hideCommandMenu();
      return true;
    },
    (context) => {
      const { event, isInEditor, isStreaming, keyTarget } = context;
      if (!shouldBlurEditorFromEscape({ key: event.key, isInEditor, isStreaming })) return false;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();

      const blurred = blurTextEntryTarget(keyTarget);
      if (blurred) {
        requestAnimationFrame(() => {
          sidebar.focusTabNavigationAnchor();
        });
      }

      return true;
    },
    (context) => {
      const { event, isStreaming, agent, keyTarget } = context;
      const isEscapeKey = event.key === "Escape" || event.key === "Esc";
      const escapeClaimedByOverlay = isEscapeKey && doesUiClaimStreamingEscape(keyTarget);

      if (
        !isEscapeKey
        || !shouldAbortFromEscape({
          isStreaming,
          hasAgent: agent !== null,
          escapeClaimedByOverlay,
        })
        || !agent
      ) {
        return false;
      }

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();

      const queueDisplay = getActiveQueueDisplay();
      markUserAborted?.(agent);
      void (async () => {
        try {
          await restoreQueuedMessagesToEditor(sidebar, queueDisplay);
        } finally {
          await agent.conversation.abort(durableContext);
        }
      })().catch(reportKeyboardError);

      return true;
    },
    (context) => {
      const { event, isInEditor, keyTarget } = context;
      if (!isRestoreQueuedMessagesShortcut(event)) {
        return false;
      }

      if (!isInEditor && isTextEntryTarget(keyTarget)) {
        return false;
      }

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();

      void restoreQueuedMessagesToEditor(sidebar, getActiveQueueDisplay()).then((restoredCount) => {
        showToast(restoredCount === 0 ? t("keyboard-shortcuts.toast.noQueuedMessages")
          : t("shortcuts.toast.restoredQueued", { count: restoredCount, plural: restoredCount === 1 ? "" : "s" }));
      }).catch(reportKeyboardError);

      return true;
    },
    (context) => {
      const { event } = context;
      if (!isCreateTabShortcut(event) || !onCreateTab) return false;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      onCreateTab();
      return true;
    },
    (context) => {
      const { event } = context;
      if (!isCloseActiveTabShortcut(event) || !onCloseActiveTab) return false;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      onCloseActiveTab();
      return true;
    },
    (context) => {
      const { event, keyTarget } = context;
      if (!isUndoCloseTabShortcut(event) || !onReopenLastClosed) return false;

      const shouldHandleUndoClose = shouldHandleUndoCloseTabShortcut({
        canUndoCloseTab: canUndoCloseTab?.() ?? true,
        isTextEntry: isTextEntryTarget(keyTarget),
        textEntryHasContent: textEntryHasContent(keyTarget),
        actionToastVisible: isActionToastVisible(),
      });

      if (!shouldHandleUndoClose) return false;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      onReopenLastClosed();
      return true;
    },
    (context) => {
      const { event } = context;
      if (!isReopenLastClosedShortcut(event) || !onReopenLastClosed) return false;

      event.preventDefault();
      onReopenLastClosed();
      return true;
    },
    (context) => {
      const { event, keyTarget } = context;
      const adjacentTabDirection = getAdjacentTabDirectionFromShortcut(event);

      if (
        !adjacentTabDirection
        || !onSwitchAdjacentTab
        || isTextEntryTarget(keyTarget)
        || isInsideSessionTabs(keyTarget)
      ) {
        return false;
      }

      event.preventDefault();
      onSwitchAdjacentTab(adjacentTabDirection);
      return true;
    },
    (context) => {
      const { event, agent } = context;
      if (!(event.shiftKey && event.key === "Tab") || !agent) return false;

      event.preventDefault();
      void cycleThinkingLevel(agent).catch(reportKeyboardError);
      return true;
    },
    (context) => {
      const { event } = context;
      if (!(event.ctrlKey || event.metaKey) || event.key !== "o") return false;

      event.preventDefault();
      const collapsed = document.body.classList.toggle("pi-hide-internals");

      // Collapse/expand tool cards to match the new mode.
      requestAnimationFrame(() => setExcelToolCardsExpanded(!collapsed));

      // Collapse/expand thinking blocks to match the new mode.
      requestAnimationFrame(() =>
        collapsed ? collapseThinkingBlocks() : expandThinkingBlocks(),
      );

      showToast(collapsed ? "Details hidden (⌃O)" : "Details shown (⌃O)", 1500);
      return true;
    },
  ];

  const onKeyDown = (event: KeyboardEvent) => {
    // Command menu takes priority
    if (isCommandMenuVisible() && handleCommandMenuKey(event)) {
      return;
    }

    const context = buildKeydownContext({
      event,
      agent: getActiveConversation(),
      sidebar,
    });

    for (const handler of shortcutHandlers) {
      if (handler(context)) {
        return;
      }
    }

    const { textarea, agent, isInEditor, isStreaming } = context;
    if (!isInEditor || !textarea || event.key !== "Enter" || event.shiftKey) return;
    const text = textarea.value.trim();
    if (text.startsWith("/")) {
      const [name = "", ...args] = text.slice(1).split(" ");
      const result = executeSlashCommand({
        name, args: args.join(" "), busy: isStreaming,
        ...(agent ? { enqueueCommand: (_name: string, instructions: string) => {
          if (agent) void agent.conversation.compact(instructions || undefined, durableContext).catch(reportKeyboardError);
        } } : {}),
        beforeExecute: () => { event.preventDefault(); event.stopImmediatePropagation(); hideCommandMenu(); sidebar.getInput()?.clear(); },
        onError: reportKeyboardError,
      });
      if (result !== "not-found") {
        event.preventDefault(); event.stopImmediatePropagation();
        if (result === "busy-blocked") showToast(t("shortcuts.toast.cantRunBusy", { command: name }));
        return;
      }
    }
    if (!isStreaming || !agent || !text) return;
    event.preventDefault(); event.stopImmediatePropagation();
    if (submitting.has(agent)) return;
    submitting.add(agent);
    const input = sidebar.getInput();
    const whenBusy = event.altKey ? "followUp" : "steer";
    void normalizePromptContent(text, agent.current.model, {}).then(async (content) => {
      await agent.conversation.submit({ type: "input", content, whenBusy }, durableContext);
      // Do not erase text entered while admission was pending.
      if (getActiveConversation() === agent && input?.value.trim() === text) input.clear();
    }).catch(reportKeyboardError).finally(() => submitting.delete(agent));
  };

  document.addEventListener("keydown", onKeyDown, true);
  return () => document.removeEventListener("keydown", onKeyDown, true);
}
