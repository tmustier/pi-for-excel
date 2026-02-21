/**
 * Keyboard shortcuts + key-driven UX.
 *
 * Extracted from taskpane.ts to keep the entrypoint thin.
 */

import type { Agent, ThinkingLevel } from "@mariozechner/pi-agent-core";
import { supportsXhigh } from "@mariozechner/pi-ai";

import type { PiSidebar } from "../ui/pi-sidebar.js";
import { moveCursorToEnd } from "../ui/input-focus.js";
import { isActionToastVisible, showToast } from "../ui/toast.js";

import { doesUiClaimStreamingEscape } from "../utils/escape-guard.js";
import { blurTextEntryTarget, isTextEntryTarget } from "../utils/text-entry.js";
import {
  handleCommandMenuKey,
  hideCommandMenu,
  isCommandMenuVisible,
} from "../commands/command-menu.js";

import { flashThinkingLevel, updateStatusBarForAgent } from "./status-bar.js";
import {
  handleSlashCommandExecution,
  handleStreamingSteerOrFollowUp,
  type ActionQueue,
  type QueueDisplay,
} from "./keyboard-shortcuts/editor-actions.js";
import { shouldHandleUndoCloseTabShortcut, textEntryHasContent } from "./keyboard-shortcuts/guards.js";
import {
  getAdjacentTabDirectionFromShortcut,
  isCloseActiveTabShortcut,
  isCreateTabShortcut,
  isFocusInputShortcut,
  isReopenLastClosedShortcut,
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
  isUndoCloseTabShortcut,
  shouldAbortFromEscape,
  shouldBlurEditorFromEscape,
} from "./keyboard-shortcuts/matchers.js";
export { shouldHandleUndoCloseTabShortcut } from "./keyboard-shortcuts/guards.js";

interface KeydownContext {
  event: KeyboardEvent;
  agent: Agent | null;
  textarea: HTMLTextAreaElement | undefined;
  keyTarget: Node | null;
  isInEditor: boolean;
  isStreaming: boolean;
}

type ShortcutHandler = (context: KeydownContext) => boolean;

const THINKING_COLORS: Record<ThinkingLevel, string> = {
  off: "#a0a0a0",
  minimal: "#767676",
  low: "#4488cc",
  medium: "#22998a",
  high: "#875f87",
  xhigh: "#8b008b",
};

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

    if (expanded) {
      body.classList.remove("max-h-0");
      body.classList.add("max-h-[2000px]", "mt-3");
    } else {
      body.classList.remove("max-h-[2000px]", "mt-3");
      body.classList.add("max-h-0");
    }

    const up = toolMessage.querySelector<HTMLElement>(".chevron-up");
    const down = toolMessage.querySelector<HTMLElement>(".chevrons-up-down");
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
  agent: Agent | null;
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
    isStreaming: agent?.state.isStreaming ?? false,
  };
}

export function getThinkingLevels(agent: Agent): ThinkingLevel[] {
  const model = agent.state.model;
  if (!model || !model.reasoning) return ["off"];

  const provider = model.provider;
  if (provider === "openai" || provider === "openai-codex") {
    const levels: ThinkingLevel[] = ["off", "minimal", "low", "medium", "high"];
    if (supportsXhigh(model)) levels.push("xhigh");
    return levels;
  }

  if (provider === "anthropic") {
    const levels: ThinkingLevel[] = ["off", "low", "medium", "high"];
    if (supportsXhigh(model)) levels.push("xhigh");
    return levels;
  }

  return ["off", "low", "medium", "high"];
}

export function cycleThinkingLevel(agent: Agent): ThinkingLevel {
  const levels = getThinkingLevels(agent);
  const current = agent.state.thinkingLevel;
  const idx = levels.indexOf(current);
  const next = levels[(idx >= 0 ? idx + 1 : 0) % levels.length];

  agent.setThinkingLevel(next);
  updateStatusBarForAgent(agent);
  flashThinkingLevel(next, THINKING_COLORS[next] || "#a0a0a0");

  return next;
}

export function installKeyboardShortcuts(opts: {
  getActiveAgent: () => Agent | null;
  getActiveQueueDisplay: () => QueueDisplay | null;
  getActiveActionQueue: () => ActionQueue | null;
  sidebar: PiSidebar;
  markUserAborted: (agent: Agent) => void;
  onCreateTab?: () => void;
  onCloseActiveTab?: () => void;
  onReopenLastClosed?: () => void;
  canUndoCloseTab?: () => boolean;
  onSwitchAdjacentTab?: (direction: -1 | 1) => void;
}): () => void {
  const {
    getActiveAgent,
    getActiveQueueDisplay,
    getActiveActionQueue,
    sidebar,
    markUserAborted,
    onCreateTab,
    onCloseActiveTab,
    onReopenLastClosed,
    canUndoCloseTab,
    onSwitchAdjacentTab,
  } = opts;

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
      markUserAborted(agent);
      agent.abort();
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
      cycleThinkingLevel(agent);
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
      agent: getActiveAgent(),
      sidebar,
    });

    for (const handler of shortcutHandlers) {
      if (handler(context)) {
        return;
      }
    }

    if (
      handleSlashCommandExecution({
        event,
        textarea: context.textarea,
        isInEditor: context.isInEditor,
        isStreaming: context.isStreaming,
        getActiveActionQueue,
        sidebar,
      })
    ) {
      return;
    }

    void handleStreamingSteerOrFollowUp({
      event,
      textarea: context.textarea,
      isInEditor: context.isInEditor,
      isStreaming: context.isStreaming,
      agent: context.agent,
      getActiveQueueDisplay,
      sidebar,
    });
  };

  document.addEventListener("keydown", onKeyDown, true);
  return () => document.removeEventListener("keydown", onKeyDown, true);
}
