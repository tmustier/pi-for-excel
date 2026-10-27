/**
 * Keyboard shortcuts + key-driven UX.
 *
 * Extracted from taskpane.ts to keep the entrypoint thin.
 */

import type { Agent, ThinkingLevel } from "@earendil-works/pi-agent-core";

import { getThinkingLevelsForModel } from "../models/thinking-levels.js";
import type { PiSidebar } from "../ui/pi-sidebar.js";
import { moveCursorToEnd } from "../ui/input-focus.js";
import { isModalOverlayOpen } from "../ui/overlay-dialog.js";
import { isActionToastVisible, showToast } from "../ui/toast.js";
import { t } from "../language/index.js";
import { formatDetailsShortcut, toggleDetailsExpanded } from "../ui/messages/transcript-state.js";

import { doesUiClaimStreamingEscape } from "../utils/escape-guard.js";
import { isDialogOverlayOpen } from "../ui/overlay-escape.js";
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
  agent: Agent | null;
  textarea: HTMLTextAreaElement | undefined;
  keyTarget: Node | null;
  isInEditor: boolean;
  /** The active tab is busy: responding, compacting or waiting to retry. */
  isBusy: boolean;
}

type ShortcutHandler = (context: KeydownContext) => boolean;

function isInsideSessionTabs(target: EventTarget | null | undefined): boolean {
  if (!(target instanceof Element)) {
    return false;
  }

  return target.closest(".pi-session-tabs") !== null;
}

/**
 * Menus and status popovers own plain arrow keys (see `menu-keyboard.ts`);
 * explicit tab chords such as Ctrl+PageUp still switch tabs from them.
 */
function isPlainArrowInsideMenu(event: KeyboardEvent, target: EventTarget | null | undefined): boolean {
  const plain = !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey;
  return plain
    && (event.key === "ArrowLeft" || event.key === "ArrowRight")
    && target instanceof Element
    && target.closest("[role=\"menu\"], .pi-status-popover") !== null;
}

function buildKeydownContext(args: {
  event: KeyboardEvent;
  agent: Agent | null;
  isBusy: boolean;
  sidebar: PiSidebar;
}): KeydownContext {
  const { event, agent, isBusy, sidebar } = args;
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
    isBusy,
  };
}

export function getThinkingLevels(agent: Agent): ThinkingLevel[] {
  return getThinkingLevelsForModel(agent.state.model);
}

export function cycleThinkingLevel(agent: Agent): ThinkingLevel {
  const levels = getThinkingLevels(agent);
  const current = agent.state.thinkingLevel;
  const idx = levels.indexOf(current);
  const next = levels[(idx >= 0 ? idx + 1 : 0) % levels.length] ?? "off";

  agent.state.thinkingLevel = next;
  // Full re-render with the real mode and lock state (rendering from the agent
  // alone would reset the mode badge to its default).
  document.dispatchEvent(new CustomEvent("pi:status-update"));
  flashThinkingLevel(next, THINKING_LEVEL_COLORS[next]);

  return next;
}

export function installKeyboardShortcuts(opts: {
  getActiveAgent: () => Agent | null;
  isActiveTabBusy: () => boolean;
  sidebar: PiSidebar;
  /** Stops the active tab, exactly like the Stop button. */
  stopActiveTab: () => void;
  /** Alt+Up: puts the last queued message back in the composer; false when there is none. */
  editLastQueued: () => boolean;
  onCreateTab?: () => void;
  onCloseActiveTab?: () => void;
  onReopenLastClosed?: () => void;
  canUndoCloseTab?: () => boolean;
  onSwitchAdjacentTab?: (direction: -1 | 1) => void;
}): () => void {
  const {
    getActiveAgent,
    isActiveTabBusy,
    sidebar,
    stopActiveTab,
    editLastQueued,
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
      event.stopImmediatePropagation();
      hideCommandMenu();
      return true;
    },
    (context) => {
      const { event, isInEditor, isBusy, keyTarget } = context;
      if (!shouldBlurEditorFromEscape({ key: event.key, isInEditor, isStreaming: isBusy })) return false;

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
      const { event, isBusy, agent, keyTarget } = context;
      const isEscapeKey = event.key === "Escape" || event.key === "Esc";
      const escapeClaimedByOverlay = isEscapeKey && doesUiClaimStreamingEscape(keyTarget);

      if (
        !isEscapeKey
        || !shouldAbortFromEscape({
          isStreaming: isBusy,
          hasAgent: agent !== null,
          escapeClaimedByOverlay,
        })
      ) {
        return false;
      }

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      stopActiveTab();
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
      editLastQueued();
      return true;
    },
    (context) => {
      const { event } = context;
      // Tab shortcuts never act on the tabs hidden behind an open dialog.
      if (!isCreateTabShortcut(event) || !onCreateTab || isDialogOverlayOpen()) return false;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      onCreateTab();
      return true;
    },
    (context) => {
      const { event } = context;
      if (!isCloseActiveTabShortcut(event) || !onCloseActiveTab || isDialogOverlayOpen()) return false;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      onCloseActiveTab();
      return true;
    },
    (context) => {
      const { event, keyTarget } = context;
      if (!isUndoCloseTabShortcut(event) || !onReopenLastClosed || isDialogOverlayOpen()) return false;

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
      if (!isReopenLastClosedShortcut(event) || !onReopenLastClosed || isDialogOverlayOpen()) return false;

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
        || isDialogOverlayOpen()
        // A focused tab handles plain arrows itself; chords still switch here.
        || (isInsideSessionTabs(keyTarget) && !(event.metaKey || event.ctrlKey || event.altKey || event.shiftKey))
        || isPlainArrowInsideMenu(event, keyTarget)
      ) {
        return false;
      }

      event.preventDefault();
      onSwitchAdjacentTab(adjacentTabDirection);
      return true;
    },
    (context) => {
      const { event, agent, textarea, keyTarget } = context;
      // Only from the composer's text field: everywhere else, including the
      // composer's own buttons, Shift+Tab moves focus backwards.
      if (!(event.shiftKey && event.key === "Tab") || !agent || !textarea || keyTarget !== textarea) return false;

      // Office on the web also moves focus on Shift+Tab from a window listener;
      // the key must do exactly one thing.
      event.preventDefault();
      event.stopPropagation();
      cycleThinkingLevel(agent);
      return true;
    },
    (context) => {
      const { event } = context;
      if (!(event.ctrlKey || event.metaKey) || event.key !== "o") return false;

      event.preventDefault();
      const expanded = toggleDetailsExpanded();
      const shortcut = formatDetailsShortcut();
      showToast(
        expanded
          ? t("transcript.details.expanded", { shortcut })
          : t("transcript.details.collapsed", { shortcut }),
        1500,
      );
      return true;
    },
  ];

  const onKeyDown = (event: KeyboardEvent) => {
    // Keys that belong to an IME composition (including the Enter that
    // commits a candidate) are the input method's, not shortcuts.
    if (event.isComposing || event.keyCode === 229) return;

    // A modal overlay owns the keyboard. Its own handlers take Escape and Tab;
    // taskpane shortcuts must not act on the tabs or composer behind it.
    if (isModalOverlayOpen()) {
      return;
    }

    const context = buildKeydownContext({
      event,
      agent: getActiveAgent(),
      isBusy: isActiveTabBusy(),
      sidebar,
    });

    // The command menu is a combobox on the composer: it only takes keys typed there.
    if (isCommandMenuVisible() && context.isInEditor && handleCommandMenuKey(event)) {
      return;
    }

    for (const handler of shortcutHandlers) {
      if (handler(context)) {
        return;
      }
    }
  };

  document.addEventListener("keydown", onKeyDown, true);
  return () => document.removeEventListener("keydown", onKeyDown, true);
}
