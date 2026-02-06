/**
 * Keyboard shortcuts + key-driven UX.
 *
 * Extracted from taskpane.ts to keep the entrypoint thin.
 */

import type { Agent, AgentMessage, ThinkingLevel } from "@mariozechner/pi-agent-core";
import { supportsXhigh } from "@mariozechner/pi-ai";

import type { PiSidebar } from "../ui/pi-sidebar.js";
import { showToast } from "../ui/toast.js";

import { commandRegistry } from "../commands/types.js";
import {
  handleCommandMenuKey,
  hideCommandMenu,
  isCommandMenuVisible,
} from "../commands/command-menu.js";

import { flashThinkingLevel, updateStatusBar } from "./status-bar.js";

type QueueDisplay = {
  add: (type: "steer" | "follow-up", text: string) => void;
};

const THINKING_COLORS: Record<ThinkingLevel, string> = {
  off: "#a0a0a0",
  minimal: "#767676",
  low: "#4488cc",
  medium: "#22998a",
  high: "#875f87",
  xhigh: "#8b008b",
};

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
  updateStatusBar(agent);
  flashThinkingLevel(next, THINKING_COLORS[next] || "#a0a0a0");

  return next;
}

export function installKeyboardShortcuts(opts: {
  agent: Agent;
  sidebar: PiSidebar;
  queueDisplay: QueueDisplay;
  markUserAborted: () => void;
}): () => void {
  const { agent, sidebar, queueDisplay, markUserAborted } = opts;

  const onKeyDown = (e: KeyboardEvent) => {
    // Command menu takes priority
    if (isCommandMenuVisible()) {
      if (handleCommandMenuKey(e)) return;
    }

    const textarea = sidebar.getTextarea();
    const isInEditor = Boolean(
      textarea && (e.target === textarea || textarea.contains(e.target as Node)),
    );
    const isStreaming = agent.state.isStreaming;

    // ESC — dismiss command menu
    if (e.key === "Escape" && isCommandMenuVisible()) {
      e.preventDefault();
      hideCommandMenu();
      return;
    }

    // ESC — abort
    if (e.key === "Escape" && isStreaming) {
      e.preventDefault();
      markUserAborted();
      agent.abort();
      return;
    }

    // Shift+Tab — cycle thinking level
    if (e.shiftKey && e.key === "Tab") {
      e.preventDefault();
      cycleThinkingLevel(agent);
      return;
    }

    // Ctrl+O — toggle thinking/tool visibility
    if ((e.ctrlKey || e.metaKey) && e.key === "o") {
      e.preventDefault();
      const collapsed = document.body.classList.toggle("pi-hide-internals");
      showToast(collapsed ? "Details hidden (⌃O)" : "Details shown (⌃O)", 1500);
      return;
    }

    // Slash command execution
    if (
      isInEditor &&
      textarea &&
      e.key === "Enter" &&
      !e.shiftKey &&
      textarea.value.startsWith("/") &&
      !isStreaming
    ) {
      const val = textarea.value.trim();
      const spaceIdx = val.indexOf(" ");
      const cmdName = spaceIdx > 0 ? val.slice(1, spaceIdx) : val.slice(1);
      const args = spaceIdx > 0 ? val.slice(spaceIdx + 1) : "";
      const cmd = commandRegistry.get(cmdName);
      if (cmd) {
        e.preventDefault();
        e.stopImmediatePropagation();
        hideCommandMenu();
        const input = sidebar.getInput();
        if (input) input.clear();
        cmd.execute(args);
        return;
      }
    }

    // Enter/Alt+Enter while streaming — steer or follow-up
    if (isInEditor && textarea && e.key === "Enter" && !e.shiftKey && isStreaming) {
      const text = textarea.value.trim();
      if (!text) return;

      e.preventDefault();
      e.stopImmediatePropagation();

      const msg: AgentMessage = {
        role: "user",
        content: [{ type: "text", text }],
        timestamp: Date.now(),
      };

      if (e.altKey) {
        agent.followUp(msg);
        queueDisplay.add("follow-up", text);
      } else {
        agent.steer(msg);
        queueDisplay.add("steer", text);
      }

      const input = sidebar.getInput();
      if (input) input.clear();
      return;
    }
  };

  document.addEventListener("keydown", onKeyDown, true);
  return () => document.removeEventListener("keydown", onKeyDown, true);
}
