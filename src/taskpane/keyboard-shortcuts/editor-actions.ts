/**
 * Editor-focused keyboard actions (slash commands + streaming steer/follow-up).
 */

import type { Agent, AgentMessage } from "@mariozechner/pi-agent-core";

import type { PiSidebar } from "../../ui/pi-sidebar.js";
import { showToast } from "../../ui/toast.js";
import { hideCommandMenu } from "../../commands/command-menu.js";
import { isBusyAllowedCommand } from "../../commands/busy-command-policy.js";
import { commandRegistry } from "../../commands/types.js";

export type QueueDisplay = {
  add: (type: "steer" | "follow-up", text: string) => void;
};

export type ActionQueue = {
  enqueueCommand: (name: string, args: string) => void;
  isBusy: () => boolean;
};

export function handleSlashCommandExecution(args: {
  event: KeyboardEvent;
  textarea: HTMLTextAreaElement | undefined;
  isInEditor: boolean;
  isStreaming: boolean;
  getActiveActionQueue: () => ActionQueue | null;
  sidebar: PiSidebar;
}): boolean {
  const {
    event,
    textarea,
    isInEditor,
    isStreaming,
    getActiveActionQueue,
    sidebar,
  } = args;

  if (
    !isInEditor
    || !textarea
    || event.key !== "Enter"
    || event.shiftKey
    || !textarea.value.startsWith("/")
  ) {
    return false;
  }

  const val = textarea.value.trim();
  const spaceIdx = val.indexOf(" ");
  const cmdName = spaceIdx > 0 ? val.slice(1, spaceIdx) : val.slice(1);
  const argsText = spaceIdx > 0 ? val.slice(spaceIdx + 1) : "";
  const cmd = commandRegistry.get(cmdName);
  if (!cmd) return false;

  const actionQueue = getActiveActionQueue();
  const busy = isStreaming || actionQueue?.isBusy() === true;

  if (busy && !isBusyAllowedCommand(cmd)) {
    event.preventDefault();
    event.stopImmediatePropagation();
    showToast(`Can't run /${cmdName} while Pi is busy`);
    return true;
  }

  event.preventDefault();
  event.stopImmediatePropagation();
  hideCommandMenu();

  const input = sidebar.getInput();
  if (input) input.clear();

  if (cmdName === "compact") {
    if (!actionQueue) {
      showToast("No active session");
      return true;
    }

    actionQueue.enqueueCommand(cmdName, argsText);
    return true;
  }

  void cmd.execute(argsText);
  return true;
}

export function handleStreamingSteerOrFollowUp(args: {
  event: KeyboardEvent;
  textarea: HTMLTextAreaElement | undefined;
  isInEditor: boolean;
  isStreaming: boolean;
  agent: Agent | null;
  getActiveQueueDisplay: () => QueueDisplay | null;
  sidebar: PiSidebar;
}): boolean {
  const {
    event,
    textarea,
    isInEditor,
    isStreaming,
    agent,
    getActiveQueueDisplay,
    sidebar,
  } = args;

  if (!isInEditor || !textarea || event.key !== "Enter" || event.shiftKey || !isStreaming || !agent) {
    return false;
  }

  const text = textarea.value.trim();
  if (!text) return false;

  event.preventDefault();
  event.stopImmediatePropagation();

  const msg: AgentMessage = {
    role: "user",
    content: [{ type: "text", text }],
    timestamp: Date.now(),
  };

  if (event.altKey) {
    agent.followUp(msg);
    getActiveQueueDisplay()?.add("follow-up", text);
  } else {
    agent.steer(msg);
    getActiveQueueDisplay()?.add("steer", text);
  }

  const input = sidebar.getInput();
  if (input) input.clear();
  return true;
}
