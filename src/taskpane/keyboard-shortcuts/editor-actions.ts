/**
 * Editor-focused keyboard actions (slash commands + streaming steer/follow-up).
 */

import type { Agent, AgentMessage } from "@mariozechner/pi-agent-core";

import type { PiSidebar } from "../../ui/pi-sidebar.js";
import { showToast } from "../../ui/toast.js";
import { hideCommandMenu } from "../../commands/command-menu.js";
import { executeSlashCommand } from "../../commands/slash-command-execution.js";

export type QueuedEditorMessage = {
  type: "steer" | "follow-up";
  text: string;
};

export type QueueDisplay = {
  add: (type: "steer" | "follow-up", text: string) => void;
  drainQueuedMessages: () => QueuedEditorMessage[];
};

export type QueuedEditorAction =
  | { type: "prompt"; text: string }
  | { type: "command"; name: string; args: string };

export type ActionQueue = {
  enqueueCommand: (name: string, args: string) => void;
  drainQueuedActions: () => QueuedEditorAction[];
  isBusy: () => boolean;
};

interface SidebarInputHost {
  getInput: () => { value: string } | undefined;
  getTextarea: () => HTMLTextAreaElement | undefined;
}

function mergeQueuedAndCurrentText(queuedTexts: string[], currentText: string): string {
  return [...queuedTexts, currentText]
    .filter((text) => text.trim().length > 0)
    .join("\n\n");
}

export function restoreQueuedMessagesToEditor(args: {
  sidebar: SidebarInputHost;
  textarea: HTMLTextAreaElement | undefined;
  agent: Agent | null;
  queueDisplay: QueueDisplay | null;
  actionQueue: ActionQueue | null;
  requeueCommands?: boolean;
}): number {
  const {
    sidebar,
    textarea,
    agent,
    queueDisplay,
    actionQueue,
    requeueCommands = true,
  } = args;

  const queuedMessages = queueDisplay?.drainQueuedMessages() ?? [];
  if (queuedMessages.length > 0) {
    agent?.clearAllQueues();
  }

  const queuedActions = actionQueue?.drainQueuedActions() ?? [];
  const queuedPromptTexts: string[] = [];
  const queuedCommands: Array<{ name: string; args: string }> = [];

  for (const action of queuedActions) {
    if (action.type === "prompt") {
      queuedPromptTexts.push(action.text);
      continue;
    }

    queuedCommands.push({ name: action.name, args: action.args });
  }

  if (requeueCommands && actionQueue) {
    for (const command of queuedCommands) {
      actionQueue.enqueueCommand(command.name, command.args);
    }
  }

  const queuedTexts = [
    ...queuedMessages.map((message) => message.text),
    ...queuedPromptTexts,
  ];

  if (queuedTexts.length === 0) {
    return 0;
  }

  const input = sidebar.getInput();
  const currentText = textarea?.value ?? input?.value ?? "";
  const mergedText = mergeQueuedAndCurrentText(queuedTexts, currentText);

  if (input) {
    input.value = mergedText;
  } else if (textarea) {
    textarea.value = mergedText;
  }

  const activeTextarea = sidebar.getTextarea() ?? textarea;
  if (activeTextarea) {
    if (typeof activeTextarea.focus === "function") {
      activeTextarea.focus();
    }

    if (typeof activeTextarea.setSelectionRange === "function") {
      const cursor = activeTextarea.value.length;
      activeTextarea.setSelectionRange(cursor, cursor);
    }
  }

  return queuedTexts.length;
}

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
  const actionQueue = getActiveActionQueue();
  const busy = isStreaming || actionQueue?.isBusy() === true;

  const result = executeSlashCommand({
    name: cmdName,
    args: argsText,
    busy,
    enqueueCommand: actionQueue
      ? (name: string, args: string) => {
        actionQueue.enqueueCommand(name, args);
      }
      : undefined,
    beforeExecute: () => {
      event.preventDefault();
      event.stopImmediatePropagation();
      hideCommandMenu();

      const input = sidebar.getInput();
      if (input) input.clear();
    },
  });

  if (result === "not-found") {
    return false;
  }

  if (result === "busy-blocked") {
    event.preventDefault();
    event.stopImmediatePropagation();
    showToast(`Can't run /${cmdName} while Pi is busy`);
    return true;
  }

  if (result === "missing-queue") {
    showToast("No active session");
    return true;
  }

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
