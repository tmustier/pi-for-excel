/**
 * What a submitted composer draft becomes. Enter, Alt+Enter, the send button
 * and a command chosen from the menu all come through here, so they behave
 * the same.
 */

import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";

import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";
import { hideCommandMenu } from "../../commands/command-menu.js";
import { executeSlashCommand } from "../../commands/slash-command-execution.js";

export type QueueDisplay = {
  add: (type: "steer" | "follow-up", text: string, message?: AgentMessage) => void;
};

export type ActionQueue = {
  enqueuePrompt: (text: string) => void;
  enqueueCommand: (name: string, args: string) => void;
  isBusy: () => boolean;
};

export interface SubmitTarget {
  agent: Agent;
  actionQueue: ActionQueue;
  queueDisplay: QueueDisplay;
}

/**
 * Handles a submitted draft for the active tab. Returns true when the draft
 * was used (the composer clears) and false when it stays for editing.
 *
 * - A leading `/` names a command. Known commands run (or are queued, for
 *   `/compact`) under the busy policy; a refused or unknown command keeps the
 *   draft and says why. A leading space sends slash text as a message.
 * - While Pi responds, Enter steers and Alt+Enter follows up.
 * - While the tab is busy otherwise (compacting, waiting to retry), the text
 *   is queued; when idle it is sent.
 */
export function submitComposerText(args: {
  text: string;
  followUp: boolean;
  target: SubmitTarget | null;
}): boolean {
  const { followUp, target } = args;
  const text = args.text.trim();
  if (!text) return false;

  if (!target) {
    showToast(t("editor-actions.toast.noActiveSession"));
    return false;
  }

  const { agent, actionQueue, queueDisplay } = target;
  const isStreaming = agent.state.isStreaming;

  if (args.text.startsWith("/")) {
    const spaceIdx = text.search(/\s/u);
    const name = spaceIdx > 0 ? text.slice(1, spaceIdx) : text.slice(1);
    const commandArgs = spaceIdx > 0 ? text.slice(spaceIdx + 1).trim() : "";

    const result = executeSlashCommand({
      name,
      args: commandArgs,
      busy: isStreaming || actionQueue.isBusy(),
      enqueueCommand: (commandName, queuedArgs) => {
        actionQueue.enqueueCommand(commandName, queuedArgs);
      },
      beforeExecute: () => {
        hideCommandMenu();
      },
      onError: () => {
        showToast(t("slash-command.toast.failed"));
      },
    });

    if (result === "not-found") {
      showToast(t("turn.toast.unknownCommand", { command: name }));
      return false;
    }
    if (result === "busy-blocked") {
      showToast(t("shortcuts.toast.cantRunBusy", { command: name }));
      return false;
    }
    return true;
  }

  if (isStreaming) {
    const message: AgentMessage = {
      role: "user",
      content: [{ type: "text", text }],
      timestamp: Date.now(),
    };
    if (followUp) {
      agent.followUp(message);
      queueDisplay.add("follow-up", text, message);
    } else {
      agent.steer(message);
      queueDisplay.add("steer", text, message);
    }
    return true;
  }

  actionQueue.enqueuePrompt(text);
  return true;
}
