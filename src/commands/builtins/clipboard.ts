/**
 * Builtin clipboard commands.
 */

import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";

import type { SlashCommand } from "../types.js";
import type { ActiveAgentProvider } from "./model.js";
import { showToast } from "../../ui/toast.js";
import { extractTextBlocks } from "../../utils/content.js";

function getLastAssistantText(messages: AgentMessage[]): string | null {
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg.role === "assistant") {
      const text = extractTextBlocks(msg.content).trim();
      return text || null;
    }
  }
  return null;
}

function resolveAgent(getActiveAgent: ActiveAgentProvider): Agent | null {
  return getActiveAgent();
}

export function createClipboardCommands(getActiveAgent: ActiveAgentProvider): SlashCommand[] {
  return [
    {
      name: "copy",
      description: "Copy last agent message to clipboard",
      source: "builtin",
      execute: () => {
        const agent = resolveAgent(getActiveAgent);
        if (!agent) {
          showToast("No active session");
          return;
        }

        const text = getLastAssistantText(agent.state.messages);
        if (text) {
          void navigator.clipboard.writeText(text).then(() => {
            showToast("Copied to clipboard");
          });
          return;
        }

        showToast("No agent message to copy");
      },
    },
  ];
}
