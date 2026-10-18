/**
 * Builtin clipboard commands.
 */

import type { AgentMessage } from "@earendil-works/pi-agent-core";

import type { SlashCommand } from "../types.js";
import type { ActiveConversationProvider } from "./model.js";
import { showToast } from "../../ui/toast.js";
import { extractTextBlocks } from "../../utils/content.js";
import { t } from "../../language/index.js";

function getLastAssistantText(messages: AgentMessage[]): string | null {
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg?.role === "assistant") {
      const text = extractTextBlocks(msg.content).trim();
      return text || null;
    }
  }
  return null;
}

export function createClipboardCommands(getActiveConversation: ActiveConversationProvider): SlashCommand[] {
  return [
    {
      name: "copy",
      description: t("command.clipboard.desc"),
      source: "builtin",
      execute: () => {
        const agent = getActiveConversation();
        if (!agent) {
          showToast(t("command.clipboard.no_session"));
          return;
        }

        const text = getLastAssistantText(agent.current.messages);
        if (text) {
          void navigator.clipboard.writeText(text).then(() => {
            showToast(t("clipboard.copied"));
          });
          return;
        }

        showToast(t("clipboard.noAgentMessage"));
      },
    },
  ];
}
