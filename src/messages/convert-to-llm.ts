/**
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono).
 */

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { ImageContent, TextContent } from "@earendil-works/pi-ai";
import type { Message } from "@earendil-works/pi-ai";

import { effectiveRecentToolResultsToKeep } from "../context/window-budgets.js";
import { autoContextToUserMessage } from "./auto-context.js";
import { interruptedResponseToLlm } from "./interrupted-response.js";
import { convertAttachments, isArtifactMessage, isUserMessageWithAttachments } from "./attachments.js";
import { compactionSummaryToUserMessage, getCompactedToolCallIds } from "./compaction.js";
import { shapeToolResultsForLlm } from "./tool-result-shaping.js";

function convertStandardMessagesToLlm(messages: AgentMessage[]): Message[] {
  const converted: Message[] = [];

  for (const message of messages) {
    if (isArtifactMessage(message)) {
      continue;
    }

    if (isUserMessageWithAttachments(message)) {
      const content: (TextContent | ImageContent)[] =
        typeof message.content === "string"
          ? [{ type: "text", text: message.content }]
          : [...message.content];

      if (message.attachments) {
        content.push(...convertAttachments(message.attachments));
      }

      converted.push({
        role: "user",
        content,
        timestamp: message.timestamp,
      });
      continue;
    }

    if (
      message.role === "system"
      || message.role === "user"
      || message.role === "assistant"
      || message.role === "toolResult"
    ) {
      converted.push(message);
    }
  }

  return converted;
}

export function createConvertToLlm(options: {
  getContextWindow?: () => number | undefined;
} = {}): (messages: AgentMessage[]) => Message[] {
  return (messages: AgentMessage[]): Message[] => {
    const normalized: AgentMessage[] = [];

    for (const message of messages) {
      if (message.role === "interruptedResponse") {
        normalized.push(...interruptedResponseToLlm(message));
        continue;
      }
      if (message.role === "archivedMessages") {
        continue;
      }

      if (message.role === "compactionSummary") {
        normalized.push(compactionSummaryToUserMessage(message));
        continue;
      }

      if (message.role === "autoContext") {
        normalized.push(autoContextToUserMessage(message));
        continue;
      }

      normalized.push(message);
    }

    const shaped = shapeToolResultsForLlm(normalized, {
      recentToolResultsToKeep: effectiveRecentToolResultsToKeep(options.getContextWindow?.()),
      compactToolCallIds: getCompactedToolCallIds(messages),
    });
    return convertStandardMessagesToLlm(shaped);
  };
}
