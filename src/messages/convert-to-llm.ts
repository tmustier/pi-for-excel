/**
 * convertToLlm() for the Excel taskpane agent.
 *
 * Standard conversion (attachments, artifact filtering) is first-party
 * (vendored from pi-web-ui 0.75.3 during the UI ownership migration —
 * see docs/ui-ownership.md), extended with:
 * - custom compaction summary message support
 * - model-facing shaping of older large tool results (scaled to the active
 *   model's context window)
 */

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { ImageContent, TextContent } from "@earendil-works/pi-ai";
import type { Message } from "@earendil-works/pi-ai/compat";

import { effectiveRecentToolResultsToKeep } from "../context/window-budgets.js";
import { convertAttachments, isArtifactMessage, isUserMessageWithAttachments } from "./attachments.js";
import { compactionSummaryToUserMessage } from "./compaction.js";
import { shapeToolResultsForLlm } from "./tool-result-shaping.js";

/**
 * Standard AgentMessage → LLM Message conversion:
 * - artifact messages are UI/session-only and filtered out
 * - user-with-attachments becomes a plain user message with content blocks
 * - user/assistant/toolResult pass through
 * - unknown custom roles are dropped
 */
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

    if (message.role === "user" || message.role === "assistant" || message.role === "toolResult") {
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
      if (message.role === "archivedMessages") {
        // UI-only history bucket, never sent to the model.
        continue;
      }

      if (message.role === "compactionSummary") {
        normalized.push(compactionSummaryToUserMessage(message));
        continue;
      }

      normalized.push(message);
    }

    const shaped = shapeToolResultsForLlm(normalized, {
      recentToolResultsToKeep: effectiveRecentToolResultsToKeep(options.getContextWindow?.()),
    });
    return convertStandardMessagesToLlm(shaped);
  };
}
