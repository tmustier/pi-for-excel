/**
 * convertToLlm() for the Excel taskpane agent.
 *
 * We mostly reuse pi-web-ui's default conversion (attachments, artifact filtering),
 * but extend it with:
 * - custom compaction summary message support
 * - model-facing shaping of older large tool results
 */

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Message } from "@earendil-works/pi-ai";
import { defaultConvertToLlm } from "@earendil-works/pi-web-ui/dist/components/Messages.js";

import { compactionSummaryToUserMessage } from "./compaction.js";
import { shapeToolResultsForLlm } from "./tool-result-shaping.js";

export function convertToLlm(messages: AgentMessage[]): Message[] {
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

  const shaped = shapeToolResultsForLlm(normalized);
  return defaultConvertToLlm(shaped);
}
