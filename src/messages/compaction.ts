/**
 * Compaction messages.
 *
 * Mirrors the approach used by pi-coding-agent: compaction becomes a first-class
 * custom AgentMessage role so we can:
 * - render it as a dedicated UI card (not an assistant text blob)
 * - keep the summary in LLM context via Agent.convertToLlm
 */

import type { UserMessage } from "@earendil-works/pi-ai";

export const COMPACTION_SUMMARY_PREFIX =
  "The conversation history before this point was compacted into the following summary:\n\n<summary>\n";

export const COMPACTION_SUMMARY_SUFFIX = "\n</summary>";

export interface CompactionSummaryMessage {
  role: "compactionSummary";
  summary: string;
  messageCountBefore: number;
  timestamp: number;
}

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    compactionSummary: CompactionSummaryMessage;
  }
}

export function createCompactionSummaryMessage(args: {
  summary: string;
  messageCountBefore: number;
  timestamp: number;
}): CompactionSummaryMessage {
  return {
    role: "compactionSummary",
    summary: args.summary,
    messageCountBefore: args.messageCountBefore,
    timestamp: args.timestamp,
  };
}

export function compactionSummaryToUserMessage(
  msg: CompactionSummaryMessage,
): UserMessage {
  return {
    role: "user",
    content: [
      {
        type: "text",
        text: COMPACTION_SUMMARY_PREFIX + msg.summary + COMPACTION_SUMMARY_SUFFIX,
      },
    ],
    timestamp: msg.timestamp,
  };
}
