/**
 * Compaction messages.
 *
 * Mirrors the approach used by pi-coding-agent: compaction becomes a first-class
 * custom AgentMessage role so we can:
 * - render it as a dedicated UI card (not an assistant text blob)
 * - keep the summary in LLM context via Agent.convertToLlm
 */

import type { CompactionSummaryMessage } from "@earendil-works/pi-agent-core";
import type { UserMessage } from "@earendil-works/pi-ai";

export type { CompactionSummaryMessage };

export const COMPACTION_SUMMARY_PREFIX =
  "The conversation history before this point was compacted into the following summary:\n\n<summary>\n";

export const COMPACTION_SUMMARY_SUFFIX = "\n</summary>";

export function createCompactionSummaryMessage(args: {
  summary: string;
  tokensBefore: number;
  timestamp: number;
}): CompactionSummaryMessage {
  return {
    role: "compactionSummary",
    summary: args.summary,
    tokensBefore: args.tokensBefore,
    timestamp: args.timestamp,
  };
}

export function formatCompactionSummaryExtent(msg: CompactionSummaryMessage): string {
  if (typeof msg.tokensBefore === "number") {
    return `${msg.tokensBefore.toLocaleString()} token${msg.tokensBefore === 1 ? "" : "s"}`;
  }

  const legacyMessageCount = (msg as { messageCountBefore?: unknown }).messageCountBefore;
  if (typeof legacyMessageCount === "number") {
    return `${legacyMessageCount.toLocaleString()} message${legacyMessageCount === 1 ? "" : "s"}`;
  }

  return "earlier context";
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
