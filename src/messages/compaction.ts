/**
 * Compaction messages.
 *
 * Mirrors the approach used by pi-coding-agent: compaction becomes a first-class
 * custom AgentMessage role so we can:
 * - render it as a dedicated UI card (not an assistant text blob)
 * - keep the summary in LLM context via Agent.convertToLlm
 *
 * pi-agent-core declared this role until 1.0, when it removed its experimental
 * harness. Persisted sessions still contain it, so the shape must not change.
 */

import { t } from "../language/index.js";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { UserMessage } from "@earendil-works/pi-ai";

export interface CompactionSummaryMessage {
  role: "compactionSummary";
  summary: string;
  tokensBefore: number;
  timestamp: number;
  compactedToolCallIds?: string[];
}

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    compactionSummary: CompactionSummaryMessage;
  }
}

export const COMPACTION_SUMMARY_PREFIX =
  "The conversation history before this point was compacted into the following summary:\n\n<summary>\n";

export const COMPACTION_SUMMARY_SUFFIX = "\n</summary>";

export function createCompactionSummaryMessage(args: {
  summary: string;
  tokensBefore: number;
  timestamp: number;
  compactedToolCallIds?: string[];
}): CompactionSummaryMessage {
  return {
    role: "compactionSummary",
    summary: args.summary,
    tokensBefore: args.tokensBefore,
    timestamp: args.timestamp,
    ...(args.compactedToolCallIds && args.compactedToolCallIds.length > 0
      ? { compactedToolCallIds: args.compactedToolCallIds }
      : {}),
  };
}

export function getCompactedToolCallIds(messages: readonly AgentMessage[]): ReadonlySet<string> {
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    const message = messages[i];
    if (message?.role === "compactionSummary") return new Set(message.compactedToolCallIds ?? []);
  }
  return new Set();
}

export function formatCompactionSummaryExtent(msg: CompactionSummaryMessage): string {
  if (typeof msg.tokensBefore === "number") {
    const key = msg.tokensBefore === 1 ? "compaction.extent.token" : "compaction.extent.tokens";
    return t(key, { count: msg.tokensBefore.toLocaleString() });
  }

  const legacyMessageCount = (msg as { messageCountBefore?: unknown }).messageCountBefore;
  if (typeof legacyMessageCount === "number") {
    const key = legacyMessageCount === 1 ? "compaction.extent.message" : "compaction.extent.messages";
    return t(key, { count: legacyMessageCount.toLocaleString() });
  }

  return t("compaction.extent.earlier");
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
