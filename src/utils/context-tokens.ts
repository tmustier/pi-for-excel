/**
 * Context window token estimation utilities.
 *
 * We use the same conservative heuristic as pi-coding-agent: tokens ≈ chars / 4.
 * When we have provider-reported `usage` for the last assistant turn, we use it
 * as an anchor because it already reflects prompt caching and provider-side
 * tokenization.
 */

import type { Usage } from "@earendil-works/pi-ai";
import type { AgentMessage, AgentState } from "@earendil-works/pi-agent-core";

const CHARS_PER_TOKEN = 4;

export function estimateTextTokens(text: string): number {
  return Math.ceil(text.length / CHARS_PER_TOKEN);
}

/**
 * Total tokens that count against the model's context window.
 *
 * For providers with prompt caching (e.g. Anthropic), `usage.input` may exclude
 * cached prompt tokens. `cacheRead`/`cacheWrite` still count towards context.
 */
export function calculateContextTokens(u: Usage): number {
  return u.totalTokens || u.input + u.output + u.cacheRead + u.cacheWrite;
}

export function estimateMessageTokens(message: AgentMessage): number {
  let chars = 0;

  if (message.role === "artifact") {
    // UI-only, not part of LLM context (defaultConvertToLlm filters it out).
    return 0;
  }

  if (message.role === "compactionSummary") {
    return estimateTextTokens(message.summary);
  }

  if (message.role === "user" || message.role === "user-with-attachments") {
    const content = message.content;
    if (typeof content === "string") {
      chars += content.length;
    } else if (Array.isArray(content)) {
      for (const block of content) {
        if (block.type === "text") chars += block.text.length;
        if (block.type === "image") chars += 4800; // ~1200 tokens
      }
    }
    return Math.ceil(chars / CHARS_PER_TOKEN);
  }

  if (message.role === "assistant") {
    for (const block of message.content) {
      if (block.type === "text") chars += block.text.length;
      else if (block.type === "thinking") chars += block.thinking.length;
      else if (block.type === "toolCall") {
        chars += block.name.length;
        try {
          chars += JSON.stringify(block.arguments).length;
        } catch {
          // ignore
        }
      }
    }
    return Math.ceil(chars / CHARS_PER_TOKEN);
  }

  if (message.role === "toolResult") {
    for (const block of message.content) {
      if (block.type === "text") chars += block.text.length;
      if (block.type === "image") chars += 4800;
    }
    return Math.ceil(chars / CHARS_PER_TOKEN);
  }

  // Unknown custom message types: ignore.
  return 0;
}

export type ContextTokenEstimate = {
  /** Estimated total tokens in context (system prompt + messages) */
  totalTokens: number;
  /** The last provider usage we anchored on (for debug display) */
  lastUsage: Usage | null;
};

export function estimateContextTokens(
  state: Pick<AgentState, "systemPrompt" | "messages">,
): ContextTokenEstimate {
  const messages = state.messages;

  let lastUsage: Usage | null = null;
  let lastUsageIndex: number | null = null;
  let lastUsageTimestamp = 0;

  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg.role !== "assistant") continue;
    if (msg.stopReason === "error" || msg.stopReason === "aborted") continue;

    const t = calculateContextTokens(msg.usage);
    if (t > 0) {
      lastUsage = msg.usage;
      lastUsageIndex = i;
      lastUsageTimestamp = msg.timestamp;
      break;
    }
  }

  let lastCompactionTimestamp = 0;
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg.role === "compactionSummary") {
      lastCompactionTimestamp = msg.timestamp;
      break;
    }
  }

  const usageIsStale = lastUsage !== null && lastCompactionTimestamp > lastUsageTimestamp;

  let totalTokens = 0;
  if (lastUsage && lastUsageIndex !== null && !usageIsStale) {
    totalTokens = calculateContextTokens(lastUsage);
    for (let i = lastUsageIndex + 1; i < messages.length; i++) {
      totalTokens += estimateMessageTokens(messages[i]);
    }
  } else {
    // No reliable usage signal (or it became stale after /compact). Estimate from scratch.
    totalTokens = estimateTextTokens(state.systemPrompt);
    for (const m of messages) {
      totalTokens += estimateMessageTokens(m);
    }
  }

  return { totalTokens, lastUsage };
}
