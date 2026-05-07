/**
 * Helpers for working with pi-ai content blocks.
 *
 * Note: `AgentMessage` supports custom message types, so many call sites
 * treat `message.content` as `unknown`. These utilities are defensive and
 * safe to use at those edges.
 */

import type { TextContent, ToolCall } from "@earendil-works/pi-ai";

import { isRecord } from "./type-guards.js";

export function isTextBlock(block: unknown): block is TextContent {
  return (
    isRecord(block) &&
    (block as { type?: unknown }).type === "text" &&
    typeof (block as { text?: unknown }).text === "string"
  );
}

function isToolCallBlock(block: unknown): block is ToolCall {
  return (
    isRecord(block) &&
    (block as { type?: unknown }).type === "toolCall" &&
    typeof (block as { name?: unknown }).name === "string" &&
    isRecord((block as { arguments?: unknown }).arguments)
  );
}

export function extractTextBlocks(content: unknown, separator = "\n"): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";

  return content
    .filter(isTextBlock)
    .map((b) => b.text)
    .join(separator);
}

/** Convenience helper: extract text blocks and join without separators. */
export function extractTextFromContent(content: unknown): string {
  return extractTextBlocks(content, "");
}

export type TranscriptSummaryLimits = {
  toolInput: number;
  toolResult: number;
};

export function summarizeContentForTranscript(
  content: unknown,
  limits: TranscriptSummaryLimits = { toolInput: 200, toolResult: 500 },
): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";

  return content
    .map((b) => {
      if (isTextBlock(b)) return b.text;

      // pi-ai tool calls
      if (isToolCallBlock(b)) {
        const rawArgs = JSON.stringify(b.arguments);
        const snippet =
          rawArgs.length > limits.toolInput
            ? rawArgs.slice(0, limits.toolInput)
            : rawArgs;
        return `[toolCall: ${b.name}(${snippet})]`;
      }

      // Backwards compatibility: Anthropic-style tool blocks
      if (isRecord(b) && (b as { type?: unknown }).type === "tool_use") {
        const name = typeof (b as { name?: unknown }).name === "string" ? (b as { name: string }).name : "tool";
        const input = JSON.stringify((b as { input?: unknown }).input);
        const snippet =
          input.length > limits.toolInput ? input.slice(0, limits.toolInput) : input;
        return `[tool_use: ${name}(${snippet})]`;
      }

      if (isRecord(b) && (b as { type?: unknown }).type === "tool_result") {
        const raw =
          typeof (b as { content?: unknown }).content === "string"
            ? (b as { content: string }).content
            : JSON.stringify((b as { content?: unknown }).content);
        const snippet =
          raw.length > limits.toolResult
            ? raw.slice(0, limits.toolResult)
            : raw;
        return `[tool_result: ${snippet}]`;
      }

      if (isRecord(b) && typeof (b as { type?: unknown }).type === "string") {
        return `[${(b as { type: string }).type}]`;
      }

      return "[block]";
    })
    .join("\n");
}
