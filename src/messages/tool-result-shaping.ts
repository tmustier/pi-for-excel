import type { AgentMessage } from "@mariozechner/pi-agent-core";

interface ToolResultShapingConfig {
  recentToolResultsToKeep: number;
  maxCharsBeforeCompaction: number;
  previewChars: number;
}

export const DEFAULT_TOOL_RESULT_SHAPING: Readonly<ToolResultShapingConfig> = {
  // Keep recent tool outputs fully intact so immediate follow-up reasoning stays high quality.
  // Primary payload safety comes from execution-time tool-output truncation.
  recentToolResultsToKeep: 6,
  // Older tool results above this size are compacted for model-facing context.
  maxCharsBeforeCompaction: 1200,
  // Keep a short deterministic preview for grounding.
  previewChars: 500,
};

type ToolResultMessage = Extract<AgentMessage, { role: "toolResult" }>;

type TextBlock = Extract<ToolResultMessage["content"][number], { type: "text" }>;
type ImageBlock = Extract<ToolResultMessage["content"][number], { type: "image" }>;

interface ToolResultPayloadStats {
  textPayload: string;
  textChars: number;
  imageCount: number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isToolResultMessage(message: AgentMessage): message is ToolResultMessage {
  return message.role === "toolResult";
}

function isTextBlock(block: unknown): block is TextBlock {
  return isRecord(block) && block.type === "text" && typeof block.text === "string";
}

function isImageBlock(block: unknown): block is ImageBlock {
  return (
    isRecord(block) &&
    block.type === "image" &&
    typeof block.data === "string" &&
    typeof block.mimeType === "string"
  );
}

function clampPositiveInteger(value: number, fallback: number): number {
  if (!Number.isFinite(value)) return fallback;
  const rounded = Math.floor(value);
  return rounded > 0 ? rounded : fallback;
}

function normalizeConfig(config: Partial<ToolResultShapingConfig> | undefined): ToolResultShapingConfig {
  return {
    recentToolResultsToKeep: clampPositiveInteger(
      config?.recentToolResultsToKeep ?? DEFAULT_TOOL_RESULT_SHAPING.recentToolResultsToKeep,
      DEFAULT_TOOL_RESULT_SHAPING.recentToolResultsToKeep,
    ),
    maxCharsBeforeCompaction: clampPositiveInteger(
      config?.maxCharsBeforeCompaction ?? DEFAULT_TOOL_RESULT_SHAPING.maxCharsBeforeCompaction,
      DEFAULT_TOOL_RESULT_SHAPING.maxCharsBeforeCompaction,
    ),
    previewChars: clampPositiveInteger(
      config?.previewChars ?? DEFAULT_TOOL_RESULT_SHAPING.previewChars,
      DEFAULT_TOOL_RESULT_SHAPING.previewChars,
    ),
  };
}

function collectToolResultIndices(messages: readonly AgentMessage[]): number[] {
  const indices: number[] = [];
  for (let i = 0; i < messages.length; i += 1) {
    if (messages[i].role === "toolResult") {
      indices.push(i);
    }
  }
  return indices;
}

function buildRecentIndexSet(indices: readonly number[], keep: number): Set<number> {
  if (indices.length <= keep) return new Set<number>(indices);
  return new Set<number>(indices.slice(indices.length - keep));
}

function normalizeToolResultContent(message: ToolResultMessage): ToolResultMessage["content"] {
  const rawContent: unknown = message.content;

  // Backwards compatibility: older persisted sessions may still carry string
  // tool-result payloads.
  if (typeof rawContent === "string") {
    return [{ type: "text", text: rawContent }];
  }

  if (!Array.isArray(rawContent)) {
    return [];
  }

  const normalized: ToolResultMessage["content"] = [];
  for (const block of rawContent) {
    if (isTextBlock(block) || isImageBlock(block)) {
      normalized.push(block);
    }
  }

  return normalized;
}

function collectToolResultPayload(content: ToolResultMessage["content"]): ToolResultPayloadStats {
  let textPayload = "";
  let textChars = 0;
  let imageCount = 0;

  for (const block of content) {
    if (block.type === "text") {
      if (textPayload.length > 0) textPayload += "\n";
      textPayload += block.text;
      textChars += block.text.length;
      continue;
    }

    if (block.type === "image") {
      imageCount += 1;
    }
  }

  return { textPayload, textChars, imageCount };
}

function summarizeToolResult(
  message: ToolResultMessage,
  payload: ToolResultPayloadStats,
  previewChars: number,
): TextBlock {
  const previewSource = payload.textPayload.trim();
  const preview = previewSource.slice(0, previewChars);
  const previewWasTruncated = preview.length < previewSource.length;

  const lines: string[] = [];
  lines.push(`[Compacted tool result] ${message.toolName}${message.isError ? " (error)" : ""}`);

  const sourceParts: string[] = [];
  sourceParts.push(`${payload.textChars.toLocaleString()} text chars`);
  if (payload.imageCount > 0) {
    sourceParts.push(`${payload.imageCount} image block${payload.imageCount === 1 ? "" : "s"}`);
  }
  lines.push(`Original payload: ${sourceParts.join(", ")}.`);

  if (preview.length > 0) {
    lines.push("Preview:");
    lines.push(preview + (previewWasTruncated ? "…" : ""));
  } else {
    lines.push("Preview: (no text payload)");
  }

  lines.push("Full output remains visible in chat history; this compact version is model-facing only.");

  return {
    type: "text",
    text: lines.join("\n\n"),
  };
}

function shouldCompactToolResult(
  payload: ToolResultPayloadStats,
  maxCharsBeforeCompaction: number,
): boolean {
  if (payload.imageCount > 0) return true;
  return payload.textChars > maxCharsBeforeCompaction;
}

export function shapeToolResultsForLlm(
  messages: AgentMessage[],
  config?: Partial<ToolResultShapingConfig>,
): AgentMessage[] {
  const resolvedConfig = normalizeConfig(config);
  const toolResultIndices = collectToolResultIndices(messages);
  if (toolResultIndices.length === 0) return messages;

  const recentSet = buildRecentIndexSet(toolResultIndices, resolvedConfig.recentToolResultsToKeep);

  return messages.map((message, index) => {
    if (!isToolResultMessage(message)) return message;

    const normalizedMessage: ToolResultMessage = {
      ...message,
      content: normalizeToolResultContent(message),
    };

    if (recentSet.has(index)) return normalizedMessage;

    const payload = collectToolResultPayload(normalizedMessage.content);
    if (!shouldCompactToolResult(payload, resolvedConfig.maxCharsBeforeCompaction)) {
      return normalizedMessage;
    }

    return {
      ...normalizedMessage,
      content: [summarizeToolResult(normalizedMessage, payload, resolvedConfig.previewChars)],
    };
  });
}
