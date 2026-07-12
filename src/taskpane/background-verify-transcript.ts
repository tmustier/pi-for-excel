/**
 * Pure transcript/usage export helpers for the background verification bridge.
 *
 * Evals need enough signal to compute reply text, tool-call counts/errors and
 * token totals per run without leaking workbook/customer content or secrets
 * into terminal artifacts. Even though the bridge is tokened and loopback-only,
 * exports minimize sensitive surface:
 *   - Raw tool-call arguments are NEVER exported (only tool names + status).
 *   - Raw user-message and tool-result text are NEVER exported (only lengths,
 *     tool name, and error flag).
 *   - Only assistant text/reply is exported, bounded by explicit caps.
 *   - Raw assistant error text is NEVER exported (only a boolean `hasError`).
 * Everything here is pure (no DOM/Office access) so it can be unit-tested and
 * asserted against sentinel-secret leakage. The DOM-facing bridge feeds it
 * `runtime.agent.state.messages`.
 */

import type { AgentMessage } from "@earendil-works/pi-agent-core";

export interface TranscriptExportOptions {
  maxReplyChars?: number;
  maxMessages?: number;
  maxMessageTextChars?: number;
  maxTools?: number;
}

export interface TranscriptToolUsage {
  name: string;
  calls: number;
  errors: number;
}

export interface TranscriptUsageTotals {
  input: number;
  output: number;
  cacheRead: number;
  cacheWrite: number;
  reasoning: number;
  totalTokens: number;
}

export interface TranscriptReply {
  text: string;
  truncated: boolean;
  fullLength: number;
}

export interface TranscriptLastToolCall {
  name: string;
  status: "ok" | "error" | "pending";
}

export interface TranscriptLastAssistant {
  provider: string;
  model: string;
  api: string;
  stopReason: string;
  hasError: boolean;
  textLength: number;
  snippet: string;
}

export interface CompactTranscriptToolCall {
  name: string;
}

export interface CompactTranscriptMessage {
  index: number;
  role: string;
  timestamp: number;
  /** Length only for user/tool-result content; raw text is never exported. */
  textLength?: number;
  /** Bounded assistant text only. */
  text?: string;
  textTruncated?: boolean;
  stopReason?: string;
  hasError?: boolean;
  toolName?: string;
  isError?: boolean;
  toolCalls?: CompactTranscriptToolCall[];
  usage?: { input: number; output: number; totalTokens: number };
}

export interface TranscriptExport {
  messageCount: number;
  userCount: number;
  assistantCount: number;
  toolResultCount: number;
  otherCount: number;
  toolCallCount: number;
  toolErrorCount: number;
  tools: TranscriptToolUsage[];
  toolsTruncated: boolean;
  usage: TranscriptUsageTotals;
  lastAssistant: TranscriptLastAssistant | null;
  lastToolCall: TranscriptLastToolCall | null;
  reply: TranscriptReply | null;
  stopReasons: Record<string, number>;
  messages: CompactTranscriptMessage[];
  messagesTruncated: boolean;
}

const DEFAULT_MAX_REPLY_CHARS = 4_000;
const DEFAULT_MAX_MESSAGES = 120;
const DEFAULT_MAX_MESSAGE_TEXT_CHARS = 400;
const DEFAULT_MAX_TOOLS = 60;
const SNIPPET_CHARS = 240;

function clampInt(value: number | undefined, fallback: number, min: number, max: number): number {
  const base = typeof value === "number" && Number.isFinite(value) ? Math.floor(value) : fallback;
  return Math.max(min, Math.min(max, base));
}

interface Truncated {
  text: string;
  truncated: boolean;
  fullLength: number;
}

function truncateText(text: string, maxChars: number): Truncated {
  const fullLength = text.length;
  if (maxChars <= 0) return { text: "", truncated: fullLength > 0, fullLength };
  if (fullLength <= maxChars) return { text, truncated: false, fullLength };
  return { text: `${text.slice(0, maxChars)}…`, truncated: true, fullLength };
}

interface AssistantTextParts {
  text: string;
  toolCallNames: string[];
}

function readAssistantParts(message: AgentMessage): AssistantTextParts {
  const parts: AssistantTextParts = { text: "", toolCallNames: [] };
  if (message.role !== "assistant") return parts;
  for (const part of message.content) {
    if (part.type === "text") {
      parts.text += part.text;
    } else if (part.type === "toolCall") {
      parts.toolCallNames.push(part.name);
    }
  }
  return parts;
}

function userTextLength(message: AgentMessage): number {
  if (message.role !== "user") return 0;
  const content = message.content;
  if (typeof content === "string") return content.length;
  let length = 0;
  for (const part of content) {
    if (part.type === "text") length += part.text.length;
  }
  return length;
}

function toolResultTextLength(message: AgentMessage): number {
  if (message.role !== "toolResult") return 0;
  let length = 0;
  for (const part of message.content) {
    if (part.type === "text") length += part.text.length;
  }
  return length;
}

/**
 * Summarize the most recent tool call and whether its result landed. Returns
 * null when the transcript contains no tool calls. Never includes tool
 * arguments.
 */
export function summarizeLastToolCall(messages: readonly AgentMessage[]): TranscriptLastToolCall | null {
  const toolResultErrors = new Map<string, boolean>();
  for (const message of messages) {
    if (message.role === "toolResult") {
      toolResultErrors.set(message.toolCallId, message.isError);
    }
  }

  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (!message || message.role !== "assistant") continue;
    for (let partIndex = message.content.length - 1; partIndex >= 0; partIndex -= 1) {
      const part = message.content[partIndex];
      if (!part || part.type !== "toolCall") continue;
      const resultError = toolResultErrors.get(part.id);
      const status: TranscriptLastToolCall["status"] = resultError === undefined
        ? "pending"
        : resultError ? "error" : "ok";
      return { name: part.name, status };
    }
  }
  return null;
}

/**
 * Return a bounded snippet of an assistant message's text content.
 */
export function assistantTextSnippet(message: AgentMessage, maxChars = SNIPPET_CHARS): string {
  const parts = readAssistantParts(message);
  return truncateText(parts.text, clampInt(maxChars, SNIPPET_CHARS, 0, 2_000)).text;
}

function findLastAssistantIndex(messages: readonly AgentMessage[]): number {
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (message && message.role === "assistant") return index;
  }
  return -1;
}

/**
 * Build a bounded transcript + usage export sufficient to compute reply text,
 * tool-call counts/errors and token totals for a single eval run. Raw tool
 * arguments and raw user/tool-result text are intentionally excluded.
 */
export function buildTranscriptExport(
  messages: readonly AgentMessage[],
  options: TranscriptExportOptions = {},
): TranscriptExport {
  const maxReplyChars = clampInt(options.maxReplyChars, DEFAULT_MAX_REPLY_CHARS, 0, 20_000);
  const maxMessages = clampInt(options.maxMessages, DEFAULT_MAX_MESSAGES, 0, 500);
  const maxMessageTextChars = clampInt(options.maxMessageTextChars, DEFAULT_MAX_MESSAGE_TEXT_CHARS, 0, 4_000);
  const maxTools = clampInt(options.maxTools, DEFAULT_MAX_TOOLS, 0, 200);

  let userCount = 0;
  let assistantCount = 0;
  let toolResultCount = 0;
  let otherCount = 0;
  let toolCallCount = 0;
  let toolErrorCount = 0;
  const usage: TranscriptUsageTotals = {
    input: 0,
    output: 0,
    cacheRead: 0,
    cacheWrite: 0,
    reasoning: 0,
    totalTokens: 0,
  };
  const stopReasons: Record<string, number> = {};
  const toolStats = new Map<string, TranscriptToolUsage>();

  const ensureTool = (name: string): TranscriptToolUsage => {
    const existing = toolStats.get(name);
    if (existing) return existing;
    const created: TranscriptToolUsage = { name, calls: 0, errors: 0 };
    toolStats.set(name, created);
    return created;
  };

  for (const message of messages) {
    if (message.role === "assistant") {
      assistantCount += 1;
      usage.input += message.usage.input;
      usage.output += message.usage.output;
      usage.cacheRead += message.usage.cacheRead;
      usage.cacheWrite += message.usage.cacheWrite;
      usage.reasoning += message.usage.reasoning ?? 0;
      usage.totalTokens += message.usage.totalTokens;
      stopReasons[message.stopReason] = (stopReasons[message.stopReason] ?? 0) + 1;
      for (const part of message.content) {
        if (part.type === "toolCall") {
          toolCallCount += 1;
          ensureTool(part.name).calls += 1;
        }
      }
    } else if (message.role === "user") {
      userCount += 1;
    } else if (message.role === "toolResult") {
      toolResultCount += 1;
      if (message.isError) {
        toolErrorCount += 1;
        ensureTool(message.toolName).errors += 1;
      } else {
        ensureTool(message.toolName);
      }
    } else {
      otherCount += 1;
    }
  }

  const allTools = Array.from(toolStats.values());
  const tools = allTools.slice(0, maxTools);
  const toolsTruncated = allTools.length > tools.length;

  const lastAssistantIndex = findLastAssistantIndex(messages);
  let lastAssistant: TranscriptLastAssistant | null = null;
  let reply: TranscriptReply | null = null;
  if (lastAssistantIndex >= 0) {
    const message = messages[lastAssistantIndex];
    if (message && message.role === "assistant") {
      const parts = readAssistantParts(message);
      const snippet = truncateText(parts.text, SNIPPET_CHARS).text;
      lastAssistant = {
        provider: message.provider,
        model: message.model,
        api: message.api,
        stopReason: message.stopReason,
        textLength: parts.text.length,
        snippet,
        hasError: message.errorMessage !== undefined,
      };
      const replyTruncated = truncateText(parts.text, maxReplyChars);
      reply = {
        text: replyTruncated.text,
        truncated: replyTruncated.truncated,
        fullLength: replyTruncated.fullLength,
      };
    }
  }

  const startIndex = Math.max(0, messages.length - maxMessages);
  const compact: CompactTranscriptMessage[] = [];
  for (let index = startIndex; index < messages.length; index += 1) {
    const message = messages[index];
    if (!message) continue;
    if (message.role === "assistant") {
      const parts = readAssistantParts(message);
      const truncatedText = truncateText(parts.text, maxMessageTextChars);
      const toolCalls: CompactTranscriptToolCall[] = parts.toolCallNames.map((name) => ({ name }));
      compact.push({
        index,
        role: message.role,
        timestamp: message.timestamp,
        textLength: parts.text.length,
        text: truncatedText.text,
        textTruncated: truncatedText.truncated,
        stopReason: message.stopReason,
        hasError: message.errorMessage !== undefined,
        ...(toolCalls.length > 0 ? { toolCalls } : {}),
        usage: {
          input: message.usage.input,
          output: message.usage.output,
          totalTokens: message.usage.totalTokens,
        },
      });
    } else if (message.role === "user") {
      compact.push({
        index,
        role: message.role,
        timestamp: message.timestamp,
        textLength: userTextLength(message),
      });
    } else if (message.role === "toolResult") {
      compact.push({
        index,
        role: message.role,
        timestamp: message.timestamp,
        toolName: message.toolName,
        isError: message.isError,
        textLength: toolResultTextLength(message),
      });
    } else {
      // Custom message roles (e.g. archived history) have heterogeneous
      // shapes; record only the safe, bounded role marker.
      compact.push({
        index,
        role: message.role,
        timestamp: 0,
      });
    }
  }

  return {
    messageCount: messages.length,
    userCount,
    assistantCount,
    toolResultCount,
    otherCount,
    toolCallCount,
    toolErrorCount,
    tools,
    toolsTruncated,
    usage,
    lastAssistant,
    lastToolCall: summarizeLastToolCall(messages),
    reply,
    stopReasons,
    messages: compact,
    messagesTruncated: startIndex > 0,
  };
}
