import type {
  AgentTool,
  AgentToolResult,
  AgentToolUpdateCallback,
} from "@mariozechner/pi-agent-core";
import type { ImageContent, TextContent } from "@mariozechner/pi-ai";

import { isRecord } from "../utils/type-guards.js";
import type {
  ToolOutputTruncationDetails,
  ToolOutputTruncationReason,
  ToolOutputTruncationStrategy,
} from "./tool-details.js";

export const DEFAULT_TOOL_OUTPUT_MAX_LINES = 2_000;
export const DEFAULT_TOOL_OUTPUT_MAX_BYTES = 50 * 1024;

const utf8Encoder = new TextEncoder();

const TAIL_TRUNCATION_TOOL_NAMES = new Set<string>([
  "python_run",
  "tmux",
  "mcp",
  "execute_office_js",
]);

export interface ToolOutputTruncationLimits {
  maxLines: number;
  maxBytes: number;
}

interface TruncationComputation {
  output: string;
  truncated: boolean;
  truncatedBy: ToolOutputTruncationReason;
  totalLines: number;
  totalBytes: number;
  outputLines: number;
  outputBytes: number;
}

export interface ToolOutputTruncationStoreArgs {
  toolName: string;
  toolCallId: string;
  fullText: string;
  truncation: ToolOutputTruncationDetails;
}

export interface ToolOutputTruncationOptions {
  limits?: Partial<ToolOutputTruncationLimits>;
  strategyForTool?: (toolName: string) => ToolOutputTruncationStrategy;
  saveTruncatedOutput?: (args: ToolOutputTruncationStoreArgs) => Promise<string | undefined>;
}

function utf8ByteLength(text: string): number {
  return utf8Encoder.encode(text).byteLength;
}

function normalizePositiveInt(value: number | undefined, fallback: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return fallback;
  const rounded = Math.floor(value);
  return rounded > 0 ? rounded : fallback;
}

function normalizeLimits(limits: Partial<ToolOutputTruncationLimits> | undefined): ToolOutputTruncationLimits {
  return {
    maxLines: normalizePositiveInt(limits?.maxLines, DEFAULT_TOOL_OUTPUT_MAX_LINES),
    maxBytes: normalizePositiveInt(limits?.maxBytes, DEFAULT_TOOL_OUTPUT_MAX_BYTES),
  };
}

function computeUntruncated(text: string): TruncationComputation {
  const lines = text.split("\n");
  const bytes = utf8ByteLength(text);

  return {
    output: text,
    truncated: false,
    truncatedBy: null,
    totalLines: lines.length,
    totalBytes: bytes,
    outputLines: lines.length,
    outputBytes: bytes,
  };
}

function truncateHead(text: string, limits: ToolOutputTruncationLimits): TruncationComputation {
  const lines = text.split("\n");
  const totalBytes = utf8ByteLength(text);
  const totalLines = lines.length;

  if (totalLines <= limits.maxLines && totalBytes <= limits.maxBytes) {
    return computeUntruncated(text);
  }

  const selected: string[] = [];
  let outputBytes = 0;
  let truncatedBy: ToolOutputTruncationReason = "lines";

  for (let i = 0; i < lines.length && i < limits.maxLines; i += 1) {
    const line = lines[i];
    const lineBytes = utf8ByteLength(line) + (i > 0 ? 1 : 0);
    if (outputBytes + lineBytes > limits.maxBytes) {
      truncatedBy = "bytes";
      break;
    }

    selected.push(line);
    outputBytes += lineBytes;
  }

  if (selected.length >= limits.maxLines && outputBytes <= limits.maxBytes) {
    truncatedBy = "lines";
  }

  const output = selected.join("\n");
  const computedOutputBytes = utf8ByteLength(output);

  return {
    output,
    truncated: true,
    truncatedBy,
    totalLines,
    totalBytes,
    outputLines: selected.length,
    outputBytes: computedOutputBytes,
  };
}

function adjustStartToCodePointBoundary(text: string, start: number): number {
  if (start <= 0 || start >= text.length) return start;

  const code = text.charCodeAt(start);
  const isLowSurrogate = code >= 0xdc00 && code <= 0xdfff;
  if (!isLowSurrogate) return start;

  return start + 1;
}

function truncateStringToBytesFromEnd(text: string, maxBytes: number): string {
  if (utf8ByteLength(text) <= maxBytes) {
    return text;
  }

  let low = 0;
  let high = text.length;

  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    const safeMid = Math.min(adjustStartToCodePointBoundary(text, mid), text.length);
    const suffix = text.slice(safeMid);

    if (utf8ByteLength(suffix) <= maxBytes) {
      high = mid;
    } else {
      low = mid + 1;
    }
  }

  const start = Math.min(adjustStartToCodePointBoundary(text, low), text.length);
  return text.slice(start);
}

function truncateTail(text: string, limits: ToolOutputTruncationLimits): TruncationComputation {
  const lines = text.split("\n");
  const totalBytes = utf8ByteLength(text);
  const totalLines = lines.length;

  if (totalLines <= limits.maxLines && totalBytes <= limits.maxBytes) {
    return computeUntruncated(text);
  }

  const selected: string[] = [];
  let outputBytes = 0;
  let truncatedBy: ToolOutputTruncationReason = "lines";

  for (let i = lines.length - 1; i >= 0 && selected.length < limits.maxLines; i -= 1) {
    const line = lines[i];
    const lineBytes = utf8ByteLength(line) + (selected.length > 0 ? 1 : 0);

    if (outputBytes + lineBytes > limits.maxBytes) {
      truncatedBy = "bytes";
      if (selected.length === 0) {
        const suffix = truncateStringToBytesFromEnd(line, limits.maxBytes);
        if (suffix.length > 0) {
          selected.unshift(suffix);
          outputBytes = utf8ByteLength(suffix);
        }
      }
      break;
    }

    selected.unshift(line);
    outputBytes += lineBytes;
  }

  if (selected.length >= limits.maxLines && outputBytes <= limits.maxBytes) {
    truncatedBy = "lines";
  }

  const output = selected.join("\n");
  const computedOutputBytes = utf8ByteLength(output);

  return {
    output,
    truncated: true,
    truncatedBy,
    totalLines,
    totalBytes,
    outputLines: selected.length,
    outputBytes: computedOutputBytes,
  };
}

function defaultStrategyForTool(toolName: string): ToolOutputTruncationStrategy {
  return TAIL_TRUNCATION_TOOL_NAMES.has(toolName)
    ? "tail"
    : "head";
}

function formatSize(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes}B`;
  }

  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)}KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}

function buildTruncationNotice(truncation: ToolOutputTruncationDetails): string {
  const window = truncation.strategy === "tail"
    ? "last"
    : "first";

  const parts: string[] = [];
  parts.push(
    `[Output truncated: showing ${window} ${truncation.outputLines.toLocaleString()} of ${truncation.totalLines.toLocaleString()} lines`,
  );

  if (truncation.truncatedBy === "bytes") {
    parts.push(` (${formatSize(truncation.maxBytes)} limit)`);
  }

  parts.push(`; limits: ${truncation.maxLines.toLocaleString()} lines / ${formatSize(truncation.maxBytes)}`);

  if (truncation.fullOutputWorkspacePath) {
    parts.push(`; full output saved to Files workspace: ${truncation.fullOutputWorkspacePath}`);
  }

  parts.push("]");

  return parts.join("");
}

function mergeTruncationIntoDetails(
  details: unknown,
  truncation: ToolOutputTruncationDetails,
): unknown {
  if (isRecord(details)) {
    return {
      ...details,
      outputTruncation: truncation,
    };
  }

  return {
    outputTruncation: truncation,
  };
}

function splitToolResultContent(content: readonly (TextContent | ImageContent)[]): {
  text: string;
  textBlockCount: number;
  nonTextBlocks: ImageContent[];
} {
  const textParts: string[] = [];
  const nonTextBlocks: ImageContent[] = [];

  for (const block of content) {
    if (block.type === "text") {
      textParts.push(block.text);
      continue;
    }

    nonTextBlocks.push(block);
  }

  return {
    text: textParts.join("\n"),
    textBlockCount: textParts.length,
    nonTextBlocks,
  };
}

function computeTruncation(
  text: string,
  strategy: ToolOutputTruncationStrategy,
  limits: ToolOutputTruncationLimits,
): TruncationComputation {
  if (strategy === "tail") {
    return truncateTail(text, limits);
  }

  return truncateHead(text, limits);
}

function buildTruncatedResult(args: {
  result: AgentToolResult<unknown>;
  computed: TruncationComputation;
  nonTextBlocks: ImageContent[];
  truncation: ToolOutputTruncationDetails;
}): AgentToolResult<unknown> {
  const notice = buildTruncationNotice(args.truncation);
  const textWithNotice = args.computed.output.trim().length > 0
    ? `${args.computed.output}\n\n${notice}`
    : notice;

  const content: (TextContent | ImageContent)[] = [
    { type: "text", text: textWithNotice },
    ...args.nonTextBlocks,
  ];

  return {
    ...args.result,
    content,
    details: mergeTruncationIntoDetails(args.result.details, args.truncation),
  };
}

function applyTruncationSync(
  result: AgentToolResult<unknown>,
  strategy: ToolOutputTruncationStrategy,
  limits: ToolOutputTruncationLimits,
): AgentToolResult<unknown> {
  const { text, textBlockCount, nonTextBlocks } = splitToolResultContent(result.content);
  if (textBlockCount === 0) {
    return result;
  }

  const computed = computeTruncation(text, strategy, limits);
  if (!computed.truncated) {
    return result;
  }

  const truncation: ToolOutputTruncationDetails = {
    version: 1,
    strategy,
    truncated: true,
    truncatedBy: computed.truncatedBy,
    totalLines: computed.totalLines,
    totalBytes: computed.totalBytes,
    outputLines: computed.outputLines,
    outputBytes: computed.outputBytes,
    maxLines: limits.maxLines,
    maxBytes: limits.maxBytes,
  };

  return buildTruncatedResult({
    result,
    computed,
    nonTextBlocks,
    truncation,
  });
}

async function applyTruncationFinal(
  result: AgentToolResult<unknown>,
  args: {
    toolName: string;
    toolCallId: string;
    strategy: ToolOutputTruncationStrategy;
    limits: ToolOutputTruncationLimits;
    saveTruncatedOutput?: (args: ToolOutputTruncationStoreArgs) => Promise<string | undefined>;
  },
): Promise<AgentToolResult<unknown>> {
  const { text, textBlockCount, nonTextBlocks } = splitToolResultContent(result.content);
  if (textBlockCount === 0) {
    return result;
  }

  const computed = computeTruncation(text, args.strategy, args.limits);
  if (!computed.truncated) {
    return result;
  }

  const truncation: ToolOutputTruncationDetails = {
    version: 1,
    strategy: args.strategy,
    truncated: true,
    truncatedBy: computed.truncatedBy,
    totalLines: computed.totalLines,
    totalBytes: computed.totalBytes,
    outputLines: computed.outputLines,
    outputBytes: computed.outputBytes,
    maxLines: args.limits.maxLines,
    maxBytes: args.limits.maxBytes,
  };

  if (args.saveTruncatedOutput) {
    try {
      const path = await args.saveTruncatedOutput({
        toolName: args.toolName,
        toolCallId: args.toolCallId,
        fullText: text,
        truncation,
      });
      if (path) {
        truncation.fullOutputWorkspacePath = path;
      }
    } catch {
      // best effort only
    }
  }

  return buildTruncatedResult({
    result,
    computed,
    nonTextBlocks,
    truncation,
  });
}

function wrapToolWithOutputTruncation(
  tool: AgentTool,
  options: {
    strategyForTool: (toolName: string) => ToolOutputTruncationStrategy;
    limits: ToolOutputTruncationLimits;
    saveTruncatedOutput?: (args: ToolOutputTruncationStoreArgs) => Promise<string | undefined>;
  },
): AgentTool {
  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      const strategy = options.strategyForTool(tool.name);

      const wrappedOnUpdate: AgentToolUpdateCallback<unknown> | undefined = onUpdate
        ? (partialResult) => {
          onUpdate(applyTruncationSync(partialResult, strategy, options.limits));
        }
        : undefined;

      const result = await tool.execute(toolCallId, params, signal, wrappedOnUpdate);
      return applyTruncationFinal(result, {
        toolName: tool.name,
        toolCallId,
        strategy,
        limits: options.limits,
        saveTruncatedOutput: options.saveTruncatedOutput,
      });
    },
  };
}

export function applyToolOutputTruncation(
  tools: AgentTool[],
  options: ToolOutputTruncationOptions = {},
): AgentTool[] {
  const limits = normalizeLimits(options.limits);
  const strategyForTool = options.strategyForTool ?? defaultStrategyForTool;

  return tools.map((tool) => wrapToolWithOutputTruncation(tool, {
    strategyForTool,
    limits,
    saveTruncatedOutput: options.saveTruncatedOutput,
  }));
}

export { saveTruncatedToolOutputToWorkspace } from "./output-truncation-storage.js";
