import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";

import {
  applyToolOutputTruncation,
  DEFAULT_TOOL_OUTPUT_MAX_BYTES,
  DEFAULT_TOOL_OUTPUT_MAX_LINES,
  type ToolOutputTruncationStoreArgs,
} from "../src/tools/output-truncation.ts";
import { getToolOutputTruncationDetails } from "../src/tools/tool-details.ts";

const emptySchema = Type.Object({});
type EmptyParams = Static<typeof emptySchema>;

function createTextTool(args: {
  name: string;
  text: string;
  details?: unknown;
  onUpdateText?: string;
}): AgentTool<typeof emptySchema, unknown> {
  return {
    name: args.name,
    label: args.name,
    description: `${args.name} test tool`,
    parameters: emptySchema,
    execute: (
      _toolCallId: string,
      _params: EmptyParams,
      _signal?: AbortSignal,
      onUpdate?: (partial: { content: Array<{ type: "text"; text: string }>; details: unknown }) => void,
    ) => {
      if (args.onUpdateText && onUpdate) {
        onUpdate({
          content: [{ type: "text", text: args.onUpdateText }],
          details: args.details,
        });
      }

      return Promise.resolve({
        content: [{ type: "text", text: args.text }],
        details: args.details,
      });
    },
  };
}

function makeLinePayload(lineCount: number): string {
  const lines: string[] = [];
  for (let i = 1; i <= lineCount; i += 1) {
    lines.push(`line-${i.toString().padStart(4, "0")}`);
  }
  return lines.join("\n");
}

void test("applies head truncation by default and records metadata", async () => {
  const tool = createTextTool({
    name: "read_range",
    text: makeLinePayload(2_100),
  });

  const wrapped = applyToolOutputTruncation([tool], {
    limits: {
      maxLines: DEFAULT_TOOL_OUTPUT_MAX_LINES,
      maxBytes: 500_000,
    },
  });

  const result = await wrapped[0].execute("call-1", {});
  const text = result.content.find((block) => block.type === "text");
  assert.ok(text);
  assert.match(text.text, /line-0001/);
  assert.doesNotMatch(text.text, /line-2100/);
  assert.match(text.text, /\[Output truncated: showing first 2,000 of 2,100 lines/);

  const truncation = getToolOutputTruncationDetails(result.details);
  assert.ok(truncation);
  assert.equal(truncation.strategy, "head");
  assert.equal(truncation.truncatedBy, "lines");
  assert.equal(truncation.outputLines, 2_000);
  assert.equal(truncation.totalLines, 2_100);
});

void test("applies tail truncation for log-style tools", async () => {
  const tool = createTextTool({
    name: "python_run",
    text: makeLinePayload(2_100),
  });

  const wrapped = applyToolOutputTruncation([tool], {
    limits: {
      maxLines: DEFAULT_TOOL_OUTPUT_MAX_LINES,
      maxBytes: 500_000,
    },
  });

  const result = await wrapped[0].execute("call-2", {});
  const text = result.content.find((block) => block.type === "text");
  assert.ok(text);
  assert.match(text.text, /line-2100/);
  assert.doesNotMatch(text.text, /line-0001/);
  assert.match(text.text, /\[Output truncated: showing last 2,000 of 2,100 lines/);

  const truncation = getToolOutputTruncationDetails(result.details);
  assert.ok(truncation);
  assert.equal(truncation.strategy, "tail");
  assert.equal(truncation.truncatedBy, "lines");
});

void test("stores full output path metadata when persistence callback succeeds", async () => {
  const fullText = "x".repeat(DEFAULT_TOOL_OUTPUT_MAX_BYTES + 1_000);
  const stored: ToolOutputTruncationStoreArgs[] = [];

  const tool = createTextTool({
    name: "search_workbook",
    text: fullText,
  });

  const wrapped = applyToolOutputTruncation([tool], {
    limits: {
      maxLines: 4_000,
      maxBytes: DEFAULT_TOOL_OUTPUT_MAX_BYTES,
    },
    saveTruncatedOutput: (args) => {
      stored.push(args);
      return Promise.resolve(".tool-output/test-full-output.txt");
    },
  });

  const result = await wrapped[0].execute("call-3", {});
  assert.equal(stored.length, 1);
  assert.equal(stored[0]?.fullText, fullText);

  const text = result.content.find((block) => block.type === "text");
  assert.ok(text);
  assert.match(text.text, /full output saved to Files workspace: \.tool-output\/test-full-output\.txt/);

  const truncation = getToolOutputTruncationDetails(result.details);
  assert.ok(truncation);
  assert.equal(truncation.truncatedBy, "bytes");
  assert.equal(truncation.fullOutputWorkspacePath, ".tool-output/test-full-output.txt");
});

void test("truncates streaming updates before forwarding onUpdate callback", async () => {
  const tool = createTextTool({
    name: "python_run",
    text: "done",
    onUpdateText: makeLinePayload(2_100),
  });

  const wrapped = applyToolOutputTruncation([tool], {
    limits: {
      maxLines: DEFAULT_TOOL_OUTPUT_MAX_LINES,
      maxBytes: 500_000,
    },
  });

  let updateText = "";
  await wrapped[0].execute("call-4", {}, undefined, (partial) => {
    const block = partial.content.find((item) => item.type === "text");
    updateText = block?.text ?? "";
  });

  assert.match(updateText, /\[Output truncated: showing last 2,000 of 2,100 lines/);
  assert.match(updateText, /line-2100/);
  assert.doesNotMatch(updateText, /line-0001/);
});
