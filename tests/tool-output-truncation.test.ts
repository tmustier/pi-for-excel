import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
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
  details?: DynamicValue;
  onUpdateText?: string;
  imageData?: string;
}): AgentTool<typeof emptySchema, DynamicValue> {
  return {
    name: args.name,
    label: args.name,
    description: `${args.name} test tool`,
    parameters: emptySchema,
    execute: (
      _toolCallId: string,
      _params: EmptyParams,
      _signal?: AbortSignal,
      onUpdate?: (partial: { content: Array<{ type: "text"; text: string }>; details: DynamicValue }) => void,
    ) => {
      if (args.onUpdateText && onUpdate) {
        onUpdate({
          content: [{ type: "text", text: args.onUpdateText }],
          details: args.details,
        });
      }

      const content: AgentToolResult<DynamicValue>["content"] = [{ type: "text", text: args.text }];
      if (args.imageData) {
        content.push({ type: "image", data: args.imageData, mimeType: "image/png" });
      }

      return Promise.resolve({
        content,
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

