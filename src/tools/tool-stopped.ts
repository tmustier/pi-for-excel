import type { AgentTool } from "@earendil-works/pi-agent-core";

import { isAbortError } from "../utils/network.js";

export class ToolStoppedError extends Error {
  constructor(message = "Stopped by the user.") {
    super(message);
    this.name = "ToolStoppedError";
  }
}

export function isStop(error: unknown, signal: AbortSignal | undefined): boolean {
  return error instanceof ToolStoppedError || (signal?.aborted === true && isAbortError(error));
}

export function withStopResults(tools: AgentTool[]): AgentTool[] {
  return tools.map((tool) => ({
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      try {
        return await tool.execute(toolCallId, params, signal, onUpdate);
      } catch (error) {
        if (!isStop(error, signal)) throw error;
        const text = error instanceof ToolStoppedError ? error.message : new ToolStoppedError().message;
        return { content: [{ type: "text", text }], details: { kind: "stopped" }, isError: true };
      }
    },
  }));
}
