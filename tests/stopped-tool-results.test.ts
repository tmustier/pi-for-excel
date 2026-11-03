import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentTool } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider, fauxToolCall, type ToolResultMessage } from "@earendil-works/pi-ai";

import { createPythonRunTool } from "../src/tools/python-run.ts";
import { markStoppedToolResult, withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";
import { buildTranscriptBlocks } from "../src/ui/messages/turn-model.ts";
import type { WorkbookCoordinator, WorkbookQueueSnapshot } from "../src/workbook/coordinator.ts";

const coordinator: WorkbookCoordinator = {
  runRead: (_ctx, fn) => fn(),
  runWrite: async (_ctx, fn) => ({ result: await fn(), revision: 1 }),
  getRevision: () => 0,
  getSnapshot: (): WorkbookQueueSnapshot => ({ revision: 0, queuedWrites: 0, activeWrite: null, queued: [] }),
  subscribe: () => () => {},
};

function untilAborted(signal: AbortSignal | undefined): Promise<void> {
  return new Promise((resolve) => {
    if (!signal || signal.aborted) resolve();
    else signal.addEventListener("abort", () => resolve(), { once: true });
  });
}

const writeCells: AgentTool = {
  name: "write_cells",
  label: "Write cells",
  description: "Write cells",
  parameters: { type: "object", properties: {} },
  execute: () => Promise.reject(new Error("must not run after Stop")),
};

const pythonRun = createPythonRunTool({
  getBridgeConfig: () => Promise.resolve(null),
  isPyodideAvailable: () => true,
  callPyodide: async (_request, signal) => {
    await untilAborted(signal);
    throw new Error("Aborted");
  },
});

async function stopDuringTool(call: ReturnType<typeof fauxToolCall>): Promise<string[]> {
  const faux = fauxProvider({ models: [{ id: "stop-fixture" }] });
  faux.setResponses([fauxAssistantMessage([call], { stopReason: "toolUse" }), fauxAssistantMessage("unused")]);
  const models = createModels();
  models.setProvider(faux.provider);
  const tools = withWorkbookCoordinator(
    [writeCells, pythonRun],
    coordinator,
    { getWorkbookId: () => Promise.resolve("workbook-1"), getSessionId: () => "session-1" },
    undefined,
    {
      getExecutionMode: () => Promise.resolve("safe"),
      requestToolApproval: async ({ signal }) => {
        await untilAborted(signal);
        return false;
      },
    },
  );
  const agent = new Agent({
    initialState: { model: faux.getModel(), tools, messages: [] },
    streamFn: (model, context, options) => models.streamSimple(model, context, options),
    afterToolCall: markStoppedToolResult,
  });
  agent.getApiKey = () => Promise.resolve("fixture-key");
  agent.subscribe((event) => {
    if (event.type === "tool_execution_start") setTimeout(() => agent.abort(), 20);
  });

  await agent.prompt("go");

  const results = new Map<string, ToolResultMessage>();
  for (const message of agent.state.messages) if (message.role === "toolResult") results.set(message.toolCallId, message);
  const blocks = buildTranscriptBlocks({
    messages: agent.state.messages,
    streamingMessage: null,
    tools,
    results,
    pendingToolCalls: new Set(),
    running: false,
  });
  return blocks.flatMap((block) => block.type === "turnGroup"
    ? block.group.steps.flatMap((step) => step.type === "tool" ? [`${step.toolCall.name}:${step.state}`] : [])
    : []);
}

void test("a tool stopped while it waits for approval or runs shows as stopped, not failed", async () => {
  const rows = [
    { name: "approval pending", call: fauxToolCall("write_cells", { start_cell: "Sheet1!A1", values: [[1]] }), expected: ["write_cells:stopped"] },
    { name: "python running", call: fauxToolCall("python_run", { code: "import time; time.sleep(60)" }), expected: ["python_run:stopped"] },
  ];
  for (const { name, call, expected } of rows) {
    assert.deepEqual(await stopDuringTool(call), expected, name);
  }
});

void test("a write whose outcome is unknown after Stop stays an error so its warning shows", async () => {
  const marked = await markStoppedToolResult({
    assistantMessage: fauxAssistantMessage("x"),
    toolCall: fauxToolCall("write_cells", {}),
    args: {},
    result: { content: [{ type: "text", text: "Stopped by the user while Excel was still working on this change. The outcome is unknown: the change may still be applied. Read the affected range back before relying on it." }], details: {} },
    isError: true,
    context: { systemPrompt: "", messages: [] },
  }, AbortSignal.abort());
  assert.equal(marked, undefined);
});
