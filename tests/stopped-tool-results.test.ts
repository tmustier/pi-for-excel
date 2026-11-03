import assert from "node:assert/strict";
import { test, type TestContext } from "node:test";

import { Agent, type AgentTool } from "@earendil-works/pi-agent-core";
import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import { createModels, fauxAssistantMessage, fauxProvider, fauxToolCall, type ToolResultMessage } from "@earendil-works/pi-ai";

import { installImageInputNormalization } from "../src/messages/image-input-limits.ts";
import { callPyodideRuntime, terminatePyodide } from "../src/python/pyodide-runtime.ts";
import { applyExperimentalToolGates } from "../src/tools/experimental-tool-gates.ts";
import { createFetchPageTool } from "../src/tools/fetch-page.ts";
import { createLibreOfficeConvertTool } from "../src/tools/libreoffice-convert.ts";
import { createMcpTool } from "../src/tools/mcp.ts";
import { applyToolOutputTruncation } from "../src/tools/output-truncation.ts";
import { createPythonRunTool } from "../src/tools/python-run.ts";
import { createTmuxTool } from "../src/tools/tmux.ts";
import { withStopResults } from "../src/tools/tool-stopped.ts";
import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";
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

function fixtureTool(name: string, execute: AgentTool["execute"]): AgentTool {
  return { name, label: name, description: name, parameters: { type: "object", properties: {} }, execute };
}

class SilentWorker {
  addEventListener(): void {}
  postMessage(): void {}
  terminate(): void {}
}

const JsonRpcRequest = Type.Object({ method: Type.String(), id: Type.Optional(Type.Union([Type.String(), Type.Number()])) });

function parseJsonRpc(body: string): Static<typeof JsonRpcRequest> {
  return Value.Parse(JsonRpcRequest, JSON.parse(body));
}

function installSilentHost(t: TestContext): void {
  const worker = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  const fetch = Object.getOwnPropertyDescriptor(globalThis, "fetch");
  terminatePyodide();
  Reflect.set(globalThis, "Worker", SilentWorker);
  Reflect.set(globalThis, "fetch", (url: string, init: RequestInit) => {
    const request = url.startsWith(mcpServer.url) && typeof init.body === "string" ? parseJsonRpc(init.body) : null;
    if (request?.method === "notifications/initialized") return Promise.resolve(new Response(null, { status: 202 }));
    if (request?.method === "initialize" || request?.method === "tools/list") {
      const result = request.method === "initialize" ? { protocolVersion: "2025-03-26" } : { tools: [{ name: "echo", inputSchema: { type: "object" } }] };
      return Promise.resolve(Response.json({ jsonrpc: "2.0", id: request.id, result }));
    }
    return new Promise((_resolve, reject) => {
      init.signal?.addEventListener("abort", () => reject(new DOMException("The operation was aborted.", "AbortError")), { once: true });
    });
  });
  t.after(() => {
    terminatePyodide();
    if (worker) Object.defineProperty(globalThis, "Worker", worker);
    else Reflect.deleteProperty(globalThis, "Worker");
    if (fetch) Object.defineProperty(globalThis, "fetch", fetch);
  });
}

const bridge = { url: "https://localhost:3340" };
const mcpServer = { id: "local", name: "local", url: "https://mcp.fixture.invalid", enabled: true };
const neverFails = fixtureTool("files", () => Promise.reject(new Error("must not run after Stop")));

const tools: AgentTool[] = [
  fixtureTool("write_cells", () => Promise.reject(new Error("must not run after Stop"))),
  fixtureTool("read_range", async (_id, _params, signal) => {
    await untilAborted(signal);
    throw new Error("Access denied to Sheet1");
  }),
  fixtureTool("get_workbook_overview", () => Promise.resolve({
    content: [{ type: "image", data: "screenshot", mimeType: "image/png" }],
    details: undefined,
  })),
  createPythonRunTool({ getBridgeConfig: () => Promise.resolve(null), isPyodideAvailable: () => true }),
  createTmuxTool({ getBridgeConfig: () => Promise.resolve(bridge) }),
  { ...neverFails },
  { ...neverFails, name: "extensions_manager" },
  createFetchPageTool({ getConfig: () => Promise.resolve({}) }),
  createMcpTool({ getRuntimeConfig: () => Promise.resolve({ servers: [mcpServer] }), requestToolApproval: () => Promise.resolve(true) }),
  createLibreOfficeConvertTool({
    getBridgeConfig: () => Promise.resolve(bridge),
    getApprovedBridgeUrl: () => Promise.resolve(bridge.url),
  }),
];

type StopAt = "tool start" | "tool running" | "image normalization";

async function stopDuring(call: ReturnType<typeof fauxToolCall>, stopAt: StopAt): Promise<string[]> {
  const faux = fauxProvider({ models: [{ id: "stop-fixture", input: ["text", "image"] }] });
  faux.setResponses([fauxAssistantMessage([call], { stopReason: "toolUse" }), fauxAssistantMessage("unused")]);
  const models = createModels();
  models.setProvider(faux.provider);
  const approveAfterStop = async ({ signal }: { signal?: AbortSignal }) => {
    await untilAborted(signal);
    return false;
  };
  const coordinated = withWorkbookCoordinator(
    tools,
    coordinator,
    { getWorkbookId: () => Promise.resolve("workbook-1"), getSessionId: () => "session-1" },
    undefined,
    {
      getExecutionMode: () => Promise.resolve("safe"),
      requestToolApproval: approveAfterStop,
    },
  );
  const wrapped = withStopResults(applyToolOutputTruncation(await applyExperimentalToolGates(coordinated, {
    getExecutionMode: () => Promise.resolve("safe"),
    requestLocalChangeApproval: approveAfterStop,
    inspectWorkspaceFile: () => Promise.resolve({ size: 1, readOnly: false }),
    listExtensions: () => [],
    getTmuxBridgeUrl: () => Promise.resolve(bridge.url),
    probeTmuxBridge: () => Promise.resolve(true),
    getPythonBridgeUrl: () => Promise.resolve(bridge.url),
    probePythonBridge: () => Promise.resolve(true),
    getApprovedPythonBridgeUrl: () => Promise.resolve(bridge.url),
  })));
  const agent = new Agent({
    initialState: { model: faux.getModel(), tools: wrapped, messages: [] },
    streamFn: (model, context, options) => models.streamSimple(model, context, options),
  });
  agent.getApiKey = () => Promise.resolve("fixture-key");
  installImageInputNormalization(agent, {
    processor: () => {
      agent.abort();
      return Promise.reject(new Error("Image processing aborted."));
    },
  });
  agent.subscribe((event) => {
    if (event.type !== "tool_execution_start") return;
    if (stopAt === "tool start") agent.abort();
    if (stopAt === "tool running") setTimeout(() => agent.abort(), 20);
  });

  await agent.prompt("go");

  const results = new Map<string, ToolResultMessage>();
  for (const message of agent.state.messages) if (message.role === "toolResult") results.set(message.toolCallId, message);
  return buildTranscriptBlocks({
    messages: agent.state.messages,
    streamingMessage: null,
    tools: wrapped,
    results,
    pendingToolCalls: new Set(),
    running: false,
  }).flatMap((block) => block.type === "turnGroup"
    ? block.group.steps.flatMap((step) => step.type === "tool" ? [`${step.toolCall.name}:${step.state}`] : [])
    : []);
}

void test("Stop shows a tool as stopped only when Stop caused its outcome", async (t) => {
  installSilentHost(t);
  const rows: { name: string; call: ReturnType<typeof fauxToolCall>; stopAt: StopAt; expected: string }[] = [
    { name: "approval pending", call: fauxToolCall("write_cells", { start_cell: "Sheet1!A1", values: [[1]] }), stopAt: "tool running", expected: "write_cells:stopped" },
    { name: "before the tool starts", call: fauxToolCall("read_range", { range: "Sheet1!A1" }), stopAt: "tool start", expected: "read_range:stopped" },
    { name: "Python running", call: fauxToolCall("python_run", { code: "import time; time.sleep(60)" }), stopAt: "tool running", expected: "python_run:stopped" },
    { name: "tmux bridge running", call: fauxToolCall("tmux", { action: "list_sessions" }), stopAt: "tool running", expected: "tmux:stopped" },
    { name: "LibreOffice bridge running", call: fauxToolCall("libreoffice_convert", { input_path: "/tmp/a.xlsx", target_format: "pdf" }), stopAt: "tool running", expected: "libreoffice_convert:stopped" },
    { name: "Files deletion approval pending", call: fauxToolCall("files", { action: "delete", path: "a.txt" }), stopAt: "tool running", expected: "files:stopped" },
    { name: "extension install approval pending", call: fauxToolCall("extensions_manager", { action: "install_code", name: "fixture", code: "export default () => {}" }), stopAt: "tool running", expected: "extensions_manager:stopped" },
    { name: "fetch_page running", call: fauxToolCall("fetch_page", { url: "https://page.fixture.invalid/" }), stopAt: "tool running", expected: "fetch_page:stopped" },
    { name: "MCP call running", call: fauxToolCall("mcp", { server: "local", tool: "echo", args: "{}" }), stopAt: "tool running", expected: "mcp:stopped" },
    { name: "a real failure after Stop", call: fauxToolCall("read_range", { range: "Sheet1!A1" }), stopAt: "tool running", expected: "read_range:failed" },
    { name: "image normalization after the tool finished", call: fauxToolCall("get_workbook_overview", {}), stopAt: "image normalization", expected: "get_workbook_overview:done" },
  ];
  for (const { name, call, stopAt, expected } of rows) {
    assert.deepEqual(await stopDuring(call, stopAt), [expected], name);
  }
});

class FailingWorker extends SilentWorker {
  override postMessage(): void {
    throw new Error("Worker dispatch failed");
  }
}

void test("a Pyodide failure that settles before Stop stays a failure", async (t) => {
  installSilentHost(t);
  Reflect.set(globalThis, "Worker", FailingWorker);
  const controller = new AbortController();
  const result = callPyodideRuntime({ code: "pass" }, controller.signal);
  controller.abort();
  assert.equal((await result).error, "Worker dispatch failed");
});
