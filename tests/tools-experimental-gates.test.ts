import assert from "node:assert/strict";
import { test } from "node:test";

import { createAllTools } from "../src/tools/index.ts";
import { applyExperimentalToolGates } from "../src/tools/experimental-tool-gates.ts";
import { decodeToolDetailsOfKind } from "../src/tools/tool-details.ts";

async function runtimeTools(dependencies: Parameters<typeof applyExperimentalToolGates>[1] = {}) {
  return applyExperimentalToolGates(createAllTools({ hostKind: "office" }), dependencies);
}

function findTool(tools: Awaited<ReturnType<typeof runtimeTools>>, name: string) {
  const tool = tools.find((candidate) => candidate.name === name);
  assert.ok(tool, `${name} should be visible to the model`);
  return tool;
}

void test("the model-visible tmux tool re-checks bridge availability for each execution", async () => {
  let bridgeUrl: string | undefined = "https://localhost:4441";
  const tools = await runtimeTools({
    getTmuxBridgeUrl: () => Promise.resolve(bridgeUrl),
    validateBridgeUrl: (url) => url,
    probeTmuxBridge: () => Promise.resolve(false),
  });
  const tmux = findTool(tools, "tmux");

  const unreachable = await tmux.execute("call-unreachable", { action: "list_sessions" });
  assert.equal(decodeToolDetailsOfKind(unreachable.details, "tmux_bridge")?.gateReason, "bridge_unreachable");

  bridgeUrl = undefined;
  const missing = await tmux.execute("call-missing", { action: "list_sessions" });
  assert.equal(decodeToolDetailsOfKind(missing.details, "tmux_bridge")?.gateReason, "missing_bridge_url");
});

void test("the model-visible files and Office.js tools remain available without feature flags", async () => {
  const tools = await runtimeTools({
    getExecutionMode: () => Promise.resolve("yolo"),
    requestOfficeJsExecuteApproval: () => Promise.reject(new Error("pure Office.js must not request approval in Auto mode")),
  });

  await assert.rejects(
    async () => findTool(tools, "files").execute("call-files", { action: "read" }),
    /'path' is required/u,
  );

  const officeResult = await findTool(tools, "execute_office_js").execute("call-office", {
    explanation: "Inspect workbook",
    code: "return Excel.run(async () => true);",
  });
  const officeText = officeResult.content[0]?.type === "text" ? officeResult.content[0].text : "";
  assert.match(officeText, /Do not call Excel\.run/u);
});

async function executeInvalidPython(tool: ReturnType<typeof findTool>, id: string): Promise<string> {
  const result = await tool.execute(id, {});
  const block = result.content[0];
  return block?.type === "text" ? block.text : "";
}

void test("model-visible Python fallback reaches its concrete tool when no bridge is configured", async () => {
  const tools = await runtimeTools({
    getPythonBridgeUrl: () => Promise.resolve(undefined),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(false),
  });

  assert.equal(await executeInvalidPython(findTool(tools, "python_run"), "call-python-fallback"), "Error: code is required.");
});

void test("an approved Python bridge call reaches the concrete registered tool", async () => {
  const tools = await runtimeTools({
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(true),
    requestPythonBridgeApproval: () => Promise.resolve(true),
  });

  assert.equal(await executeInvalidPython(findTool(tools, "python_run"), "call-python-approved"), "Error: code is required.");
});

void test("approval persisted for one Python bridge URL applies to its next concrete call", async () => {
  let storedUrl: string | undefined;
  let shouldApprove = true;
  const tools = await runtimeTools({
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(true),
    requestPythonBridgeApproval: () => Promise.resolve(shouldApprove),
    getApprovedPythonBridgeUrl: () => Promise.resolve(storedUrl),
    setApprovedPythonBridgeUrl: (url) => {
      storedUrl = url;
      return Promise.resolve();
    },
  });
  const python = findTool(tools, "python_run");

  assert.equal(await executeInvalidPython(python, "call-python-first"), "Error: code is required.");
  shouldApprove = false;
  assert.equal(await executeInvalidPython(python, "call-python-second"), "Error: code is required.");
  assert.equal(storedUrl, "https://localhost:3340");
});
