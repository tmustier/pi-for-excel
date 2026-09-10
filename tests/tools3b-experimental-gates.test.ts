import assert from "node:assert/strict";
import { test } from "node:test";

import { applyExperimentalToolGates } from "../src/tools/experimental-tool-gates.ts";
import { createAllTools } from "../src/tools/index.ts";

async function gatedTools(dependencies: Parameters<typeof applyExperimentalToolGates>[1]) {
  return applyExperimentalToolGates(createAllTools({ hostKind: "office" }), dependencies);
}

function pythonRun(tools: Awaited<ReturnType<typeof gatedTools>>) {
  const tool = tools.find((candidate) => candidate.name === "python_run");
  if (!tool) throw new Error("python_run should be model-visible");
  return tool;
}

async function executeInvalidPython(tool: ReturnType<typeof pythonRun>, id: string): Promise<string> {
  const result = await tool.execute(id, {});
  const block = result.content[0];
  return block?.type === "text" ? block.text : "";
}

void test("model-visible Python fallback reaches its concrete tool when no bridge is configured", async () => {
  const tools = await gatedTools({
    getPythonBridgeUrl: () => Promise.resolve(undefined),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(false),
  });

  assert.equal(await executeInvalidPython(pythonRun(tools), "call-python-fallback"), "Error: code is required.");
});

void test("an approved Python bridge call reaches the concrete registered tool", async () => {
  const tools = await gatedTools({
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(true),
    requestPythonBridgeApproval: () => Promise.resolve(true),
  });

  assert.equal(await executeInvalidPython(pythonRun(tools), "call-python-approved"), "Error: code is required.");
});

void test("approval persisted for one Python bridge URL applies to its next concrete call", async () => {
  let storedUrl: string | undefined;
  let shouldApprove = true;
  const tools = await gatedTools({
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
  const python = pythonRun(tools);

  assert.equal(await executeInvalidPython(python, "call-python-first"), "Error: code is required.");
  shouldApprove = false;
  assert.equal(await executeInvalidPython(python, "call-python-second"), "Error: code is required.");
  assert.equal(storedUrl, "https://localhost:3340");
});
