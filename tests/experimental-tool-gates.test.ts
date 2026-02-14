import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool } from "@mariozechner/pi-agent-core";
import { Type } from "@sinclair/typebox";

import {
  applyExperimentalToolGates,
  buildPythonBridgeGateErrorMessage,
  buildTmuxBridgeGateErrorMessage,
  evaluatePythonBridgeGate,
  evaluateTmuxBridgeGate,
} from "../src/tools/experimental-tool-gates.ts";

const emptySchema = Type.Object({});

function createTestTool(
  name: string,
  onExecute?: () => void,
): AgentTool<typeof emptySchema, undefined> {
  return {
    label: `${name} tool`,
    name,
    description: `${name} description`,
    parameters: emptySchema,
    execute: () => {
      onExecute?.();
      return Promise.resolve({
        content: [{ type: "text", text: `${name}:ok` }],
        details: undefined,
      });
    },
  };
}

void test("keeps tmux tool registered when experiment is disabled", async () => {
  let probeCalled = false;

  const tools = [createTestTool("tmux"), createTestTool("read_range")];
  const gated = await applyExperimentalToolGates(tools, {
    isTmuxExperimentEnabled: () => false,
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3337"),
    validateBridgeUrl: () => "https://localhost:3337",
    probeTmuxBridge: () => {
      probeCalled = true;
      return Promise.resolve(true);
    },
  });

  assert.deepEqual(gated.map((tool) => tool.name), ["tmux", "read_range"]);
  assert.equal(probeCalled, false);

  const tmuxTool = gated.find((tool) => tool.name === "tmux");
  assert.ok(tmuxTool);

  await assert.rejects(
    () => tmuxTool.execute("call-1", {}),
    /\/experimental on tmux-bridge/,
  );

  assert.equal(probeCalled, false);
});

void test("tmux hard gate re-checks execution on every call", async () => {
  let enabled = true;
  let bridgeHealthy = true;
  let executeCount = 0;

  const [gatedTmux] = await applyExperimentalToolGates([createTestTool("tmux", () => {
    executeCount += 1;
  })], {
    isTmuxExperimentEnabled: () => enabled,
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3337"),
    validateBridgeUrl: () => "https://localhost:3337",
    probeTmuxBridge: () => Promise.resolve(bridgeHealthy),
  });

  assert.ok(gatedTmux);

  await gatedTmux.execute("call-1", {});
  assert.equal(executeCount, 1);

  enabled = false;

  await assert.rejects(
    () => gatedTmux.execute("call-2", {}),
    /\/experimental on tmux-bridge/,
  );
  assert.equal(executeCount, 1);

  enabled = true;
  bridgeHealthy = false;

  await assert.rejects(
    () => gatedTmux.execute("call-3", {}),
    /not reachable/i,
  );
  assert.equal(executeCount, 1);
});

void test("evaluateTmuxBridgeGate reports explicit reason codes", async () => {
  const missingUrl = await evaluateTmuxBridgeGate({
    isTmuxExperimentEnabled: () => true,
    getTmuxBridgeUrl: () => Promise.resolve(undefined),
  });

  assert.equal(missingUrl.allowed, false);
  assert.equal(missingUrl.reason, "missing_bridge_url");

  const unreachable = await evaluateTmuxBridgeGate({
    isTmuxExperimentEnabled: () => true,
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3337"),
    validateBridgeUrl: () => "https://localhost:3337",
    probeTmuxBridge: () => Promise.resolve(false),
  });

  assert.equal(unreachable.allowed, false);
  assert.equal(unreachable.reason, "bridge_unreachable");
  assert.match(buildTmuxBridgeGateErrorMessage(unreachable.reason), /not reachable/i);
});

void test("files tool passes through without any gate", async () => {
  let executeCount = 0;

  const [filesTool] = await applyExperimentalToolGates([
    createTestTool("files", () => {
      executeCount += 1;
    }),
  ], {});

  // All actions pass through directly
  await filesTool.execute("call-files-list", { action: "list" });
  await filesTool.execute("call-files-read", { action: "read", path: "notes.md" });
  await filesTool.execute("call-files-write", { action: "write", path: "notes.md", content: "hello" });
  await filesTool.execute("call-files-delete", { action: "delete", path: "notes.md" });

  assert.equal(executeCount, 4);
});

void test("execute_office_js is available without experimental feature gates", async () => {
  let executeCount = 0;

  const [officeTool] = await applyExperimentalToolGates([
    createTestTool("execute_office_js", () => {
      executeCount += 1;
    }),
  ], {
    requestOfficeJsExecuteApproval: () => Promise.resolve(true),
  });

  await officeTool.execute("call-office", {
    explanation: "Update workbook metadata",
    code: "return { ok: true };",
  });

  assert.equal(executeCount, 1);
});

void test("execute_office_js requires explicit user approval", async () => {
  let executeCount = 0;

  const [officeTool] = await applyExperimentalToolGates([
    createTestTool("execute_office_js", () => {
      executeCount += 1;
    }),
  ], {
    requestOfficeJsExecuteApproval: ({ explanation, code }) => {
      assert.equal(explanation, "Rebuild totals");
      assert.equal(code, "return { ok: true };");
      return Promise.resolve(false);
    },
  });

  await assert.rejects(
    () => officeTool.execute("call-office", {
      explanation: "Rebuild totals",
      code: "return { ok: true };",
    }),
    /cancelled by user/i,
  );

  assert.equal(executeCount, 0);
});

void test("execute_office_js fails closed when confirmation UI is unavailable", async () => {
  let executeCount = 0;

  const [officeTool] = await applyExperimentalToolGates([
    createTestTool("execute_office_js", () => {
      executeCount += 1;
    }),
  ], {});

  await assert.rejects(
    () => officeTool.execute("call-office", {
      explanation: "Rebuild totals",
      code: "return { ok: true };",
    }),
    /approval.*unavailable|confirmation UI is unavailable/i,
  );

  assert.equal(executeCount, 0);
});

void test("python bridge tools require URL + probe but no experiment flag", async () => {
  let executeCount = 0;

  const tools = [
    createTestTool("python_run", () => { executeCount += 1; }),
    createTestTool("libreoffice_convert", () => { executeCount += 1; }),
    createTestTool("python_transform_range", () => { executeCount += 1; }),
    createTestTool("read_range"),
  ];

  // No bridge URL → gate fails with missing_bridge_url
  const gatedTools = await applyExperimentalToolGates(tools, {
    getPythonBridgeUrl: () => Promise.resolve(undefined),
  });

  assert.deepEqual(gatedTools.map((tool) => tool.name), [
    "python_run",
    "libreoffice_convert",
    "python_transform_range",
    "read_range",
  ]);

  const pythonTool = gatedTools.find((tool) => tool.name === "python_run");
  assert.ok(pythonTool);

  await assert.rejects(
    () => pythonTool.execute("call-python", { code: "print('hi')" }),
    /not configured/i,
  );

  // Bridge unreachable → gate fails
  const gatedTools2 = await applyExperimentalToolGates(tools, {
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: () => "https://localhost:3340",
    probePythonBridge: () => Promise.resolve(false),
  });

  const pythonTool2 = gatedTools2.find((tool) => tool.name === "python_run");
  assert.ok(pythonTool2);

  await assert.rejects(
    () => pythonTool2.execute("call-python", { code: "print('hi')" }),
    /not reachable/i,
  );

  assert.equal(executeCount, 0);
});

void test("python bridge tools require explicit user approval", async () => {
  let approvalCalls = 0;
  let executeCount = 0;

  const [pythonTool] = await applyExperimentalToolGates([
    createTestTool("python_run", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: () => "https://localhost:3340",
    probePythonBridge: () => Promise.resolve(true),
    requestPythonBridgeApproval: ({ toolName, params }) => {
      approvalCalls += 1;
      assert.equal(toolName, "python_run");
      assert.deepEqual(params, { code: "print('hello')" });
      return Promise.resolve(false);
    },
  });

  await assert.rejects(
    () => pythonTool.execute("call-python", { code: "print('hello')" }),
    /cancelled by user/i,
  );

  assert.equal(approvalCalls, 1);
  assert.equal(executeCount, 0);
});

void test("approved python bridge calls proceed to tool execution", async () => {
  let executeCount = 0;

  const [pythonTool] = await applyExperimentalToolGates([
    createTestTool("python_run", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: () => "https://localhost:3340",
    probePythonBridge: () => Promise.resolve(true),
    requestPythonBridgeApproval: () => Promise.resolve(true),
  });

  await pythonTool.execute("call-python", { code: "print('hello')" });
  assert.equal(executeCount, 1);
});

void test("python bridge approval is cached per bridge URL", async () => {
  let approvalCalls = 0;
  let executeCount = 0;
  let currentBridgeUrl = "https://localhost:3340";
  let approvedBridgeUrl: string | undefined;

  const [pythonTool] = await applyExperimentalToolGates([
    createTestTool("python_run", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve(currentBridgeUrl),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(true),
    requestPythonBridgeApproval: ({ bridgeUrl }) => {
      approvalCalls += 1;
      assert.equal(bridgeUrl, currentBridgeUrl);
      return Promise.resolve(true);
    },
    getApprovedPythonBridgeUrl: () => Promise.resolve(approvedBridgeUrl),
    setApprovedPythonBridgeUrl: (bridgeUrl) => {
      approvedBridgeUrl = bridgeUrl;
      return Promise.resolve();
    },
  });

  await pythonTool.execute("call-1", { code: "print('first')" });
  await pythonTool.execute("call-2", { code: "print('second')" });

  assert.equal(approvalCalls, 1);
  assert.equal(executeCount, 2);
  assert.equal(approvedBridgeUrl, "https://localhost:3340");

  currentBridgeUrl = "https://localhost:3350";
  await pythonTool.execute("call-3", { code: "print('third')" });

  assert.equal(approvalCalls, 2);
  assert.equal(executeCount, 3);
  assert.equal(approvedBridgeUrl, "https://localhost:3350");
});

void test("evaluatePythonBridgeGate reports explicit reason codes", async () => {
  const missingUrl = await evaluatePythonBridgeGate({
    getPythonBridgeUrl: () => Promise.resolve(undefined),
  });

  assert.equal(missingUrl.allowed, false);
  assert.equal(missingUrl.reason, "missing_bridge_url");

  const unreachable = await evaluatePythonBridgeGate({
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: () => "https://localhost:3340",
    probePythonBridge: () => Promise.resolve(false),
  });

  assert.equal(unreachable.allowed, false);
  assert.equal(unreachable.reason, "bridge_unreachable");
  assert.match(buildPythonBridgeGateErrorMessage(unreachable.reason), /not reachable/i);
});
