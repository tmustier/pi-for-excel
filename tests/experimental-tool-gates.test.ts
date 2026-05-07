import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool } from "@earendil-works/pi-agent-core";
import { Type } from "@sinclair/typebox";

import {
  applyExperimentalToolGates,
  buildPythonBridgeGateErrorMessage,
  buildTmuxBridgeGateErrorMessage,
  evaluatePythonBridgeGate,
  evaluateTmuxBridgeGate,
} from "../src/tools/experimental-tool-gates.ts";
import {
  isBridgeGateError,
  isLibreOfficeBridgeDetails,
  isPythonBridgeDetails,
  isPythonTransformRangeDetails,
  isTmuxBridgeDetails,
} from "../src/tools/tool-details.ts";

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

function assertTmuxGateError(
  details: unknown,
  reason: "missing_bridge_url" | "bridge_unreachable",
): void {
  assert.ok(isTmuxBridgeDetails(details));
  assert.equal(details.ok, false);
  assert.equal(details.gateReason, reason);
  assert.equal(details.skillHint, "tmux-bridge");
}

function assertPythonGateError(details: unknown): void {
  assert.ok(isPythonBridgeDetails(details));
  assert.equal(details.ok, false);
  assert.equal(details.gateReason, "bridge_unreachable");
  assert.equal(details.skillHint, "python-bridge");
}

function assertPythonTransformRangeGateError(details: unknown): void {
  assert.ok(isPythonTransformRangeDetails(details));
  assert.equal(details.blocked, false);
  assert.equal(details.gateReason, "bridge_unreachable");
  assert.equal(details.skillHint, "python-bridge");
  assert.match(details.error ?? "", /not reachable/i);
}

function assertLibreOfficeGateError(
  details: unknown,
  reason: "missing_bridge_url" | "bridge_unreachable",
): void {
  assert.ok(isLibreOfficeBridgeDetails(details));
  assert.equal(details.ok, false);
  assert.equal(details.gateReason, reason);
  assert.equal(details.skillHint, "python-bridge");
}

void test("keeps tmux tool registered and returns structured gate errors", async () => {
  let probeCalled = false;

  const tools = [createTestTool("tmux"), createTestTool("read_range")];
  const gated = await applyExperimentalToolGates(tools, {
    getTmuxBridgeUrl: () => Promise.resolve(undefined),
    validateBridgeUrl: (url) => url,
    probeTmuxBridge: () => {
      probeCalled = true;
      return Promise.resolve(false);
    },
  });

  assert.deepEqual(gated.map((tool) => tool.name), ["tmux", "read_range"]);

  const tmuxTool = gated.find((tool) => tool.name === "tmux");
  assert.ok(tmuxTool);

  const result = await tmuxTool.execute("call-1", {
    action: "capture_pane",
    session: "dev",
  });

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.match(text, /Terminal access is not available/i);
  assert.match(text, /default URL|URL override/i);
  assert.match(text, /Skill: tmux-bridge/i);

  const resultDetails: unknown = result.details;
  assertTmuxGateError(resultDetails, "missing_bridge_url");
  assert.ok(isTmuxBridgeDetails(resultDetails));
  assert.equal(resultDetails.action, "capture_pane");

  assert.equal(probeCalled, true);
});

void test("tmux hard gate re-checks execution on every call", async () => {
  let bridgeUrl: string | undefined = "https://localhost:4441";
  let configuredBridgeHealthy = true;
  let defaultBridgeHealthy = false;
  let executeCount = 0;

  const [gatedTmux] = await applyExperimentalToolGates([createTestTool("tmux", () => {
    executeCount += 1;
  })], {
    getTmuxBridgeUrl: () => Promise.resolve(bridgeUrl),
    validateBridgeUrl: (url) => url,
    probeTmuxBridge: (url) => Promise.resolve(
      url === "https://localhost:3341" ? defaultBridgeHealthy : configuredBridgeHealthy,
    ),
  });

  assert.ok(gatedTmux);

  await gatedTmux.execute("call-1", {});
  assert.equal(executeCount, 1);

  bridgeUrl = undefined;

  const missingResult = await gatedTmux.execute("call-2", {});
  const missingText = missingResult.content[0]?.type === "text" ? missingResult.content[0].text : "";
  assert.match(missingText, /Terminal access is not available/i);
  assert.match(missingText, /default URL|URL override/i);
  assert.match(missingText, /Skill: tmux-bridge/i);
  const missingDetails: unknown = missingResult.details;
  assertTmuxGateError(missingDetails, "missing_bridge_url");
  assert.equal(executeCount, 1);

  bridgeUrl = "https://localhost:4441";
  configuredBridgeHealthy = false;
  defaultBridgeHealthy = true;

  const unreachableResult = await gatedTmux.execute("call-3", {});
  const unreachableText = unreachableResult.content[0]?.type === "text" ? unreachableResult.content[0].text : "";
  assert.match(unreachableText, /Terminal access is not available/i);
  assert.match(unreachableText, /not reachable/i);
  assert.match(unreachableText, /Skill: tmux-bridge/i);
  const unreachableDetails: unknown = unreachableResult.details;
  assertTmuxGateError(unreachableDetails, "bridge_unreachable");
  assert.equal(executeCount, 1);
});

void test("evaluateTmuxBridgeGate reports explicit reason codes", async () => {
  const missingUrl = await evaluateTmuxBridgeGate({
    getTmuxBridgeUrl: () => Promise.resolve(undefined),
    validateBridgeUrl: (url) => url,
    probeTmuxBridge: () => Promise.resolve(false),
  });

  assert.equal(missingUrl.allowed, false);
  assert.equal(missingUrl.reason, "missing_bridge_url");

  const unreachable = await evaluateTmuxBridgeGate({
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3341"),
    validateBridgeUrl: () => "https://localhost:3341",
    probeTmuxBridge: () => Promise.resolve(false),
  });

  assert.equal(unreachable.allowed, false);
  assert.equal(unreachable.reason, "bridge_unreachable");
  assert.match(buildTmuxBridgeGateErrorMessage(unreachable.reason), /Terminal access is not available/i);
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

void test("execute_office_js aborts if cancellation happens during approval", async () => {
  let executeCount = 0;

  const abortController = new AbortController();
  const [officeTool] = await applyExperimentalToolGates([
    createTestTool("execute_office_js", () => {
      executeCount += 1;
    }),
  ], {
    requestOfficeJsExecuteApproval: () => {
      abortController.abort();
      return Promise.resolve(true);
    },
  });

  await assert.rejects(
    () => officeTool.execute("call-office", {
      explanation: "Rebuild totals",
      code: "return { ok: true };",
    }, abortController.signal),
    /aborted/i,
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

void test("python bridge approvals fail open when no approval handler is configured", async () => {
  let executeCount = 0;

  const [pythonTool] = await applyExperimentalToolGates([
    createTestTool("python_run", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(true),
  });

  await pythonTool.execute("call-python-no-approval-handler", {
    code: "print('hello')",
  });

  assert.equal(executeCount, 1);
});

void test("python fallback tools execute even when bridge gate fails", async () => {
  let executeCount = 0;
  let approvalCalls = 0;

  const tools = [
    createTestTool("python_run", () => { executeCount += 1; }),
    createTestTool("python_transform_range", () => { executeCount += 1; }),
  ];

  const gatedTools = await applyExperimentalToolGates(tools, {
    getPythonBridgeUrl: () => Promise.resolve(undefined),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(false),
    requestPythonBridgeApproval: () => {
      approvalCalls += 1;
      return Promise.resolve(true);
    },
  });

  const pythonRun = gatedTools.find((tool) => tool.name === "python_run");
  const pythonTransform = gatedTools.find((tool) => tool.name === "python_transform_range");

  assert.ok(pythonRun);
  assert.ok(pythonTransform);

  await pythonRun.execute("call-python-run", { code: "print('hi')" });
  await pythonTransform.execute("call-python-transform", {
    range: "Sheet1!A1:A2",
    code: "result = [[1], [2]]",
  });

  assert.equal(executeCount, 2);
  assert.equal(approvalCalls, 0);
});

void test("python fallback tools return structured gate errors when configured bridge is unreachable", async () => {
  let executeCount = 0;

  const [pythonTool] = await applyExperimentalToolGates([
    createTestTool("python_run", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: () => "https://localhost:3340",
    probePythonBridge: () => Promise.resolve(false),
  });

  const result = await pythonTool.execute("call-python-unreachable", { code: "print('hi')" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.match(text, /Native Python is not available/i);
  assert.match(text, /not reachable/i);
  assert.match(text, /Skill: python-bridge/i);

  const resultDetails: unknown = result.details;
  assertPythonGateError(resultDetails);

  assert.equal(executeCount, 0);
});

void test("python_transform_range gate errors keep transform detail kind", async () => {
  let executeCount = 0;

  const [tool] = await applyExperimentalToolGates([
    createTestTool("python_transform_range", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: () => "https://localhost:3340",
    probePythonBridge: () => Promise.resolve(false),
  });

  const result = await tool.execute("call-python-transform-unreachable", {
    range: "Sheet1!A1:A2",
    code: "result = [[1], [2]]",
  });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.match(text, /Native Python is not available/i);
  assert.match(text, /not reachable/i);
  assert.match(text, /Skill: python-bridge/i);

  const details: unknown = result.details;
  assertPythonTransformRangeGateError(details);
  assert.equal(isBridgeGateError(details), true);

  assert.equal(executeCount, 0);
});

void test("libreoffice_convert still requires configured + reachable bridge", async () => {
  let executeCount = 0;

  const [toolWhenMissing] = await applyExperimentalToolGates([
    createTestTool("libreoffice_convert", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve(undefined),
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(false),
  });

  const missingResult = await toolWhenMissing.execute("call-libreoffice-missing", {
    input_path: "/tmp/source.xlsx",
    target_format: "csv",
  });
  const missingText = missingResult.content[0]?.type === "text" ? missingResult.content[0].text : "";
  assert.match(missingText, /Native Python is not available/i);
  assert.match(missingText, /default URL|URL override|not configured/i);
  assert.match(missingText, /Skill: python-bridge/i);

  const missingDetails: unknown = missingResult.details;
  assertLibreOfficeGateError(missingDetails, "missing_bridge_url");

  const [toolWhenUnreachable] = await applyExperimentalToolGates([
    createTestTool("libreoffice_convert", () => {
      executeCount += 1;
    }),
  ], {
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
    validatePythonBridgeUrl: () => "https://localhost:3340",
    probePythonBridge: () => Promise.resolve(false),
  });

  const unreachableResult = await toolWhenUnreachable.execute("call-libreoffice-unreachable", {
    input_path: "/tmp/source.xlsx",
    target_format: "csv",
  });
  const unreachableText = unreachableResult.content[0]?.type === "text" ? unreachableResult.content[0].text : "";
  assert.match(unreachableText, /Native Python is not available/i);
  assert.match(unreachableText, /not reachable/i);
  assert.match(unreachableText, /Skill: python-bridge/i);

  const unreachableDetails: unknown = unreachableResult.details;
  assertLibreOfficeGateError(unreachableDetails, "bridge_unreachable");

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
    validatePythonBridgeUrl: (url) => url,
    probePythonBridge: () => Promise.resolve(false),
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
  assert.match(buildPythonBridgeGateErrorMessage(unreachable.reason), /Native Python is not available/i);
  assert.match(buildPythonBridgeGateErrorMessage(unreachable.reason), /not reachable/i);
});

void test("bridge gate helper detects only gate-shaped bridge details", () => {
  const gateError = {
    kind: "python_bridge",
    ok: false,
    action: "run_python",
    error: "Native Python is not available right now because the Python bridge is not reachable at the configured URL.",
    gateReason: "bridge_unreachable",
    skillHint: "python-bridge",
  };

  const nonGateError = {
    kind: "python_bridge",
    ok: false,
    action: "run_python",
    error: "NameError: x is not defined",
  };

  assert.equal(isBridgeGateError(gateError), true);
  assert.equal(isBridgeGateError(nonGateError), false);
});
