import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool } from "@earendil-works/pi-agent-core";
import { Type } from "typebox";

import {
  applyExperimentalToolGates,
  buildOfficeJsExecuteApprovalMessage,
} from "../src/tools/experimental-tool-gates.ts";
import { decodeToolDetails, decodeToolDetailsOfKind, isBridgeGateError } from "../src/tools/tool-details.ts";

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
  raw: unknown,
  reason: "missing_bridge_url" | "bridge_unreachable",
): void {
  const details = decodeToolDetailsOfKind(raw, "tmux_bridge");
  assert.ok(details);
  assert.equal(details.ok, false);
  assert.equal(details.gateReason, reason);
  assert.equal(details.skillHint, "tmux-bridge");
}

function assertPythonGateError(raw: unknown): void {
  const details = decodeToolDetailsOfKind(raw, "python_bridge");
  assert.ok(details);
  assert.equal(details.ok, false);
  assert.equal(details.gateReason, "bridge_unreachable");
  assert.equal(details.skillHint, "python-bridge");
}

function assertPythonTransformRangeGateError(raw: unknown): void {
  const details = decodeToolDetailsOfKind(raw, "python_transform_range");
  assert.ok(details);
  assert.equal(details.blocked, false);
  assert.equal(details.gateReason, "bridge_unreachable");
  assert.equal(details.skillHint, "python-bridge");
  assert.match(details.error ?? "", /not reachable/i);
}

function assertLibreOfficeGateError(
  raw: unknown,
  reason: "missing_bridge_url" | "bridge_unreachable",
): void {
  const details = decodeToolDetailsOfKind(raw, "libreoffice_bridge");
  assert.ok(details);
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

  assertTmuxGateError(result.details, "missing_bridge_url");
  assert.equal(decodeToolDetailsOfKind(result.details, "tmux_bridge")?.action, "capture_pane");

  assert.equal(probeCalled, true);
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


void test("execute_wps_js uses the same direct-JS approval gate with WPS labeling", async () => {
  let executeCount = 0;

  const [wpsTool] = await applyExperimentalToolGates([
    createTestTool("execute_wps_js", () => {
      executeCount += 1;
    }),
  ], {
    requestOfficeJsExecuteApproval: ({ explanation, code, apiName }) => {
      assert.equal(explanation, "Inspect WPS workbook");
      assert.equal(code, "return Application.ActiveWorkbook.Name;");
      assert.equal(apiName, "WPS JSAPI");
      return Promise.resolve(true);
    },
  });

  await wpsTool.execute("call-wps", {
    explanation: "Inspect WPS workbook",
    code: "return Application.ActiveWorkbook.Name;",
  });

  assert.equal(executeCount, 1);
});

void test("execute_office_js requires approval in Auto mode when code references ambient browser authority", async () => {
  let executeCount = 0;
  let approvalCount = 0;

  const [officeTool] = await applyExperimentalToolGates([
    createTestTool("execute_office_js", () => {
      executeCount += 1;
    }),
  ], {
    getExecutionMode: () => Promise.resolve("yolo" as const),
    requestOfficeJsExecuteApproval: (request) => {
      approvalCount += 1;
      assert.deepEqual(request.riskIdentifiers, ["fetch"]);

      const message = buildOfficeJsExecuteApprovalMessage(request);
      assert.match(message, /beyond the Excel API: fetch/u);

      return Promise.resolve(true);
    },
  });

  await officeTool.execute("call-office", {
    explanation: "Post data to service",
    code: "await fetch(\"https://example.com\", { method: \"POST\" });",
  });

  assert.equal(executeCount, 1);
  assert.equal(approvalCount, 1);
});

void test("execute_office_js denial of risky code cancels execution in Auto mode", async () => {
  let executeCount = 0;

  const [officeTool] = await applyExperimentalToolGates([
    createTestTool("execute_office_js", () => {
      executeCount += 1;
    }),
  ], {
    getExecutionMode: () => Promise.resolve("yolo" as const),
    requestOfficeJsExecuteApproval: () => Promise.resolve(false),
  });

  await assert.rejects(
    () => officeTool.execute("call-office", {
      explanation: "Read settings",
      code: "return localStorage.getItem(\"connections.store.v1\");",
    }),
    /cancelled by user/i,
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

  assertPythonTransformRangeGateError(result.details);
  const details = decodeToolDetails(result.details);
  assert.ok(details);
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

