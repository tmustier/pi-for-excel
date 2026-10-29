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


void test("files writes and deletes use the same ownership consent in Auto and Confirm", async () => {
  const existing = new Map<string, { size: number; readOnly: boolean }>([
    ["notes/user.txt", { size: 40, readOnly: false }],
    ["assistant-docs/README.md", { size: 10, readOnly: true }],
  ]);
  const calls: string[] = [];
  const asked: string[] = [];
  let mode: "yolo" | "safe" = "yolo";
  let answer = true;
  const tool: AgentTool = {
    ...createTestTool("files"),
    execute: (_id, params) => {
      const { action, path } = params as { action: string; path: string };
      calls.push(`${action}:${path}`);
      if (action === "write") existing.set(path, { size: 1, readOnly: false });
      if (action === "delete") existing.delete(path);
      return Promise.resolve({ content: [{ type: "text", text: "ok" }], details: undefined });
    },
  };
  const [gated] = await applyExperimentalToolGates([tool], {
    getExecutionMode: () => Promise.resolve(mode),
    inspectWorkspaceFile: (path) => Promise.resolve(existing.get(path) ?? null),
    assistantCreatedFiles: new Set<string>(),
    requestLocalChangeApproval: (request) => {
      asked.push(request.title);
      return Promise.resolve(answer);
    },
  });
  assert.ok(gated);
  const run = (params: Record<string, string>) => gated.execute("call", params);

  await run({ action: "write", path: "out/report.csv" });
  await run({ action: "write", path: "out/report.csv" });
  await run({ action: "delete", path: "out/report.csv" });
  await run({ action: "read", path: "notes/user.txt" });
  await run({ action: "write", path: "assistant-docs/README.md" });
  assert.deepEqual(asked, []);

  answer = false;
  await assert.rejects(run({ action: "write", path: "notes/user.txt" }), /cancelled by user/);
  await assert.rejects(run({ action: "delete", path: "notes/user.txt" }), /cancelled by user/);
  assert.deepEqual(asked, ["Replace a file?", "Delete a file?"]);
  assert.ok(existing.has("notes/user.txt"));

  mode = "safe";
  answer = true;
  await run({ action: "write", path: "out/new.csv" });
  await run({ action: "write", path: "out/new.csv" });
  await run({ action: "delete", path: "out/new.csv" });
  assert.equal(asked.length, 2);
  answer = false;
  await assert.rejects(run({ action: "write", path: "notes/user.txt" }), /cancelled by user/);
  await assert.rejects(run({ action: "delete", path: "notes/user.txt" }), /cancelled by user/);
  assert.deepEqual(asked, ["Replace a file?", "Delete a file?", "Replace a file?", "Delete a file?"]);
  assert.deepEqual(calls, [
    "write:out/report.csv", "write:out/report.csv", "delete:out/report.csv",
    "read:notes/user.txt", "write:assistant-docs/README.md", "write:out/new.csv",
    "write:out/new.csv", "delete:out/new.csv",
  ]);
});

void test("files changes fail closed when no approval UI is available", async () => {
  let ran = false;
  const [gated] = await applyExperimentalToolGates([createTestTool("files", () => { ran = true; })], {
    getExecutionMode: () => Promise.resolve("yolo"),
    inspectWorkspaceFile: () => Promise.resolve({ size: 5, readOnly: false }),
  });
  assert.ok(gated);
  await assert.rejects(gated.execute("call", { action: "delete", path: "a.txt" }), /approval/);
  assert.equal(ran, false);
});

void test("extensions_manager asks before running new code or uninstalling, in Auto mode too", async () => {
  const requests: Array<{ title: string; message: string }> = [];
  const actions: string[] = [];
  const tool: AgentTool = {
    ...createTestTool("extensions_manager"),
    execute: (_id, params) => {
      actions.push((params as { action: string }).action);
      return Promise.resolve({ content: [{ type: "text", text: "ok" }], details: undefined });
    },
  };
  const [gated] = await applyExperimentalToolGates([tool], {
    getExecutionMode: () => Promise.resolve("yolo"),
    listExtensions: () => [{ id: "ext.a", name: "Quick KPI", grantedCapabilities: ["commands.register"] }],
    requestLocalChangeApproval: (request) => {
      requests.push({ title: request.title, message: request.message });
      return Promise.resolve(true);
    },
  });
  assert.ok(gated);

  await gated.execute("c1", { action: "list" });
  await gated.execute("c2", { action: "set_enabled", extension_id: "ext.a", enabled: false });
  await gated.execute("c2b", { action: "reload", extension_id: "ext.a" });
  assert.equal(requests.length, 0, "listing, disabling and reloading in Auto mode need no approval");

  await gated.execute("c3", {
    action: "install_code",
    name: "Exfil",
    code: "export function activate(api) { api.registerTool('x', {}); }",
    capabilities: ["tools.register", "http.fetch"],
  });
  await gated.execute("c4", { action: "set_enabled", extension_id: "ext.a", enabled: true });
  await gated.execute("c5", { action: "uninstall", extension_id: "ext.a" });

  assert.deepEqual(requests.map((r) => r.title), ["Install an extension?", "Enable an extension?", "Uninstall an extension?"]);
  const install = requests[0]?.message ?? "";
  assert.match(install, /tools\.register, http\.fetch/);
  assert.match(install, /api\.registerTool\('x'/);
  assert.match(requests[1]?.message ?? "", /Quick KPI/);
  assert.deepEqual(actions, ["list", "set_enabled", "reload", "install_code", "set_enabled", "uninstall"]);
});

void test("extensions_manager reload asks in Confirm mode, and a declined reload does not run", async () => {
  const titles: string[] = [];
  let answer = false;
  const actions: string[] = [];
  const tool: AgentTool = {
    ...createTestTool("extensions_manager"),
    execute: (_id, params) => {
      actions.push((params as { action: string }).action);
      return Promise.resolve({ content: [{ type: "text", text: "ok" }], details: undefined });
    },
  };
  const [gated] = await applyExperimentalToolGates([tool], {
    getExecutionMode: () => Promise.resolve("safe"),
    listExtensions: () => [{ id: "ext.a", name: "Quick KPI", grantedCapabilities: [] }],
    requestLocalChangeApproval: (request) => {
      titles.push(request.title);
      return Promise.resolve(answer);
    },
  });
  assert.ok(gated);

  await assert.rejects(gated.execute("c1", { action: "reload", extension_id: "ext.a" }), /Reloading extension "Quick KPI" cancelled/);
  answer = true;
  await gated.execute("c2", { action: "reload", extension_id: "ext.a" });

  assert.deepEqual(titles, ["Reload an extension?", "Reload an extension?"]);
  assert.deepEqual(actions, ["reload"]);
});

void test("stopping the turn withdraws a pending files consent request and the change does not run", async () => {
  let ran = false;
  const [gated] = await applyExperimentalToolGates([createTestTool("files", () => { ran = true; })], {
    getExecutionMode: () => Promise.resolve("safe"),
    inspectWorkspaceFile: () => Promise.resolve({ size: 5, readOnly: false }),
    requestLocalChangeApproval: (request) => new Promise<boolean>((resolve) => {
      request.signal?.addEventListener("abort", () => resolve(false), { once: true });
      markAsked();
    }),
  });
  assert.ok(gated);
  let markAsked: () => void = () => {};
  const asked = new Promise<void>((resolve) => { markAsked = resolve; });
  const controller = new AbortController();
  const pending = gated.execute("call", { action: "delete", path: "a.txt" }, controller.signal);
  await asked;
  controller.abort();
  await assert.rejects(pending, /Aborted/);
  assert.equal(ran, false);
});
