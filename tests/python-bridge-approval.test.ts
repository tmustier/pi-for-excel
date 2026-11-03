import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import type { AgentTool } from "@earendil-works/pi-agent-core";
import { applyExperimentalToolGates } from "../src/tools/experimental-tool-gates.ts";
import { createPythonRunTool } from "../src/tools/python-run.ts";
import { createPythonTransformRangeTool } from "../src/tools/python-transform-range.ts";
import { createLibreOfficeConvertTool } from "../src/tools/libreoffice-convert.ts";

const approvedUrl = "https://localhost:3340";
const replacementUrl = "https://localhost:4340";

for (const name of ["python_run", "python_transform_range", "libreoffice_convert"]) {
  void test(`${name} refuses a bridge URL changed during approval`, async () => {
    let configuredUrl = approvedUrl;
    let persistedApproval: string | undefined;
    let nativeCalls = 0;
    const getBridgeConfig = () => Promise.resolve({ url: configuredUrl, configured: true });
    const getApprovedBridgeUrl = () => Promise.resolve(persistedApproval);
    const callBridge = () => {
      nativeCalls += 1;
      return Promise.resolve({ ok: true as const, action: "run_python" as const, result_json: "[[1]]" });
    };
    let tool: AgentTool;
    let params: { code: string } | { range: string; code: string } | { input_path: string; target_format: string };
    if (name === "python_run") {
      tool = createPythonRunTool({ getBridgeConfig, getApprovedBridgeUrl, callBridge });
      params = { code: "result = [[1]]" };
    } else if (name === "python_transform_range") {
      tool = createPythonTransformRangeTool({
        getBridgeConfig, getApprovedBridgeUrl, callBridge,
        readInputRange: () => Promise.resolve({ sheetName: "Sheet1", address: "A1", values: [[1]], text: [["1"]] }),
        writeOutputValues: () => Promise.reject(new Error("must not write")),
      });
      params = { range: "Sheet1!A1", code: "result = [[1]]" };
    } else {
      tool = createLibreOfficeConvertTool({
        getBridgeConfig, getApprovedBridgeUrl,
        callBridge: () => { nativeCalls += 1; return Promise.resolve({ ok: true, action: "convert" }); },
      });
      params = { input_path: "/tmp/source.xlsx", target_format: "csv" };
    }
    const [gated] = await applyExperimentalToolGates([tool], {
      getPythonBridgeUrl: () => Promise.resolve(configuredUrl),
      probePythonBridge: () => Promise.resolve(true),
      getApprovedPythonBridgeUrl: getApprovedBridgeUrl,
      setApprovedPythonBridgeUrl: (url) => { persistedApproval = url; return Promise.resolve(); },
      requestPythonBridgeApproval: () => { configuredUrl = replacementUrl; return Promise.resolve(true); },
    });
    assert.ok(gated);
    const result = await gated.execute("changed-url", params);
    assert.equal(nativeCalls, 0);
    const text = result.content[0];
    assert.ok(text?.type === "text");
    assert.match(text.text, /requires user approval/u);
    assert.equal(persistedApproval, approvedUrl);
  });
}

void test("concurrent Python tools share one pending approval per bridge URL", async () => {
  let prompts = 0;
  let executions = 0;
  let stored: string | undefined;
  let answer: (approved: boolean) => void = () => {};
  const approval = new Promise<boolean>((resolve) => { answer = resolve; });
  const tools: AgentTool[] = ["python_run", "python_transform_range"].map((name) => ({
    name, label: name, description: name, parameters: Type.Object({}),
    execute: () => { executions += 1; return Promise.resolve({ content: [], details: undefined }); },
  }));
  const gateDependencies = {
    getPythonBridgeUrl: () => Promise.resolve(approvedUrl),
    probePythonBridge: () => Promise.resolve(true),
    getApprovedPythonBridgeUrl: () => Promise.resolve(stored),
    setApprovedPythonBridgeUrl: (url: string) => { stored = url; return Promise.resolve(); },
    requestPythonBridgeApproval: () => { prompts += 1; return approval; },
  };
  const gated = (await Promise.all(tools.map((tool) => applyExperimentalToolGates([tool], gateDependencies)))).flat();
  const requests = gated.map((tool) => tool.execute(tool.name, {}));
  await new Promise<void>((resolve) => { setImmediate(resolve); });
  answer(true);
  await Promise.all(requests);
  assert.equal(prompts, 1);
  assert.equal(executions, 2);
  assert.equal(stored, approvedUrl);
});

void test("an unapproved default bridge is skipped without executing native Python", async () => {
  let nativeCalls = 0;
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve({ url: approvedUrl, configured: false }),
    getApprovedBridgeUrl: () => Promise.resolve(undefined),
    callBridge: () => { nativeCalls += 1; return Promise.resolve({ ok: true, action: "run_python" }); },
    isPyodideAvailable: () => true,
    callPyodide: () => Promise.resolve({ ok: true, action: "run_python", stdout: "browser" }),
  });
  const result = await tool.execute("default-unapproved", { code: "print('browser')" });
  assert.equal(nativeCalls, 0);
  assert.equal(result.details?.ok, true);
  assert.equal(result.details?.bridgeUrl, undefined);
  assert.equal(result.details?.stdoutPreview, "browser");
});
