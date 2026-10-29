import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import {
  createPythonRunTool,
  type PythonBridgeConfig,
  type PythonBridgeRequest,
  type PythonBridgeResponse,
  type PythonRunToolDetails,
} from "../src/tools/python-run.ts";

function firstText(result: AgentToolResult<PythonRunToolDetails>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") {
    throw new Error("Expected first content block to be text");
  }
  return first.text;
}

void test("python_run returns guidance when no backend is available", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => false,
  });

  const result = await tool.execute("tc-missing", { code: "print('hello')" });

  assert.match(firstText(result), /Python is unavailable/u);
  assert.match(firstText(result), /Skill: python-bridge/u);
  assert.equal(result.details?.kind, "python_bridge");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "no_python_runtime");
  assert.equal(result.details?.skillHint, "python-bridge");
});

void test("python_run validates input_json before bridge call", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", configured: true }),
    callBridge: () => Promise.resolve({ ok: true, action: "run_python" }),
  });

  const result = await tool.execute("tc-validate", {
    code: "print('x')",
    input_json: "{not-json}",
  });

  assert.match(firstText(result), /input_json must be valid JSON/u);
  assert.equal(result.details?.ok, false);
});

void test("python_run sends v1 bridge payload and renders stdout/result", async () => {
  let capturedRequest: PythonBridgeRequest | null = null;
  let capturedConfig: PythonBridgeConfig | null = null;

  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
      token: "secret-token",
      configured: true,
    }),
    callBridge: (
      request: PythonBridgeRequest,
      config: PythonBridgeConfig,
    ): Promise<PythonBridgeResponse> => {
      capturedRequest = request;
      capturedConfig = config;

      return Promise.resolve({
        ok: true,
        action: "run_python",
        exit_code: 0,
        stdout: "done",
        result_json: "{\"rows\": [[1,2],[3,4]]}",
      });
    },
  });

  const result = await tool.execute("tc-contract", {
    code: "result = {'rows': [[1,2],[3,4]]}",
    input_json: "{\"source\":\"A1:B2\"}",
    timeout_ms: 5000,
  });

  assert.ok(capturedRequest);
  assert.equal(capturedRequest?.code, "result = {'rows': [[1,2],[3,4]]}");
  assert.equal(capturedRequest?.input_json, "{\"source\":\"A1:B2\"}");
  assert.equal(capturedRequest?.timeout_ms, 5000);

  assert.ok(capturedConfig);
  assert.equal(capturedConfig?.url, "https://localhost:3340");
  assert.equal(capturedConfig?.token, "secret-token");

  const text = firstText(result);
  assert.match(text, /Ran Python snippet/u);
  assert.match(text, /Result JSON:/u);
  assert.match(text, /Stdout:/u);
  assert.equal(result.details?.ok, true);
  assert.equal(result.details?.action, "run_python");
  assert.equal(result.details?.exitCode, 0);
});

void test("python_run bridge errors are surfaced", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", configured: true }),
    callBridge: () => Promise.reject(new Error("bridge unavailable")),
  });

  const result = await tool.execute("tc-error", {
    code: "print('x')",
  });

  assert.match(firstText(result), /Error: bridge unavailable/u);
  assert.match(firstText(result), /Skill: python-bridge/u);
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "bridge unavailable");
  assert.equal(result.details?.skillHint, "python-bridge");
});

void test("python_run falls back to Pyodide only when the default bridge is unreachable", async (t) => {
  const traceback = "Traceback (most recent call last):\n  File \"<string>\", line 1, in <module>\nModuleNotFoundError: No module named 'pandas'";
  const cases: Array<{ name: string; fetch: () => Promise<Response>; fallback: boolean; text: RegExp }> = [
    {
      name: "connection refused",
      fetch: () => Promise.reject(new TypeError("Failed to fetch")),
      fallback: true,
      text: /pyodide fallback/u,
    },
    {
      name: "python error on the bridge",
      fetch: () => Promise.resolve(new Response(JSON.stringify({
        ok: false, error: traceback, action: "run_python", exit_code: 1, stderr: traceback,
      }), { status: 400 })),
      fallback: false,
      text: /Python failed in local Python \(bridge at https:\/\/localhost:3340\) \(exit code 1\): ModuleNotFoundError: No module named 'pandas'/u,
    },
    {
      name: "bridge timeout",
      fetch: () => Promise.resolve(new Response(JSON.stringify({
        ok: false, error: "Process timed out after 2000ms.",
      }), { status: 504 })),
      fallback: false,
      text: /timed out after 2000ms/u,
    },
  ];

  const originalFetch = globalThis.fetch;
  t.after(() => {
    globalThis.fetch = originalFetch;
  });

  for (const testCase of cases) {
    globalThis.fetch = testCase.fetch;
    let pyodideCalls = 0;
    const tool = createPythonRunTool({
      getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", configured: false }),
      isPyodideAvailable: () => true,
      callPyodide: () => {
        pyodideCalls += 1;
        return Promise.resolve({ ok: true, action: "run_python", exit_code: 0, stdout: "pyodide fallback" });
      },
    });

    const result = await tool.execute("tc-fallback", { code: "import pandas", timeout_ms: 2000 });

    assert.equal(pyodideCalls, testCase.fallback ? 1 : 0, testCase.name);
    assert.match(firstText(result), testCase.text, testCase.name);
    assert.equal(result.details?.ok, testCase.fallback, testCase.name);
  }
});

void test("python_run handles explicit bridge-level rejection payloads", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", configured: true }),
    callBridge: () => Promise.resolve({
      ok: false,
      action: "run_python",
      error: "execution denied",
    }),
  });

  const result = await tool.execute("tc-reject", {
    code: "print('x')",
  });

  assert.equal(firstText(result), "Error: execution denied");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "execution denied");
});
