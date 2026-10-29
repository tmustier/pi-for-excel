import assert from "node:assert/strict";
import { test } from "node:test";

import { createPythonRunTool } from "../src/tools/python-run.ts";
import type { PythonBridgeResponse } from "../src/tools/python-run.ts";

function firstText(result: { content: Array<{ type: string; text?: string }> }): string {
  const item = result.content.find((c) => c.type === "text");
  return item?.text ?? "";
}

function makePyodideResponse(overrides: Partial<PythonBridgeResponse> = {}): PythonBridgeResponse {
  return {
    ok: true,
    action: "run_python",
    exit_code: 0,
    stdout: "hello from pyodide",
    metadata: { backend: "pyodide" },
    ...overrides,
  };
}

void test("python_run reports Pyodide failures with stderr and stdout when the error is empty", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => true,
    callPyodide: () => Promise.resolve(makePyodideResponse({
      ok: false,
      exit_code: 1,
      error: "",
      stdout: "before\n",
      stderr: "Traceback (most recent call last):\n  File \"<exec>\", line 1, in <module>\n" +
        "ModuleNotFoundError: No module named 'pandas'\nThe module 'pandas' is included in the Pyodide distribution, but it is not installed.",
    })),
  });

  const result = await tool.execute("tc-4", { code: "print('before')\nimport pandas" });
  const text = firstText(result);

  assert.equal(result.details?.ok, false);
  assert.match(text, /^Error: Python failed in Pyodide \(in-browser\) \(exit code 1\): ModuleNotFoundError: No module named 'pandas'/u);
  assert.match(text, /Stderr:[\s\S]*is not installed/u);
  assert.match(text, /Stdout:[\s\S]*before/u);
});

void test("python_run Pyodide returns result_json correctly", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => true,
    callPyodide: () => Promise.resolve(makePyodideResponse({
      result_json: '{"answer": 42}',
    })),
  });

  const result = await tool.execute("tc-5", { code: "result = {'answer': 42}" });

  assert.equal(result.details?.ok, true);
  assert.match(firstText(result), /Result JSON/);
  assert.match(firstText(result), /42/);
});

void test("python_run validates params before checking backend", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => true,
    callPyodide: () => {
      throw new Error("should not be called");
    },
  });

  const result = await tool.execute("tc-6", { code: "" });

  assert.equal(result.details?.ok, false);
  assert.match(firstText(result), /code is required/i);
});

void test("python_run passes input_json to Pyodide", async () => {
  let receivedRequest: unknown = null;

  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => true,
    callPyodide: (request) => {
      receivedRequest = request;
      return Promise.resolve(makePyodideResponse());
    },
  });

  await tool.execute("tc-7", {
    code: "x = input_data['values']",
    input_json: '{"values": [1, 2, 3]}',
  });

  assert.ok(receivedRequest);
  assert.equal(
    (receivedRequest as Record<string, unknown>).input_json,
    '{"values": [1, 2, 3]}',
  );
});
