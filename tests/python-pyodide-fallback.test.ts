/**
 * Tests for Pyodide fallback behavior in python_run and python_transform_range.
 *
 * These tests verify:
 * - Native bridge is preferred when configured
 * - Pyodide fallback is used when no bridge is configured
 * - Proper error when neither backend is available
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { createPythonRunTool } from "../src/tools/python-run.ts";
import type { PythonBridgeResponse } from "../src/tools/python-run.ts";

function firstText(result: { content: Array<{ type: string; text?: string }> }): string {
  const item = result.content.find((c) => c.type === "text");
  return item?.text ?? "";
}

function makeBridgeResponse(overrides: Partial<PythonBridgeResponse> = {}): PythonBridgeResponse {
  return {
    ok: true,
    action: "run_python",
    exit_code: 0,
    stdout: "hello from bridge",
    ...overrides,
  };
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

void test("python_run prefers native bridge when configured", async () => {
  let bridgeCalled = false;
  let pyodideCalled = false;

  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340" }),
    callBridge: () => {
      bridgeCalled = true;
      return Promise.resolve(makeBridgeResponse());
    },
    isPyodideAvailable: () => {
      return true;
    },
    callPyodide: () => {
      pyodideCalled = true;
      return Promise.resolve(makePyodideResponse());
    },
  });

  const result = await tool.execute("tc-1", { code: "print('hello')" });

  assert.equal(bridgeCalled, true);
  assert.equal(pyodideCalled, false);
  assert.match(firstText(result), /hello from bridge/);
});

void test("python_run falls back to Pyodide when no bridge configured", async () => {
  let pyodideCalled = false;

  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => true,
    callPyodide: (_request) => {
      pyodideCalled = true;
      return Promise.resolve(makePyodideResponse());
    },
  });

  const result = await tool.execute("tc-2", { code: "print('hello')" });

  assert.equal(pyodideCalled, true);
  assert.match(firstText(result), /hello from pyodide/);
  assert.equal(result.details?.ok, true);
});

void test("python_run returns error when neither bridge nor Pyodide available", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => false,
  });

  const result = await tool.execute("tc-3", { code: "print('hello')" });

  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "no_python_runtime");
  assert.match(firstText(result), /Python is unavailable/);
});

void test("python_run propagates Pyodide execution errors", async () => {
  const tool = createPythonRunTool({
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => true,
    callPyodide: () => Promise.resolve(makePyodideResponse({
      ok: false,
      error: "NameError: name 'foo' is not defined",
    })),
  });

  const result = await tool.execute("tc-4", { code: "foo()" });

  assert.equal(result.details?.ok, false);
  assert.match(firstText(result), /NameError/);
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
