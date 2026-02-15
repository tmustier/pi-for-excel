import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@mariozechner/pi-agent-core";

import {
  createPythonTransformRangeTool,
} from "../src/tools/python-transform-range.ts";
import type {
  PythonTransformRangeDetails,
} from "../src/tools/tool-details.ts";
import type {
  PythonBridgeConfig,
  PythonBridgeRequest,
  PythonBridgeResponse,
} from "../src/tools/python-run.ts";

function firstText(result: AgentToolResult<PythonTransformRangeDetails>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") {
    throw new Error("Expected first content block to be text");
  }
  return first.text;
}

void test("python_transform_range returns bridge setup guidance when URL is missing", async () => {
  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({
      sheetName: "Sheet1",
      address: "A1:B2",
      values: [[1, 2], [3, 4]],
    }),
    getBridgeConfig: () => Promise.resolve(null),
    isPyodideAvailable: () => false,
  });

  const result = await tool.execute("tc-missing", {
    range: "Sheet1!A1:B2",
    code: "result = input_data['values']",
  });

  assert.match(firstText(result), /Python is unavailable/u);
  assert.equal(result.details?.kind, "python_transform_range");
  assert.equal(result.details?.blocked, false);
  assert.equal(result.details?.error, "no_python_runtime");
});

void test("python_transform_range reads source, runs python, and writes transformed output", async () => {
  let capturedBridgeRequest: PythonBridgeRequest | null = null;
  let capturedBridgeConfig: PythonBridgeConfig | null = null;
  let capturedWriteRequest: {
    outputStartCell: string;
    values: unknown[][];
    allowOverwrite: boolean;
  } | null = null;

  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({
      sheetName: "Sheet1",
      address: "A1:B2",
      values: [[1, 2], [3, 4]],
    }),
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
      token: "secret-token",
    }),
    callBridge: (
      request: PythonBridgeRequest,
      config: PythonBridgeConfig,
    ): Promise<PythonBridgeResponse> => {
      capturedBridgeRequest = request;
      capturedBridgeConfig = config;

      return Promise.resolve({
        ok: true,
        action: "run_python",
        exit_code: 0,
        result_json: "{\"values\":[[10,20],[30,40]]}",
      });
    },
    writeOutputValues: (request) => {
      capturedWriteRequest = request;
      return Promise.resolve({
        blocked: false,
        outputAddress: "Sheet1!D1:E2",
        rowsWritten: 2,
        colsWritten: 2,
        formulaErrorCount: 0,
      });
    },
  });

  const result = await tool.execute("tc-success", {
    range: "Sheet1!A1:B2",
    code: "result = {'values': [[10,20],[30,40]]}",
    output_start_cell: "D1",
  });

  assert.ok(capturedBridgeRequest);
  assert.equal(capturedBridgeRequest?.code, "result = {'values': [[10,20],[30,40]]}");

  assert.equal(
    capturedBridgeRequest?.input_json,
    "{\"range\":\"Sheet1!A1:B2\",\"values\":[[1,2],[3,4]]}",
  );

  assert.ok(capturedBridgeConfig);
  assert.equal(capturedBridgeConfig?.url, "https://localhost:3340");
  assert.equal(capturedBridgeConfig?.token, "secret-token");

  assert.ok(capturedWriteRequest);
  assert.equal(capturedWriteRequest?.outputStartCell, "Sheet1!D1");
  assert.deepEqual(capturedWriteRequest?.values, [[10, 20], [30, 40]]);
  assert.equal(capturedWriteRequest?.allowOverwrite, false);

  const text = firstText(result);
  assert.match(text, /Transformed/u);
  assert.match(text, /Sheet1!D1:E2/u);
  assert.equal(result.details?.blocked, false);
  assert.equal(result.details?.rowsWritten, 2);
  assert.equal(result.details?.colsWritten, 2);
});

void test("python_transform_range surfaces overwrite blocking from destination write", async () => {
  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({
      sheetName: "Sheet1",
      address: "A1:A2",
      values: [[1], [2]],
    }),
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
    }),
    callBridge: () => Promise.resolve({
      ok: true,
      action: "run_python",
      result_json: "[[100],[200]]",
    }),
    writeOutputValues: () => Promise.resolve({
      blocked: true,
      outputAddress: "Sheet1!A1:A2",
      existingCount: 2,
    }),
  });

  const result = await tool.execute("tc-blocked", {
    range: "Sheet1!A1:A2",
    code: "result = [[100],[200]]",
  });

  const text = firstText(result);
  assert.match(text, /Transform blocked/u);
  assert.match(text, /allow_overwrite/u);
  assert.equal(result.details?.blocked, true);
  assert.equal(result.details?.existingCount, 2);
  assert.equal(result.details?.outputAddress, "Sheet1!A1:A2");
});

void test("python_transform_range requires result_json from python execution", async () => {
  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({
      sheetName: "Sheet1",
      address: "A1:B2",
      values: [[1, 2], [3, 4]],
    }),
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
    }),
    callBridge: () => Promise.resolve({
      ok: true,
      action: "run_python",
      stdout: "no result",
    }),
    writeOutputValues: () => Promise.resolve({
      blocked: false,
      outputAddress: "Sheet1!A1:B2",
      rowsWritten: 2,
      colsWritten: 2,
      formulaErrorCount: 0,
    }),
  });

  const result = await tool.execute("tc-no-result", {
    range: "Sheet1!A1:B2",
    code: "print('hello')",
  });

  assert.match(firstText(result), /returned no result_json/u);
  assert.equal(result.details?.blocked, false);
  assert.match(result.details?.error ?? "", /result_json/u);
});
