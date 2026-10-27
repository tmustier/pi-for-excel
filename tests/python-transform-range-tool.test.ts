import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import {
  createPythonTransformRangeTool,
} from "../src/tools/python-transform-range.ts";
import type {
  PythonTransformRangeDetails,
} from "../src/tools/tool-details.ts";
import {
  PythonBridgeUnreachableError,
  type PythonBridgeConfig,
  type PythonBridgeRequest,
  type PythonBridgeResponse,
} from "../src/tools/python-run.ts";
import { FakeExcelHost, withFakeExcel } from "./fixtures/excel-write-host.ts";

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
  assert.match(firstText(result), /Skill: python-bridge/u);
  assert.equal(result.details?.kind, "python_transform_range");
  assert.equal(result.details?.blocked, false);
  assert.equal(result.details?.error, "no_python_runtime");
  assert.equal(result.details?.skillHint, "python-bridge");
});

void test("python_transform_range falls back to Pyodide when default bridge URL is unavailable", async () => {
  let bridgeCalls = 0;
  let pyodideCalls = 0;

  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({
      sheetName: "Sheet1",
      address: "A1:B2",
      values: [[1, 2], [3, 4]],
    }),
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
      source: "default",
    }),
    callBridge: () => {
      bridgeCalls += 1;
      return Promise.reject(new PythonBridgeUnreachableError("https://localhost:3340", new TypeError("Failed to fetch")));
    },
    isPyodideAvailable: () => true,
    callPyodide: () => {
      pyodideCalls += 1;
      return Promise.resolve({
        ok: true,
        action: "run_python",
        exit_code: 0,
        result_json: "[[10,20],[30,40]]",
      });
    },
    writeOutputValues: (request) => {
      assert.deepEqual(request.values, [[10, 20], [30, 40]]);
      return Promise.resolve({
        blocked: false,
        outputAddress: "Sheet1!C1:D2",
        rowsWritten: 2,
        colsWritten: 2,
        formulaErrorCount: 0,
      });
    },
  });

  const result = await tool.execute("tc-default-fallback", {
    range: "Sheet1!A1:B2",
    code: "result = [[10,20],[30,40]]",
    output_start_cell: "C1",
  });

  assert.equal(bridgeCalls, 1);
  assert.equal(pyodideCalls, 1);
  assert.equal(result.details?.blocked, false);
  assert.equal(result.details?.bridgeUrl, undefined);
  assert.match(firstText(result), /Sheet1!C1:D2/u);
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

void test("python_transform_range writes through the checked write path and shows stdout/stderr", async () => {
  const host = new FakeExcelHost(["Py"]);
  host.sheet("Py").merge("F1:G1");
  let result_json = "";
  const tool = createPythonTransformRangeTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", source: "default" }),
    callBridge: () => Promise.resolve({
      ok: true,
      action: "run_python",
      exit_code: 0,
      stdout: "hello stdout",
      stderr: "warn",
      result_json,
    }),
  });

  await withFakeExcel(host, async () => {
    host.sheet("Py").set("A1", "source");

    // A merged destination: Excel would keep only F1, so nothing is written.
    result_json = "[[\"m1\",\"m2\"],[\"m3\",\"m4\"]]";
    const merged = await tool.execute("tc-merged", { range: "Py!A1", code: "result = ...", output_start_cell: "Py!F1" });
    assert.equal(merged.details?.blocked, true);
    assert.equal(merged.isError, true);
    assert.match(firstText(merged), /Py!F1:G1 is merged; only its top-left cell holds a value, so G1 would be dropped/u);
    assert.match(firstText(merged), /Stdout:[\s\S]*hello stdout[\s\S]*Stderr:[\s\S]*warn/u);
    assert.equal(host.sheet("Py").value("F1"), "");
    assert.equal(host.sheet("Py").value("F2"), "");

    // A ragged result with a text code Excel converts.
    result_json = "[[\"00123\",2,3],[4]]";
    const ragged = await tool.execute("tc-ragged", { range: "Py!A1", code: "result = ...", output_start_cell: "Py!R1" });
    const text = firstText(ragged);
    assert.equal(ragged.details?.blocked, false);
    assert.equal(ragged.details?.outputAddress, "Py!R1:T2");
    assert.equal(ragged.details?.paddedRowCount, 1);
    assert.equal(ragged.details?.convertedCount, 1);
    assert.match(text, /1 of 2 row\(s\) had fewer than 3 values \(result rows 2\)/u);
    assert.match(text, /Py!R1: sent "00123", stored 123/u);
    assert.match(text, /Stdout:[\s\S]*hello stdout/u);

    // Overwriting a formula with a constant is called out.
    host.sheet("Py").set("Z1", "=A1");
    result_json = "[[\"fixed\"]]";
    const overwrite = await tool.execute("tc-lost-formula", {
      range: "Py!A1", code: "result = ...", output_start_cell: "Py!Z1", allow_overwrite: true,
    });
    assert.equal(overwrite.details?.formulasReplacedCount, 1);
    assert.match(firstText(overwrite), /1 formula\(s\) on sheet "Py" replaced by values: Z1\. Z1 was `=A1`/u);
  });
});

void test("python_transform_range refuses input too large for a running bridge and says how many rows fit", async () => {
  const rows = Array.from({ length: 5000 }, () => ["x".repeat(98)]);
  const range = "Sheet1!A1:A5000";
  // Each row serializes to 102 characters plus a comma; input_json may hold 200,000 characters.
  const overhead = JSON.stringify({ range, values: [] }).length;
  const expectedFit = Math.floor((200_000 - overhead + 1) / 103);

  let bridgeCalls = 0;
  let pyodideCalls = 0;
  const makeTool = (reachable: boolean) => createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({ sheetName: "Sheet1", address: "A1:A5000", values: rows }),
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", source: "default" }),
    probeBridge: () => Promise.resolve(reachable),
    callBridge: () => {
      bridgeCalls += 1;
      return Promise.reject(new Error("should not be sent"));
    },
    isPyodideAvailable: () => true,
    callPyodide: () => {
      pyodideCalls += 1;
      return Promise.resolve({ ok: true, action: "run_python", exit_code: 0, result_json: "[[5000]]" });
    },
    writeOutputValues: () => Promise.resolve({
      blocked: false, outputAddress: "Sheet1!C1", rowsWritten: 1, colsWritten: 1, formulaErrorCount: 0,
    }),
  });

  const refused = await makeTool(true).execute("tc-big", { range, code: "result = [[len(input_data['values'])]]", output_start_cell: "C1" });
  assert.equal(refused.isError, true);
  assert.match(firstText(refused), /Input too large for the Python bridge/u);
  assert.match(firstText(refused), new RegExp(`About ${expectedFit.toLocaleString("en-US")} rows fit per call`, "u"));
  assert.match(firstText(refused), new RegExp(`Sheet1!A1:A${expectedFit}\\b`, "u"));
  assert.equal(bridgeCalls, 0);
  assert.equal(pyodideCalls, 0);

  // No bridge answering at the default URL: Pyodide has no such limit and runs it.
  const viaPyodide = await makeTool(false).execute("tc-big-pyodide", { range, code: "result = [[1]]", output_start_cell: "C1" });
  assert.equal(viaPyodide.details?.blocked, false);
  assert.equal(bridgeCalls, 0);
  assert.equal(pyodideCalls, 1);
});

void test("python_transform_range explains the bridge output limit with a row estimate", async () => {
  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({ sheetName: "Sheet1", address: "A1:A2", values: [[1], [2]] }),
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", source: "default" }),
    callBridge: () => Promise.reject(new Error("Python bridge request failed (413): Process output exceeded 262144 bytes.")),
    isPyodideAvailable: () => true,
    callPyodide: () => Promise.reject(new Error("should not fall back")),
  });

  const result = await tool.execute("tc-big-out", { range: "Sheet1!A1:A2", code: "result = [[i] for i in range(50000)]" });
  // The input "[[1],[2]]" is 9 bytes for 2 rows, about 5 bytes a row: 262,144 / 5 rows.
  assert.match(firstText(result), /fit about 52,428 per call/u);
  assert.equal(result.details?.blocked, false);
});
