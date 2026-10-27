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
      text: [["1", "2"], ["3", "4"]],
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
      text: [["1", "2"], ["3", "4"]],
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
  const host = new FakeExcelHost();
  const sheet = host.sheet("Sheet1");
  sheet.set("A1", 1);
  sheet.set("B1", 2);
  sheet.set("A2", 3);
  sheet.set("B2", 4);
  const tool = createPythonTransformRangeTool({
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
      token: "secret-token",
    }),
    callBridge: (request, config) => {
      assert.equal(request.input_json, '{"range":"Sheet1!A1:B2","values":[[1,2],[3,4]],"text":[["1","2"],["3","4"]]}');
      assert.equal(config.token, "secret-token");
      return Promise.resolve({
        ok: true,
        action: "run_python",
        exit_code: 0,
        result_json: "{\"values\":[[10,20],[30,40]]}",
      });
    },
  });

  const result = await withFakeExcel(host, () => tool.execute("tc-success", {
    range: "Sheet1!A1:B2",
    code: "result = {'values': [[10,20],[30,40]]}",
    output_start_cell: "D1",
  }));

  const text = firstText(result);
  assert.match(text, /Transformed/u);
  assert.match(text, /Sheet1!D1:E2/u);
  assert.equal(result.details?.blocked, false);
  assert.equal(result.details?.rowsWritten, 2);
  assert.equal(result.details?.colsWritten, 2);
  assert.deepEqual(["D1", "E1", "D2", "E2"].map((address) => sheet.value(address)), [10, 20, 30, 40]);
});

void test("python_transform_range passes raw dates and numbers alongside their distinct displayed text", async () => {
  const host = new FakeExcelHost();
  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({
      sheetName: "Sheet1", address: "A1:B1",
      values: [[45292, 0.125]], text: [["2024-01-01", "12.50%"]],
    }),
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340" }),
    callBridge: (request) => {
      assert.equal(request.input_json, '{"range":"Sheet1!A1:B1","values":[[45292,0.125]],"text":[["2024-01-01","12.50%"]]}');
      return Promise.resolve({ ok: true, action: "run_python", exit_code: 0, result_json: '[["2024-01-01",0.125]]' });
    },
  });
  const result = await withFakeExcel(host, () => tool.execute("tc-displayed", {
    range: "Sheet1!A1:B1", output_start_cell: "D1",
    code: "result = [[input_data['text'][0][0], input_data['values'][0][1]]]",
  }));
  assert.equal(result.isError, undefined);
  assert.deepEqual([host.sheet("Sheet1").value("D1"), host.sheet("Sheet1").value("E1")], ["2024-01-01", 0.125]);
});

void test("python_transform_range surfaces overwrite blocking from destination write", async () => {
  const host = new FakeExcelHost();
  host.sheet("Sheet1").set("A1", 1);
  host.sheet("Sheet1").set("A2", 2);
  const tool = createPythonTransformRangeTool({
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
    }),
    callBridge: () => Promise.resolve({
      ok: true,
      action: "run_python",
      result_json: "[[100],[200]]",
    }),
  });

  const result = await withFakeExcel(host, () => tool.execute("tc-blocked", {
    range: "Sheet1!A1:A2",
    code: "result = [[100],[200]]",
  }));

  const text = firstText(result);
  assert.match(text, /Transform blocked/u);
  assert.match(text, /If the user asked to change these cells, retry with `allow_overwrite: true`\. Otherwise show the user the existing data and ask before overwriting\./u);
  assert.equal(result.details?.blocked, true);
  assert.equal(result.details?.existingCount, 2);
  assert.equal(result.details?.outputAddress, "Sheet1!A1:A2");
  assert.deepEqual(["A1", "A2"].map((address) => host.sheet("Sheet1").value(address)), [1, 2]);
});

void test("python_transform_range requires result_json from python execution", async () => {
  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({
      sheetName: "Sheet1",
      address: "A1:B2",
      values: [[1, 2], [3, 4]],
      text: [["1", "2"], ["3", "4"]],
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
  // Each of the raw/displayed rows is 103 characters with its separator.
  const overhead = JSON.stringify({ range, values: [], text: [] }).length;
  const expectedFit = Math.floor((200_000 - overhead + 2) / 206);

  let bridgeCalls = 0;
  let pyodideCalls = 0;
  const makeTool = (reachable: boolean) => createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({ sheetName: "Sheet1", address: "A1:A5000", values: rows, text: rows }),
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
    readInputRange: () => Promise.resolve({ sheetName: "Sheet1", address: "A1:A2", values: [[1], [2]], text: [["1"], ["2"]] }),
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

void test("python_transform_range answers Stop and unconfirmed chunks like write_cells", async () => {
  const host = new FakeExcelHost(["Py"]);
  // Two rows over 1M characters each are written in two chunks.
  const big = "y".repeat(1_100_000);
  let result_json = "";
  let controller = new AbortController();
  let abortDuringPython = false;
  const tool = createPythonTransformRangeTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", source: "default" }),
    callBridge: () => {
      if (abortDuringPython) controller.abort();
      return Promise.resolve({ ok: true, action: "run_python", exit_code: 0, result_json });
    },
  });
  const sync = host.context.sync;

  await withFakeExcel(host, async () => {
    host.sheet("Py").set("A1", "source");

    // Stop while Python ran: nothing is written.
    abortDuringPython = true;
    result_json = "[[1,2],[3,4]]";
    const before = await tool.execute("tc-stop-before", { range: "Py!A1", code: "result = ...", output_start_cell: "Py!R1" }, controller.signal);
    assert.equal(firstText(before), "⏹ **Stopped before writing** — nothing changed in Py!R1:S2.");
    assert.equal(before.details?.stopped, true);
    assert.equal(host.sheet("Py").value("R1"), "");

    // Stop after the first chunk landed: that chunk stands, the rest is not written.
    abortDuringPython = false;
    controller = new AbortController();
    result_json = JSON.stringify([[big], [big]]);
    let writes = 0;
    host.context.sync = () => {
      const pending = (Reflect.get(host, "queued") as unknown[]).length;
      return sync().then(() => {
        if (pending > 0 && ++writes === 1) controller.abort();
      });
    };
    const after = await tool.execute("tc-stop-after", { range: "Py!A1", code: "result = ...", output_start_cell: "Py!U1" }, controller.signal);
    assert.match(firstText(after), /^⏹ \*\*Stopped after writing Py!U1\*\* \(row 1 of Py!U1:U2\) — the change stands\.\nThe remaining 1 row\(s\) were not written\./u);
    assert.equal(after.details?.stopped, true);
    assert.equal(host.sheet("Py").value("U1"), big);
    assert.equal(host.sheet("Py").value("U2"), "");

    // Excel never answered the second chunk: it may or may not have landed.
    host.context.sync = sync;
    controller = new AbortController();
    host.failWrite = { at: 3, error: new Error("Excel did not respond") };
    const unknown = await tool.execute("tc-unconfirmed", { range: "Py!A1", code: "result = ...", output_start_cell: "Py!W1" }, controller.signal);
    assert.match(firstText(unknown), /Partly written\*\* — Py!W1 \(row 1 of Py!W1:W2\) was written and stands\.\nrow 2 \(Py!W2\) may or may not have been written: Excel did not confirm that step\./u);
    assert.equal(unknown.isError, true);
  });
});

void test("python_transform_range: after an Excel.run timeout mid-write, the abandoned run writes nothing more than the result reports", async () => {
  const host = new FakeExcelHost(["Big"]);
  // ~300 characters a row: 15,000 rows are written in several chunks.
  const rows = Array.from({ length: 15_000 }, (_, index) => [`${String(index).padStart(6, "0")}${"x".repeat(294)}`]);
  const tool = createPythonTransformRangeTool({
    readInputRange: () => Promise.resolve({ sheetName: "Big", address: "Z1", values: [[1]], text: [["1"]] }),
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340", source: "default" }),
    callBridge: () => Promise.resolve({ ok: true, action: "run_python", exit_code: 0, result_json: JSON.stringify(rows) }),
  });

  await withFakeExcel(host, async () => {
    // As #792's excelRun does: race the whole Excel.run against a timeout. The run itself goes on.
    const fakeRun = Reflect.get(globalThis, "Excel") as { run: (callback: (context: unknown) => Promise<unknown>) => Promise<unknown> };
    Reflect.set(globalThis, "Excel", {
      run: (callback: (context: unknown) => Promise<unknown>) => Promise.race([
        fakeRun.run(callback),
        new Promise((_, reject) => { setTimeout(() => {
          reject(Object.assign(new Error("Excel did not answer within 60 s."), { code: "excel_timeout", outcomeUnknown: true }));
        }, 30); }),
      ]),
    });
    // The third sync (the second chunk) is held past the timeout, then lands.
    host.delaySync = { at: 3, ms: 80 };
    const result = await tool.execute("tc-timeout", { range: "Big!Z1", code: "result = ...", output_start_cell: "Big!A1" });
    const reported = firstText(result);
    await new Promise((resolve) => { setTimeout(resolve, 200); });
    let landed = 0;
    while (landed < rows.length && host.sheet("Big").value(`A${landed + 1}`) !== "") landed += 1;

    const written = /Partly written\*\* — Big!A1:A(\d+)/u.exec(reported);
    const uncertain = /rows (\d+)–(\d+) \(Big!A\d+:A\d+\) may or may not have been written/u.exec(reported);
    assert.ok(written?.[1] && uncertain?.[2], reported);
    // Everything that landed, even after the result was returned, is within what the result allowed for.
    assert.ok(landed <= Number(uncertain[2]), `landed ${landed} rows; result: ${reported}`);
    assert.ok(landed >= Number(written[1]));

    // A timeout while inspecting the target: nothing was sent, and nothing lands later.
    host.delaySync = { at: host.syncCount + 1, ms: 80 };
    const early = await tool.execute("tc-timeout-early", { range: "Big!Z1", code: "result = ...", output_start_cell: "Big!C1" });
    await new Promise((resolve) => { setTimeout(resolve, 200); });
    assert.match(firstText(early), /Nothing was written to Big!C1:C15000\. Excel did not answer before the change was sent/u);
    assert.doesNotMatch(firstText(early), /unknown/u);
    assert.equal(host.sheet("Big").value("C1"), "");
  });
});
