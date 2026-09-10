import assert from "node:assert/strict";
import { test } from "node:test";

import { createExecuteOfficeJsTool } from "../src/tools/execute-office-js.ts";
import { applyExperimentalToolGates } from "../src/tools/experimental-tool-gates.ts";

const BLOCKED_VECTORS = [
  { name: "network egress", code: "return fetch(url);", identifiers: ["fetch"] },
  { name: "storage", code: "return localStorage.getItem('key');", identifiers: ["localStorage"] },
  { name: "window handle", code: "return window.location.href;", identifiers: ["window"] },
  { name: "computed global handle", code: "return globalThis['fe' + 'tch'];", identifiers: ["globalThis"] },
  { name: "document handle", code: "return document.cookie;", identifiers: ["document"] },
  { name: "navigator handle", code: "navigator.sendBeacon(url, data);", identifiers: ["navigator"] },
  { name: "eval", code: "eval(payload);", identifiers: ["eval"] },
  { name: "Function constructor", code: "new Function('return 1')();", identifiers: ["Function"] },
  { name: "dynamic import", code: "const mod = await import(url);", identifiers: ["import"] },
  { name: "constructor realm escape", code: "return ({}).constructor.constructor('return 1')();", identifiers: ["constructor"] },
  { name: "Office host surface", code: "return Office.context.document.url;", identifiers: ["Office"] },
  { name: "identifiers in comments and strings", code: "// eval later\nreturn 'fetch';", identifiers: ["fetch", "eval"] },
  { name: "four-digit unicode escape", code: "return \\u0066etch(url);", identifiers: ["fetch"] },
  { name: "braced unicode escape", code: "return \\u{66}etch(url);", identifiers: ["fetch"] },
  { name: "embedded unicode escape", code: "return fe\\u0074ch(url);", identifiers: ["fetch"] },
  { name: "nested unicode escape", code: "\\u005cu0066etch(url);", identifiers: ["fetch"] },
] as const;

void test("execute_office_js blocks every ambient-authority vector in Auto mode", async () => {
  for (const vector of BLOCKED_VECTORS) {
    const [tool] = await applyExperimentalToolGates([
      createExecuteOfficeJsTool({ runCode: () => Promise.resolve({ executed: true }) }),
    ], {
      getExecutionMode: () => Promise.resolve("yolo" as const),
      requestOfficeJsExecuteApproval: (request) => {
        assert.deepEqual(request.riskIdentifiers, vector.identifiers, vector.name);
        return Promise.resolve(false);
      },
    });
    if (!tool) throw new Error("Expected execute_office_js tool.");

    await assert.rejects(
      () => tool.execute(`risk-${vector.name}`, { explanation: vector.name, code: vector.code }),
      /cancelled by user/iu,
      vector.name,
    );
  }
});

const ALLOWED_CODE = [
  {
    name: "pure Excel API",
    code: "const sheet = context.workbook.worksheets.getActiveWorksheet(); return Excel.ChartType.columnClustered;",
  },
  { name: "Excel object member names", code: "chart.top = 50; return comment.parent;" },
  { name: "longer look-alike identifiers", code: "const refetchCount = 1; const windowless = 2; return refetchCount;" },
  { name: "plain text unicode escapes", code: "range.values = [['caf\\u00e9', '\\u00fcber']]; return true;" },
] as const;

void test("execute_office_js allows Excel code and risk-name look-alikes in Auto mode", async () => {
  for (const example of ALLOWED_CODE) {
    const [tool] = await applyExperimentalToolGates([
      createExecuteOfficeJsTool({ runCode: () => Promise.resolve({ executed: example.name }) }),
    ], {
      getExecutionMode: () => Promise.resolve("yolo" as const),
      requestOfficeJsExecuteApproval: () => Promise.reject(new Error(`unexpected approval for ${example.name}`)),
    });
    if (!tool) throw new Error("Expected execute_office_js tool.");

    const result = await tool.execute(`allowed-${example.name}`, {
      explanation: example.name,
      code: example.code,
    });
    const block = result.content[0];
    if (!block || block.type !== "text") throw new Error("Expected text tool result.");
    assert.match(block.text, new RegExp(`Executed Office\\.js: ${example.name}`, "u"));
  }
});
