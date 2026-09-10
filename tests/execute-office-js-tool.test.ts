import assert from "node:assert/strict";
import { test } from "node:test";

import { createExecuteOfficeJsTool } from "../src/tools/execute-office-js.ts";

function firstText(result: { content: Array<{ type: string; text: string }> }): string {
  const block = result.content[0];
  if (!block || block.type !== "text") {
    throw new Error("Expected first content block to be text.");
  }

  return block.text;
}

void test("execute_office_js runs code and serializes result", async () => {
  const tool = createExecuteOfficeJsTool({
    runCode: () => Promise.resolve({
      ok: true,
      sheet: "Sheet1",
      changedCells: 4,
    }),
  });

  const result = await tool.execute("tool-call-1", {
    explanation: "Recalculate dashboard totals",
    code: "return { ok: true };",
  });

  const text = firstText(result);
  assert.match(text, /Executed Office\.js: Recalculate dashboard totals/u);
  assert.match(text, /```json/u);
  assert.match(text, /"ok": true/u);
});

void test("execute_office_js blocks nested Excel.run usage", async () => {
  let runCalled = false;

  const tool = createExecuteOfficeJsTool({
    runCode: () => {
      runCalled = true;
      return Promise.resolve(null);
    },
  });

  const result = await tool.execute("tool-call-2", {
    explanation: "Update workbook",
    code: "return Excel.run(async (context) => { await context.sync(); });",
  });

  const text = firstText(result);
  assert.match(text, /Do not call Excel\.run\(\)/u);
  assert.equal(runCalled, false);
});

void test("execute_office_js reports non-serializable result payloads", async () => {
  const circular: { self?: unknown } = {};
  circular.self = circular;

  const tool = createExecuteOfficeJsTool({
    runCode: () => Promise.resolve(circular),
  });

  const result = await tool.execute("tool-call-3", {
    explanation: "Inspect workbook state",
    code: "return {};",
  });

  const text = firstText(result);
  assert.match(text, /Result is not JSON-serializable/u);
});
