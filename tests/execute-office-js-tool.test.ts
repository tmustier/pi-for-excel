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
  assert.match(text, /\{"ok":true,"sheet":"Sheet1","changedCells":4\}/u);
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

void test("execute_office_js on a protected sheet is an error that names the protection", async () => {
  const { FakeExcelHost, OfficeError, withFakeExcel } = await import("./fixtures/excel-write-host.ts");
  const host = new FakeExcelHost(["Sheet1", "Locked"]);
  host.sheet("Locked").protect();
  await withFakeExcel(host, async () => {
    const tool = createExecuteOfficeJsTool({
      runCode: () => Promise.reject(new OfficeError("AccessDenied", "You cannot perform the requested operation.", "Range.values")),
    });
    const result = await tool.execute("ojs-protected", { explanation: "Write to Locked", code: "return 1;" });

    assert.equal(result.isError, true);
    assert.match(firstText(result), /Protected sheet\(s\): "Locked"\. Excel refused the change.*Review → Unprotect Sheet/u);
    assert.doesNotMatch(firstText(result), /You cannot perform the requested operation/u);
  });
});

void test("execute_office_js after a timeout: the code's later Excel calls never reach Excel, and the result says so", async () => {
  const { FakeExcelHost, withFakeExcel } = await import("./fixtures/excel-write-host.ts");
  const { runMutation } = await import("../src/tools/mutation/run-mutation.ts");
  const host = new FakeExcelHost();
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
    host.delaySync = { at: 1, ms: 80 }; // the code's first sync is answered after the timeout
    const tool = createExecuteOfficeJsTool({
      // User code that writes in two steps, run the way the default runner runs it.
      runCode: () => runMutation(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Sheet1");
        sheet.getRange("A1").values = [["first"]];
        await context.sync();
        sheet.getRange("A2").values = [["second"]];
        await context.sync();
        return 1;
      }),
    });
    const result = await tool.execute("ojs-timeout", { explanation: "two steps", code: "return 1;" });
    await new Promise((resolve) => { setTimeout(resolve, 200); });

    assert.equal(result.isError, true);
    assert.match(firstText(result), /Pi cannot stop Office\.js code once it runs\. It blocked this code's later Excel calls/u);
    assert.equal(host.sheet("Sheet1").value("A1"), "first"); // the unanswered step still landed
    assert.equal(host.sheet("Sheet1").value("A2"), ""); // nothing after it reached Excel
  });
});
