import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { createFillFormulaTool } from "../src/tools/fill-formula.ts";
import { createWriteCellsTool } from "../src/tools/write-cells.ts";
import { findErrors } from "../src/utils/format.ts";
import { colName, FakeExcelHost, OfficeError, withFakeExcel } from "./fixtures/excel-write-host.ts";

function text<T>(result: AgentToolResult<T>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("Expected a text result.");
  return first.text;
}

const noBackups = {
  appendAuditEntry: (): Promise<void> => Promise.resolve(),
  appendRecoverySnapshot: (): Promise<null> => Promise.resolve(null),
};

void test("write_cells reports value and formula changes with a bounded sample", async () => {
  const host = new FakeExcelHost();
  await withFakeExcel(host, async () => {
    const values = [["=2+2", 20, ...Array.from({ length: 11 }, (_, index) => index + 30)]];
    const result = await createWriteCellsTool(noBackups).execute("write-1", {
      start_cell: "Sheet1!A1",
      values,
    });

    assert.equal(result.details.changes?.changedCount, 13);
    assert.equal(result.details.changes?.truncated, true);
    assert.equal(result.details.changes?.sample.length, 12);
    assert.deepEqual(result.details.changes?.sample[1], {
      address: "Sheet1!B1",
      beforeValue: "",
      afterValue: "20",
    });
    assert.equal(result.details.changes?.sample.at(-1)?.address, `Sheet1!${colName(11)}1`);
  });
});

void test("write_cells accepts a quoted sheet name containing an exclamation mark", async () => {
  const host = new FakeExcelHost(["Q1!Ops"]);
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("write-quoted-bang", {
      start_cell: "'Q1!Ops'!A1",
      values: [["ready"]],
    });

    assert.doesNotMatch(text(result), /^Error/u);
    assert.equal(result.isError, undefined);
    assert.equal(host.sheet("Q1!Ops").value("A1"), "ready");
  });
});

void test("write_cells blocks values that a merged area would drop, and allows its top-left cell", async () => {
  const host = new FakeExcelHost(["Merged"]);
  const sheet = host.sheet("Merged");
  sheet.set("A1", "Title");
  sheet.merge("A1:D1");
  sheet.merge("B4:C6");

  await withFakeExcel(host, async () => {
    const tool = createWriteCellsTool(noBackups);

    const inside = await tool.execute("merge-inside", { start_cell: "Merged!B1", values: [["hello"]] });
    assert.equal(inside.isError, true);
    assert.equal(inside.details.blocked, true);
    assert.match(text(inside), /Merged!A1:D1 is merged/u);
    assert.match(text(inside), /B1 would be dropped/u);

    // The 2×2 block starts on the merge's top-left cell; only the other three are lost.
    const block = await tool.execute("merge-block", { start_cell: "Merged!B4", values: [["a", "b"], ["c", "d"]] });
    assert.equal(block.isError, true);
    assert.match(text(block), /Merged!B4:C6 is merged; only its top-left cell holds a value, so C4, B5, C5 would be dropped/u);
    assert.equal(sheet.value("B4"), "", "nothing is written when the write is blocked");

    const anchor = await tool.execute("merge-anchor", {
      start_cell: "Merged!A1",
      values: [["New title", "", "", ""]],
      allow_overwrite: true,
    });
    assert.equal(anchor.isError, undefined);
    assert.equal(sheet.value("A1"), "New title");
  });
});

void test("write_cells refuses to write into a spilled range and names the anchor", async () => {
  const host = new FakeExcelHost(["Spill"]);
  const sheet = host.sheet("Spill");
  sheet.spill("A1", "=SEQUENCE(5)", [1, 2, 3, 4, 5]);

  await withFakeExcel(host, async () => {
    const tool = createWriteCellsTool(noBackups);
    const child = await tool.execute("spill-child", { start_cell: "Spill!A3", values: [["x"]], allow_overwrite: true });

    assert.equal(child.isError, true);
    assert.match(text(child), /A3 is spilled from Spill!A1 \(=SEQUENCE\(5\)\)/u);
    assert.equal(sheet.value("A1"), 1, "the anchor is not broken");

    // Replacing the anchor itself is an ordinary overwrite.
    const anchor = await tool.execute("spill-anchor", { start_cell: "Spill!A1", values: [["=A9"]], allow_overwrite: true });
    assert.equal(anchor.isError, undefined);
  });
});

void test("write_cells reports text Excel converted, but not literal text or plain numbers", async () => {
  const host = new FakeExcelHost();
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("coerce", {
      start_cell: "A1",
      values: [["00123", "'007", "42", "label"]],
    });

    assert.equal(result.isError, undefined);
    assert.equal(result.details.convertedCount, 1);
    assert.match(text(result), /Sheet1!A1: sent "00123", stored 123/u);
    assert.doesNotMatch(text(result), /Sheet1!(B1|C1|D1): sent/u);
    assert.match(text(result), /leading apostrophe/u);
    assert.equal(host.sheet("Sheet1").value("B1"), "007");
  });
});

void test("write_cells reports values that violate data validation", async () => {
  const host = new FakeExcelHost(["Valid"]);
  host.sheet("Valid").listValidation("A2:A5", ["Yes", "No"]);

  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("validation", {
      start_cell: "Valid!A2",
      values: [["Maybe"], ["Yes"]],
    });

    assert.equal(result.details.invalidCount, 1);
    assert.match(text(result), /1 value\(s\) violate the cells' data validation \(list: Yes,No\):\*\* Valid!A2\./u);
  });
});

void test("write_cells calculates written formulas under manual calculation and says so", async () => {
  const host = new FakeExcelHost(["Calc"]);
  host.sheet("Calc").set("A1", 1);
  host.sheet("Calc").set("A2", 2);
  host.calculationMode = "Manual";

  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("manual", {
      start_cell: "Calc!A3",
      values: [["=A1+A2"]],
    });

    assert.match(text(result), /\| 3 \|/u);
    assert.match(text(result), /Calculation is set to Manual/u);
    assert.equal(result.details.calculationMode, "Manual");
  });
});

void test("under manual calculation, changing an input says its dependents are stale", async () => {
  const host = new FakeExcelHost(["Calc"]);
  host.sheet("Calc").set("A1", 1);
  host.sheet("Calc").set("B1", "=A1");
  host.calculationMode = "Manual";

  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("manual-input", {
      start_cell: "Calc!A1",
      values: [[10]],
      allow_overwrite: true,
    });

    assert.match(text(result), /Calculation is set to Manual\. Formulas that depend on them still show old results/u);
    assert.equal(result.details.calculationMode, "Manual");
  });
});

void test("write_cells refuses locked cells on a protected sheet and writes unlocked ones", async () => {
  const host = new FakeExcelHost(["Protected"]);
  host.sheet("Protected").protect(["B2"]);

  await withFakeExcel(host, async () => {
    const tool = createWriteCellsTool(noBackups);
    const locked = await tool.execute("locked", { start_cell: "Protected!C1", values: [["x"]] });
    assert.equal(locked.isError, true);
    assert.match(text(locked), /Sheet "Protected" is protected/u);

    const unlocked = await tool.execute("unlocked", { start_cell: "Protected!B2", values: [["ok"]] });
    assert.equal(unlocked.isError, undefined);
    assert.equal(host.sheet("Protected").value("B2"), "ok");
  });
});

void test("write_cells names the missing sheet and lists the sheets that exist", async () => {
  const host = new FakeExcelHost(["Data", "Summary"]);
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("missing", { start_cell: "Nope!A1", values: [["x"]] });

    assert.equal(result.isError, true);
    assert.match(text(result), /nothing was written/u);
    assert.match(text(result), /Sheet "Nope" does not exist\. Sheets in this workbook: "Data", "Summary"\./u);
  });
});

void test("a write too large to read back in one call is applied, verified in part, and says it has no backup", async () => {
  const host = new FakeExcelHost(["Big"]);
  const rows = Array.from({ length: 60_000 }, (_, index) => [`r${index}`]);

  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool({
      appendAuditEntry: () => Promise.resolve(),
      appendRecoverySnapshot: () => Promise.reject(new Error("a large write must not try to back up")),
    }).execute("large", { start_cell: "Big!A1", values: rows });

    assert.equal(result.isError, undefined, text(result));
    assert.equal(result.details.verification, "sampled");
    assert.equal(host.sheet("Big").value("A60000"), "r59999");
    assert.match(text(result), /Verification is partial/u);
    assert.equal(result.details.recovery?.status, "not_available");
    assert.match(text(result), /No backup: this change covers 60,000 cells/u);
  });
});

void test("a large overwrite of existing data is refused because it could not be backed up", async () => {
  const host = new FakeExcelHost(["Big"]);
  host.sheet("Big").set("A30000", "keep me");
  const rows = Array.from({ length: 60_000 }, (_, index) => [`r${index}`]);

  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("large-overwrite", {
      start_cell: "Big!A1",
      values: rows,
      allow_overwrite: true,
    });

    assert.equal(result.isError, true);
    assert.match(text(result), /Write refused/u);
    assert.equal(host.sheet("Big").value("A30000"), "keep me");
    assert.equal(host.sheet("Big").value("A1"), "");
  });
});

void test("fill_formula over a range too large to read back reports the fill instead of a payload error", async () => {
  const host = new FakeExcelHost(["Big"]);
  await withFakeExcel(host, async () => {
    const result = await createFillFormulaTool().execute("fill-large", {
      range: "Big!F2:F60001",
      formula: "=C2*E2",
    });

    assert.equal(result.isError, undefined, text(result));
    assert.equal(result.details.verification, "sampled");
    assert.equal(host.sheet("Big").formula("F60001"), "=C60001*E60001");
    assert.match(text(result), /bottom-right `=C60001\*E60001`/u);
    assert.match(text(result), /No backup: this change covers 60,000 cells.*clearing it undoes the change/u);
    const recovery = result.details.recovery;
    assert.equal(recovery?.status, "not_available");
    assert.match(recovery?.status === "not_available" ? recovery.reason : "", /Too large to back up: 60,000 cells exceed the 20,000-cell backup limit/u);
  });
});

void test("fill_formula next to a table says Excel extended the table and named the new column", async () => {
  const host = new FakeExcelHost(["Sales"]);
  const sales = host.sheet("Sales");
  sales.set("A1", "Qty");
  sales.set("B1", "Price");
  for (let row = 2; row <= 4; row += 1) {
    sales.set(`A${row}`, row);
    sales.set(`B${row}`, 10);
  }
  sales.table("SalesTable", "A1:B4");

  await withFakeExcel(host, async () => {
    const result = await createFillFormulaTool().execute("fill-table", { range: "Sales!C2:C4", formula: "=A2*B2" });

    assert.equal(result.isError, undefined, text(result));
    assert.equal(sales.value("C4"), 40);
    assert.match(text(result), /Excel added the written cells to table \*\*SalesTable\*\*, which now covers A1:C4\. It named the new column\(s\) "Column1"/u);
  });
});

void test("write_cells over formulas says which cells lost a formula, not for formulas rewritten or cleared", async () => {
  const host = new FakeExcelHost(["Model"]);
  const model = host.sheet("Model");
  model.set("C5", 100);
  model.set("C6", 200);
  model.set("D5", "=C5*C5");
  model.set("D6", "=C6+C5");
  model.set("D7", "=C5");
  model.set("D8", 7);
  model.set("D9", "=C5");

  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("hardcode", {
      start_cell: "Model!D5",
      values: [[110], [220], ["=C6"], [8], [""]],
      allow_overwrite: true,
    });

    assert.equal(result.isError, undefined, text(result));
    assert.match(text(result), /⚠️ 2 formula\(s\) on sheet "Model" replaced by values: D5, D6\. D5 was `=C5\*C5`, D6 was `=C6\+C5`\./u);
    assert.equal(result.details.formulasReplacedCount, 2);
  });
});

void test("findErrors reports Excel errors but not text that starts with #", () => {
  assert.deepEqual(findErrors([["#1 priority", "#DIV/0!", "#N/A", "#hashtag!"]], "B2"), [
    { address: "C2", error: "#DIV/0!" },
    { address: "D2", error: "#N/A" },
  ]);
});

void test("a chunked write interrupted by cell-edit mode reports exactly which rows landed", async () => {
  const host = new FakeExcelHost(["Big"]);
  // ~4.5 MB of text: written in several chunks, each synced on its own.
  const rows = Array.from({ length: 15_000 }, (_, index) => [`${String(index).padStart(6, "0")}${"x".repeat(294)}`]);
  host.failWrite = { at: 2, error: new OfficeError("InvalidOperationInCellEditMode", "Excel is in cell-editing mode.") };
  const audits: Array<{ blocked: boolean; outputAddress?: string }> = [];

  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool({
      appendAuditEntry: (entry) => { audits.push(entry); return Promise.resolve(); },
      appendRecoverySnapshot: () => Promise.resolve(null),
    }).execute("edit-mode-mid-write", { start_cell: "Big!A1", values: rows });

    const landed = host.sheet("Big");
    const lastWritten = [...Array(15_000).keys()].filter((index) => landed.value(`A${index + 1}`) !== "").length;
    assert.ok(lastWritten > 0 && lastWritten < 15_000, "the first chunk landed and the rest did not");
    assert.equal(result.isError, true);
    assert.match(text(result), new RegExp(`Partly written\\*\\* — Big!A1:A${lastWritten} \\(rows 1–${lastWritten} of Big!A1:A15000\\) was written and stands`, "u"));
    assert.match(text(result), /The remaining \d+ row\(s\) were not written\./u);
    assert.doesNotMatch(text(result), /may or may not/u, "Excel refused the chunk, so its outcome is known");
    assert.match(text(result), /Excel is in cell-edit mode/u);
    assert.match(text(result), /Nothing was rolled back/u);
    assert.equal(audits.at(-1)?.blocked, false, "the rows that landed are audited as a change");
  });
});

function bigRows(count: number): string[][] {
  // ~300 characters a row: 15,000 rows are written in several chunks.
  return Array.from({ length: count }, (_, index) => [`${String(index).padStart(6, "0")}${"x".repeat(294)}`]);
}

function writtenRowCount(host: FakeExcelHost, sheet: string, rows: number): number {
  return [...Array(rows).keys()].filter((index) => host.sheet(sheet).value(`A${index + 1}`) !== "").length;
}

void test("a write of long text under the cell limit is verified in part, not read back in one oversized call", async () => {
  const host = new FakeExcelHost(["Big"]);
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("long-text", { start_cell: "Big!A1", values: bigRows(15_000) });
    assert.equal(result.isError, undefined, text(result));
    assert.equal(result.details.verification, "sampled");
    assert.match(text(result), /Verification is partial/u);
  });
});

void test("Stop before the first chunk writes nothing and says so", async () => {
  const host = new FakeExcelHost(["Big"]);
  const stop = new AbortController();
  stop.abort();
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("stop-before", { start_cell: "Big!A1", values: [["x"]] }, stop.signal);
    assert.equal(result.isError, true);
    assert.equal(result.details.stopped, true);
    assert.equal(result.details.writtenCount, 0);
    assert.match(text(result), /Stopped before writing\*\* — nothing changed in Big!A1\./u);
    assert.equal(host.sheet("Big").value("A1"), "");
  });
});

void test("Stop between chunks keeps the rows written and names them", async () => {
  const host = new FakeExcelHost(["Big"]);
  const stop = new AbortController();
  host.afterWriteSync = (count) => { if (count === 1) stop.abort(); };
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("stop-mid", { start_cell: "Big!A1", values: bigRows(15_000) }, stop.signal);
    const landed = writtenRowCount(host, "Big", 15_000);
    assert.ok(landed > 0 && landed < 15_000);
    assert.equal(result.isError, true);
    assert.equal(result.details.stopped, true);
    assert.equal(result.details.writtenCount, landed);
    assert.match(text(result), new RegExp(`Stopped after writing Big!A1:A${landed}\\*\\* \\(rows 1–${landed} of Big!A1:A15000\\) — the change stands\\.`, "u"));
    assert.match(text(result), new RegExp(`The remaining ${15_000 - landed} row\\(s\\) were not written\\.`, "u"));
  });
});

void test("Stop after the last chunk says the whole write stands", async () => {
  const host = new FakeExcelHost();
  const stop = new AbortController();
  host.afterWriteSync = () => stop.abort();
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("stop-after", { start_cell: "A1", values: [["a"], ["b"]] }, stop.signal);
    assert.equal(result.isError, undefined);
    assert.match(text(result), /Stopped after writing Sheet1!A1:A2\*\* \(2×1\) — the change stands/u);
    assert.equal(result.details.writtenCount, 2);
    assert.equal(host.sheet("Sheet1").value("A2"), "b");
  });
});

void test("a chunk Excel never confirmed is reported as unknown, not as unwritten", async () => {
  const host = new FakeExcelHost(["Big"]);
  // Shape #792's excelRun timeout is expected to throw: no Office code, outcome unknown.
  host.failWrite = { at: 2, error: Object.assign(new Error("Excel did not answer in 60 s"), { code: "excel_timeout", outcomeUnknown: true }) };
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("timeout-mid", { start_cell: "Big!A1", values: bigRows(15_000) });
    const landed = writtenRowCount(host, "Big", 15_000);
    assert.equal(result.isError, true);
    assert.match(text(result), new RegExp(`rows ${landed + 1}–\\d+ \\(Big!A${landed + 1}:A\\d+\\) may or may not have been written`, "u"));
    assert.match(text(result), /Excel did not answer/u);
    assert.doesNotMatch(text(result), /nothing was written/u);
  });
});

void test("a timeout while reading the target says nothing was written, without calling the outcome unknown", async () => {
  const host = new FakeExcelHost();
  host.failSync = { at: 1, error: Object.assign(new Error("Excel did not answer within 60 s."), { code: "excel_timeout", outcomeUnknown: true }) };
  host.failSync.error.message = "Excel did not answer within 60 s. The outcome is unknown: the request may or may not have been applied.";
  await withFakeExcel(host, async () => {
    const result = await createWriteCellsTool(noBackups).execute("timeout-inspect", { start_cell: "Sheet1!H7", values: [["x"]] });

    assert.equal(result.isError, true);
    assert.match(text(result), /nothing was written to Sheet1!H7\. Excel did not answer before the change was sent, so nothing was written\./u);
    assert.doesNotMatch(text(result), /unknown/u);
    assert.equal(host.sheet("Sheet1").value("H7"), "");
  });
});

void test("after an Excel.run timeout mid-write, the abandoned run writes nothing more than the result reports", async () => {
  const host = new FakeExcelHost(["Big"]);
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
    const result = await createWriteCellsTool(noBackups).execute("timeout-abandoned", { start_cell: "Big!A1", values: bigRows(15_000) });
    const reported = text(result);
    await new Promise((resolve) => { setTimeout(resolve, 200); });
    const landed = writtenRowCount(host, "Big", 15_000);

    const written = /Partly written\*\* — Big!A1:A(\d+)/u.exec(reported);
    const uncertain = /rows (\d+)–(\d+) \(Big!A\d+:A\d+\) may or may not have been written/u.exec(reported);
    assert.ok(written?.[1] && uncertain?.[2], reported);
    // Everything that landed, even after the result was returned, is within what the result allowed for.
    assert.ok(landed <= Number(uncertain[2]), `landed ${landed} rows; result: ${reported}`);
    assert.ok(landed >= Number(written[1]));
  });
});

void test("fill_formula that times out while inspecting never fills afterwards", async () => {
  const host = new FakeExcelHost();
  await withFakeExcel(host, async () => {
    const fakeRun = Reflect.get(globalThis, "Excel") as { run: (callback: (context: unknown) => Promise<unknown>) => Promise<unknown> };
    Reflect.set(globalThis, "Excel", {
      run: (callback: (context: unknown) => Promise<unknown>) => Promise.race([
        fakeRun.run(callback),
        new Promise((_, reject) => { setTimeout(() => {
          reject(Object.assign(new Error("Excel did not answer within 60 s."), { code: "excel_timeout", outcomeUnknown: true }));
        }, 30); }),
      ]),
    });
    host.delaySync = { at: 1, ms: 80 }; // the first read is held past the timeout
    const result = await createFillFormulaTool().execute("fill-timeout", { range: "Sheet1!B1:B3", formula: "=A1" });
    await new Promise((resolve) => { setTimeout(resolve, 200); });

    assert.match(text(result), /nothing was written to Sheet1!B1:B3/u);
    assert.equal(host.sheet("Sheet1").formula("B3"), "");
  });
});

void test("fill_formula stopped after the fill says the fill stands", async () => {
  const host = new FakeExcelHost();
  const stop = new AbortController();
  host.afterWriteSync = () => stop.abort();
  await withFakeExcel(host, async () => {
    const result = await createFillFormulaTool().execute("fill-stop", { range: "B1:B3", formula: "=A1" }, stop.signal);
    assert.match(text(result), /Stopped after filling Sheet1!B1:B3\*\* — the change stands/u);
    assert.equal(result.details.stopped, true);
    assert.equal(result.details.writtenCount, 3);
    assert.equal(host.sheet("Sheet1").formula("B3"), "=A3");
  });
});
