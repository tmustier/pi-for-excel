import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { createFillFormulaTool } from "../src/tools/fill-formula.ts";
import { createWriteCellsTool } from "../src/tools/write-cells.ts";
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
    assert.match(text(result), /No backup/u);
  });
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
    assert.match(text(result), new RegExp(`Partly written\\*\\* — rows 1–${lastWritten} of Big!A1:A15000`, "u"));
    assert.match(text(result), /Excel is in cell-edit mode/u);
    assert.match(text(result), /Nothing was rolled back/u);
    assert.equal(audits.at(-1)?.blocked, false, "the rows that landed are audited as a change");
  });
});
