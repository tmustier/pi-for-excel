import assert from "node:assert/strict";
import { test } from "node:test";

import { buildWorkbookCellChangeSummary } from "../src/audit/cell-diff.ts";

void test("buildWorkbookCellChangeSummary captures value and formula changes", () => {
  const summary = buildWorkbookCellChangeSummary({
    sheetName: "Sheet1",
    startCell: "A1",
    beforeValues: [[1, 2], [3, 4]],
    beforeFormulas: [["", ""], ["", ""]],
    afterValues: [[1, 20], [3, 4]],
    afterFormulas: [["", ""], ["=A1*2", ""]],
  });

  assert.equal(summary.changedCount, 2);
  assert.equal(summary.truncated, false);
  assert.equal(summary.sample.length, 2);

  const byAddress = new Map(summary.sample.map((change) => [change.address, change]));

  const valueChange = byAddress.get("Sheet1!B1");
  assert.ok(valueChange);
  assert.equal(valueChange?.beforeValue, "2");
  assert.equal(valueChange?.afterValue, "20");

  const formulaChange = byAddress.get("Sheet1!A2");
  assert.ok(formulaChange);
  assert.equal(formulaChange?.beforeFormula, undefined);
  assert.equal(formulaChange?.afterFormula, "=A1*2");
});

void test("buildWorkbookCellChangeSummary truncates sample to limit", () => {
  const summary = buildWorkbookCellChangeSummary({
    sheetName: "Sheet1",
    startCell: "A1",
    beforeValues: [[1, 2, 3]],
    beforeFormulas: [["", "", ""]],
    afterValues: [[10, 20, 30]],
    afterFormulas: [["", "", ""]],
    sampleLimit: 2,
  });

  assert.equal(summary.changedCount, 3);
  assert.equal(summary.truncated, true);
  assert.equal(summary.sample.length, 2);
});
