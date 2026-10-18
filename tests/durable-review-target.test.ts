import assert from "node:assert/strict";
import { test } from "node:test";
import type { MutationRecord } from "../src/durable/mutations.ts";
import { t } from "../src/language/index.ts";
import { mutationTargetLabel } from "../src/ui/durable-controls.ts";

function record(action: string, ranges: string[] = [], sheetId = "sheet:original"): MutationRecord {
  return {
    version: 3, mutationId: "task:review", toolCallId: "call:review", toolName: "modify_structure",
    workbookId: "workbook:review", sessionId: "session:review", requiresReview: false, phase: "dispatched",
    intent: { action }, target: { kind: "sheet", sheetId, ranges, chartId: null, action },
    before: { workbookId: "workbook:review", structure: [{ id: "sheet:original", name: "Original", position: 0,
      visibility: "Visible", active: true, tables: [] }], state: null, complete: false, limitation: null },
    expected: null, after: null, resolution: null,
  };
}

void test("sheet deletion review names the original stable target without a cell range", () => {
  assert.equal(mutationTargetLabel(record("delete_sheet")), "Original");
});

void test("sheet duplication review names the recorded source sheet", () => {
  assert.equal(mutationTargetLabel(record("duplicate_sheet")), "Original");
});

void test("a new sheet is not confused with the existing context sheet", () => {
  assert.equal(mutationTargetLabel(record("add_sheet")), t("durable.targetUnknown"));
});

void test("review target names never substitute another sheet for an absent stable ID", () => {
  assert.equal(mutationTargetLabel(record("delete_sheet", [], "sheet:missing")), t("durable.targetUnknown"));
});

void test("range review labels retain their recorded sheet identity and full address", () => {
  assert.equal(mutationTargetLabel(record("delete_rows", ["3:5"])), "Original!3:5");
});
