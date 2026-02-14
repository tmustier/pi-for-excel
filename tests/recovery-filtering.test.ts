import assert from "node:assert/strict";
import { test } from "node:test";

import {
  applyRecoveryFilters,
  buildToolFilterOptions,
  DEFAULT_FILTER_STATE,
  type RecoveryFilterState,
} from "../src/commands/builtins/recovery-filtering.ts";
import type { RecoveryCheckpointSummary } from "../src/commands/builtins/recovery-overlay.ts";

function makeCheckpoint(
  overrides: Partial<RecoveryCheckpointSummary> & { id: string },
): RecoveryCheckpointSummary {
  return {
    at: Date.now(),
    toolName: "write_cells",
    address: "Sheet1!A1:B10",
    changedCount: 5,
    ...overrides,
  };
}

const FIXTURES: RecoveryCheckpointSummary[] = [
  makeCheckpoint({ id: "cp1", at: 1000, toolName: "write_cells", address: "Sheet1!A1:B10" }),
  makeCheckpoint({ id: "cp2", at: 2000, toolName: "format_cells", address: "Sheet1!C1:C20" }),
  makeCheckpoint({ id: "cp3", at: 3000, toolName: "write_cells", address: "Sheet2!A1:A5" }),
  makeCheckpoint({ id: "cp4", at: 4000, toolName: "restore_snapshot", address: "Sheet1!A1:B10", restoredFromSnapshotId: "cp1" }),
  makeCheckpoint({ id: "cp5", at: 5000, toolName: "modify_structure", address: "Sheet3" }),
];

void test("default filter returns all checkpoints sorted newest first", () => {
  const result = applyRecoveryFilters(FIXTURES, DEFAULT_FILTER_STATE);
  assert.equal(result.length, 5);
  assert.equal(result[0].id, "cp5");
  assert.equal(result[4].id, "cp1");
});

void test("sort oldest first reverses order", () => {
  const state: RecoveryFilterState = { ...DEFAULT_FILTER_STATE, sortOrder: "oldest" };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result[0].id, "cp1");
  assert.equal(result[4].id, "cp5");
});

void test("tool filter narrows to matching tool", () => {
  const state: RecoveryFilterState = { ...DEFAULT_FILTER_STATE, toolFilter: "write_cells" };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result.length, 2);
  assert.ok(result.every((c) => c.toolName === "write_cells"));
});

void test("search matches by id", () => {
  const state: RecoveryFilterState = { ...DEFAULT_FILTER_STATE, search: "cp3" };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result.length, 1);
  assert.equal(result[0].id, "cp3");
});

void test("search matches by address (case insensitive)", () => {
  const state: RecoveryFilterState = { ...DEFAULT_FILTER_STATE, search: "sheet2" };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result.length, 1);
  assert.equal(result[0].id, "cp3");
});

void test("search matches by tool label", () => {
  const state: RecoveryFilterState = { ...DEFAULT_FILTER_STATE, search: "Modify structure" };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result.length, 1);
  assert.equal(result[0].toolName, "modify_structure");
});

void test("search and tool filter combine", () => {
  const state: RecoveryFilterState = {
    search: "Sheet1",
    toolFilter: "write_cells",
    sortOrder: "newest",
  };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result.length, 1);
  assert.equal(result[0].id, "cp1");
});

void test("empty search matches everything", () => {
  const state: RecoveryFilterState = { ...DEFAULT_FILTER_STATE, search: "   " };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result.length, 5);
});

void test("no matches returns empty array", () => {
  const state: RecoveryFilterState = { ...DEFAULT_FILTER_STATE, search: "nonexistent" };
  const result = applyRecoveryFilters(FIXTURES, state);
  assert.equal(result.length, 0);
});

void test("buildToolFilterOptions includes all with counts", () => {
  const options = buildToolFilterOptions(FIXTURES);
  const all = options.find((o) => o.value === "all");
  assert.ok(all);
  assert.equal(all.count, 5);

  const write = options.find((o) => o.value === "write_cells");
  assert.ok(write);
  assert.equal(write.count, 2);

  const format = options.find((o) => o.value === "format_cells");
  assert.ok(format);
  assert.equal(format.count, 1);
});

void test("buildToolFilterOptions omits tools with zero count", () => {
  const options = buildToolFilterOptions(FIXTURES);
  const python = options.find((o) => o.value === "python_transform_range");
  assert.equal(python, undefined);
});

void test("buildToolFilterOptions handles empty input", () => {
  const options = buildToolFilterOptions([]);
  assert.equal(options.length, 1);
  assert.equal(options[0].value, "all");
  assert.equal(options[0].count, 0);
});
