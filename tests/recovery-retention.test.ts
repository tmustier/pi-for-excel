import assert from "node:assert/strict";
import { test } from "node:test";

import {
  clampRetentionLimit,
  MAX_RECOVERY_ENTRIES,
  MIN_RETENTION_LIMIT,
} from "../src/workbook/recovery/constants.ts";

void test("clampRetentionLimit returns default for non-number input", () => {
  assert.equal(clampRetentionLimit(undefined), MAX_RECOVERY_ENTRIES);
  assert.equal(clampRetentionLimit(null), MAX_RECOVERY_ENTRIES);
  assert.equal(clampRetentionLimit("50"), MAX_RECOVERY_ENTRIES);
  assert.equal(clampRetentionLimit(NaN), MAX_RECOVERY_ENTRIES);
  assert.equal(clampRetentionLimit(Infinity), MAX_RECOVERY_ENTRIES);
});

void test("clampRetentionLimit floors to integer", () => {
  assert.equal(clampRetentionLimit(50.7), 50);
  assert.equal(clampRetentionLimit(5.1), 5);
});

void test("clampRetentionLimit clamps to min", () => {
  assert.equal(clampRetentionLimit(1), MIN_RETENTION_LIMIT);
  assert.equal(clampRetentionLimit(0), MIN_RETENTION_LIMIT);
  assert.equal(clampRetentionLimit(-10), MIN_RETENTION_LIMIT);
});

void test("clampRetentionLimit clamps to max", () => {
  assert.equal(clampRetentionLimit(999), MAX_RECOVERY_ENTRIES);
  assert.equal(clampRetentionLimit(MAX_RECOVERY_ENTRIES + 1), MAX_RECOVERY_ENTRIES);
});

void test("clampRetentionLimit passes through valid values", () => {
  assert.equal(clampRetentionLimit(5), 5);
  assert.equal(clampRetentionLimit(50), 50);
  assert.equal(clampRetentionLimit(120), 120);
});
