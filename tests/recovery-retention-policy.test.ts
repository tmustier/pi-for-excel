import assert from "node:assert/strict";
import { test } from "node:test";

import {
  clampRetentionLimit,
  MAX_RECOVERY_ENTRIES,
  MIN_RETENTION_LIMIT,
} from "../src/workbook/recovery/constants.ts";

void test("retention limit normalizes default, floor, min, max, and valid values", () => {
  const cases: ReadonlyArray<{ input: DynamicValue; expected: number; label: string }> = [
    { input: undefined, expected: MAX_RECOVERY_ENTRIES, label: "undefined defaults" },
    { input: null, expected: MAX_RECOVERY_ENTRIES, label: "null defaults" },
    { input: "50", expected: MAX_RECOVERY_ENTRIES, label: "string defaults" },
    { input: NaN, expected: MAX_RECOVERY_ENTRIES, label: "NaN defaults" },
    { input: Infinity, expected: MAX_RECOVERY_ENTRIES, label: "infinity defaults" },
    { input: 50.7, expected: 50, label: "fraction floors" },
    { input: 1, expected: MIN_RETENTION_LIMIT, label: "below minimum clamps" },
    { input: -10, expected: MIN_RETENTION_LIMIT, label: "negative clamps" },
    { input: 999, expected: MAX_RECOVERY_ENTRIES, label: "above maximum clamps" },
    { input: 5, expected: 5, label: "minimum passes through" },
    { input: 50, expected: 50, label: "middle passes through" },
    { input: 120, expected: 120, label: "maximum passes through" },
  ];

  for (const scenario of cases) {
    assert.equal(clampRetentionLimit(scenario.input), scenario.expected, scenario.label);
  }
});
