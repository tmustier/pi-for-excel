import assert from "node:assert/strict";
import { test } from "node:test";

import {
  effectiveRecentToolResultsToKeep,
  effectiveToolOutputLimits,
} from "../src/context/window-budgets.ts";
import {
  DEFAULT_TOOL_OUTPUT_MAX_BYTES,
  DEFAULT_TOOL_OUTPUT_MAX_LINES,
} from "../src/tools/output-truncation.ts";
import { DEFAULT_TOOL_RESULT_SHAPING } from "../src/messages/tool-result-shaping.ts";

void test("windows >= 128k keep full default budgets", () => {
  for (const contextWindow of [128_000, 200_000, 1_000_000]) {
    assert.deepEqual(effectiveToolOutputLimits(contextWindow), {
      maxBytes: DEFAULT_TOOL_OUTPUT_MAX_BYTES,
      maxLines: DEFAULT_TOOL_OUTPUT_MAX_LINES,
    });
    assert.equal(
      effectiveRecentToolResultsToKeep(contextWindow),
      DEFAULT_TOOL_RESULT_SHAPING.recentToolResultsToKeep,
    );
  }
});

void test("unknown or invalid windows are treated as large (defaults)", () => {
  for (const contextWindow of [undefined, 0, -1, Number.NaN]) {
    assert.deepEqual(effectiveToolOutputLimits(contextWindow), {
      maxBytes: DEFAULT_TOOL_OUTPUT_MAX_BYTES,
      maxLines: DEFAULT_TOOL_OUTPUT_MAX_LINES,
    });
    assert.equal(
      effectiveRecentToolResultsToKeep(contextWindow),
      DEFAULT_TOOL_RESULT_SHAPING.recentToolResultsToKeep,
    );
  }
});

void test("a 65k window scales tool output caps to ~half", () => {
  const limits = effectiveToolOutputLimits(65_536);

  // 65,536 / 128,000 = 0.512
  assert.equal(limits.maxBytes, Math.floor(DEFAULT_TOOL_OUTPUT_MAX_BYTES * (65_536 / 128_000)));
  assert.equal(limits.maxLines, Math.floor(DEFAULT_TOOL_OUTPUT_MAX_LINES * (65_536 / 128_000)));
  assert.ok(limits.maxBytes < DEFAULT_TOOL_OUTPUT_MAX_BYTES);
  assert.ok(limits.maxLines < DEFAULT_TOOL_OUTPUT_MAX_LINES);

  assert.equal(effectiveRecentToolResultsToKeep(65_536), 3);
});

void test("tiny windows hit the floors", () => {
  const limits = effectiveToolOutputLimits(4_096);

  assert.equal(limits.maxBytes, 8 * 1024);
  assert.equal(limits.maxLines, 200);
  assert.equal(effectiveRecentToolResultsToKeep(4_096), 2);
});
