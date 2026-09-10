import assert from "node:assert/strict";
import { test } from "node:test";

import { getCompactionThresholds } from "../src/compaction/defaults.ts";

// Policy table: published compaction budgets. The 200k consumer contract is
// independently exercised through the agent request seam in agent-request.test.ts.
void test("compaction thresholds follow the published context-window policy", () => {
  const cases = [
    {
      name: "200k quality cap",
      contextWindow: 200_000,
      expected: { contextWindow: 200_000, reserveTokens: 16_384, hardTriggerTokens: 170_000, softWarningTokens: 160_000 },
    },
    {
      name: "small-window reserve",
      contextWindow: 32_768,
      expected: { contextWindow: 32_768, reserveTokens: 16_384, hardTriggerTokens: 16_384, softWarningTokens: 14_336 },
    },
    {
      name: "invalid-window fallback",
      contextWindow: Number.NaN,
      expected: { contextWindow: 200_000, reserveTokens: 16_384, hardTriggerTokens: 170_000, softWarningTokens: 160_000 },
    },
  ];

  for (const entry of cases) {
    assert.deepEqual(getCompactionThresholds(entry.contextWindow), entry.expected, entry.name);
  }
});
