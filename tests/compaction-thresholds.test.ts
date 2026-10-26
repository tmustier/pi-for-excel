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

// The context meter turns red only once compaction is near, on every window size.
void test("context meter colours follow the compaction budget", async () => {
  const { getStatusContextHealth } = await import("../src/taskpane/status-context.ts");
  const { shouldAutoCompactForProjectedTokens } = await import("../src/compaction/auto-compaction.ts");
  const cases: Array<[tokens: number, window: number, colour: string]> = [
    [150_000, 200_000, "pi-status-ctx--yellow"], // 75% of a 200k window, 20k before compaction
    [161_000, 200_000, "pi-status-ctx--red"],
    [500_000, 1_000_000, ""], // 50% of 1M, far from compaction
    [15_000, 32_768, "pi-status-ctx--red"], // 46% of a small window, just before compaction
    [40_000, 32_768, "pi-status-ctx--red"],
  ];
  for (const [tokens, window, colour] of cases) {
    assert.equal(getStatusContextHealth(tokens, window).colorClass, colour, `${tokens}/${window}`);
  }
  // Red always comes before the compaction trigger.
  for (const window of [32_768, 128_000, 200_000, 1_000_000]) {
    const { softWarningTokens } = getCompactionThresholds(window);
    assert.equal(getStatusContextHealth(softWarningTokens + 1, window).colorClass, "pi-status-ctx--red");
    assert.equal(shouldAutoCompactForProjectedTokens({ projectedTokens: softWarningTokens + 1, contextWindow: window }), false);
  }
});
