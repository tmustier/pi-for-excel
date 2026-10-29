import assert from "node:assert/strict";
import { test } from "node:test";

import { getCompactionThresholds } from "../src/compaction/defaults.ts";

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

void test("context meter colours follow the compaction budget", async () => {
  const { getStatusContextHealth } = await import("../src/taskpane/status-context.ts");
  const { shouldAutoCompactForProjectedTokens } = await import("../src/compaction/auto-compaction.ts");
  const cases: Array<[tokens: number, window: number, colour: string]> = [
    [150_000, 200_000, "pi-status-ctx--yellow"],
    [161_000, 200_000, "pi-status-ctx--red"],
    [500_000, 1_000_000, ""],
    [15_000, 32_768, "pi-status-ctx--red"],
    [40_000, 32_768, "pi-status-ctx--red"],
  ];
  for (const [tokens, window, colour] of cases) {
    assert.equal(getStatusContextHealth(tokens, window).colorClass, colour, `${tokens}/${window}`);
  }
  for (const window of [32_768, 128_000, 200_000, 1_000_000]) {
    const { softWarningTokens } = getCompactionThresholds(window);
    assert.equal(getStatusContextHealth(softWarningTokens + 1, window).colorClass, "pi-status-ctx--red");
    assert.equal(shouldAutoCompactForProjectedTokens({ projectedTokens: softWarningTokens + 1, contextWindow: window }), false);
  }
});
