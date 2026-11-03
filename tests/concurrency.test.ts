import assert from "node:assert/strict";
import test from "node:test";

import { createConcurrencyLimit, mapWithConcurrency } from "../src/utils/concurrency.ts";

void test("mapWithConcurrency keeps input order and never exceeds the limit", async () => {
  let active = 0;
  let peak = 0;
  const result = await mapWithConcurrency([30, 5, 20, 1, 10, 15], 2, async (delay) => {
    active += 1;
    peak = Math.max(peak, active);
    await new Promise((resolve) => setTimeout(resolve, delay));
    active -= 1;
    return delay * 2;
  });
  assert.deepEqual(result, [60, 10, 40, 2, 20, 30]);
  assert.equal(peak, 2);
});

void test("a rejected task frees its slot for the next one", async () => {
  const limit = createConcurrencyLimit(1);
  const failed = limit(() => Promise.reject(new Error("boom")));
  const next = limit(() => Promise.resolve("ran"));
  await assert.rejects(failed, /boom/);
  assert.equal(await next, "ran");
});
