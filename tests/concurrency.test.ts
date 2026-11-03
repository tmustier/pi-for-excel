import assert from "node:assert/strict";
import test from "node:test";

import { createAsyncCoalescer, createConcurrencyLimit, mapWithConcurrency } from "../src/utils/concurrency.ts";

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

void test("createAsyncCoalescer does not lose a request as the previous task settles", async () => {
  let release: () => void = () => {};
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  let runCount = 0;
  const run = createAsyncCoalescer(() => {
    runCount += 1;
    return gate;
  });

  const first = run();
  await Promise.resolve();
  const atSettlement = gate.then(run);
  release();
  await Promise.all([first, atSettlement]);

  assert.equal(runCount, 2);
});

void test("createAsyncCoalescer can run again after a task throws before returning a promise", async () => {
  let attempts = 0;
  const run = createAsyncCoalescer(() => {
    attempts += 1;
    if (attempts === 1) throw new Error("failed to start");
    return Promise.resolve();
  });

  await assert.rejects(run(), /failed to start/u);
  await run();
  assert.equal(attempts, 2);
});

void test("createAsyncCoalescer coalesces overlapping calls into a single rerun", async () => {
  let runCount = 0;
  const blockers: Array<() => void> = [];

  const run = createAsyncCoalescer(async () => {
    runCount += 1;
    await new Promise<void>((resolve) => {
      blockers.push(resolve);
    });
  });

  const first = run();
  await Promise.resolve();

  assert.equal(runCount, 1);
  assert.equal(blockers.length, 1);

  const second = run();
  const third = run();
  await Promise.resolve();

  assert.equal(runCount, 1);

  const releaseFirst = blockers.shift();
  if (!releaseFirst) {
    throw new Error("Expected first blocker");
  }
  releaseFirst();

  await Promise.resolve();
  await Promise.resolve();

  assert.equal(runCount, 2);
  assert.equal(blockers.length, 1);

  const releaseSecond = blockers.shift();
  if (!releaseSecond) {
    throw new Error("Expected second blocker");
  }
  releaseSecond();

  await Promise.all([first, second, third]);
  assert.equal(runCount, 2);
});
