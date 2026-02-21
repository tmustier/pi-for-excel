import assert from "node:assert/strict";
import { test } from "node:test";

import {
  awaitWithTimeout,
  createAsyncCoalescer,
  isLikelyCorsErrorMessage,
  normalizeRuntimeTools,
} from "../src/taskpane/runtime-utils.ts";

void test("isLikelyCorsErrorMessage detects known cors/network signatures", () => {
  assert.equal(isLikelyCorsErrorMessage("Failed to fetch"), true);
  assert.equal(isLikelyCorsErrorMessage("Load failed"), true);
  assert.equal(isLikelyCorsErrorMessage("CORS requests are not allowed"), true);
  assert.equal(isLikelyCorsErrorMessage("Cross-Origin policy blocked request"), true);
  assert.equal(isLikelyCorsErrorMessage("provider overloaded"), false);
});

void test("normalizeRuntimeTools drops invalid and duplicate entries", () => {
  const firstTool = {
    name: "alpha",
    label: "Alpha",
    description: "alpha tool",
    parameters: { type: "object", properties: {} },
    execute: () => ({ content: [{ type: "text", text: "ok" }] }),
  };

  const duplicateByName = {
    name: "alpha",
    label: "Alpha duplicate",
    description: "duplicate",
    parameters: { type: "object", properties: {} },
    execute: () => ({ content: [{ type: "text", text: "dup" }] }),
  };

  const invalid = {
    name: "missing-execute",
    label: "Invalid",
    description: "invalid",
    parameters: { type: "object", properties: {} },
  };

  const normalized = normalizeRuntimeTools([
    invalid,
    firstTool,
    duplicateByName,
  ]);

  assert.equal(normalized.length, 1);
  assert.equal(normalized[0]?.name, "alpha");
  assert.equal(normalized[0]?.description, "alpha tool");
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

void test("awaitWithTimeout resolves when task finishes in time", async () => {
  const value = await awaitWithTimeout("quick task", 50, Promise.resolve("ok"));
  assert.equal(value, "ok");
});

void test("awaitWithTimeout rejects with label on timeout", async () => {
  await assert.rejects(
    awaitWithTimeout(
      "slow task",
      5,
      new Promise<string>(() => {
        // Never resolves; timeout controls completion.
      }),
    ),
    /slow task timed out after 5ms/,
  );
});
