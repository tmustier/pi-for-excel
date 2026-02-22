import assert from "node:assert/strict";
import { test } from "node:test";

import {
  awaitWithTimeout,
  createAsyncCoalescer,
  createRuntimeToolFingerprint,
  isLikelyCorsErrorMessage,
  normalizeRuntimeTools,
  shouldApplyRuntimeToolUpdate,
} from "../src/taskpane/runtime-utils.ts";

function createFingerprintTestTool(args: {
  name: string;
  description: string;
  parameters?: unknown;
}): {
  name: string;
  label: string;
  description: string;
  parameters: unknown;
  execute: () => Promise<{ content: Array<{ type: "text"; text: string }>; details: null }>;
} {
  return {
    name: args.name,
    label: args.name,
    description: args.description,
    parameters: args.parameters ?? { type: "object", properties: {} },
    execute: () => Promise.resolve({
      content: [{ type: "text", text: "ok" }],
      details: null,
    }),
  };
}

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

void test("createRuntimeToolFingerprint is stable for equivalent tool metadata", () => {
  const first = [
    createFingerprintTestTool({ name: "alpha", description: "alpha tool" }),
    createFingerprintTestTool({
      name: "beta",
      description: "beta tool",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string" },
        },
        required: ["query"],
      },
    }),
  ];

  const second = [
    createFingerprintTestTool({ name: "alpha", description: "alpha tool" }),
    createFingerprintTestTool({
      name: "beta",
      description: "beta tool",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string" },
        },
        required: ["query"],
      },
    }),
  ];

  assert.equal(createRuntimeToolFingerprint(first), createRuntimeToolFingerprint(second));
});

void test("createRuntimeToolFingerprint changes when tool metadata changes", () => {
  const baseline = [
    createFingerprintTestTool({ name: "alpha", description: "alpha tool" }),
    createFingerprintTestTool({ name: "beta", description: "beta tool" }),
  ];

  const changedDescription = [
    createFingerprintTestTool({ name: "alpha", description: "alpha tool (v2)" }),
    createFingerprintTestTool({ name: "beta", description: "beta tool" }),
  ];

  const reordered = [
    createFingerprintTestTool({ name: "beta", description: "beta tool" }),
    createFingerprintTestTool({ name: "alpha", description: "alpha tool" }),
  ];

  const baselineFingerprint = createRuntimeToolFingerprint(baseline);

  assert.notEqual(baselineFingerprint, createRuntimeToolFingerprint(changedDescription));
  assert.notEqual(baselineFingerprint, createRuntimeToolFingerprint(reordered));
});

void test("shouldApplyRuntimeToolUpdate applies updates when fingerprint changes", () => {
  assert.equal(
    shouldApplyRuntimeToolUpdate({
      hasExtensionTools: false,
      previousFingerprint: "aaaa",
      nextFingerprint: "bbbb",
    }),
    true,
  );

  assert.equal(
    shouldApplyRuntimeToolUpdate({
      hasExtensionTools: false,
      previousFingerprint: "same",
      nextFingerprint: "same",
    }),
    false,
  );
});

void test("shouldApplyRuntimeToolUpdate forces updates when extension tools are present", () => {
  assert.equal(
    shouldApplyRuntimeToolUpdate({
      hasExtensionTools: true,
      previousFingerprint: "same",
      nextFingerprint: "same",
    }),
    true,
  );
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
