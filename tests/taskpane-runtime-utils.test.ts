import assert from "node:assert/strict";
import { test } from "node:test";

import {
  awaitCredentialRestoreForStartup,
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
  parameters?: DynamicValue;
}): {
  name: string;
  label: string;
  description: string;
  parameters: DynamicValue;
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
      previousFingerprint: "aaaa",
      nextFingerprint: "bbbb",
      previousExtensionToolRevision: 1,
      nextExtensionToolRevision: 1,
    }),
    true,
  );

  assert.equal(
    shouldApplyRuntimeToolUpdate({
      previousFingerprint: "same",
      nextFingerprint: "same",
      previousExtensionToolRevision: 3,
      nextExtensionToolRevision: 3,
    }),
    false,
  );
});

void test("shouldApplyRuntimeToolUpdate applies updates when extension tool revision changes", () => {
  assert.equal(
    shouldApplyRuntimeToolUpdate({
      previousFingerprint: "same",
      nextFingerprint: "same",
      previousExtensionToolRevision: 2,
      nextExtensionToolRevision: 3,
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

void test("credential restore adds no duplicate fast refresh and one late refresh after timeout", async () => {
  let refreshCount = 0;
  await awaitCredentialRestoreForStartup(Promise.resolve(), 50, () => {
    refreshCount += 1;
  });
  refreshCount += 1; // The normal provider lookup performed by taskpane startup.
  assert.equal(refreshCount, 1);

  refreshCount = 0;
  let resolveRestore: (() => void) | undefined;
  const delayedRestore = new Promise<void>((resolve) => {
    resolveRestore = resolve;
  });
  await assert.rejects(
    awaitCredentialRestoreForStartup(delayedRestore, 5, () => {
      refreshCount += 1;
    }),
    /timed out/,
  );
  refreshCount += 1; // First-paint provider lookup proceeds after the timeout.
  resolveRestore?.();
  await delayedRestore;
  await Promise.resolve();
  assert.equal(refreshCount, 2);
});

void test("rejected credential restore does not schedule refresh", async () => {
  let refreshCount = 0;
  await assert.rejects(
    awaitCredentialRestoreForStartup(Promise.reject(new Error("restore failed")), 50, () => {
      refreshCount += 1;
    }),
    /restore failed/,
  );
  await Promise.resolve();
  assert.equal(refreshCount, 0);
});

void test("credential restore rejection after timeout remains observable", async (t) => {
  const warnings = t.mock.method(console, "warn", () => {});
  let rejectRestore: ((error: Error) => void) | undefined;
  const promise = new Promise<void>((_resolve, reject) => { rejectRestore = reject; });
  await assert.rejects(awaitCredentialRestoreForStartup(promise, 5, () => {
    assert.fail("Rejected credentials must not trigger a refresh");
  }), /timed out/);

  const failure = new Error("late restore failure");
  assert.ok(rejectRestore);
  rejectRestore(failure);
  await promise.catch(() => {});
  await Promise.resolve();
  assert.equal(warnings.mock.callCount(), 1);
  assert.deepEqual(warnings.mock.calls[0]?.arguments, [
    "[auth] Credential restore failed after timeout:", failure,
  ]);
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
