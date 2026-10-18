import assert from "node:assert/strict";
import { test } from "node:test";
import type { QueuedInput } from "../src/durable/inbox.ts";
import { restoreQueuedMessagesToEditor } from "../src/taskpane/keyboard-shortcuts.ts";

void test("restore waits for atomic withdrawal and leaves the latest draft alone when placement won", async () => {
  const input = { value: "original draft" };
  let complete: (items: QueuedInput[]) => void = () => {};
  const taken = new Promise<QueuedInput[]>((resolve) => { complete = resolve; });
  const restoring = restoreQueuedMessagesToEditor({ getInput: () => input, getTextarea: () => undefined }, {
    takeQueuedInputs: () => taken,
  });
  assert.equal(input.value, "original draft");
  input.value = "new draft";
  // Empty withdrawal is the native result when placement won the race.
  complete([]);
  assert.equal(await restoring, 0);
  assert.equal(input.value, "new draft");
});

void test("failed withdrawal never changes the composer", async () => {
  const input = { value: "draft" };
  await assert.rejects(restoreQueuedMessagesToEditor({ getInput: () => input, getTextarea: () => undefined }, {
    takeQueuedInputs: () => Promise.reject(new Error("persistence failed")),
  }), /persistence failed/);
  assert.equal(input.value, "draft");
});
