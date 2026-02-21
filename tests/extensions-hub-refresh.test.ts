import assert from "node:assert/strict";
import { test } from "node:test";

import { createDeferredConnectionsRefreshController } from "../src/commands/builtins/extensions-hub-refresh.ts";

function nextTick(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

void test("refresh controller runs immediately when no input is active", () => {
  let refreshCount = 0;

  const controller = createDeferredConnectionsRefreshController({
    isDisposed: () => false,
    hasActiveSecretInput: () => false,
    refresh: () => {
      refreshCount += 1;
    },
  });

  controller.requestRefresh();
  assert.equal(refreshCount, 1);
});

void test("refresh controller defers while input is active and flushes when editing ends", async () => {
  let refreshCount = 0;
  let activeInput = true;

  const controller = createDeferredConnectionsRefreshController({
    isDisposed: () => false,
    hasActiveSecretInput: () => activeInput,
    refresh: () => {
      refreshCount += 1;
    },
  });

  controller.requestRefresh();
  assert.equal(refreshCount, 0);

  // User tabs from one input to another — still editing, no flush.
  controller.onConnectionsFocusOut();
  await nextTick();
  assert.equal(refreshCount, 0);

  // Editing ends, next focusout flushes deferred refresh.
  activeInput = false;
  controller.onConnectionsFocusOut();
  await nextTick();
  assert.equal(refreshCount, 1);
});

void test("refresh controller dispose cancels pending flush", async () => {
  let refreshCount = 0;

  const controller = createDeferredConnectionsRefreshController({
    isDisposed: () => false,
    hasActiveSecretInput: () => true,
    refresh: () => {
      refreshCount += 1;
    },
  });

  controller.requestRefresh();
  controller.onConnectionsFocusOut();
  controller.dispose();

  await nextTick();
  assert.equal(refreshCount, 0);
});
