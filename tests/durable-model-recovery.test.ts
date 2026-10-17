import assert from "node:assert/strict";
import { test } from "node:test";
import { createAssistantMessageEventStream, isRetryableAssistantError, type AssistantMessage } from "@earendil-works/pi-ai";
import { fauxAssistantMessage } from "@earendil-works/pi-ai/providers/faux";
import {
  createModelRecoveryLifecycle,
  prepareDurableReload,
  protectModelStream,
} from "../src/durable/model-recovery.ts";

function transportFailure(): AssistantMessage {
  return { ...fauxAssistantMessage(""), stopReason: "error", errorMessage: "Load failed" };
}

for (const event of ["beforeunload", "pagehide", "explicit reload"]) {
  void test(`Durable seals commit admission before ${event} cancels provider requests`, () => {
    const target = new EventTarget();
    let closes = 0;
    const lifecycle = createModelRecoveryLifecycle(() => {
      closes++;
      return Promise.resolve();
    }, target);
    if (event === "explicit reload") prepareDurableReload(target);
    else target.dispatchEvent(new Event(event));
    assert.equal(closes, 1, "closing must start synchronously, not after navigation");
    assert.equal(lifecycle.suspending, true);
    target.dispatchEvent(new Event("pagehide"));
    assert.equal(closes, 1);
    lifecycle.dispose();
  });
}

void test("WebKit's terminal fetch error yields to an already pending unload", async () => {
  const target = new EventTarget();
  let sealed = false;
  const lifecycle = createModelRecoveryLifecycle(() => {
    sealed = true;
    return Promise.resolve();
  }, target);
  const source = createAssistantMessageEventStream();
  const protectedStream = protectModelStream(source, lifecycle, transportFailure);
  const failure = transportFailure();
  source.push({ type: "error", reason: "error", error: failure });
  // Model the task queued by the browser for unloading after fetch cancellation.
  setTimeout(() => target.dispatchEvent(new Event("pagehide")), 0);
  const events = [];
  for await (const event of protectedStream) events.push(event);
  assert.equal(sealed, true);
  assert.equal(events.length, 1);
  assert.equal(events[0]?.type, "error");
  assert.equal((await protectedStream.result()).stopReason, "aborted");
  lifecycle.dispose();
});

void test("ordinary provider errors are not suppressed or reclassified", async () => {
  const lifecycle = createModelRecoveryLifecycle(() => Promise.resolve(), new EventTarget());
  const source = createAssistantMessageEventStream();
  const protectedStream = protectModelStream(source, lifecycle, transportFailure);
  const failure = { ...transportFailure(), errorMessage: "Forbidden" };
  source.push({ type: "error", reason: "error", error: failure });
  for await (const event of protectedStream) assert.equal(event.type, "error");
  assert.deepEqual(await protectedStream.result(), failure);
  assert.equal(lifecycle.suspending, false);
  lifecycle.dispose();
});

void test("successful streams and tool calls keep their terminal message", async () => {
  const lifecycle = createModelRecoveryLifecycle(() => Promise.resolve(), new EventTarget());
  const source = createAssistantMessageEventStream();
  const protectedStream = protectModelStream(source, lifecycle, transportFailure);
  const success = fauxAssistantMessage("Done");
  source.push({ type: "done", reason: "stop", message: success });
  for await (const event of protectedStream) assert.equal(event.type, "done");
  assert.deepEqual(await protectedStream.result(), success);
  lifecycle.dispose();
});

void test("disposing a closed harness removes navigation handlers", () => {
  const target = new EventTarget();
  let closes = 0;
  const lifecycle = createModelRecoveryLifecycle(() => { closes++; return Promise.resolve(); }, target);
  lifecycle.dispose();
  prepareDurableReload(target);
  assert.equal(closes, 0);
});

void test("native WebKit fetch cancellation remains a retryable pending request without an unload event", async () => {
  const lifecycle = createModelRecoveryLifecycle(() => Promise.resolve(), new EventTarget());
  const source = createAssistantMessageEventStream();
  const stream = protectModelStream(source, lifecycle, transportFailure);
  const error = transportFailure();
  assert.equal(isRetryableAssistantError(error), false, "upstream does not recognize WebKit wording");
  source.push({ type: "error", reason: "error", error });
  for await (const event of stream) assert.equal(event.type, "error");
  const result = await stream.result();
  assert.equal(result.errorMessage, "Network error: Load failed");
  assert.equal(isRetryableAssistantError(result), true);
  assert.equal(lifecycle.suspending, false);
  lifecycle.dispose();
});
