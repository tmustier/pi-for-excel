import assert from "node:assert/strict";
import { test } from "node:test";

import { processImageInput } from "../src/messages/image-input-limits.ts";

void test("a synchronous image-worker send failure terminates the worker and removes abort listeners", async () => {
  const previousWorker: unknown = Reflect.get(globalThis, "Worker");
  const hadWorker = Reflect.has(globalThis, "Worker");
  let terminated = 0;
  class ThrowingWorker {
    postMessage(): void { throw new DOMException("Cannot clone image request", "DataCloneError"); }
    terminate(): void { terminated += 1; }
  }
  Reflect.set(globalThis, "Worker", ThrowingWorker);
  const controller = new AbortController();
  let removed = 0;
  const remove = controller.signal.removeEventListener.bind(controller.signal);
  controller.signal.removeEventListener = (
    type: string,
    listener: EventListenerOrEventListenerObject | null,
    options?: boolean | EventListenerOptions,
  ) => { removed += 1; remove(type, listener, options); };
  try {
    const result = await processImageInput({ type: "image", data: "AAAA", mimeType: "image/png" }, {
      signal: controller.signal,
    }).catch(() => null);
    assert.equal(terminated, 1, "send failure must terminate its worker");
    assert.equal(removed, 1, "send failure must detach its abort listener");
    assert.equal(result?.ok, false);
    controller.abort();
    assert.equal(terminated, 1, "abort after completion must not terminate twice");
  } finally {
    if (hadWorker) Reflect.set(globalThis, "Worker", previousWorker);
    else Reflect.deleteProperty(globalThis, "Worker");
  }
});
