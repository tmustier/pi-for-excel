import { authDeferred } from "./fixtures/auth-deferred.ts";
import "fake-indexeddb/auto";
import assert from "node:assert/strict";
import { test } from "node:test";

import { initAppStorage } from "../src/storage/init-app-storage.ts";
import { pollOAuthCallbackCapture } from "../src/auth/oauth-callback-capture.ts";

void test("OAuth callback deadline aborts a pending fetch and returns no capture", async (t) => {
  const { settings } = initAppStorage("oauth-callback-deadline");
  await settings.set("proxy.enabled", true);
  await settings.set("proxy.url", "https://localhost:3003");
  t.mock.timers.enable({ apis: ["setTimeout", "Date"], now: 100_000 });
  const originalFetch = globalThis.fetch;
  const started = authDeferred<void>();
  const cleanup = new AbortController();
  let aborted = false;
  globalThis.fetch = (_input, init) => {
    const signal = init?.signal;
    started.resolve();
    return new Promise((_resolve, reject) => {
      signal?.addEventListener("abort", () => {
        aborted = true;
        reject(signal.reason instanceof Error ? signal.reason : new Error("Aborted"));
      }, { once: true });
    });
  };
  t.after(() => { cleanup.abort(); globalThis.fetch = originalFetch; });
  const poll = pollOAuthCallbackCapture({
    providerId: "anthropic", state: "fixture-state", timeoutMs: 100, signal: cleanup.signal,
  });
  void poll.catch(() => {});
  await started.promise;
  t.mock.timers.tick(100);
  assert.equal(aborted, true, "capture deadline must abort the in-flight fetch");
  assert.equal(await poll, null);
});
