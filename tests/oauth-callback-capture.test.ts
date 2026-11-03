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

void test("OAuth callback polling retries pending captures and returns the matching callback", async (t) => {
  const { settings } = initAppStorage("oauth-callback-ready");
  await settings.set("proxy.enabled", true);
  await settings.set("proxy.url", "https://localhost:3003");
  const originalFetch = globalThis.fetch;
  const requests: string[] = [];
  const capture = {
    status: "ready", providerId: "anthropic", state: "fixture-state",
    code: "fixture-code", url: "https://example.test/callback", receivedAt: 123,
  };
  globalThis.fetch = (input) => {
    requests.push(typeof input === "string" ? input : input instanceof URL ? input.href : input.url);
    return Promise.resolve(Response.json(requests.length === 1 ? { status: "pending" } : capture));
  };
  t.after(() => { globalThis.fetch = originalFetch; });
  assert.deepEqual(await pollOAuthCallbackCapture({
    providerId: "anthropic", state: "fixture-state", intervalMs: 0,
  }), capture);
  assert.deepEqual(requests, [
    "https://localhost:3003/oauth/callback/anthropic?state=fixture-state",
    "https://localhost:3003/oauth/callback/anthropic?state=fixture-state",
  ]);
});

for (const { name, response } of [
  { name: "a different state", response: { status: "ready", providerId: "anthropic", state: "another-state", code: "fixture-code", url: "fixture-url", receivedAt: 123 } },
  { name: "an incomplete capture", response: { status: "ready", providerId: "anthropic", state: "fixture-state" } },
  { name: "an unrecognized status", response: { status: "unrecognized" } },
]) {
  void test(`OAuth callback polling rejects ${name}`, async (t) => {
    const { settings } = initAppStorage("oauth-callback-invalid");
    await settings.set("proxy.enabled", true);
    await settings.set("proxy.url", "https://localhost:3003");
    const originalFetch = globalThis.fetch;
    globalThis.fetch = () => Promise.resolve(Response.json(response));
    t.after(() => { globalThis.fetch = originalFetch; });
    assert.equal(await pollOAuthCallbackCapture({ providerId: "anthropic", state: "fixture-state" }), null);
  });
}

void test("OAuth callback cancellation interrupts the polling interval without another request", async (t) => {
  const { settings } = initAppStorage("oauth-callback-cancel");
  await settings.set("proxy.enabled", true);
  await settings.set("proxy.url", "https://localhost:3003");
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const originalFetch = globalThis.fetch;
  const started = authDeferred<void>();
  let requests = 0;
  globalThis.fetch = () => {
    requests += 1;
    started.resolve();
    return Promise.resolve(Response.json({ status: "pending" }));
  };
  t.after(() => { globalThis.fetch = originalFetch; });
  const abort = new AbortController();
  const poll = pollOAuthCallbackCapture({ providerId: "anthropic", state: "fixture-state", signal: abort.signal });
  const rejected = assert.rejects(poll, { name: "AbortError" });
  await started.promise;
  await new Promise<void>((resolve) => setImmediate(resolve));
  abort.abort();
  await rejected;
  t.mock.timers.tick(1000);
  assert.equal(requests, 1);
});
