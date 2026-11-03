import { authDeferred } from "./fixtures/auth-deferred.ts";
import assert from "node:assert/strict";
import { test } from "node:test";

void test("fetch interceptor installation is idempotent", async (t) => {
  const previous = Reflect.get(globalThis, "window");
  const host = { fetch: (_input: RequestInfo | URL) => Promise.resolve(new Response()), location: { protocol: "https:" } };
  Reflect.set(globalThis, "window", host);
  t.after(() => { Reflect.set(globalThis, "window", previous); });
  const path = "../src/auth/cors-proxy.ts?idempotent";
  const { installFetchInterceptor } = await import(path) as typeof import("../src/auth/cors-proxy.ts");
  installFetchInterceptor();
  const installed = host.fetch;
  installFetchInterceptor();
  assert.equal(host.fetch, installed, "second install must keep the original interceptor");
  await host.fetch("https://example.test");
});

void test("concurrent OAuth fetches wait for one proxy settings read and invalid URLs never poison the cache", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: 100_000 });
  const previous = Reflect.get(globalThis, "window");
  const requests: string[] = [];
  const host = {
    fetch: (input: RequestInfo | URL) => {
      requests.push(typeof input === "string" ? input : input instanceof URL ? input.href : input.url);
      return Promise.resolve(new Response());
    },
    location: { protocol: "https:" },
  };
  Reflect.set(globalThis, "window", host);
  t.after(() => { Reflect.set(globalThis, "window", previous); });
  const path = "../src/auth/cors-proxy.ts?settings";
  const { installFetchInterceptor } = await import(path) as typeof import("../src/auth/cors-proxy.ts");
  const read = authDeferred<void>();
  let reads = 0;
  let proxyUrl = "https://localhost:3003";
  installFetchInterceptor({
    async get(key) {
      reads += 1;
      await read.promise;
      return key === "proxy.enabled" ? true : proxyUrl;
    },
  });
  const endpoint = "https://auth.openai.com/oauth/token";
  const first = host.fetch(endpoint);
  const second = host.fetch(endpoint);
  assert.equal(requests.length, 0, "concurrent requests must wait for proxy settings");
  read.resolve();
  await Promise.all([first, second]);
  assert.equal(reads, 2, "one settings snapshot reads enabled and URL once each");
  assert.ok(requests.every((url) => url.startsWith("https://localhost:3003/?url=")));
  t.mock.timers.tick(3001);
  proxyUrl = "not-a-url";
  await assert.rejects(host.fetch(endpoint), /proxy/i);
  await assert.rejects(host.fetch(endpoint), /proxy/i);
  assert.equal(requests.length, 2, "invalid settings must not reuse the previous proxy URL");
});

void test("an invalid proxy URL never publishes an enabled cache with the previous URL", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: 100_000 });
  const previous = Reflect.get(globalThis, "window");
  const host = {
    fetch: (_input: RequestInfo | URL) => Promise.resolve(new Response()),
    location: { protocol: "https:" },
  };
  Reflect.set(globalThis, "window", host);
  t.after(() => { Reflect.set(globalThis, "window", previous); });
  const path = "../src/auth/cors-proxy.ts?invalid";
  const { installFetchInterceptor } = await import(path) as typeof import("../src/auth/cors-proxy.ts");
  let proxyUrl = "https://localhost:3003";
  installFetchInterceptor({ get: (key) => Promise.resolve(key === "proxy.enabled" ? true : proxyUrl) });
  const endpoint = "https://auth.openai.com/oauth/token";
  await host.fetch(endpoint);
  t.mock.timers.tick(3001);
  proxyUrl = "not-a-url";
  await assert.rejects(host.fetch(endpoint), /proxy/i);
  await assert.rejects(host.fetch(endpoint), /proxy/i, "invalid settings must not reuse an earlier validated URL");
});
