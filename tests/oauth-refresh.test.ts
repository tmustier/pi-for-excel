import assert from "node:assert/strict";
import test from "node:test";

import type { OAuthCredentials } from "@earendil-works/pi-ai";

import type { BrowserOAuthProvider } from "../src/auth/browser-oauth-types.ts";
import { createOAuthRefresher, type OAuthRefreshLock } from "../src/auth/oauth-refresh.ts";

const NOW = 1_800_000_000_000;

class MemorySettings {
  readonly values = new Map<string, unknown>();
  get(key: string): Promise<unknown> {
    return Promise.resolve(structuredClone(this.values.get(key) ?? null));
  }
  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, structuredClone(value));
    return Promise.resolve();
  }
  delete(key: string): Promise<void> {
    this.values.delete(key);
    return Promise.resolve();
  }
}

class MemoryKeys {
  readonly values = new Map<string, string>();
  set(provider: string, key: string): Promise<void> {
    this.values.set(provider, key);
    return Promise.resolve();
  }
  delete(provider: string): Promise<void> {
    this.values.delete(provider);
    return Promise.resolve();
  }
}

/** A fake Anthropic OAuth provider whose refresh rotates the refresh token, like the real ones. */
function fakeProvider(opts: { fail?: Error } = {}) {
  let refreshCalls = 0;
  const provider = {
    id: "anthropic",
    name: "Fake Anthropic",
    getApiKey: (credentials: OAuthCredentials) => `key:${credentials.access}`,
    refreshToken: async (credentials: OAuthCredentials): Promise<OAuthCredentials> => {
      refreshCalls += 1;
      await new Promise((resolve) => setTimeout(resolve, 5));
      if (opts.fail) throw opts.fail;
      return { access: `access-${refreshCalls}`, refresh: `${credentials.refresh}+`, expires: NOW + 3_600_000 };
    },
  } as unknown as BrowserOAuthProvider;
  return { provider, refreshCalls: () => refreshCalls };
}

/** One in-memory lock shared by several "taskpanes", standing in for navigator.locks. */
function sharedLock(): OAuthRefreshLock {
  let chain: Promise<unknown> = Promise.resolve();
  return <T>(_name: string, run: () => Promise<T>): Promise<T> => {
    const next = chain.then(run, run);
    chain = next.catch(() => {});
    return next;
  };
}

function seed(settings: MemorySettings, expires: number): void {
  settings.values.set("oauth.anthropic", { access: "access-0", refresh: "refresh-0", expires });
}

void test("an expiring token is refreshed once for concurrent requests in one taskpane", async () => {
  const settings = new MemorySettings();
  const keys = new MemoryKeys();
  seed(settings, NOW + 30_000);
  const fake = fakeProvider();
  const refresher = createOAuthRefresher({
    settings, providerKeys: keys, getOAuthProvider: () => fake.provider, now: () => NOW,
  });

  const outcomes = await Promise.all([1, 2, 3].map(() => refresher.ensureFresh("anthropic")));
  assert.deepEqual(outcomes, ["refreshed", "refreshed", "refreshed"]);
  assert.equal(fake.refreshCalls(), 1);
  assert.equal(keys.values.get("anthropic"), "key:access-1");
});

void test("a token that expires while the pane is open is refreshed before the next request", async () => {
  const settings = new MemorySettings();
  const keys = new MemoryKeys();
  // Valid when the pane opens, expires 3.5 minutes later.
  seed(settings, NOW + 210_000);
  let clock = NOW;
  const fake = fakeProvider();
  const refresher = createOAuthRefresher({
    settings, providerKeys: keys, getOAuthProvider: () => fake.provider, now: () => clock,
  });

  assert.equal(await refresher.ensureFresh("anthropic"), "fresh");
  assert.equal(fake.refreshCalls(), 0);

  clock = NOW + 240_000;
  assert.equal(await refresher.ensureFresh("anthropic"), "refreshed");
  assert.equal(fake.refreshCalls(), 1);
  assert.equal(keys.values.get("anthropic"), "key:access-1");
});

void test("two taskpanes refreshing at once rotate the grant once and both keep a working token", async () => {
  const settings = new MemorySettings();
  seed(settings, NOW - 1);
  const fake = fakeProvider();
  const lock = sharedLock();
  const panes = [new MemoryKeys(), new MemoryKeys()].map((keys) => ({
    keys,
    refresher: createOAuthRefresher({
      settings, providerKeys: keys, getOAuthProvider: () => fake.provider, lock, now: () => NOW,
    }),
  }));

  const outcomes = await Promise.all(panes.map((pane) => pane.refresher.ensureFresh("anthropic")));
  assert.deepEqual([...outcomes].sort(), ["fresh", "refreshed"]);
  assert.equal(fake.refreshCalls(), 1);
  for (const pane of panes) assert.equal(pane.keys.values.get("anthropic"), "key:access-1");
  assert.equal((settings.values.get("oauth.anthropic") as OAuthCredentials).refresh, "refresh-0+");
});

void test("a valid token is not refreshed and non-OAuth providers are untouched", async () => {
  const settings = new MemorySettings();
  seed(settings, NOW + 3_600_000);
  const fake = fakeProvider();
  const refresher = createOAuthRefresher({
    settings, providerKeys: new MemoryKeys(), getOAuthProvider: () => fake.provider, now: () => NOW,
  });

  assert.equal(await refresher.ensureFresh("anthropic"), "fresh");
  assert.equal(await refresher.ensureFresh("openai"), "not_oauth");
  assert.equal(fake.refreshCalls(), 0);
});

void test("a failed refresh keeps the stored grant so the next request can retry", async () => {
  const settings = new MemorySettings();
  seed(settings, NOW + 30_000);
  const fake = fakeProvider({ fail: new Error("offline") });
  const refresher = createOAuthRefresher({
    settings, providerKeys: new MemoryKeys(), getOAuthProvider: () => fake.provider, now: () => NOW, warn: () => {},
  });

  assert.equal(await refresher.ensureFresh("anthropic"), "failed");
  assert.equal((settings.values.get("oauth.anthropic") as OAuthCredentials).refresh, "refresh-0");
  assert.equal(await refresher.ensureFresh("anthropic"), "failed");
  assert.equal(fake.refreshCalls(), 2);
});

void test("a failed refresh of an expired token reports the login expired instead of keeping the token", async () => {
  const settings = new MemorySettings();
  seed(settings, NOW - 1);
  const fake = fakeProvider({ fail: new Error("invalid_grant") });
  const refresher = createOAuthRefresher({
    settings, providerKeys: new MemoryKeys(), getOAuthProvider: () => fake.provider, now: () => NOW, warn: () => {},
  });

  assert.equal(await refresher.ensureFresh("anthropic"), "expired");
  // The grant is kept so a later request can retry once the provider answers again.
  assert.equal((settings.values.get("oauth.anthropic") as OAuthCredentials).refresh, "refresh-0");
});
