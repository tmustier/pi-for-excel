import assert from "node:assert/strict";
import test from "node:test";

import { restoreCredentials } from "../src/auth/restore.ts";
import { ProviderKeysStore } from "../src/storage/local/provider-keys-store.ts";
import { SettingsStore } from "../src/storage/local/settings-store.ts";

class TestProviderKeysStore extends ProviderKeysStore {
  readonly values = new Map<string, string>();
  failProvider: string | undefined;

  override set(provider: string, key: string): Promise<void> {
    if (provider === this.failProvider) {
      this.failProvider = undefined;
      return Promise.reject(new Error("write failed"));
    }
    this.values.set(provider, key);
    return Promise.resolve();
  }

  override delete(provider: string): Promise<void> {
    this.values.delete(provider);
    return Promise.resolve();
  }
}

class TestSettingsStore extends SettingsStore {
  readonly values = new Map<string, DynamicValue>();

  override get<T = DynamicValue>(key: string): Promise<T | null> {
    const value = this.values.get(key);
    // Test fixtures control every stored value and callers validate OAuth payloads.
    return Promise.resolve(value === undefined ? null : value as T);
  }

  override set<T = DynamicValue>(key: string, value: T): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  override delete(key: string): Promise<void> {
    this.values.delete(key);
    return Promise.resolve();
  }
}

const validBrowserGrant = {
  access: "browser-access-fixture",
  refresh: "browser-refresh-fixture",
  expires: Date.now() + 60_000,
};

async function restoreBrowserAfterDevPayload(
  payload: DynamicValue,
  providerKeys = new TestProviderKeysStore(),
): Promise<TestProviderKeysStore> {
  const settings = new TestSettingsStore();
  settings.values.set("oauth.anthropic", validBrowserGrant);
  await restoreCredentials(
    providerKeys,
    settings,
    () => Promise.resolve(new Response(JSON.stringify(payload))),
  );
  return providerKeys;
}

void test("empty and unrelated dev auth fall back to a valid browser grant", async () => {
  for (const payload of [
    {},
    { unrelated: { type: "other" } },
    { anthropic: { ...validBrowserGrant, type: "oauth", access: "" } },
  ]) {
    const providerKeys = await restoreBrowserAfterDevPayload(payload);
    assert.equal(providerKeys.values.get("anthropic"), validBrowserGrant.access);
  }
});

void test("failed dev credential write falls back to a valid browser grant", async () => {
  const providerKeys = new TestProviderKeysStore();
  providerKeys.failProvider = "anthropic";
  const restored = await restoreBrowserAfterDevPayload(
    { anthropic: { type: "api_key", key: "dev-fixture" } },
    providerKeys,
  );
  assert.equal(restored.values.get("anthropic"), validBrowserGrant.access);
});

void test("dev credentials take precedence across provider aliases", async () => {
  const providerKeys = new TestProviderKeysStore();
  const settings = new TestSettingsStore();
  settings.values.set("oauth.google-gemini-cli", validBrowserGrant);

  await restoreCredentials(
    providerKeys,
    settings,
    () => Promise.resolve(new Response(JSON.stringify({
      "gemini-cli": { type: "api_key", key: "dev-alias-fixture" },
    }))),
  );
  assert.equal(providerKeys.values.get("google-gemini-cli"), "dev-alias-fixture");
});
