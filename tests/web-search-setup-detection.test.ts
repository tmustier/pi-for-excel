import assert from "node:assert/strict";
import { test } from "node:test";

import type { WebSearchDetails } from "../src/tools/tool-details.ts";
import {
  detectWebSearchSetupContext,
} from "../src/tools/web-search-setup-detection.ts";
import {
  saveWebSearchApiKey,
  saveWebSearchProvider,
  type WebSearchConfigStore,
} from "../src/tools/web-search-config.ts";

class MemorySettingsStore implements WebSearchConfigStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    const value = this.values.has(key) ? this.values.get(key) ?? null : null;
    return Promise.resolve(value);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  delete(key: string): Promise<void> {
    this.values.delete(key);
    return Promise.resolve();
  }
}

function createFailedDetails(overrides?: Partial<WebSearchDetails>): WebSearchDetails {
  return {
    kind: "web_search",
    ok: false,
    provider: "jina",
    query: "latest rates",
    sentQuery: "latest rates",
    maxResults: 5,
    ...overrides,
  };
}

void test("detectWebSearchSetupContext returns needs_key and skips probing when proxy is not enabled", async () => {
  const settings = new MemorySettingsStore();
  let probeCalls = 0;

  const context = await detectWebSearchSetupContext(
    createFailedDetails(),
    settings,
    {
      isDev: false,
      probeProxyReachability: () => {
        probeCalls += 1;
        return Promise.resolve(true);
      },
    },
  );

  assert.equal(context.mode.type, "needs_key");
  assert.equal(context.provider, "jina");
  assert.equal(probeCalls, 0);
});

void test("detectWebSearchSetupContext returns needs_both when proxy is enabled but unreachable", async () => {
  const settings = new MemorySettingsStore();
  await settings.set("proxy.enabled", true);

  const context = await detectWebSearchSetupContext(
    createFailedDetails(),
    settings,
    {
      isDev: false,
      probeProxyReachability: () => Promise.resolve(false),
    },
  );

  assert.equal(context.mode.type, "needs_both");
});

void test("detectWebSearchSetupContext returns wrong_provider when another provider key is configured", async () => {
  const settings = new MemorySettingsStore();
  await saveWebSearchProvider(settings, "serper");
  await saveWebSearchApiKey(settings, "brave", "br-key-1234567890");

  const context = await detectWebSearchSetupContext(
    createFailedDetails({ provider: "serper" }),
    settings,
    {
      isDev: false,
      probeProxyReachability: () => Promise.resolve(true),
    },
  );

  assert.equal(context.mode.type, "wrong_provider");
  if (context.mode.type !== "wrong_provider") {
    throw new Error("Expected wrong_provider mode");
  }

  assert.equal(context.mode.availableProvider, "brave");
});

void test("detectWebSearchSetupContext returns needs_proxy when result marks proxyDown", async () => {
  const settings = new MemorySettingsStore();
  await saveWebSearchApiKey(settings, "jina", "jina_1234567890abcdef");

  let probeCalls = 0;

  const context = await detectWebSearchSetupContext(
    createFailedDetails({ proxyDown: true }),
    settings,
    {
      isDev: false,
      probeProxyReachability: () => {
        probeCalls += 1;
        return Promise.resolve(true);
      },
    },
  );

  assert.equal(context.mode.type, "needs_proxy");
  assert.equal(probeCalls, 0);
});

void test("detectWebSearchSetupContext skips proxy probe in dev mode", async () => {
  const settings = new MemorySettingsStore();

  let probeCalls = 0;

  const context = await detectWebSearchSetupContext(
    createFailedDetails(),
    settings,
    {
      isDev: true,
      probeProxyReachability: () => {
        probeCalls += 1;
        return Promise.resolve(false);
      },
    },
  );

  assert.equal(context.mode.type, "needs_key");
  assert.equal(probeCalls, 0);
});
