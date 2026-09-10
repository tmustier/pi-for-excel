import assert from "node:assert/strict";
import { test } from "node:test";

import { createFetchPageTool } from "../src/tools/fetch-page.ts";
import { AppStorage, setAppStorage } from "../src/storage/local/app-storage.ts";
import { CustomProvidersStore } from "../src/storage/local/custom-providers-store.ts";
import { IndexedDBStorageBackend } from "../src/storage/local/indexeddb-storage-backend.ts";
import { ModelCatalogsStore } from "../src/storage/local/model-catalogs-store.ts";
import { ProviderKeysStore } from "../src/storage/local/provider-keys-store.ts";
import { SessionsStore } from "../src/storage/local/sessions-store.ts";
import { SettingsStore } from "../src/storage/local/settings-store.ts";

class ProxySettingsStore extends SettingsStore {
  private readonly values: ReadonlyMap<string, unknown>;

  constructor(values: ReadonlyMap<string, unknown>) {
    super();
    this.values = values;
  }

  override get<T = unknown>(key: string): Promise<T | null> {
    const value = this.values.get(key);
    // The fixture values are controlled at the settings storage boundary.
    return Promise.resolve(value === undefined ? null : value as T);
  }
}

function installProxySettings(values: ReadonlyMap<string, unknown>): void {
  setAppStorage(new AppStorage(
    new ProxySettingsStore(values),
    new ProviderKeysStore(),
    new SessionsStore(),
    new CustomProvidersStore(),
    new ModelCatalogsStore(),
    new IndexedDBStorageBackend({ dbName: "tools-fetch-page-test", version: 1, stores: [] }),
  ));
}

void test("fetch_page uses the default local proxy when proxy is enabled without a URL", async () => {
  installProxySettings(new Map([["proxy.enabled", true]]));
  let requestedUrl = "";
  const tool = createFetchPageTool({
    executeFetch: (url) => {
      requestedUrl = url;
      return Promise.resolve({
        status: 200,
        ok: true,
        contentType: "text/plain",
        body: "Proxy response",
      });
    },
    now: () => 101_000,
  });

  const result = await tool.execute("call-proxy", { url: "https://example.com/resource?q=1" });

  assert.equal(
    requestedUrl,
    "https://localhost:3003/?url=https%3A%2F%2Fexample.com%2Fresource%3Fq%3D1",
  );
  assert.equal(result.details.proxied, true);
  assert.equal(result.details.proxyBaseUrl, "https://localhost:3003");
});

void test("fetch_page ignores a configured proxy URL when proxy is disabled", async () => {
  installProxySettings(new Map<string, unknown>([
    ["proxy.enabled", false],
    ["proxy.url", "https://localhost:3004"],
  ]));
  let requestedUrl = "";
  const tool = createFetchPageTool({
    executeFetch: (url) => {
      requestedUrl = url;
      return Promise.resolve({
        status: 200,
        ok: true,
        contentType: "text/plain",
        body: "Direct response",
      });
    },
    now: () => 103_000,
  });

  const result = await tool.execute("call-direct", { url: "https://example.org/resource" });

  assert.equal(requestedUrl, "https://example.org/resource");
  assert.equal(result.details.proxied, false);
  assert.equal(result.details.proxyBaseUrl, undefined);
});
