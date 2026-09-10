import assert from "node:assert/strict";
import test from "node:test";

import {
  CONNECTION_STORE_KEY,
  loadConnectionStoreDocument,
  saveConnectionStoreDocument,
} from "../src/connections/store.ts";
import {
  SKILL_ACTIVATION_STORAGE_KEY,
  loadDisabledSkillNamesFromSettings,
  setSkillEnabledInSettings,
} from "../src/skills/activation-store.ts";
import {
  EXTERNAL_TOOLS_ENABLED_SETTING_KEY,
  getExternalToolsEnabled,
  getSessionIntegrationIds,
  sessionIntegrationsKey,
} from "../src/integrations/store.ts";
import {
  getExtensionStorageValue,
  setExtensionStorageValue,
} from "../src/extensions/storage-store.ts";
import {
  EXTENSIONS_REGISTRY_STORAGE_KEY,
  loadStoredExtensions,
} from "../src/extensions/store.ts";
import {
  DEFAULT_WEB_SEARCH_PROVIDER,
  loadWebSearchProviderConfig,
  saveWebSearchApiKey,
} from "../src/tools/web-search-config.ts";
import { loadMcpServers } from "../src/tools/mcp-config.ts";
import { getEnabledProxyBaseUrl } from "../src/tools/external-fetch.ts";
import { readCorsProxySettings } from "../src/auth/cors-proxy.ts";
import {
  loadOAuthCredentials,
  saveOAuthCredentials,
} from "../src/auth/oauth-storage.ts";
import {
  getSessionWorkbookId,
  linkSessionToWorkbook,
  sessionWorkbookKey,
} from "../src/workbook/session-association.ts";
import { readBridgeUrls } from "../src/commands/builtins/extensions-hub-connections.ts";

class MemorySettingsStore {
  readonly values = new Map<string, unknown>();

  async get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.get(key) ?? null);
  }

  async set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    await Promise.resolve();
  }

  async delete(key: string): Promise<void> {
    this.values.delete(key);
    await Promise.resolve();
  }
}

class ThrowingSettingsStore extends MemorySettingsStore {
  private readsFail = true;

  allowReads(): void {
    this.readsFail = false;
  }

  override async get(key: string): Promise<unknown> {
    if (this.readsFail) {
      return Promise.reject(new Error("seeded read failure"));
    }
    return super.get(key);
  }
}

void test("feature-owned readers preserve documented corrupt-value and read-failure behavior", async () => {
  const corrupt = new MemorySettingsStore();
  corrupt.values.set(SKILL_ACTIVATION_STORAGE_KEY, { version: "bad", disabledNames: 3 });
  corrupt.values.set(EXTERNAL_TOOLS_ENABLED_SETTING_KEY, { enabled: true });
  corrupt.values.set(sessionIntegrationsKey("session-a"), { ids: ["web_search"] });
  corrupt.values.set(CONNECTION_STORE_KEY, { version: 1, items: "bad" });
  corrupt.values.set(EXTENSIONS_REGISTRY_STORAGE_KEY, { version: 2, items: "bad" });
  corrupt.values.set("extensions.registry.v1", { version: 1, items: "bad" });
  corrupt.values.set("extensions.storage.v1", { version: 1, items: "bad" });
  corrupt.values.set("mcp.servers.v1", { version: 1, servers: "bad" });
  corrupt.values.set("proxy.enabled", { enabled: true });
  corrupt.values.set("proxy.url", 42);
  corrupt.values.set("oauth.github", { refresh: 1, access: "token", expires: "later" });
  corrupt.values.set(sessionWorkbookKey("session-a"), { workbookId: "book-a" });
  corrupt.values.set("python.bridge.url", 42);
  corrupt.values.set("tmux.bridge.url", false);

  assert.deepEqual([...await loadDisabledSkillNamesFromSettings(corrupt)], []);
  assert.equal(await getExternalToolsEnabled(corrupt), false);
  assert.deepEqual(await getSessionIntegrationIds(corrupt, "session-a", ["web_search"]), []);
  assert.deepEqual(await loadConnectionStoreDocument(corrupt), {});
  assert.equal((await loadStoredExtensions(corrupt))[0]?.id, "builtin.snake");
  assert.equal(await getExtensionStorageValue(corrupt, "extension-a", "key"), undefined);
  assert.deepEqual(await loadMcpServers(corrupt), []);
  assert.equal((await loadWebSearchProviderConfig(corrupt)).provider, DEFAULT_WEB_SEARCH_PROVIDER);
  assert.equal(await getEnabledProxyBaseUrl(corrupt), undefined);
  assert.deepEqual(await readCorsProxySettings(corrupt), { enabled: true, url: "" });
  assert.equal(await loadOAuthCredentials(corrupt, "github"), null);
  assert.equal(await getSessionWorkbookId(corrupt, "session-a"), null);
  assert.deepEqual(await readBridgeUrls(corrupt), { pythonUrl: "", tmuxUrl: "" });

  const failing = new ThrowingSettingsStore();
  await assert.rejects(loadDisabledSkillNamesFromSettings(failing), /seeded read failure/u);
  await assert.rejects(getExternalToolsEnabled(failing), /seeded read failure/u);
  assert.deepEqual(await loadConnectionStoreDocument(failing), {});
  assert.deepEqual(await loadMcpServers(failing), []);
  assert.equal((await loadWebSearchProviderConfig(failing)).provider, DEFAULT_WEB_SEARCH_PROVIDER);
  assert.equal(await getEnabledProxyBaseUrl(failing), undefined);
  assert.deepEqual(await readCorsProxySettings(failing), { enabled: false, url: "" });
  assert.equal(await loadOAuthCredentials(failing, "github"), null);
  await assert.rejects(getSessionWorkbookId(failing, "session-a"), /seeded read failure/u);
  assert.deepEqual(await readBridgeUrls(failing), { pythonUrl: "", tmuxUrl: "" });
});

void test("CORS proxy settings preserve baseline truthiness for legacy flags", async () => {
  const settings = new MemorySettingsStore();
  settings.values.set("proxy.url", "https://proxy.example.com");

  for (const value of ["false", "0", "no", { enabled: false }]) {
    settings.values.set("proxy.enabled", value);
    assert.equal((await readCorsProxySettings(settings)).enabled, true);
  }

  for (const value of [false, 0, "", null]) {
    settings.values.set("proxy.enabled", value);
    assert.equal((await readCorsProxySettings(settings)).enabled, false);
  }
});

void test("valid feature settings round-trip through unchanged public keys and formats", async () => {
  const settings = new MemorySettingsStore();

  await setSkillEnabledInSettings({ settings, name: "Analysis", enabled: false });
  assert.deepEqual([...await loadDisabledSkillNamesFromSettings(settings)], ["analysis"]);

  await saveConnectionStoreDocument(settings, {
    sibling: { status: "connected", secrets: { token: "keep" } },
  });
  assert.equal((await loadConnectionStoreDocument(settings)).sibling?.secrets?.token, "keep");

  await saveWebSearchApiKey(settings, "serper", "serper-key");
  const connectionItems = await loadConnectionStoreDocument(settings);
  assert.equal(connectionItems.sibling?.secrets?.token, "keep");
  assert.equal(connectionItems["builtin.web_search.providers"]?.secrets?.serper_api_key, "serper-key");

  await setExtensionStorageValue(settings, "extension-a", "first", "one");
  await setExtensionStorageValue(settings, "extension-a", "second", "two");
  await setExtensionStorageValue(settings, "extension-b", "sibling", "keep");
  assert.equal(await getExtensionStorageValue(settings, "extension-a", "first"), "one");
  assert.equal(await getExtensionStorageValue(settings, "extension-a", "second"), "two");
  assert.equal(await getExtensionStorageValue(settings, "extension-b", "sibling"), "keep");

  await saveOAuthCredentials(settings, "github", {
    refresh: "refresh-token",
    access: "access-token",
    expires: 123,
  });
  assert.deepEqual(await loadOAuthCredentials(settings, "github"), {
    refresh: "refresh-token",
    access: "access-token",
    expires: 123,
  });

  await linkSessionToWorkbook(settings, "session-a", "workbook-a");
  await linkSessionToWorkbook(settings, "session-a", "workbook-b");
  assert.equal(await getSessionWorkbookId(settings, "session-a"), "workbook-a");
});

void test("mutation reads fail closed instead of dropping recoverable sibling records", async () => {
  const settings = new ThrowingSettingsStore();
  settings.values.set(SKILL_ACTIVATION_STORAGE_KEY, {
    version: 1,
    disabledNames: ["existing-disabled-skill"],
  });
  await assert.rejects(
    () => setSkillEnabledInSettings({ settings, name: "new-disabled-skill", enabled: false }),
    /seeded read failure/u,
  );
  settings.allowReads();
  assert.deepEqual(
    [...await loadDisabledSkillNamesFromSettings(settings)],
    ["existing-disabled-skill"],
  );

  const connectionSettings = new ThrowingSettingsStore();
  connectionSettings.values.set(CONNECTION_STORE_KEY, {
    version: 1,
    items: { sibling: { status: "connected", secrets: { token: "keep" } } },
  });

  await assert.rejects(
    () => saveWebSearchApiKey(connectionSettings, "serper", "new-key"),
    /seeded read failure/u,
  );
  connectionSettings.allowReads();
  assert.equal(
    (await loadConnectionStoreDocument(connectionSettings)).sibling?.secrets?.token,
    "keep",
  );

  const extensionSettings = new ThrowingSettingsStore();
  extensionSettings.values.set("extensions.storage.v1", {
    version: 1,
    items: { sibling: { key: "keep" } },
  });
  await assert.rejects(
    () => setExtensionStorageValue(extensionSettings, "extension-a", "key", "value"),
    /seeded read failure/u,
  );
  extensionSettings.allowReads();
  assert.equal(await getExtensionStorageValue(extensionSettings, "sibling", "key"), "keep");
});
