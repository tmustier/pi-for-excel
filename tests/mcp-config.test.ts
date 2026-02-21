import assert from "node:assert/strict";
import { test } from "node:test";

import { CONNECTION_STORE_KEY } from "../src/connections/store.ts";
import {
  createMcpServerConfig,
  loadMcpServers,
  MCP_SERVER_TOKENS_CONNECTION_ID,
  MCP_SERVERS_SETTING_KEY,
  migrateLegacyMcpTokensToConnectionStore,
  saveMcpServers,
  validateMcpServerUrl,
} from "../src/tools/mcp-config.ts";

class MemorySettingsStore {
  protected readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  peek(key: string): unknown {
    return this.values.get(key);
  }
}

class FailingConnectionStoreSettings extends MemorySettingsStore {
  private failConnectionStoreWrite = true;

  override set(key: string, value: unknown): Promise<void> {
    if (this.failConnectionStoreWrite && key === CONNECTION_STORE_KEY) {
      this.failConnectionStoreWrite = false;
      return Promise.reject(new Error("simulated connection store failure"));
    }

    return super.set(key, value);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function readConnectionStoreTokenMap(settings: MemorySettingsStore): Record<string, string> | undefined {
  const rawStore = settings.peek(CONNECTION_STORE_KEY);
  if (!isRecord(rawStore)) return undefined;

  const rawItems = rawStore.items;
  if (!isRecord(rawItems)) return undefined;

  const rawRecord = rawItems[MCP_SERVER_TOKENS_CONNECTION_ID];
  if (!isRecord(rawRecord)) return undefined;

  const rawSecrets = rawRecord.secrets;
  if (!isRecord(rawSecrets)) return undefined;

  const tokens: Record<string, string> = {};
  for (const [serverId, token] of Object.entries(rawSecrets)) {
    if (typeof token !== "string") continue;
    tokens[serverId] = token;
  }

  return tokens;
}

function readStoredServerEntries(settings: MemorySettingsStore): Array<Record<string, unknown>> {
  const rawDoc = settings.peek(MCP_SERVERS_SETTING_KEY);
  if (!isRecord(rawDoc)) return [];

  const rawServers = rawDoc.servers;
  if (!Array.isArray(rawServers)) return [];

  return rawServers.filter(isRecord);
}

void test("validateMcpServerUrl accepts http(s) and rejects invalid schemes", () => {
  assert.equal(validateMcpServerUrl("https://example.com/mcp/"), "https://example.com/mcp");
  assert.equal(validateMcpServerUrl("http://localhost:4010"), "http://localhost:4010");
  assert.throws(() => validateMcpServerUrl("ftp://example.com"), /must use http:\/\//);
});

void test("mcp config store round-trips normalized server entries", async () => {
  const settings = new MemorySettingsStore();

  const first = createMcpServerConfig({
    name: "local",
    url: "https://localhost:4010/mcp",
    token: "secret",
  });

  await saveMcpServers(settings, [first]);
  const loaded = await loadMcpServers(settings);

  assert.equal(loaded.length, 1);
  assert.equal(loaded[0].name, "local");
  assert.equal(loaded[0].url, "https://localhost:4010/mcp");
  assert.equal(loaded[0].token, "secret");
  assert.equal(loaded[0].enabled, true);
});

void test("saveMcpServers stores bearer tokens in connection store", async () => {
  const settings = new MemorySettingsStore();

  const first = createMcpServerConfig({
    name: "local",
    url: "https://localhost:4010/mcp",
    token: "secret-token",
  });

  await saveMcpServers(settings, [first]);

  const storedServers = readStoredServerEntries(settings);
  assert.equal(storedServers.length, 1);
  assert.equal(storedServers[0].token, undefined);

  const tokens = readConnectionStoreTokenMap(settings);
  assert.ok(tokens);
  assert.equal(tokens[first.id], "secret-token");
});

void test("saveMcpServers does not strip legacy tokens when token-store write fails", async () => {
  const settings = new FailingConnectionStoreSettings();

  await settings.set(MCP_SERVERS_SETTING_KEY, {
    version: 1,
    servers: [
      {
        id: "mcp-local",
        name: "local",
        url: "https://localhost:4010/mcp",
        enabled: true,
        token: "legacy-token",
      },
    ],
  });

  await assert.rejects(
    saveMcpServers(settings, [{
      id: "mcp-local",
      name: "local",
      url: "https://localhost:4010/mcp",
      enabled: true,
      token: "new-token",
    }]),
    /simulated connection store failure/,
  );

  const storedServers = readStoredServerEntries(settings);
  assert.equal(storedServers.length, 1);
  assert.equal(storedServers[0].token, "legacy-token");

  const tokens = readConnectionStoreTokenMap(settings);
  assert.equal(tokens, undefined);
});

void test("loadMcpServers falls back to legacy token when connection store token is absent", async () => {
  const settings = new MemorySettingsStore();

  await settings.set(MCP_SERVERS_SETTING_KEY, {
    version: 1,
    servers: [
      {
        id: "mcp-local",
        name: "local",
        url: "https://localhost:4010/mcp",
        enabled: true,
        token: "legacy-token",
      },
    ],
  });

  const loaded = await loadMcpServers(settings);

  assert.equal(loaded.length, 1);
  assert.equal(loaded[0].token, "legacy-token");
});

void test("legacy MCP tokens migrate into connection store", async () => {
  const settings = new MemorySettingsStore();

  await settings.set(MCP_SERVERS_SETTING_KEY, {
    version: 1,
    servers: [
      {
        id: "mcp-local",
        name: "local",
        url: "https://localhost:4010/mcp",
        enabled: true,
        token: "legacy-token",
      },
    ],
  });

  const migrated = await migrateLegacyMcpTokensToConnectionStore(settings);
  assert.equal(migrated, true);

  const tokens = readConnectionStoreTokenMap(settings);
  assert.ok(tokens);
  assert.equal(tokens["mcp-local"], "legacy-token");

  const storedServers = readStoredServerEntries(settings);
  assert.equal(storedServers.length, 1);
  assert.equal(storedServers[0].token, undefined);
});

void test("legacy MCP migration does not overwrite existing connection-store tokens", async () => {
  const settings = new MemorySettingsStore();

  await settings.set(CONNECTION_STORE_KEY, {
    version: 1,
    items: {
      [MCP_SERVER_TOKENS_CONNECTION_ID]: {
        status: "connected",
        secrets: {
          "mcp-local": "new-token",
        },
      },
    },
  });

  await settings.set(MCP_SERVERS_SETTING_KEY, {
    version: 1,
    servers: [
      {
        id: "mcp-local",
        name: "local",
        url: "https://localhost:4010/mcp",
        enabled: true,
        token: "legacy-token",
      },
    ],
  });

  const migrated = await migrateLegacyMcpTokensToConnectionStore(settings);
  assert.equal(migrated, true);

  const loaded = await loadMcpServers(settings);
  assert.equal(loaded.length, 1);
  assert.equal(loaded[0].token, "new-token");
});
