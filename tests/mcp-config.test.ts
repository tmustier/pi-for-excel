import assert from "node:assert/strict";
import { test } from "node:test";

import {
  CONNECTION_STORE_KEY,
  loadConnectionStoreDocument,
} from "../src/connections/store.ts";
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
  protected readonly values = new Map<string, DynamicValue>();

  get(key: string): Promise<DynamicValue> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: DynamicValue): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

}

class TransientReadSettings extends MemorySettingsStore {
  private failNextRead = false;

  armReadFailure(): void {
    this.failNextRead = true;
  }

  override get(key: string): Promise<DynamicValue> {
    if (this.failNextRead) {
      this.failNextRead = false;
      return Promise.reject(new Error("transient MCP read failure"));
    }
    return super.get(key);
  }
}

class FailingConnectionStoreSettings extends MemorySettingsStore {
  private failConnectionStoreWrite = true;

  override set(key: string, value: DynamicValue): Promise<void> {
    if (this.failConnectionStoreWrite && key === CONNECTION_STORE_KEY) {
      this.failConnectionStoreWrite = false;
      return Promise.reject(new Error("simulated connection store failure"));
    }

    return super.set(key, value);
  }
}

class FailingServerSettings extends MemorySettingsStore {
  private failServerDocumentWrite = false;

  armServerDocumentFailure(): void {
    this.failServerDocumentWrite = true;
  }

  override set(key: string, value: DynamicValue): Promise<void> {
    if (this.failServerDocumentWrite && key === MCP_SERVERS_SETTING_KEY) {
      this.failServerDocumentWrite = false;
      return Promise.reject(new Error("simulated mcp.servers write failure"));
    }

    return super.set(key, value);
  }
}

class ConcurrentServerFailureSettings extends MemorySettingsStore {
  private firstServerWriteReject: ((reason?: DynamicValue) => void) | null = null;
  private firstServerWriteStartedResolve: (() => void) | null = null;
  private readonly firstServerWriteStarted: Promise<void>;
  private shouldInterceptServerWrite = false;

  constructor() {
    super();
    this.firstServerWriteStarted = new Promise<void>((resolve) => {
      this.firstServerWriteStartedResolve = resolve;
    });
  }

  armFirstServerWriteFailure(): void {
    this.shouldInterceptServerWrite = true;
  }

  waitForFirstServerWrite(): Promise<void> {
    return this.firstServerWriteStarted;
  }

  failFirstServerWrite(): void {
    const reject = this.firstServerWriteReject;
    if (!reject) {
      throw new Error("First server write is not pending.");
    }

    this.firstServerWriteReject = null;
    reject(new Error("simulated concurrent mcp.servers write failure"));
  }

  override set(key: string, value: DynamicValue): Promise<void> {
    if (key === MCP_SERVERS_SETTING_KEY && this.shouldInterceptServerWrite) {
      this.shouldInterceptServerWrite = false;

      const resolve = this.firstServerWriteStartedResolve;
      if (resolve) {
        this.firstServerWriteStartedResolve = null;
        resolve();
      }

      return new Promise<void>((_resolve, reject) => {
        this.firstServerWriteReject = reject;
      });
    }

    return super.set(key, value);
  }
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

void test("a failed MCP update leaves the previous server available and a retry installs the replacement", async () => {
  const settings = new TransientReadSettings();
  await settings.set(CONNECTION_STORE_KEY, {
    version: 1,
    items: {
      sibling: { status: "connected", secrets: { token: "keep" } },
      [MCP_SERVER_TOKENS_CONNECTION_ID]: {
        status: "connected",
        secrets: { "mcp-existing": "existing-token" },
      },
    },
  });
  await settings.set(MCP_SERVERS_SETTING_KEY, {
    version: 1,
    servers: [{
      id: "mcp-existing",
      name: "existing",
      url: "https://existing.example.com/mcp",
      enabled: true,
    }],
  });
  const nextServers = [{
    id: "mcp-new",
    name: "new",
    url: "https://example.com/mcp",
    enabled: true,
    token: "new-token",
  }];
  settings.armReadFailure();

  await assert.rejects(() => saveMcpServers(settings, nextServers), /transient MCP read failure/u);
  assert.equal((await loadMcpServers(settings))[0]?.token, "existing-token");

  await saveMcpServers(settings, nextServers);
  assert.deepEqual(await loadMcpServers(settings), [{
    id: "mcp-new",
    name: "new",
    url: "https://example.com/mcp",
    enabled: true,
    token: "new-token",
  }]);
  const connections = await loadConnectionStoreDocument(settings);
  assert.equal(connections.sibling?.secrets?.token, "keep");
});

void test("saved MCP credentials remain resolvable after reloading configuration", async () => {
  const settings = new MemorySettingsStore();

  const first = createMcpServerConfig({
    name: "local",
    url: "https://localhost:4010/mcp",
    token: "secret-token",
  });

  await saveMcpServers(settings, [first]);

  const reloaded = await loadMcpServers(settings);
  assert.equal(reloaded[0]?.token, "secret-token");
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

  const reloaded = await loadMcpServers(settings);
  assert.equal(reloaded[0]?.token, "legacy-token");
});

void test("saveMcpServers rolls back token changes when server-document write fails", async () => {
  const settings = new FailingServerSettings();

  await settings.set(CONNECTION_STORE_KEY, {
    version: 1,
    items: {
      [MCP_SERVER_TOKENS_CONNECTION_ID]: {
        status: "connected",
        secrets: {
          "mcp-local": "existing-token",
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
      },
    ],
  });

  settings.armServerDocumentFailure();

  await assert.rejects(
    saveMcpServers(settings, [{
      id: "mcp-local",
      name: "local",
      url: "https://localhost:4010/mcp",
      enabled: true,
      token: "new-token",
    }]),
    /simulated mcp\.servers write failure/,
  );

  const reloaded = await loadMcpServers(settings);
  assert.equal(reloaded[0]?.token, "existing-token");
});

void test("saveMcpServers does not roll back newer token writes from overlapping saves", async () => {
  const settings = new ConcurrentServerFailureSettings();

  await settings.set(CONNECTION_STORE_KEY, {
    version: 1,
    items: {
      [MCP_SERVER_TOKENS_CONNECTION_ID]: {
        status: "connected",
        secrets: {
          "mcp-local": "initial-token",
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
      },
    ],
  });

  settings.armFirstServerWriteFailure();

  const firstSavePromise = saveMcpServers(settings, [{
    id: "mcp-local",
    name: "local",
    url: "https://localhost:4010/mcp",
    enabled: true,
    token: "token-a",
  }]);

  await settings.waitForFirstServerWrite();

  await saveMcpServers(settings, [{
    id: "mcp-local",
    name: "local",
    url: "https://localhost:4010/mcp",
    enabled: true,
    token: "token-b",
  }]);

  settings.failFirstServerWrite();

  await assert.rejects(
    firstSavePromise,
    /simulated concurrent mcp\.servers write failure/,
  );

  const reloaded = await loadMcpServers(settings);
  assert.equal(reloaded[0]?.token, "token-b");
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

void test("legacy MCP migration retries transient reads before stripping tokens", async () => {
  const settings = new TransientReadSettings();
  await settings.set(CONNECTION_STORE_KEY, {
    version: 1,
    items: { sibling: { status: "connected", secrets: { token: "keep" } } },
  });
  await settings.set(MCP_SERVERS_SETTING_KEY, {
    version: 1,
    servers: [{
      id: "mcp-local",
      name: "local",
      url: "https://localhost:4010/mcp",
      enabled: true,
      token: "legacy-token",
    }],
  });
  settings.armReadFailure();

  await assert.rejects(
    () => migrateLegacyMcpTokensToConnectionStore(settings),
    /transient MCP read failure/u,
  );
  assert.equal((await loadMcpServers(settings))[0]?.token, "legacy-token");

  assert.equal(await migrateLegacyMcpTokensToConnectionStore(settings), true);
  assert.equal((await loadMcpServers(settings))[0]?.token, "legacy-token");
  assert.equal((await loadConnectionStoreDocument(settings)).sibling?.secrets?.token, "keep");
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

  assert.equal((await loadMcpServers(settings))[0]?.token, "legacy-token");
  assert.equal(await migrateLegacyMcpTokensToConnectionStore(settings), false);
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
