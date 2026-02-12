import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createMcpServerConfig,
  loadMcpServers,
  saveMcpServers,
  validateMcpServerUrl,
} from "../src/tools/mcp-config.ts";

class MemorySettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
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
