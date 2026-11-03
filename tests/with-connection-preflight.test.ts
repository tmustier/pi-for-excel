import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import { ConnectionManager } from "../src/connections/manager.ts";
import type { ConnectionAwareAgentTool } from "../src/tools/connection-requirements.ts";
import { withConnectionPreflight } from "../src/tools/with-connection-preflight.ts";

void test("a late authentication failure cannot invalidate replacement connection credentials", async () => {
  const values = new Map<string, unknown>();
  const settings = {
    get: (key: string): Promise<unknown> => Promise.resolve(values.get(key)),
    set: (key: string, value: unknown): Promise<void> => { values.set(key, value); return Promise.resolve(); },
  };
  const manager = new ConnectionManager({ settings });
  const id = "ext.vendor.api";
  manager.registerDefinition("vendor", {
    id, title: "Vendor", capability: "lookup", authKind: "api_key",
    secretFields: [{ id: "key", label: "Key", required: true }],
  });
  await manager.setSecrets("vendor", id, { key: "original-key" });
  let started: () => void = () => {};
  const entered = new Promise<void>((resolve) => { started = resolve; });
  let rejectRequest: (error: Error) => void = () => {};
  const request = new Promise<never>((_resolve, reject) => { rejectRequest = reject; });
  const tool: ConnectionAwareAgentTool = {
    name: "lookup", label: "Lookup", description: "Lookup", parameters: Type.Object({}), requiresConnection: id,
    execute: () => { started(); return request; },
  };
  const wrapped = withConnectionPreflight([tool], { connectionManager: manager })[0];
  assert.ok(wrapped);
  const pending = wrapped.execute("lookup", {});
  await entered;
  await manager.updateSecretsFromHost(id, { key: "replacement-key" });
  rejectRequest(new Error("401 unauthorized"));
  const result = await pending;
  assert.ok(result.details);
  assert.equal((await manager.getSnapshot(id))?.status, "connected");
  assert.deepEqual(await manager.getSecretsForOwner("vendor", id), { key: "replacement-key" });
});
