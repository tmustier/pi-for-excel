import assert from "node:assert/strict";
import { test } from "node:test";

import { ConnectionManager } from "../src/connections/manager.ts";
import type { ConnectionDefinition } from "../src/connections/types.ts";

// ── In-memory settings store ────────────────────────

function createMemorySettings(): {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
  delete(key: string): Promise<void>;
} {
  const data = new Map<string, unknown>();
  return {
    get: (key) => Promise.resolve(data.get(key)),
    set: (key, value) => { data.set(key, value); return Promise.resolve(); },
    delete: (key) => { data.delete(key); return Promise.resolve(); },
  };
}

const APOLLO_DEFINITION: ConnectionDefinition = {
  id: "ext.apollo.apollo",
  title: "Apollo",
  capability: "company enrichment via Apollo API",
  authKind: "api_key",
  secretFields: [
    { id: "apiKey", label: "API key", required: true },
  ],
};

const MULTI_FIELD_DEFINITION: ConnectionDefinition = {
  id: "ext.vendor.vendor",
  title: "Vendor API",
  capability: "procurement data pull",
  authKind: "custom",
  secretFields: [
    { id: "clientId", label: "Client ID", required: true },
    { id: "clientSecret", label: "Client secret", required: true },
    { id: "endpoint", label: "Endpoint URL", required: false },
  ],
};

// ── getDefinition ───────────────────────────────────

void test("getDefinition returns null for unregistered connection", () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  assert.equal(manager.getDefinition("nonexistent"), null);
});

void test("getDefinition returns definition without ownerId", () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  const def = manager.getDefinition("ext.apollo.apollo");
  assert.ok(def);
  assert.equal(def.id, "ext.apollo.apollo");
  assert.equal(def.title, "Apollo");
  assert.equal(def.capability, "company enrichment via Apollo API");
  assert.equal(Reflect.get(def, "ownerId"), undefined);
});

void test("getDefinition returns detached objects", () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  const first = manager.getDefinition("ext.apollo.apollo");
  assert.ok(first);
  first.title = "Changed";
  first.secretFields[0].label = "Changed field";

  const second = manager.getDefinition("ext.apollo.apollo");
  assert.ok(second);
  assert.equal(second.title, "Apollo");
  assert.equal(second.secretFields[0].label, "API key");
});

// ── listDefinitions ─────────────────────────────────

void test("listDefinitions returns empty array when none registered", () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  assert.deepEqual(manager.listDefinitions(), []);
});

void test("listDefinitions returns definitions sorted by title", () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.vendor", MULTI_FIELD_DEFINITION);
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  const defs = manager.listDefinitions();
  assert.equal(defs.length, 2);
  assert.equal(defs[0].title, "Apollo");
  assert.equal(defs[1].title, "Vendor API");
  // No ownerId leaked
  assert.equal(Reflect.get(defs[0], "ownerId"), undefined);
});

void test("listDefinitions returns detached objects", () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  const defs = manager.listDefinitions();
  assert.equal(defs.length, 1);
  defs[0].title = "Changed";
  defs[0].secretFields[0].label = "Changed field";

  const next = manager.listDefinitions();
  assert.equal(next[0].title, "Apollo");
  assert.equal(next[0].secretFields[0].label, "API key");
});

// ── getSecretFieldPresence ──────────────────────────

void test("getSecretFieldPresence returns all false when no secrets stored", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  const presence = await manager.getSecretFieldPresence("ext.apollo.apollo");
  assert.deepEqual(presence, { apiKey: false });
});

void test("getSecretFieldPresence returns true for stored fields", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);
  await manager.setSecrets("ext.apollo", "ext.apollo.apollo", { apiKey: "sk-test-123" });

  const presence = await manager.getSecretFieldPresence("ext.apollo.apollo");
  assert.deepEqual(presence, { apiKey: true });
});

void test("getSecretFieldPresence handles multi-field partial secrets", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.vendor", MULTI_FIELD_DEFINITION);
  await manager.setSecrets("ext.vendor", "ext.vendor.vendor", {
    clientId: "id-123",
    clientSecret: "",
  });

  const presence = await manager.getSecretFieldPresence("ext.vendor.vendor");
  assert.equal(presence.clientId, true);
  assert.equal(presence.clientSecret, false);
  assert.equal(presence.endpoint, false);
});

// ── redactMessageForConnection ──────────────────────

void test("redactMessageForConnection masks stored secrets", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.vendor", MULTI_FIELD_DEFINITION);

  await manager.setSecrets("ext.vendor", "ext.vendor.vendor", {
    clientId: "client-123",
    clientSecret: "top-secret-token",
  });

  const redacted = await manager.redactMessageForConnection(
    "ext.vendor.vendor",
    "403 Forbidden: invalid credential top-secret-token for client-123",
  );

  assert.ok(!redacted.includes("top-secret-token"));
  assert.ok(!redacted.includes("client-123"));
  assert.match(redacted, /••••/);
});

void test("markInvalid redacts stored secret values from failure reasons", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.vendor", MULTI_FIELD_DEFINITION);

  await manager.setSecrets("ext.vendor", "ext.vendor.vendor", {
    clientId: "client-123",
    clientSecret: "top-secret-token",
  });

  await manager.markInvalid(
    "ext.vendor",
    "ext.vendor.vendor",
    "Credential top-secret-token for client-123 is revoked",
  );

  const snapshot = await manager.getSnapshot("ext.vendor.vendor");
  assert.equal(snapshot?.status, "invalid");
  assert.ok(!snapshot?.lastError?.includes("top-secret-token"));
  assert.ok(!snapshot?.lastError?.includes("client-123"));
});

// ── updateSecretsFromHost ───────────────────────────

void test("updateSecretsFromHost merges into empty store", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  await manager.updateSecretsFromHost("ext.apollo.apollo", { apiKey: "sk-new" });

  const snapshot = await manager.getSnapshot("ext.apollo.apollo");
  assert.ok(snapshot);
  assert.equal(snapshot.status, "connected");
  const presence = await manager.getSecretFieldPresence("ext.apollo.apollo");
  assert.equal(presence.apiKey, true);
});

void test("updateSecretsFromHost merges partial patch preserving existing fields", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.vendor", MULTI_FIELD_DEFINITION);

  // Set initial secrets
  await manager.setSecrets("ext.vendor", "ext.vendor.vendor", {
    clientId: "id-old",
    clientSecret: "secret-old",
  });

  // Patch only clientSecret
  await manager.updateSecretsFromHost("ext.vendor.vendor", { clientSecret: "secret-new" });

  const presence = await manager.getSecretFieldPresence("ext.vendor.vendor");
  assert.equal(presence.clientId, true);   // preserved
  assert.equal(presence.clientSecret, true); // updated
});

void test("updateSecretsFromHost ignores empty values in patch", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);
  await manager.setSecrets("ext.apollo", "ext.apollo.apollo", { apiKey: "sk-existing" });

  // Patch with empty string — should be no-op
  await manager.updateSecretsFromHost("ext.apollo.apollo", { apiKey: "" });

  const presence = await manager.getSecretFieldPresence("ext.apollo.apollo");
  assert.equal(presence.apiKey, true); // unchanged
});

void test("updateSecretsFromHost clears error status on save", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);
  await manager.setSecrets("ext.apollo", "ext.apollo.apollo", { apiKey: "sk-old" });
  await manager.markRuntimeAuthFailure("ext.apollo.apollo", { message: "401 unauthorized" });

  const beforeSnapshot = await manager.getSnapshot("ext.apollo.apollo");
  assert.equal(beforeSnapshot?.status, "error");

  await manager.updateSecretsFromHost("ext.apollo.apollo", { apiKey: "sk-fixed" });

  const afterSnapshot = await manager.getSnapshot("ext.apollo.apollo");
  assert.equal(afterSnapshot?.status, "connected");
  assert.equal(afterSnapshot?.lastError, undefined);
});

void test("updateSecretsFromHost rejects unknown secret fields", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  await assert.rejects(
    () => manager.updateSecretsFromHost("ext.apollo.apollo", { bogus: "value" }),
    /Unknown secret field "bogus"/,
  );
});

void test("updateSecretsFromHost notifies listeners", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

  let notified = false;
  manager.subscribe(() => { notified = true; });

  await manager.updateSecretsFromHost("ext.apollo.apollo", { apiKey: "sk-test" });
  assert.ok(notified);
});

// ── clearSecretsFromHost ────────────────────────────

void test("clearSecretsFromHost wipes secrets and sets status to missing", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);
  await manager.setSecrets("ext.apollo", "ext.apollo.apollo", { apiKey: "sk-test" });

  await manager.clearSecretsFromHost("ext.apollo.apollo");

  const snapshot = await manager.getSnapshot("ext.apollo.apollo");
  assert.ok(snapshot);
  assert.equal(snapshot.status, "missing");
  const presence = await manager.getSecretFieldPresence("ext.apollo.apollo");
  assert.equal(presence.apiKey, false);
});

void test("clearSecretsFromHost clears error state", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);
  await manager.setSecrets("ext.apollo", "ext.apollo.apollo", { apiKey: "sk-test" });
  await manager.markRuntimeAuthFailure("ext.apollo.apollo", { message: "401" });

  await manager.clearSecretsFromHost("ext.apollo.apollo");

  const snapshot = await manager.getSnapshot("ext.apollo.apollo");
  assert.ok(snapshot);
  assert.equal(snapshot.status, "missing");
  assert.equal(snapshot.lastError, undefined);
});

void test("clearSecretsFromHost notifies listeners", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);
  await manager.setSecrets("ext.apollo", "ext.apollo.apollo", { apiKey: "sk-test" });

  let notified = false;
  manager.subscribe(() => { notified = true; });

  await manager.clearSecretsFromHost("ext.apollo.apollo");
  assert.ok(notified);
});

void test("clearSecretsFromHost normalizes connection id", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });
  manager.registerDefinition("ext.apollo", APOLLO_DEFINITION);
  await manager.setSecrets("ext.apollo", "ext.apollo.apollo", { apiKey: "sk-test" });

  // Use uppercase variant — should still clear the right record
  await manager.clearSecretsFromHost("EXT.APOLLO.APOLLO");

  const snapshot = await manager.getSnapshot("ext.apollo.apollo");
  assert.ok(snapshot);
  assert.equal(snapshot.status, "missing");
});

void test("clearSecretsFromHost throws for unregistered connection", async () => {
  const manager = new ConnectionManager({ settings: createMemorySettings() });

  await assert.rejects(
    () => manager.clearSecretsFromHost("nonexistent"),
    /not registered/,
  );
});
