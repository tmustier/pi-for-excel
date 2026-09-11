/**
 * Persisted files-workspace records: what `FilesWorkspace` reads back from the
 * two settings keys it writes. Each payload goes through the real product path
 * (construct the workspace over the settings, then list) and the test states
 * which entries survive.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { MemoryBackend } from "../src/files/backend.ts";
import { FilesWorkspace } from "../src/files/workspace.ts";

const METADATA_KEY = "files.workspace.metadata.v1";
const AUDIT_KEY = "files.workspace.audit.v1";

class MemorySettings {
  readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    const value = this.values.get(key);
    return Promise.resolve(value === undefined ? null : structuredClone(value));
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

function validAuditEntry(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    id: "entry",
    at: 100,
    action: "write",
    actor: "user",
    source: "test",
    backend: "memory",
    path: "notes/a.txt",
    ...overrides,
  };
}

async function workspaceWithFiles(settings: MemorySettings, paths: readonly string[]): Promise<FilesWorkspace> {
  const backend = new MemoryBackend();
  for (const path of paths) {
    await backend.writeBytes(path, new TextEncoder().encode("x"));
  }
  return new FilesWorkspace({ initialBackend: backend, settings });
}

void test("an audit entry is loaded as written, or dropped; it is not repaired", async () => {
  const settings = new MemorySettings();
  settings.values.set(AUDIT_KEY, {
    version: 1,
    entries: [
      validAuditEntry({ id: "complete" }),
      validAuditEntry({ id: "rename", action: "rename", path: undefined, fromPath: "a.txt", toPath: "b/c.txt", bytes: 3, workbookId: "w", workbookLabel: "W.xlsx" }),
      validAuditEntry({ id: "unknown-action", action: "compile" }),
      validAuditEntry({ id: "unknown-actor", actor: "robot" }),
      validAuditEntry({ id: "unknown-backend", backend: "s3" }),
      validAuditEntry({ id: "empty-source", source: "" }),
      validAuditEntry({ id: "string-at", at: "100" }),
      validAuditEntry({ id: "" }),
      validAuditEntry({ id: "no-id", ...{ id: undefined } }),
      validAuditEntry({ id: "absolute-path", path: "/etc/passwd" }),
      validAuditEntry({ id: "dot-segment", path: "notes/../a.txt" }),
      validAuditEntry({ id: "untrimmed-path", path: " notes/a.txt " }),
      validAuditEntry({ id: "empty-path", path: "" }),
      validAuditEntry({ id: "string-bytes", bytes: "3" }),
      "not an object",
      null,
    ],
  });

  const workspace = await workspaceWithFiles(settings, []);
  const survivors = (await workspace.listAuditEntries(100)).map((entry) => entry.id);
  assert.deepEqual(survivors, ["complete", "rename"]);
});

void test("audit entries load newest first and are capped at 300", async () => {
  const settings = new MemorySettings();
  const entries = [];
  for (let index = 0; index < 305; index += 1) {
    entries.push(validAuditEntry({ id: `e${index}`, at: index }));
  }
  settings.values.set(AUDIT_KEY, { version: 1, entries });

  const workspace = await workspaceWithFiles(settings, []);
  const loaded = await workspace.listAuditEntries(300);
  assert.equal(loaded.length, 300);
  assert.equal(loaded[0]?.id, "e304");
  assert.equal(loaded[299]?.id, "e5");
});

void test("undeclared properties on a persisted audit entry are dropped on load and not written back", async () => {
  const settings = new MemorySettings();
  settings.values.set(AUDIT_KEY, {
    version: 1,
    entries: [validAuditEntry({ id: "extra", futureField: { nested: true } })],
  });

  const workspace = await workspaceWithFiles(settings, ["notes/a.txt"]);
  const [loaded] = await workspace.listAuditEntries(10);
  assert.ok(loaded);
  assert.equal("futureField" in loaded, false);

  await workspace.writeTextFile("notes/b.txt", "b", undefined, { audit: { actor: "user", source: "test" } });
  const written = JSON.stringify(settings.values.get(AUDIT_KEY));
  assert.ok(written.includes('"id":"extra"'));
  assert.equal(written.includes("futureField"), false);
});

void test("an audit trail without the version envelope this module writes is not loaded", async () => {
  for (const payload of [
    { entries: [validAuditEntry()] },
    { version: 2, entries: [validAuditEntry()] },
    { version: 1, entries: "none" },
    [validAuditEntry()],
    "entries",
  ]) {
    const settings = new MemorySettings();
    settings.values.set(AUDIT_KEY, payload);
    const workspace = await workspaceWithFiles(settings, []);
    assert.deepEqual(await workspace.listAuditEntries(10), [], JSON.stringify(payload));
  }
});

void test("a workbook tag is loaded per path, or dropped; the rest of the map survives", async () => {
  const settings = new MemorySettings();
  const tag = { workbookId: "url_sha256:abc", workbookLabel: "Model.xlsx", taggedAt: 100 };
  settings.values.set(METADATA_KEY, {
    version: 1,
    byPath: {
      "notes/a.txt": tag,
      "b.txt": { ...tag, workbookLabel: "" },
      "c.txt": { workbookId: "url_sha256:abc", workbookLabel: "Model.xlsx" },
      "d.txt": { ...tag, taggedAt: "100" },
      "e.txt": "tag",
      "/f.txt": tag,
      " g.txt": tag,
      "": tag,
    },
  });

  const workspace = await workspaceWithFiles(settings, [
    "notes/a.txt", "b.txt", "c.txt", "d.txt", "e.txt", "f.txt", "g.txt",
  ]);
  const files = await workspace.listFiles();
  const tagged = files.filter((file) => file.workbookTag).map((file) => file.path);
  assert.deepEqual(tagged, ["notes/a.txt"]);
  assert.deepEqual(files.find((file) => file.path === "notes/a.txt")?.workbookTag, tag);
});

void test("a metadata map without the version envelope this module writes is not loaded", async () => {
  const tag = { workbookId: "url_sha256:abc", workbookLabel: "Model.xlsx", taggedAt: 100 };
  for (const payload of [
    { byPath: { "a.txt": tag } },
    { version: "1", byPath: { "a.txt": tag } },
    { version: 1, byPath: [tag] },
    { version: 1 },
  ]) {
    const settings = new MemorySettings();
    settings.values.set(METADATA_KEY, payload);
    const workspace = await workspaceWithFiles(settings, ["a.txt"]);
    const files = await workspace.listFiles();
    assert.equal(files.find((file) => file.path === "a.txt")?.workbookTag, undefined, JSON.stringify(payload));
  }
});
