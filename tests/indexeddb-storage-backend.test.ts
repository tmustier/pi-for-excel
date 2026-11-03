import "fake-indexeddb/auto";
import assert from "node:assert/strict";
import { test } from "node:test";

import { IndexedDBStorageBackend } from "../src/storage/local/indexeddb-storage-backend.ts";

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("IndexedDB request failed"));
  });
}

void test("a failed IndexedDB open can be retried after the host recovers", async () => {
  const dbName = `open-retry-${crypto.randomUUID()}`;
  const incompatible = await requestResult(indexedDB.open(dbName, 2));
  incompatible.close();
  const backend = new IndexedDBStorageBackend({ dbName, version: 1, stores: [{ name: "settings" }] });
  await assert.rejects(backend.get("settings", "key"), { name: "VersionError" });
  await requestResult(indexedDB.deleteDatabase(dbName));
  await backend.set("settings", "key", "recovered");
  assert.equal(await backend.get("settings", "key"), "recovered");
});

for (const operation of ["set", "delete", "clear"] as const) {
  void test(`standalone IndexedDB ${operation} rejects a commit-time abort`, async () => {
    const backend = new IndexedDBStorageBackend({
      dbName: `commit-${operation}-${crypto.randomUUID()}`, version: 1, stores: [{ name: "settings" }],
    });
    await backend.set("settings", "key", "before");
    const original = IDBDatabase.prototype.transaction;
    IDBDatabase.prototype.transaction = function (storeNames, mode, options) {
      const tx = original.call(this, storeNames, mode, options);
      if (mode === "readwrite") {
        queueMicrotask(() => {
          const afterWrite = tx.objectStore("settings").get("key");
          afterWrite.addEventListener("success", () => tx.abort(), { once: true });
        });
      }
      return tx;
    };
    try {
      const result = operation === "set"
        ? backend.set("settings", "key", "after")
        : operation === "delete" ? backend.delete("settings", "key") : backend.clear("settings");
      await assert.rejects(result, /IndexedDB request failed|abort/iu, "success requires transaction commit");
    } finally {
      IDBDatabase.prototype.transaction = original;
    }
    assert.equal(await backend.get("settings", "key"), "before", "aborted write must remain rolled back");
  });
}

void test("an upgrade-closed IndexedDB backend stays reload-required", async () => {
  const dbName = `upgrade-${crypto.randomUUID()}`;
  const backend = new IndexedDBStorageBackend({ dbName, version: 1, stores: [{ name: "settings" }] });
  await backend.set("settings", "key", "before");
  const upgraded = await requestResult(indexedDB.open(dbName, 2));
  try {
    await assert.rejects(backend.get("settings", "key"), /reload required/u);
  } finally {
    upgraded.close();
  }
});
