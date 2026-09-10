import assert from "node:assert/strict";
import { test } from "node:test";

import {
  clearExtensionStorage,
  deleteExtensionStorageValue,
  getExtensionStorageValue,
  listExtensionStorageKeys,
  setExtensionStorageValue,
} from "../src/extensions/storage-store.ts";

class MemorySettingsStore {
  private readonly store = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.store.get(key));
  }

  set(key: string, value: unknown): Promise<void> {
    this.store.set(key, value);
    return Promise.resolve();
  }
}

void test("extension storage store keeps values scoped by extension id", async () => {
  const settings = new MemorySettingsStore();

  await setExtensionStorageValue(settings, "ext.one", "token", "a");
  await setExtensionStorageValue(settings, "ext.two", "token", "b");

  assert.equal(await getExtensionStorageValue(settings, "ext.one", "token"), "a");
  assert.equal(await getExtensionStorageValue(settings, "ext.two", "token"), "b");

  assert.deepEqual(await listExtensionStorageKeys(settings, "ext.one"), ["token"]);

  await deleteExtensionStorageValue(settings, "ext.one", "token");
  assert.equal(await getExtensionStorageValue(settings, "ext.one", "token"), undefined);

  await clearExtensionStorage(settings, "ext.two");
  assert.equal(await getExtensionStorageValue(settings, "ext.two", "token"), undefined);
});
