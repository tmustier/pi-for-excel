import assert from "node:assert/strict";
import { test } from "node:test";
import { readFile } from "node:fs/promises";

import {
  BUILTIN_SNAKE_EXTENSION_ID,
  EXTENSIONS_REGISTRY_STORAGE_KEY,
  loadStoredExtensions,
  saveStoredExtensions,
} from "../src/extensions/store.ts";

class MemorySettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  readRaw(key: string): unknown {
    return this.values.has(key) ? this.values.get(key) ?? null : null;
  }
}

void test("builtins registry wires /experimental and /extensions command registration", async () => {
  const source = await readFile(new URL("../src/commands/builtins/index.ts", import.meta.url), "utf8");

  assert.match(source, /createExperimentalCommands/);
  assert.match(source, /\.\.\.createExperimentalCommands\(\)/);

  assert.match(source, /createExtensionsCommands/);
  assert.match(source, /\.\.\.createExtensionsCommands\(context\)/);

  const extensionApiSource = await readFile(new URL("../src/commands/extension-api.ts", import.meta.url), "utf8");
  assert.match(extensionApiSource, /import\.meta\.glob\("\.\.\/extensions\/\*\.\{ts,js\}"\)/);
  assert.match(extensionApiSource, /Local extension module/);
});

void test("extension registry seeds default snake extension when storage is empty", async () => {
  const settings = new MemorySettingsStore();

  const entries = await loadStoredExtensions(settings);
  assert.equal(entries.length, 1);
  assert.equal(entries[0].id, BUILTIN_SNAKE_EXTENSION_ID);

  const raw = settings.readRaw(EXTENSIONS_REGISTRY_STORAGE_KEY);
  assert.ok(raw);
});

void test("extension registry preserves explicit empty saved entries", async () => {
  const settings = new MemorySettingsStore();

  await saveStoredExtensions(settings, []);
  const entries = await loadStoredExtensions(settings);
  assert.deepEqual(entries, []);
});

