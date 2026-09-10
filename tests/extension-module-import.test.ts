import assert from "node:assert/strict";
import { test } from "node:test";

import {
  getLocalExtensionImportCandidates,
  importExtensionModule,
  readRemoteExtensionOptInFromStorage,
  resolveBundledLocalExtensionImporters,
} from "../src/commands/extension-module-import.ts";

void test("resolveBundledLocalExtensionImporters keeps function importers only", () => {
  const resolved = resolveBundledLocalExtensionImporters({
    "../extensions/alpha.ts": () => Promise.resolve({ alpha: true }),
    "../extensions/beta.ts": "not-a-function",
  });

  assert.deepEqual(Object.keys(resolved), ["../extensions/alpha.ts"]);
});

void test("getLocalExtensionImportCandidates expands ts/js variants", () => {
  assert.deepEqual(
    getLocalExtensionImportCandidates("../extensions/alpha.ts"),
    ["../extensions/alpha.ts", "../extensions/alpha.js"],
  );

  assert.deepEqual(
    getLocalExtensionImportCandidates("../extensions/beta"),
    ["../extensions/beta", "../extensions/beta.ts", "../extensions/beta.js"],
  );
});

void test("importExtensionModule prefers bundled local importer candidates", async () => {
  const module = await importExtensionModule("../extensions/alpha.js", "local-module", {
    bundledImporters: {
      "../extensions/alpha.ts": () => Promise.resolve({ from: "bundled-ts" }),
    },
    dynamicImport: (_specifier) => Promise.resolve({ from: "dynamic" }),
    isDev: true,
  });

  assert.deepEqual(module, { from: "bundled-ts" });
});

void test("importExtensionModule falls back to dynamic import only in dev for missing local bundle", async () => {
  const module = await importExtensionModule("../extensions/missing", "local-module", {
    bundledImporters: {},
    dynamicImport: (specifier) => Promise.resolve({ from: "dynamic", specifier }),
    isDev: true,
  });

  assert.deepEqual(module, { from: "dynamic", specifier: "../extensions/missing" });
});

void test("importExtensionModule uses dynamic import for blob and remote sources", async () => {
  const seenSpecifiers: string[] = [];
  const dynamicImport = (specifier: string) => {
    seenSpecifiers.push(specifier);
    return Promise.resolve({ specifier });
  };

  const blobModule = await importExtensionModule("blob:https://example.test/ext", "blob-url", {
    dynamicImport,
    isDev: false,
  });
  const remoteModule = await importExtensionModule("https://example.test/ext.js", "remote-url", {
    dynamicImport,
    isDev: false,
  });

  assert.deepEqual(blobModule, { specifier: "blob:https://example.test/ext" });
  assert.deepEqual(remoteModule, { specifier: "https://example.test/ext.js" });
  assert.deepEqual(seenSpecifiers, ["blob:https://example.test/ext", "https://example.test/ext.js"]);
});

void test("readRemoteExtensionOptInFromStorage handles true/false and storage failures", () => {
  const optedInStorage: Storage = {
    getItem: (_key: string) => "1",
    setItem: () => {
      throw new Error("unused");
    },
    removeItem: () => {
      throw new Error("unused");
    },
    clear: () => {
      throw new Error("unused");
    },
    key: () => null,
    length: 0,
  };

  const throwingStorage: Storage = {
    getItem: () => {
      throw new Error("blocked");
    },
    setItem: () => {
      throw new Error("unused");
    },
    removeItem: () => {
      throw new Error("unused");
    },
    clear: () => {
      throw new Error("unused");
    },
    key: () => null,
    length: 0,
  };

  assert.equal(readRemoteExtensionOptInFromStorage(optedInStorage), true);
  assert.equal(readRemoteExtensionOptInFromStorage(null), false);
  assert.equal(readRemoteExtensionOptInFromStorage(throwingStorage), false);
});
