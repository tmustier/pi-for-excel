import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createOfficeDocumentInstanceIdentity,
  OFFICE_DOCUMENT_INSTANCE_SETTING,
  type OfficeDocumentSettingsLike,
} from "../src/host/office-document-instance.ts";

const TOKEN_A = "11111111-2222-4333-8444-555555555555";
const TOKEN_B = "66666666-7777-4888-8999-aaaaaaaaaaaa";

/** In-memory stand-in for `Office.context.document.settings`. */
function createSettings(options: { saveSucceeds?: boolean; initial?: Record<string, string> } = {}) {
  const values = new Map<string, string>(Object.entries(options.initial ?? {}));
  let saveCalls = 0;
  const settings: OfficeDocumentSettingsLike = {
    get: (name) => values.get(name),
    set: (name, value) => { values.set(name, value); },
    remove: (name) => { values.delete(name); },
    saveAsync: (callback) => {
      saveCalls += 1;
      // Office invokes the callback asynchronously.
      setTimeout(() => callback({ succeeded: options.saveSucceeds ?? true }), 0);
    },
  };
  return {
    settings,
    get saveCalls() { return saveCalls; },
    get stored() { return values.get(OFFICE_DOCUMENT_INSTANCE_SETTING); },
  };
}

void test("ensure mints one token, persists it, and read returns it afterwards", async () => {
  const doc = createSettings();
  const tokens = [TOKEN_A, TOKEN_B];
  const identity = createOfficeDocumentInstanceIdentity({
    getSettings: () => doc.settings,
    createToken: () => tokens.shift() ?? null,
  });

  assert.equal(identity.read(), null);
  assert.equal(await identity.ensure(), TOKEN_A);
  assert.equal(doc.stored, TOKEN_A);
  assert.equal(doc.saveCalls, 1);

  assert.equal(identity.read(), TOKEN_A);
  assert.equal(await identity.ensure(), TOKEN_A, "an existing token is reused");
  assert.equal(doc.saveCalls, 1);
});

void test("concurrent ensure calls share one token and one save", async () => {
  const doc = createSettings();
  const tokens = [TOKEN_A, TOKEN_B];
  const identity = createOfficeDocumentInstanceIdentity({
    getSettings: () => doc.settings,
    createToken: () => tokens.shift() ?? null,
  });

  const results = await Promise.all([identity.ensure(), identity.ensure(), identity.ensure()]);
  assert.deepEqual(results, [TOKEN_A, TOKEN_A, TOKEN_A]);
  assert.equal(doc.saveCalls, 1);
});

void test("a failed save leaves no token behind", async () => {
  const doc = createSettings({ saveSucceeds: false });
  const identity = createOfficeDocumentInstanceIdentity({
    getSettings: () => doc.settings,
    createToken: () => TOKEN_A,
  });

  assert.equal(await identity.ensure(), null);
  assert.equal(doc.stored, undefined, "unpersisted token removed from the in-memory bag");
  assert.equal(identity.read(), null);
});

void test("a non-token value in the slot is treated as absent and replaced", async () => {
  const doc = createSettings({ initial: { [OFFICE_DOCUMENT_INSTANCE_SETTING]: "not a uuid" } });
  const identity = createOfficeDocumentInstanceIdentity({
    getSettings: () => doc.settings,
    createToken: () => TOKEN_B,
  });

  assert.equal(identity.read(), null);
  assert.equal(await identity.ensure(), TOKEN_B);
  assert.equal(doc.stored, TOKEN_B);
});

void test("without an attached document nothing is read or written", async () => {
  const identity = createOfficeDocumentInstanceIdentity({
    getSettings: () => null,
    createToken: () => TOKEN_A,
  });

  assert.equal(identity.read(), null);
  assert.equal(await identity.ensure(), null);
});

void test("without a cryptographic token source no token is stored", async () => {
  const doc = createSettings();
  const identity = createOfficeDocumentInstanceIdentity({
    getSettings: () => doc.settings,
    createToken: () => null,
  });

  assert.equal(await identity.ensure(), null);
  assert.equal(doc.stored, undefined);
  assert.equal(doc.saveCalls, 0);
});
