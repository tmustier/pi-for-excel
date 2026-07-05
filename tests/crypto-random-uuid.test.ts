import assert from "node:assert/strict";
import { test } from "node:test";

import { installCryptoRandomUuidPatch } from "../src/compat/crypto-random-uuid.js";

void test("crypto.randomUUID patch installs an RFC 4122 v4 fallback when WebView crypto lacks it", () => {
  const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "crypto");
  let nextByte = 0;

  const cryptoWithoutRandomUuid = {
    getRandomValues<T extends ArrayBufferView | null>(array: T): T {
      if (array instanceof Uint8Array) {
        for (let index = 0; index < array.length; index += 1) {
          array[index] = nextByte;
          nextByte = (nextByte + 1) & 0xff;
        }
      }
      return array;
    },
  };

  try {
    Object.defineProperty(globalThis, "crypto", {
      configurable: true,
      value: cryptoWithoutRandomUuid,
    });

    installCryptoRandomUuidPatch();

    assert.equal(typeof globalThis.crypto.randomUUID, "function");
    const uuid = globalThis.crypto.randomUUID();
    assert.match(uuid, /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/u);
  } finally {
    if (originalDescriptor) {
      Object.defineProperty(globalThis, "crypto", originalDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, "crypto");
    }
  }
});
