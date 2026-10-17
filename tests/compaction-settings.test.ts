import assert from "node:assert/strict";
import { test } from "node:test";

import {
  COMPACTION_ENABLED_SETTING_KEY,
  readAutoCompactionEnabled,
} from "../src/compaction/settings.ts";

void test("compaction setting accepts only booleans and defaults enabled", async () => {
  const read = (value: unknown) => readAutoCompactionEnabled({
    get: (key) => {
      assert.equal(key, COMPACTION_ENABLED_SETTING_KEY);
      return Promise.resolve(value);
    },
  });

  assert.equal(await read(false), false);
  assert.equal(await read(true), true);
  assert.equal(await read(null), true);
  assert.equal(await read("false"), true);
  assert.equal(await read(0), true);
});

void test("compaction setting defaults enabled when storage cannot be read", async () => {
  const enabled = await readAutoCompactionEnabled({
    get: () => Promise.reject(new Error("storage unavailable")),
  });

  assert.equal(enabled, true);
});
