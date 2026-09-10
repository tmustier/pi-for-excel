import assert from "node:assert/strict";
import { test } from "node:test";

import {
  applyRuleAction,
  getUserRules,
  getWorkbookRules,
  hasAnyRules,
  setUserRules,
  setWorkbookRules,
  type RulesStore,
} from "../src/rules/store.ts";

class MemoryStore implements RulesStore {
  private values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.get(key));
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

void test("append action adds new text on a new line", () => {
  const updated = applyRuleAction({
    currentValue: "Use EUR",
    action: "append",
    content: "Check circular refs",
  });

  assert.equal(updated, "Use EUR\nCheck circular refs");
});

void test("replace action clears rules when content is blank", () => {
  const updated = applyRuleAction({
    currentValue: "Use EUR",
    action: "replace",
    content: "   ",
  });

  assert.equal(updated, null);
});

void test("user and workbook rules round-trip through storage", async () => {
  const store = new MemoryStore();

  await setUserRules(store, "Always use dd-mmm-yyyy");
  await setWorkbookRules(store, "url_sha256:abc", "Summary sheet is read-only");

  assert.equal(await getUserRules(store), "Always use dd-mmm-yyyy");
  assert.equal(
    await getWorkbookRules(store, "url_sha256:abc"),
    "Summary sheet is read-only",
  );
});

void test("hasAnyRules reports active state correctly", () => {
  assert.equal(
    hasAnyRules({
      userRules: null,
      workbookRules: null,
    }),
    false,
  );

  assert.equal(
    hasAnyRules({
      userRules: "Use EUR",
      workbookRules: null,
    }),
    true,
  );
});
