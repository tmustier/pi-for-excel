import assert from "node:assert/strict";
import { test } from "node:test";

import { resolveTabTitle } from "../src/taskpane/session-title.ts";

void test("resolveTabTitle falls back to Chat N when explicit title is blank", () => {
  assert.equal(
    resolveTabTitle({
      hasExplicitTitle: true,
      sessionTitle: "   ",
      defaultTabNumber: 4,
    }),
    "Chat 4",
  );
});

void test("resolveTabTitle falls back to Chat 1 for invalid default numbers", () => {
  assert.equal(
    resolveTabTitle({
      hasExplicitTitle: false,
      sessionTitle: "",
      defaultTabNumber: 0,
    }),
    "Chat 1",
  );
});
