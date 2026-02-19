import assert from "node:assert/strict";
import { test } from "node:test";

import { resolveRenameDestinationPath } from "../src/ui/files-dialog-paths.ts";

test("resolveRenameDestinationPath keeps extension when user omits it", () => {
  assert.equal(
    resolveRenameDestinationPath("reports/q1/revenue.xlsx", "revenue-final"),
    "reports/q1/revenue-final.xlsx",
  );

  assert.equal(
    resolveRenameDestinationPath("reports/q1/revenue.xlsx", "archive/revenue-final"),
    "archive/revenue-final.xlsx",
  );
});

test("resolveRenameDestinationPath respects explicit target extensions", () => {
  assert.equal(
    resolveRenameDestinationPath("reports/q1/revenue.xlsx", "revenue-final.csv"),
    "reports/q1/revenue-final.csv",
  );

  assert.equal(
    resolveRenameDestinationPath("reports/q1/revenue.xlsx", ".hidden"),
    "reports/q1/.hidden",
  );

  assert.equal(
    resolveRenameDestinationPath("reports/q1/revenue.xlsx", "revenue."),
    "reports/q1/revenue.",
  );
});

test("resolveRenameDestinationPath handles empty and trailing-slash input", () => {
  assert.equal(
    resolveRenameDestinationPath("reports/q1/revenue.xlsx", ""),
    "reports/q1/revenue.xlsx",
  );

  assert.equal(
    resolveRenameDestinationPath("reports/q1/revenue.xlsx", "archive/"),
    "reports/q1/revenue.xlsx",
  );
});
