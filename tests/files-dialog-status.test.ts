import assert from "node:assert/strict";
import { test } from "node:test";

import { buildFilesDialogStatusMessage } from "../src/ui/files-dialog-status.ts";

void test("status message explains built-in docs when write/delete is gated", () => {
  const message = buildFilesDialogStatusMessage({
    filesExperimentEnabled: false,
    totalCount: 4,
    filteredCount: 2,
    selectedFilter: "builtin",
    activeFilterLabel: "Built-in docs",
    builtinDocsCount: 3,
    workspaceFilesCount: 1,
  });

  assert.match(message, /Built-in docs stay available/i);
  assert.match(message, /Enable write access/i);
});

void test("status message reports active filtered count when enabled", () => {
  const message = buildFilesDialogStatusMessage({
    filesExperimentEnabled: true,
    totalCount: 10,
    filteredCount: 3,
    selectedFilter: "builtin",
    activeFilterLabel: "Built-in docs",
    builtinDocsCount: 4,
    workspaceFilesCount: 6,
  });

  assert.equal(message, "3 of 10 files shown · Built-in docs.");
});
