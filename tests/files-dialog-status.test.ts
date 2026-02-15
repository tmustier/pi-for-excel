import assert from "node:assert/strict";
import { test } from "node:test";

import { buildFilesDialogStatusMessage } from "../src/ui/files-dialog-status.ts";

void test("status message reports total count when showing all files", () => {
  const message = buildFilesDialogStatusMessage({
    totalCount: 10,
    filteredCount: 10,
    selectedFilter: "all",
    activeFilterLabel: "All files",
  });

  assert.equal(message, "10 files available to the agent.");
});

void test("status message reports active filtered count", () => {
  const message = buildFilesDialogStatusMessage({
    totalCount: 10,
    filteredCount: 3,
    selectedFilter: "builtin",
    activeFilterLabel: "Built-in docs",
  });

  assert.equal(message, "3 of 10 files shown · Built-in docs.");
});
