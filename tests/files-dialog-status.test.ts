import assert from "node:assert/strict";
import { test } from "node:test";

import { buildFilesDialogStatusMessage } from "../src/ui/files-dialog-status.ts";

void test("status line appends connected directory name when available", () => {
  const message = buildFilesDialogStatusMessage({
    totalCount: 4,
    totalSizeBytes: 43_008,
    backendLabel: "Local folder",
    nativeDirectoryName: "Project Docs",
  });

  assert.equal(message, "4 files · 42.0 KB · Local folder: Project Docs");
});
