import assert from "node:assert/strict";
import { test } from "node:test";

import { buildFilesDialogStatusMessage } from "../src/ui/files-dialog-status.ts";

// Retained: Chromium cannot safely seed a persisted FileSystemDirectoryHandle;
// this covers the connected-native-directory suffix until that host boundary is available.
void test("status line appends connected directory name when available", () => {
  const message = buildFilesDialogStatusMessage({
    totalCount: 4,
    totalSizeBytes: 43_008,
    backendLabel: "Local folder",
    nativeDirectoryName: "Project Docs",
  });

  assert.equal(message, "4 files · 42.0 KB · Local folder: Project Docs");
});
