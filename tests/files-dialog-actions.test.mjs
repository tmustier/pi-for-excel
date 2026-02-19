import assert from "node:assert/strict";
import { test } from "node:test";

import { createFilesDialogDetailActions } from "../src/ui/files-dialog-actions.ts";
import { installFakeDom } from "./fake-dom.test.ts";

function makeFile(path, overrides = {}) {
  return {
    path,
    name: path.split("/").at(-1) ?? path,
    size: 10,
    modifiedAt: 0,
    mimeType: "text/plain",
    kind: "text",
    sourceKind: "workspace",
    readOnly: false,
    ...overrides,
  };
}

function buildReadResult(path) {
  return {
    ...makeFile(path),
    text: "hello",
  };
}

function createWorkspaceStub() {
  return {
    readFile: (path) => Promise.resolve(buildReadResult(path)),
    downloadFile: () => Promise.resolve(),
    renameFile: () => Promise.resolve(),
    deleteFile: () => Promise.resolve(),
  };
}

function getButtonLabels(root) {
  return Array.from(root.querySelectorAll("button"))
    .map((button) => button.textContent?.trim() ?? "")
    .filter((text) => text.length > 0);
}

test("detail actions show open/download only for read-only files", () => {
  const { restore } = installFakeDom();

  try {
    const actions = createFilesDialogDetailActions({
      file: makeFile("assistant-docs/docs/README.md", {
        sourceKind: "builtin-doc",
        locationKind: "builtin-doc",
        readOnly: true,
      }),
      fileRef: {
        path: "assistant-docs/docs/README.md",
        locationKind: "builtin-doc",
      },
      workspace: createWorkspaceStub(),
      auditContext: {
        actor: "user",
        source: "test",
      },
      onAfterRename: () => Promise.resolve(),
      onAfterDelete: () => Promise.resolve(),
    });

    assert.deepEqual(getButtonLabels(actions), ["Open ↗", "Download"]);
  } finally {
    restore();
  }
});

test("detail actions include rename/delete for writable files", () => {
  const { restore } = installFakeDom();

  try {
    const actions = createFilesDialogDetailActions({
      file: makeFile("notes/plan.md"),
      fileRef: {
        path: "notes/plan.md",
        locationKind: "workspace",
      },
      workspace: createWorkspaceStub(),
      auditContext: {
        actor: "user",
        source: "test",
      },
      onAfterRename: () => Promise.resolve(),
      onAfterDelete: () => Promise.resolve(),
    });

    assert.deepEqual(getButtonLabels(actions), ["Open ↗", "Download", "Rename", "Delete"]);
  } finally {
    restore();
  }
});
