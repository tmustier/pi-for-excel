import assert from "node:assert/strict";
import { test } from "node:test";

import type { WorkspaceFileEntry } from "../src/files/types.ts";
import { createFilesDialogDetailActions } from "../src/ui/files-dialog-actions.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

void test("Copy content falls back to execCommand when the clipboard API is refused (Excel on the web, Excel for Mac)", async () => {
  const { document: fakeDocument, restore } = installFakeDom();
  const previousNavigator = Reflect.get(globalThis, "navigator");
  const copied: string[] = [];
  let selected = "";
  const createElement = fakeDocument.createElement.bind(fakeDocument);
  Reflect.set(fakeDocument, "createElement", (tag: string) => {
    const element = createElement(tag);
    if (tag === "textarea") Reflect.set(element, "select", () => { selected = String(Reflect.get(element, "value")); });
    return element;
  });
  Reflect.set(fakeDocument, "execCommand", (command: string) => {
    if (command === "copy") copied.push(selected);
    return command === "copy";
  });
  Reflect.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: { clipboard: { writeText: () => Promise.reject(new Error("The request is not allowed by the user agent.")) } },
  });

  try {
    const file: WorkspaceFileEntry = {
      path: "assistant-docs/guide.md",
      name: "guide.md",
      size: 12,
      mimeType: "text/markdown",
      kind: "text",
      modifiedAt: 0,
      sourceKind: "builtin-doc",
      locationKind: "builtin-doc",
      readOnly: true,
    };
    const actions = createFilesDialogDetailActions({
      file,
      fileRef: { path: file.path, locationKind: "builtin-doc" },
      workspace: {
        readFile: () => Promise.resolve({ ...file, text: "# Guide body", truncated: false }),
        downloadFile: () => Promise.resolve(),
        renameFile: () => Promise.resolve(),
        deleteFile: () => Promise.resolve(),
      },
      auditContext: { actor: "user", source: "test" },
      onAfterRename: () => Promise.resolve(),
      onAfterDelete: () => Promise.resolve(),
    });

    const copyButton = Array.from(actions.querySelectorAll<HTMLButtonElement>("button")).find((button) => button.textContent === "Copy content");
    assert.ok(copyButton);
    copyButton.dispatchEvent(new Event("click"));
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.deepEqual(copied, ["# Guide body"]);
  } finally {
    Reflect.defineProperty(globalThis, "navigator", { configurable: true, value: previousNavigator });
    restore();
  }
});
