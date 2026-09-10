import assert from "node:assert/strict";
import { test } from "node:test";

import { getFilesWorkspace } from "../src/files/workspace.ts";
import { createFilesTool } from "../src/tools/files.ts";

function getOfficeGlobal(): unknown {
  return Reflect.get(globalThis, "Office");
}

function setOfficeGlobal(value: unknown): void {
  Reflect.set(globalThis, "Office", value);
}

function deleteOfficeGlobal(): void {
  Reflect.deleteProperty(globalThis, "Office");
}

async function withOfficeDocumentUrl(url: string, run: () => Promise<void>): Promise<void> {
  const previousOffice = getOfficeGlobal();

  setOfficeGlobal({
    context: {
      document: {
        url,
      },
    },
  });

  try {
    await run();
  } finally {
    if (previousOffice === undefined) {
      deleteOfficeGlobal();
    } else {
      setOfficeGlobal(previousOffice);
    }
  }
}

async function clearWorkspace(): Promise<void> {
  const workspace = getFilesWorkspace();
  const files = await workspace.listFiles();

  for (const file of files) {
    if (file.sourceKind !== "workspace") {
      continue;
    }

    await workspace.deleteFile(file.path);
  }

  await workspace.clearAuditTrail();
}

void test("files tool lists built-in docs even without user workspace files", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  const result = await tool.execute("call-list-empty", { action: "list" });
  const details = result.details;

  assert.ok(details && details.kind === "files_list");
  assert.ok(details.count > 0);

  const builtinDoc = details.files.find((file) => file.path === "assistant-docs/docs/extensions.md");
  assert.ok(builtinDoc);
  assert.equal(builtinDoc.sourceKind, "builtin-doc");
  assert.equal(builtinDoc.readOnly, true);

  const listText = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.match(listText, /built-in doc/i);
});

void test("files tool reads built-in docs as read-only", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  const readResult = await tool.execute("call-read-builtin", {
    action: "read",
    path: "assistant-docs/docs/extensions.md",
    mode: "text",
  });

  const details = readResult.details;
  assert.ok(details && details.kind === "files_read");
  assert.equal(details.sourceKind, "builtin-doc");
  assert.equal(details.readOnly, true);

  const text = readResult.content[0]?.type === "text" ? readResult.content[0].text : "";
  assert.match(text, /Extensions \(MVP authoring guide\)/i);

  await assert.rejects(
    () => tool.execute("call-write-builtin", {
      action: "write",
      path: "assistant-docs/docs/extensions.md",
      content: "nope",
      encoding: "text",
    }),
    /built-in doc/i,
  );

  await assert.rejects(
    () => tool.execute("call-delete-builtin", {
      action: "delete",
      path: "assistant-docs/docs/extensions.md",
    }),
    /built-in doc/i,
  );
});

void test("files tool write/read/delete round-trip for text", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  await tool.execute("call-write", {
    action: "write",
    path: "notes.md",
    content: "hello from files tool",
    encoding: "text",
  });

  const read = await tool.execute("call-read", {
    action: "read",
    path: "notes.md",
    mode: "text",
  });

  const readText = read.content[0]?.type === "text" ? read.content[0].text : "";
  assert.match(readText, /hello from files tool/);

  const readDetails = read.details;
  assert.ok(readDetails && readDetails.kind === "files_read");
  assert.equal(readDetails.mode, "text");

  await tool.execute("call-delete", {
    action: "delete",
    path: "notes.md",
  });

  const listed = await tool.execute("call-list-post-delete", { action: "list" });
  const listDetails = listed.details;
  assert.ok(listDetails && listDetails.kind === "files_list");

  const workspaceEntries = listDetails.files.filter((file) => file.sourceKind !== "builtin-doc");
  assert.equal(workspaceEntries.length, 0);
});

void test("files tool includes workbook tag metadata in list details", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  await withOfficeDocumentUrl("https://contoso.example/workbooks/Sales.xlsx", async () => {
    await tool.execute("call-write-tagged", {
      action: "write",
      path: "tagged.md",
      content: "hello",
      encoding: "text",
    });
  });

  const listResult = await tool.execute("call-list-tagged", {
    action: "list",
  });

  const details = listResult.details;
  assert.ok(details && details.kind === "files_list");

  const tagged = details.files.find((file) => file.path === "tagged.md");
  assert.ok(tagged);
  assert.equal(tagged.workbookTag?.workbookLabel, "Sales.xlsx");
});

void test("files tool rejects text-mode reads of binary files", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  await tool.execute("call-write-binary", {
    action: "write",
    path: "archive.bin",
    content: "AAEC",
    encoding: "base64",
    mime_type: "application/octet-stream",
  });

  await assert.rejects(
    () => tool.execute("call-read-binary-as-text", {
      action: "read",
      path: "archive.bin",
      mode: "text",
    }),
    /binary/i,
  );
});

void test("files tool list filters by folder prefix", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  await tool.execute("w1", { action: "write", path: "notes/index.md", content: "# Index", encoding: "text" });
  await tool.execute("w2", { action: "write", path: "notes/budget.md", content: "Budget notes", encoding: "text" });
  await tool.execute("w3", { action: "write", path: "scratch/temp.txt", content: "temp", encoding: "text" });
  await tool.execute("w4", { action: "write", path: "top-level.md", content: "root", encoding: "text" });

  // Filter to notes/
  const notesResult = await tool.execute("list-notes", { action: "list", path: "notes/" });
  const notesDetails = notesResult.details;
  assert.ok(notesDetails && notesDetails.kind === "files_list");
  assert.equal(notesDetails.count, 2);
  assert.ok(notesDetails.files.every((f) => f.path.startsWith("notes/")));

  // Filter to scratch/
  const scratchResult = await tool.execute("list-scratch", { action: "list", path: "scratch" });
  const scratchDetails = scratchResult.details;
  assert.ok(scratchDetails && scratchDetails.kind === "files_list");
  assert.equal(scratchDetails.count, 1);
  assert.equal(scratchDetails.files[0].path, "scratch/temp.txt");

  // No filter returns all (workspace + built-in docs)
  const allResult = await tool.execute("list-all", { action: "list" });
  const allDetails = allResult.details;
  assert.ok(allDetails && allDetails.kind === "files_list");
  assert.ok(allDetails.count >= 4); // at least 4 workspace files + built-in docs

  // Filter to nonexistent folder returns empty
  const emptyResult = await tool.execute("list-empty", { action: "list", path: "nonexistent/" });
  const emptyDetails = emptyResult.details;
  assert.ok(emptyDetails && emptyDetails.kind === "files_list");
  assert.equal(emptyDetails.count, 0);
});

void test("files tool list folder filter output includes count context", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  await tool.execute("w1", { action: "write", path: "notes/a.md", content: "a", encoding: "text" });
  await tool.execute("w2", { action: "write", path: "other/b.md", content: "b", encoding: "text" });

  const result = await tool.execute("list-notes", { action: "list", path: "notes/" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.match(text, /folder: notes\//);
  assert.match(text, /1 of/); // "1 of N total"
});

void test("files tool blocks path traversal", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  await assert.rejects(
    () => tool.execute("call-write-traversal", {
      action: "write",
      path: "../secret.txt",
      content: "nope",
      encoding: "text",
    }),
    /relative|cannot contain|Path/,
  );
});
