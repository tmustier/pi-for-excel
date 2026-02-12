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
    await workspace.deleteFile(file.path);
  }

  await workspace.clearAuditTrail();
}

void test("files tool lists an empty workspace", async () => {
  await clearWorkspace();
  const tool = createFilesTool();

  const result = await tool.execute("call-list-empty", { action: "list" });
  const details = result.details;

  assert.ok(details && details.kind === "files_list");
  assert.equal(details.count, 0);
  assert.match(result.content[0]?.type === "text" ? result.content[0].text : "", /No files yet/i);
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
  assert.equal(listDetails.count, 0);
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
