import assert from "node:assert/strict";
import { test } from "node:test";

import { MemoryBackend } from "../src/files/backend.ts";
import {
  buildWorkspaceContextSummary,
  FilesWorkspace,
  getFilesWorkspace,
} from "../src/files/workspace.ts";
import type { WorkspaceSnapshot } from "../src/files/types.ts";

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

async function resetWorkspace(): Promise<void> {
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

function createWorkspaceSnapshot(paths: Array<{ path: string; workbookId?: string }>): WorkspaceSnapshot {
  const files = paths.map((entry, index) => ({
    path: entry.path,
    name: entry.path.split("/").pop() ?? entry.path,
    size: 10 + index,
    modifiedAt: 1_000 + index,
    mimeType: "text/plain",
    kind: "text",
    sourceKind: "workspace",
    readOnly: false,
    workbookTag: entry.workbookId
      ? {
        workbookId: entry.workbookId,
        workbookLabel: `${entry.workbookId}.xlsx`,
        taggedAt: 1_000 + index,
      }
      : undefined,
  }));

  return {
    backend: {
      kind: "memory",
      label: "Session memory",
      nativeSupported: false,
      nativeConnected: false,
    },
    files,
    signature: "test",
  };
}

void test("files workspace tags files with active workbook metadata", async () => {
  await resetWorkspace();
  const workspace = getFilesWorkspace();

  await withOfficeDocumentUrl("https://contoso.example/workbooks/Quarterly-Plan.xlsx", async () => {
    await workspace.writeTextFile("notes.md", "hello", undefined, {
      audit: { actor: "user", source: "test" },
    });
  });

  const files = await workspace.listFiles();
  const entry = files.find((file) => file.path === "notes.md");

  assert.ok(entry);
  assert.ok(entry.workbookTag);
  assert.match(entry.workbookTag?.workbookId ?? "", /^url_sha256:/);
  assert.equal(entry.workbookTag?.workbookLabel, "Quarterly-Plan.xlsx");
});

void test("files workspace records read/write actions in audit trail", async () => {
  await resetWorkspace();
  const workspace = getFilesWorkspace();

  await workspace.writeTextFile("audit.md", "hello", undefined, {
    audit: { actor: "user", source: "test-write" },
  });

  await workspace.readFile("audit.md", {
    mode: "text",
    audit: { actor: "assistant", source: "tool:files" },
  });

  const entries = await workspace.listAuditEntries(20);

  const hasWrite = entries.some((entry) =>
    entry.action === "write" &&
    entry.path === "audit.md" &&
    entry.actor === "user" &&
    entry.source === "test-write"
  );

  const hasRead = entries.some((entry) =>
    entry.action === "read" &&
    entry.path === "audit.md" &&
    entry.actor === "assistant" &&
    entry.source === "tool:files"
  );

  assert.equal(hasWrite, true);
  assert.equal(hasRead, true);
});

void test("files workspace exposes built-in docs as read-only entries", async () => {
  await resetWorkspace();
  const workspace = getFilesWorkspace();

  const files = await workspace.listFiles();
  const builtin = files.find((entry) => entry.path === "assistant-docs/docs/extensions.md");

  assert.ok(builtin);
  assert.equal(builtin.sourceKind, "builtin-doc");
  assert.equal(builtin.readOnly, true);

  const read = await workspace.readFile("assistant-docs/docs/extensions.md", {
    mode: "text",
  });

  assert.equal(read.sourceKind, "builtin-doc");
  assert.equal(read.readOnly, true);
  assert.match(read.text ?? "", /Extensions \(MVP authoring guide\)/i);

  await assert.rejects(
    () => workspace.deleteFile("assistant-docs/docs/extensions.md"),
    /built-in doc/i,
  );
});

void test("legacy workspace collisions on assistant-docs paths stay reachable", async () => {
  const backend = new MemoryBackend();
  await backend.writeBytes(
    "assistant-docs/docs/extensions.md",
    new TextEncoder().encode("legacy collision payload"),
    "text/plain",
  );

  const workspace = new FilesWorkspace({
    initialBackend: backend,
  });

  const listWithCollision = await workspace.listFiles();
  const collisionEntry = listWithCollision.find((entry) => entry.path === "assistant-docs/docs/extensions.md");

  assert.ok(collisionEntry);
  assert.equal(collisionEntry.sourceKind, "workspace");

  const readCollision = await workspace.readFile("assistant-docs/docs/extensions.md", {
    mode: "text",
  });
  assert.match(readCollision.text ?? "", /legacy collision payload/);

  await workspace.deleteFile("assistant-docs/docs/extensions.md");

  const readBuiltinAfterDelete = await workspace.readFile("assistant-docs/docs/extensions.md", {
    mode: "text",
  });
  assert.equal(readBuiltinAfterDelete.sourceKind, "builtin-doc");
  assert.match(readBuiltinAfterDelete.text ?? "", /Extensions \(MVP authoring guide\)/i);
});

void test("workspace context summary includes only relevant folders and current workbook artifacts", () => {
  const snapshot = createWorkspaceSnapshot([
    { path: "notes/index.md", workbookId: "wb-a" },
    { path: "notes/budget.md", workbookId: "wb-a" },
    { path: "imports/source.csv", workbookId: "wb-a" },
    { path: "workbooks/budget-2026/extract.csv", workbookId: "wb-a" },
    { path: "workbooks/forecast-q3/data.csv", workbookId: "wb-b" },
    { path: "scratch/temp.txt", workbookId: "wb-a" },
    { path: "assistant-docs/docs/README.md" },
  ]);

  const summary = buildWorkspaceContextSummary({
    snapshot,
    currentWorkbookId: "wb-a",
  });

  assert.equal(summary.hasRelevantFiles, true);
  assert.match(summary.summary, /^### Workspace/m);
  assert.match(summary.summary, /notes\/: 2 files\. Read notes\/index\.md first\./);
  assert.match(summary.summary, /Current workbook artifacts: 1 file \(workbooks\/budget-2026\/extract\.csv\)\./);
  assert.match(summary.summary, /imports\/: 1 file \(imports\/source\.csv\)\./);
  assert.doesNotMatch(summary.summary, /scratch\/temp\.txt/);
  assert.doesNotMatch(summary.summary, /forecast-q3/);
});

void test("workspace context relevance signature ignores scratch-only changes", () => {
  const baseSnapshot = createWorkspaceSnapshot([
    { path: "notes/index.md", workbookId: "wb-a" },
    { path: "notes/budget.md", workbookId: "wb-a" },
    { path: "imports/source.csv", workbookId: "wb-a" },
    { path: "workbooks/budget-2026/extract.csv", workbookId: "wb-a" },
    { path: "scratch/temp-a.txt", workbookId: "wb-a" },
  ]);

  const scratchChangedSnapshot = createWorkspaceSnapshot([
    { path: "notes/index.md", workbookId: "wb-a" },
    { path: "notes/budget.md", workbookId: "wb-a" },
    { path: "imports/source.csv", workbookId: "wb-a" },
    { path: "workbooks/budget-2026/extract.csv", workbookId: "wb-a" },
    { path: "scratch/temp-b.txt", workbookId: "wb-a" },
  ]);

  const importChangedSnapshot = createWorkspaceSnapshot([
    { path: "notes/index.md", workbookId: "wb-a" },
    { path: "notes/budget.md", workbookId: "wb-a" },
    { path: "imports/new-source.csv", workbookId: "wb-a" },
    { path: "workbooks/budget-2026/extract.csv", workbookId: "wb-a" },
    { path: "scratch/temp-a.txt", workbookId: "wb-a" },
  ]);

  const noteChangedSnapshot = createWorkspaceSnapshot([
    { path: "notes/index.md", workbookId: "wb-a" },
    { path: "notes/budget-v2.md", workbookId: "wb-a" },
    { path: "imports/source.csv", workbookId: "wb-a" },
    { path: "workbooks/budget-2026/extract.csv", workbookId: "wb-a" },
    { path: "scratch/temp-a.txt", workbookId: "wb-a" },
  ]);

  const base = buildWorkspaceContextSummary({
    snapshot: baseSnapshot,
    currentWorkbookId: "wb-a",
  });

  const scratchChanged = buildWorkspaceContextSummary({
    snapshot: scratchChangedSnapshot,
    currentWorkbookId: "wb-a",
  });

  const importChanged = buildWorkspaceContextSummary({
    snapshot: importChangedSnapshot,
    currentWorkbookId: "wb-a",
  });

  const noteChanged = buildWorkspaceContextSummary({
    snapshot: noteChangedSnapshot,
    currentWorkbookId: "wb-a",
  });

  assert.equal(scratchChanged.relevantSignature, base.relevantSignature);
  assert.notEqual(importChanged.relevantSignature, base.relevantSignature);
  assert.notEqual(noteChanged.relevantSignature, base.relevantSignature);
});
