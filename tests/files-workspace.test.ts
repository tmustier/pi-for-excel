import assert from "node:assert/strict";
import { test } from "node:test";

import { MemoryBackend, type WorkspaceBackend } from "../src/files/backend.ts";
import { normalizeWorkspacePath } from "../src/files/path.ts";
import {
  buildWorkspaceContextSummary,
  FilesWorkspace,
  getFilesWorkspace,
} from "../src/files/workspace.ts";
import type {
  WorkspaceBackendKind,
  WorkspaceFileEntry,
  WorkspaceFileReadResult,
  WorkspaceSnapshot,
} from "../src/files/types.ts";

function getOfficeGlobal(): DynamicValue {
  return Reflect.get(globalThis, "Office");
}

function setOfficeGlobal(value: DynamicValue): void {
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

    await workspace.deleteFile(file.path, {
      locationKind: file.locationKind ?? "workspace",
    });
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

interface CleanupTestFile {
  path: string;
  modifiedAt: number;
  content?: string;
}

class CleanupTestBackend implements WorkspaceBackend {
  readonly kind = "memory";
  readonly label = "Session memory";

  private readonly files = new Map<string, CleanupTestFile>();
  private readonly failDeletePaths: Set<string>;

  constructor(files: readonly CleanupTestFile[], failDeletePaths: readonly string[] = []) {
    for (const file of files) {
      const normalizedPath = normalizeWorkspacePath(file.path);
      this.files.set(normalizedPath, {
        path: normalizedPath,
        modifiedAt: file.modifiedAt,
        content: file.content,
      });
    }

    this.failDeletePaths = new Set(failDeletePaths.map((path) => normalizeWorkspacePath(path)));
  }

  listFiles(): Promise<WorkspaceFileEntry[]> {
    const entries: WorkspaceFileEntry[] = Array.from(this.files.values()).map((file) => ({
      path: file.path,
      name: file.path.split("/").pop() ?? file.path,
      size: (file.content ?? "").length,
      modifiedAt: file.modifiedAt,
      mimeType: "text/plain",
      kind: "text",
      sourceKind: "workspace",
      readOnly: false,
    }));

    return Promise.resolve(entries.sort((left, right) => left.path.localeCompare(right.path)));
  }

  readFile(path: string): Promise<WorkspaceFileReadResult> {
    const normalizedPath = normalizeWorkspacePath(path);
    const file = this.files.get(normalizedPath);

    if (!file) {
      return Promise.reject(new Error(`File not found: ${normalizedPath}`));
    }

    return Promise.resolve({
      path: file.path,
      name: file.path.split("/").pop() ?? file.path,
      size: (file.content ?? "").length,
      modifiedAt: file.modifiedAt,
      mimeType: "text/plain",
      kind: "text",
      sourceKind: "workspace",
      readOnly: false,
      text: file.content ?? "",
    });
  }

  writeBytes(path: string, bytes: Uint8Array): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    this.files.set(normalizedPath, {
      path: normalizedPath,
      modifiedAt: Date.now(),
      content: new TextDecoder().decode(bytes),
    });

    return Promise.resolve();
  }

  deleteFile(path: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);

    if (this.failDeletePaths.has(normalizedPath)) {
      return Promise.reject(new Error(`forced delete failure: ${normalizedPath}`));
    }

    this.files.delete(normalizedPath);
    return Promise.resolve();
  }

  renameFile(oldPath: string, newPath: string): Promise<void> {
    const normalizedOldPath = normalizeWorkspacePath(oldPath);
    const normalizedNewPath = normalizeWorkspacePath(newPath);

    const existing = this.files.get(normalizedOldPath);
    if (!existing) {
      return Promise.reject(new Error(`File not found: ${normalizedOldPath}`));
    }

    this.files.delete(normalizedOldPath);
    this.files.set(normalizedNewPath, {
      path: normalizedNewPath,
      modifiedAt: existing.modifiedAt,
      content: existing.content,
    });

    return Promise.resolve();
  }
}

class SourceBackend implements WorkspaceBackend {
  readonly kind: WorkspaceBackendKind;
  readonly label: string;

  private readonly files = new Map<string, { bytes: Uint8Array; modifiedAt: number }>();
  private readonly missingMessageTemplate: string;
  private readonly missingErrorName?: string;

  constructor(args: {
    kind: WorkspaceBackendKind;
    label: string;
    files: Array<{ path: string; text: string; modifiedAt: number }>;
    missingMessageTemplate?: string;
    missingErrorName?: string;
  }) {
    this.kind = args.kind;
    this.label = args.label;
    this.missingMessageTemplate = args.missingMessageTemplate ?? "File not found: {path}";
    this.missingErrorName = args.missingErrorName;

    for (const file of args.files) {
      const normalizedPath = normalizeWorkspacePath(file.path);
      this.files.set(normalizedPath, {
        bytes: new TextEncoder().encode(file.text),
        modifiedAt: file.modifiedAt,
      });
    }
  }

  private buildMissingError(path: string): Error {
    const message = this.missingMessageTemplate.replace("{path}", path);
    const error = new Error(message);

    if (this.missingErrorName) {
      error.name = this.missingErrorName;
    }

    return error;
  }

  listFiles(): Promise<WorkspaceFileEntry[]> {
    const out: WorkspaceFileEntry[] = Array.from(this.files.entries()).map(([path, file]) => ({
      path,
      name: path.split("/").pop() ?? path,
      size: file.bytes.byteLength,
      modifiedAt: file.modifiedAt,
      mimeType: "text/plain",
      kind: "text",
      sourceKind: "workspace",
      readOnly: false,
    }));

    return Promise.resolve(out.sort((left, right) => left.path.localeCompare(right.path)));
  }

  readFile(path: string): Promise<WorkspaceFileReadResult> {
    const normalizedPath = normalizeWorkspacePath(path);
    const file = this.files.get(normalizedPath);
    if (!file) {
      return Promise.reject(this.buildMissingError(normalizedPath));
    }

    const text = new TextDecoder().decode(file.bytes);

    return Promise.resolve({
      path: normalizedPath,
      name: normalizedPath.split("/").pop() ?? normalizedPath,
      size: file.bytes.byteLength,
      modifiedAt: file.modifiedAt,
      mimeType: "text/plain",
      kind: "text",
      sourceKind: "workspace",
      readOnly: false,
      text,
    });
  }

  writeBytes(path: string, bytes: Uint8Array): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    this.files.set(normalizedPath, {
      bytes,
      modifiedAt: Date.now(),
    });

    return Promise.resolve();
  }

  deleteFile(path: string): Promise<void> {
    const normalizedPath = normalizeWorkspacePath(path);
    this.files.delete(normalizedPath);
    return Promise.resolve();
  }

  renameFile(oldPath: string, newPath: string): Promise<void> {
    const normalizedOldPath = normalizeWorkspacePath(oldPath);
    const normalizedNewPath = normalizeWorkspacePath(newPath);
    const existing = this.files.get(normalizedOldPath);

    if (!existing) {
      return Promise.reject(this.buildMissingError(normalizedOldPath));
    }

    this.files.delete(normalizedOldPath);
    this.files.set(normalizedNewPath, {
      bytes: existing.bytes,
      modifiedAt: Date.now(),
    });

    return Promise.resolve();
  }
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

void test("workspace lists uploaded and connected-folder sources together when native is active", async () => {
  const workspaceBackend = new SourceBackend({
    kind: "opfs",
    label: "Sandboxed workspace",
    files: [
      { path: "uploads/plan.md", text: "workspace", modifiedAt: 50 },
    ],
  });

  const nativeBackend = new SourceBackend({
    kind: "native-directory",
    label: "Local folder",
    files: [
      { path: "folder/model-spec.md", text: "native", modifiedAt: 100 },
    ],
  });

  const workspace = new FilesWorkspace({
    initialBackend: nativeBackend,
    initialWorkspaceBackend: workspaceBackend,
  });

  const files = await workspace.listFiles();

  const uploaded = files.find((file) => file.path === "uploads/plan.md");
  assert.ok(uploaded);
  assert.equal(uploaded?.locationKind, "workspace");

  const connected = files.find((file) => file.path === "folder/model-spec.md");
  assert.ok(connected);
  assert.equal(connected?.locationKind, "native-directory");

  const builtin = files.find((file) => file.path === "assistant-docs/docs/extensions.md");
  assert.ok(builtin);
  assert.equal(builtin?.locationKind, "builtin-doc");
});

void test("workspace read/rename/delete can target a specific source when paths collide", async () => {
  const workspaceBackend = new SourceBackend({
    kind: "opfs",
    label: "Sandboxed workspace",
    files: [
      { path: "shared.txt", text: "workspace copy", modifiedAt: 10 },
    ],
  });

  const nativeBackend = new SourceBackend({
    kind: "native-directory",
    label: "Local folder",
    files: [
      { path: "shared.txt", text: "native copy", modifiedAt: 20 },
    ],
  });

  const workspace = new FilesWorkspace({
    initialBackend: nativeBackend,
    initialWorkspaceBackend: workspaceBackend,
  });

  const defaultRead = await workspace.readFile("shared.txt", {
    mode: "text",
  });
  assert.equal(defaultRead.text, "native copy");
  assert.equal(defaultRead.locationKind, "native-directory");

  const explicitWorkspaceRead = await workspace.readFile("shared.txt", {
    mode: "text",
    locationKind: "workspace",
  });
  assert.equal(explicitWorkspaceRead.text, "workspace copy");
  assert.equal(explicitWorkspaceRead.locationKind, "workspace");

  await workspace.renameFile("shared.txt", "workspace-renamed.txt", {
    locationKind: "workspace",
  });

  await workspace.deleteFile("shared.txt", {
    locationKind: "native-directory",
  });

  const files = await workspace.listFiles();

  const workspaceRenamed = files.find((file) => file.path === "workspace-renamed.txt");
  assert.ok(workspaceRenamed);
  assert.equal(workspaceRenamed?.locationKind, "workspace");

  const nativeOriginal = files.find((file) => file.path === "shared.txt" && file.locationKind === "native-directory");
  assert.equal(nativeOriginal, undefined);
});

void test("workspace read falls back to uploaded files when native missing message uses WebKit wording", async () => {
  const workspaceBackend = new SourceBackend({
    kind: "opfs",
    label: "Sandboxed workspace",
    files: [
      { path: "imports/source.csv", text: "company\nAcme", modifiedAt: 10 },
    ],
  });

  const nativeBackend = new SourceBackend({
    kind: "native-directory",
    label: "Local folder",
    files: [],
    missingMessageTemplate: "The object can not be found here.",
  });

  const workspace = new FilesWorkspace({
    initialBackend: nativeBackend,
    initialWorkspaceBackend: workspaceBackend,
  });

  const read = await workspace.readFile("imports/source.csv", {
    mode: "text",
  });

  assert.equal(read.locationKind, "workspace");
  assert.equal(read.text, "company\nAcme");
});

void test("path-only mutations follow the file's source when native is connected", async () => {
  const workspaceBackend = new SourceBackend({
    kind: "opfs",
    label: "Sandboxed workspace",
    files: [
      { path: "workspace-only.txt", text: "workspace", modifiedAt: 10 },
    ],
  });

  const nativeBackend = new SourceBackend({
    kind: "native-directory",
    label: "Local folder",
    files: [
      { path: "native-only.txt", text: "native", modifiedAt: 20 },
    ],
  });

  const workspace = new FilesWorkspace({
    initialBackend: nativeBackend,
    initialWorkspaceBackend: workspaceBackend,
  });

  await workspace.deleteFile("workspace-only.txt");

  const files = await workspace.listFiles();
  assert.equal(files.some((file) => file.path === "workspace-only.txt"), false);
  assert.equal(files.some((file) => file.path === "native-only.txt" && file.locationKind === "native-directory"), true);
});

void test("path-only mutation rejects ambiguous duplicates across workspace and native", async () => {
  const workspaceBackend = new SourceBackend({
    kind: "opfs",
    label: "Sandboxed workspace",
    files: [
      { path: "shared.txt", text: "workspace", modifiedAt: 10 },
    ],
  });

  const nativeBackend = new SourceBackend({
    kind: "native-directory",
    label: "Local folder",
    files: [
      { path: "shared.txt", text: "native", modifiedAt: 20 },
    ],
  });

  const workspace = new FilesWorkspace({
    initialBackend: nativeBackend,
    initialWorkspaceBackend: workspaceBackend,
  });

  await assert.rejects(
    () => workspace.deleteFile("shared.txt"),
    /exists in both uploaded files and the connected folder/i,
  );

  const files = await workspace.listFiles();
  assert.equal(files.some((file) => file.path === "shared.txt" && file.locationKind === "workspace"), true);
  assert.equal(files.some((file) => file.path === "shared.txt" && file.locationKind === "native-directory"), true);
});

void test("importFiles defaults to workspace source when native folder is connected", async () => {
  const workspaceBackend = new SourceBackend({
    kind: "opfs",
    label: "Sandboxed workspace",
    files: [],
  });

  const nativeBackend = new SourceBackend({
    kind: "native-directory",
    label: "Local folder",
    files: [],
  });

  const workspace = new FilesWorkspace({
    initialBackend: nativeBackend,
    initialWorkspaceBackend: workspaceBackend,
  });

  const file = new File(["uploaded content"], "uploaded.txt", {
    type: "text/plain",
  });

  const importedCount = await workspace.importFiles([file]);
  assert.equal(importedCount, 1);

  const files = await workspace.listFiles();
  const uploaded = files.find((entry) => entry.path === "uploaded.txt");
  assert.ok(uploaded);
  assert.equal(uploaded?.locationKind, "workspace");

  const uploadedRead = await workspace.readFile("uploaded.txt", {
    mode: "text",
    locationKind: "workspace",
  });
  assert.equal(uploadedRead.text, "uploaded content");
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

void test("workspace removes stale scratch files older than 24h", async () => {
  const now = Date.now();
  const staleTimestamp = now - (24 * 60 * 60 * 1000) - 1000;
  const freshTimestamp = now - (60 * 1000);

  const backend = new CleanupTestBackend([
    { path: "scratch/stale.txt", modifiedAt: staleTimestamp, content: "old" },
    { path: "scratch/fresh.txt", modifiedAt: freshTimestamp, content: "new" },
    { path: "notes/keep.md", modifiedAt: staleTimestamp, content: "keep" },
  ]);

  const workspace = new FilesWorkspace({ initialBackend: backend });
  const files = await workspace.listFiles();
  const workspacePaths = files
    .filter((file) => file.sourceKind === "workspace")
    .map((file) => file.path);

  assert.ok(!workspacePaths.includes("scratch/stale.txt"));
  assert.ok(workspacePaths.includes("scratch/fresh.txt"));
  assert.ok(workspacePaths.includes("notes/keep.md"));
});

void test("workspace scratch cleanup fails open when delete throws", async () => {
  const staleTimestamp = Date.now() - (24 * 60 * 60 * 1000) - 1000;

  const backend = new CleanupTestBackend(
    [
      { path: "scratch/stale.txt", modifiedAt: staleTimestamp, content: "old" },
      { path: "notes/keep.md", modifiedAt: staleTimestamp, content: "keep" },
    ],
    ["scratch/stale.txt"],
  );

  const workspace = new FilesWorkspace({ initialBackend: backend });
  const files = await workspace.listFiles();
  const workspacePaths = files
    .filter((file) => file.sourceKind === "workspace")
    .map((file) => file.path);

  assert.ok(workspacePaths.includes("notes/keep.md"));
  assert.ok(workspacePaths.includes("scratch/stale.txt"));
});
