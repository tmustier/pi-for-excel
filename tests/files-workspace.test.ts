import assert from "node:assert/strict";
import { test } from "node:test";

import { getFilesWorkspace } from "../src/files/workspace.ts";

interface CollisionSeedBackend {
  writeBytes(path: string, bytes: Uint8Array, mimeTypeHint?: string): Promise<void>;
}

function isCollisionSeedBackend(value: unknown): value is CollisionSeedBackend {
  if (typeof value !== "object" || value === null) {
    return false;
  }

  return typeof Reflect.get(value, "writeBytes") === "function";
}

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
  await resetWorkspace();
  const workspace = getFilesWorkspace();

  await workspace.listFiles();
  const backend: unknown = Reflect.get(workspace, "backend");
  assert.ok(isCollisionSeedBackend(backend));
  if (!isCollisionSeedBackend(backend)) return;

  await backend.writeBytes(
    "assistant-docs/docs/extensions.md",
    new TextEncoder().encode("legacy collision payload"),
    "text/plain",
  );

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
