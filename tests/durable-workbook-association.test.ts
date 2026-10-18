import assert from "node:assert/strict";
import { test } from "node:test";
import type { WorkbookContext } from "../src/host/workbook-context.ts";
import type { SettingsWriter } from "../src/storage/local/settings-store.ts";
import { createDurableWorkbookAssociationResolver, listRetainedDurableAssociations } from "../src/workbook/durable-association.ts";

function settingsBag(): SettingsWriter & { clear(): void } {
  const values = new Map<string, unknown>();
  return { get: (key) => Promise.resolve(values.get(key)), set: (key, value) => { values.set(key, structuredClone(value)); return Promise.resolve(); },
    clear: () => values.clear() };
}

void test("unsaved chat identity retains its generation through first save and reload, not saved copies", async () => {
  const settings = settingsBag();
  let workbook: WorkbookContext = { workbookId: null, workbookName: "Book1", source: "unknown" };
  let documentValue: string | null = null;
  const document = { read: () => documentValue, write: (value: string) => { documentValue = value; return Promise.resolve(true); } };
  const host = { kind: "office" as const, getWorkbookContext: () => Promise.resolve(workbook),
    documentInstance: { read: () => "one", ensure: () => Promise.resolve("one") } };
  const resolve = createDurableWorkbookAssociationResolver({ host, settings, document });
  const unsaved = await resolve();
  assert.equal(unsaved.scopeId, "doc_instance:one");
  await unsaved.rememberGeneration("generation-one");
  const reopened = await createDurableWorkbookAssociationResolver({ host, settings, document })();
  assert.equal(reopened.expectedDatabaseId, "generation-one");
  workbook = { workbookId: "url_sha256:first", workbookName: "Saved.xlsx", source: "document.url" };
  const saved = await resolve();
  assert.equal(saved.databaseName, unsaved.databaseName);
  assert.equal(saved.expectedDatabaseId, "generation-one");
  assert.equal((await createDurableWorkbookAssociationResolver({ host, settings, document })()).databaseName, unsaved.databaseName);
  // Browser-origin loss must not turn a known generation into a fresh database.
  settings.clear();
  assert.equal((await createDurableWorkbookAssociationResolver({ host, settings, document })()).expectedDatabaseId, "generation-one");
  workbook = { ...workbook, workbookId: "url_sha256:copy" };
  const copy = await createDurableWorkbookAssociationResolver({ host, settings, document })();
  assert.notEqual(copy.databaseName, unsaved.databaseName);
  assert.equal(copy.expectedDatabaseId, null);
  await copy.rememberGeneration("copy-generation");
  const reloadedCopy = await createDurableWorkbookAssociationResolver({ host, settings, document })();
  assert.equal(reloadedCopy.expectedDatabaseId, "copy-generation");
  assert.deepEqual((await listRetainedDurableAssociations(settings)).map((row) => row.databaseId).sort(),
    ["copy-generation", "generation-one"]);
});

void test("distinct unsaved books and identity-less panes do not share pending work; browser fallback remains usable", async () => {
  const settings = settingsBag();
  const base = { kind: "office" as const, getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: null, workbookName: null, source: "unknown",
  }), documentInstance: null };
  const one = await createDurableWorkbookAssociationResolver({ host: base, settings, createId: () => "one" })();
  const two = await createDurableWorkbookAssociationResolver({ host: base, settings, createId: () => "two" })();
  assert.notEqual(one.databaseName, two.databaseName);
  await one.rememberGeneration("retained");
  assert.equal(one.documentBacked, false);
  assert.equal((await listRetainedDurableAssociations(settings)).find((row) => row.scopeId === one.scopeId)?.databaseId, "retained");
  const browser = await createDurableWorkbookAssociationResolver({ host: { ...base, kind: "browser" }, settings })();
  await browser.rememberGeneration("browser-generation");
  assert.equal((await createDurableWorkbookAssociationResolver({ host: { ...base, kind: "browser" }, settings })()).expectedDatabaseId,
    "browser-generation");
});

void test("association corruption and generation replacement fail rather than forgetting expected data", async () => {
  const settings = settingsBag();
  const host = { kind: "browser" as const, getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: null, workbookName: null, source: "unknown",
  }), documentInstance: null };
  const resolve = createDurableWorkbookAssociationResolver({ host, settings });
  await (await resolve()).rememberGeneration("original");
  await assert.rejects((await resolve()).rememberGeneration("replacement"), /generation changed/);
  await settings.set("workbook.durable.v1.browser-library", { wrong: true });
  await assert.rejects(resolve(), /Invalid local/);
});

void test("unchanged document settings are not rewritten and one handle cannot replace its generation", async () => {
  const settings = settingsBag();
  let value: string | null = null;
  let writes = 0;
  const document = { read: () => value, write: (text: string) => { writes++; value = text; return Promise.resolve(true); } };
  const host = { kind: "office" as const, getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: "url_sha256:book", workbookName: "https://secret.invalid/Book.xlsx?token=secret", source: "document.url",
  }) };
  const resolve = createDurableWorkbookAssociationResolver({ host, settings, document });
  const association = await resolve();
  await association.rememberGeneration("first");
  await association.rememberGeneration("first");
  await resolve();
  assert.equal(writes, 2);
  await assert.rejects(association.rememberGeneration("second"), /generation changed/);
  assert.equal((await listRetainedDurableAssociations(settings))[0]?.name, "Book.xlsx");
});
