import assert from "node:assert/strict";
import { test } from "node:test";

import "./files-dialog-actions.test.mjs";
import "./files-dialog-mime.test.mjs";
import "./files-dialog-paths.test.mjs";

import type { WorkspaceBackendStatus, WorkspaceFileEntry } from "../src/files/types.ts";
import {
  buildFilesDialogSections,
  filterFilesDialogEntries,
  fileMatchesFilesDialogFilter,
  normalizeFilesDialogFilterText,
  resolveFilesDialogBadge,
  resolveFilesDialogConnectFolderButtonState,
  resolveFilesDialogSourceLabel,
} from "../src/ui/files-dialog-filtering.ts";

function makeFile(path: string, overrides: Partial<WorkspaceFileEntry> = {}): WorkspaceFileEntry {
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

const connectedBackendStatus: WorkspaceBackendStatus = {
  kind: "native-directory",
  label: "Local folder",
  nativeSupported: true,
  nativeConnected: true,
  nativeDirectoryName: "Project Docs",
};

void test("normalizeFilesDialogFilterText trims and lowercases", () => {
  assert.equal(normalizeFilesDialogFilterText("  Notes/Q1  "), "notes/q1");
  assert.equal(normalizeFilesDialogFilterText("   "), "");
});

void test("fileMatchesFilesDialogFilter uses case-insensitive path substring", () => {
  const file = makeFile("notes/Quarterly-Plan.md");

  assert.equal(fileMatchesFilesDialogFilter({ file, filterText: "quarterly" }), true);
  assert.equal(fileMatchesFilesDialogFilter({ file, filterText: "NOTES/QU" }), true);
  assert.equal(fileMatchesFilesDialogFilter({ file, filterText: "missing" }), false);
});

void test("filterFilesDialogEntries returns all files when filter is empty", () => {
  const files = [
    makeFile("notes/index.md"),
    makeFile("imports/budget.csv"),
  ];

  const result = filterFilesDialogEntries({
    files,
    filterText: "   ",
  });

  assert.deepEqual(result.map((file) => file.path), [
    "notes/index.md",
    "imports/budget.csv",
  ]);
});

void test("filterFilesDialogEntries returns only matching paths", () => {
  const files = [
    makeFile("notes/index.md"),
    makeFile("notes/meeting-notes.md"),
    makeFile("imports/raw.csv"),
  ];

  const result = filterFilesDialogEntries({
    files,
    filterText: "notes/meeting",
  });

  assert.deepEqual(result.map((file) => file.path), ["notes/meeting-notes.md"]);
});

void test("buildFilesDialogSections groups files by category with folder groups", () => {
  const files = [
    makeFile("report.xlsx", { modifiedAt: 35 }),
    makeFile("data/raw.csv", { modifiedAt: 30 }),
    makeFile("notes/summary.md", { modifiedAt: 12 }),
    makeFile("notes/index.md", { modifiedAt: 20 }),
    makeFile("skills/my-skill/SKILL.md", { modifiedAt: 10 }),
    makeFile("skills/external/data-clean/SKILL.md", { modifiedAt: 8 }),
    makeFile("src/main.ts", { modifiedAt: 16, locationKind: "native-directory" }),
    makeFile("src/utils.ts", { modifiedAt: 14, locationKind: "native-directory" }),
    makeFile("README.md", { modifiedAt: 18, locationKind: "native-directory" }),
    makeFile("assistant-docs/docs/extensions.md", {
      sourceKind: "builtin-doc",
      locationKind: "builtin-doc",
      readOnly: true,
      modifiedAt: 1,
    }),
  ];

  const sections = buildFilesDialogSections({
    files,
    filterText: "",
    backendStatus: connectedBackendStatus,
  });

  assert.deepEqual(sections.map((s) => s.label), [
    "YOUR FILES",
    "PI'S NOTES",
    "SKILLS",
    "FROM PROJECT DOCS",
    "BUILT-IN DOCS",
  ]);

  // YOUR FILES: root file + folder group "data"
  const yourFiles = sections[0];
  assert.ok(yourFiles);
  assert.deepEqual(yourFiles.files.map((f) => f.path), ["report.xlsx"]);
  assert.equal(yourFiles.folders.length, 1);
  assert.equal(yourFiles.folders[0]?.name, "data");
  assert.deepEqual(yourFiles.folders[0]?.files.map((f) => f.path), ["data/raw.csv"]);

  // PI'S NOTES: all flat (notes/ prefix stripped)
  const notes = sections[1];
  assert.ok(notes);
  assert.deepEqual(notes.files.map((f) => f.path), ["notes/index.md", "notes/summary.md"]);
  assert.equal(notes.folders.length, 0);

  // SKILLS: two folder groups (external/ prefix stripped for data-clean)
  const skills = sections[2];
  assert.ok(skills);
  assert.equal(skills.files.length, 0);
  assert.equal(skills.folders.length, 2);
  assert.equal(skills.folders[0]?.name, "data-clean");
  assert.equal(skills.folders[1]?.name, "my-skill");

  // Connected folder: root file + folder group "src"
  const connected = sections[3];
  assert.ok(connected);
  assert.deepEqual(connected.files.map((f) => f.path), ["README.md"]);
  assert.equal(connected.folders.length, 1);
  assert.equal(connected.folders[0]?.name, "src");
  assert.deepEqual(connected.folders[0]?.files.map((f) => f.path), ["src/main.ts", "src/utils.ts"]);

  // BUILT-IN DOCS: always flat
  const builtin = sections[4];
  assert.ok(builtin);
  assert.equal(builtin.folders.length, 0);
});

void test("resolveFilesDialogBadge follows priority rules", () => {
  const builtInWithTag = makeFile("assistant-docs/docs/README.md", {
    sourceKind: "builtin-doc",
    locationKind: "builtin-doc",
    readOnly: true,
    workbookTag: {
      workbookId: "wb-a",
      workbookLabel: "Budget.xlsx",
      taggedAt: 1,
    },
  });
  assert.deepEqual(resolveFilesDialogBadge(builtInWithTag), { tone: "muted", label: "Read only" });

  const tagged = makeFile("uploads/plan.md", {
    workbookTag: {
      workbookId: "wb-a",
      workbookLabel: "Budget.xlsx",
      taggedAt: 1,
    },
  });
  assert.deepEqual(resolveFilesDialogBadge(tagged), { tone: "ok", label: "Budget.xlsx" });

  const agentNotes = makeFile("notes/today.md");
  assert.deepEqual(resolveFilesDialogBadge(agentNotes), { tone: "muted", label: "Agent" });

  const connected = makeFile("folder/readme.md", { locationKind: "native-directory" });
  assert.deepEqual(resolveFilesDialogBadge(connected), { tone: "info", label: "Folder" });

  const uploaded = makeFile("uploads/raw.csv");
  assert.equal(resolveFilesDialogBadge(uploaded), null);
});

void test("resolveFilesDialogSourceLabel maps each source", () => {
  assert.equal(resolveFilesDialogSourceLabel(makeFile("assistant-docs/docs/README.md", {
    sourceKind: "builtin-doc",
    locationKind: "builtin-doc",
    readOnly: true,
  })), "Pi documentation");

  assert.equal(resolveFilesDialogSourceLabel(makeFile("notes/today.md")), "Written by agent");
  assert.equal(resolveFilesDialogSourceLabel(makeFile("folder/readme.md", { locationKind: "native-directory" })), "Local file");
  assert.equal(resolveFilesDialogSourceLabel(makeFile("uploads/raw.csv")), "Uploaded");
});

void test("resolveFilesDialogConnectFolderButtonState reflects backend status", () => {
  assert.deepEqual(resolveFilesDialogConnectFolderButtonState(null), {
    hidden: true,
    disabled: true,
    label: "Connect folder",
    title: "",
  });

  assert.deepEqual(resolveFilesDialogConnectFolderButtonState({
    kind: "opfs",
    label: "Sandboxed workspace",
    nativeSupported: false,
    nativeConnected: false,
  }), {
    hidden: true,
    disabled: true,
    label: "Connect folder",
    title: "",
  });

  assert.deepEqual(resolveFilesDialogConnectFolderButtonState({
    kind: "opfs",
    label: "Sandboxed workspace",
    nativeSupported: true,
    nativeConnected: false,
  }), {
    hidden: false,
    disabled: false,
    label: "Connect folder",
    title: "Connect local folder",
  });

  assert.deepEqual(resolveFilesDialogConnectFolderButtonState(connectedBackendStatus), {
    hidden: false,
    disabled: true,
    label: "Connected ✓",
    title: "Folder already connected",
  });
});

