/**
 * Built-in, read-only docs exposed through the files workspace.
 *
 * Every non-archive markdown doc is auto-discovered and bundled so the
 * assistant can answer setup/runtime questions accurately.
 */

import { loadRawMarkdownFromTestGlob } from "../utils/test-raw-markdown-glob.js";
import { normalizeWorkspacePath } from "./path.js";
import type { WorkspaceFileEntry, WorkspaceFileReadResult } from "./types.js";

interface BuiltinDocSource {
  path: string;
  markdown: string;
}

interface BuiltinWorkspaceDoc {
  path: string;
  name: string;
  markdown: string;
  size: number;
}

const BUILTIN_DOCS_PREFIX = "assistant-docs";
const BUILTIN_DOC_TIMESTAMP = Date.now();
const EXCLUDED_DOC_PREFIXES: readonly string[] = [
  "docs/archive/",
  "docs/release-smoke-runs/",
];

function toByteLength(text: string): number {
  return new TextEncoder().encode(text).byteLength;
}

function toRepoRelativePath(globPath: string): string {
  const normalized = globPath.replaceAll("\\", "/");
  if (normalized.startsWith("./")) {
    return normalized.slice(2);
  }

  let cursor = normalized;
  while (cursor.startsWith("../")) {
    cursor = cursor.slice(3);
  }

  return cursor;
}

function isBundledDocPath(repoRelativePath: string): boolean {
  if (repoRelativePath === "README.md") {
    return true;
  }

  if (!repoRelativePath.startsWith("docs/")) {
    return false;
  }

  return !EXCLUDED_DOC_PREFIXES.some((prefix) => repoRelativePath.startsWith(prefix));
}

function toBuiltinDocSource(globPath: string, markdown: string): BuiltinDocSource | null {
  const repoRelativePath = toRepoRelativePath(globPath);
  if (!isBundledDocPath(repoRelativePath)) {
    return null;
  }

  return {
    path: `${BUILTIN_DOCS_PREFIX}/${repoRelativePath}`,
    markdown,
  };
}

function isBuiltinDocSource(value: BuiltinDocSource | null): value is BuiltinDocSource {
  return value !== null;
}

function readBundledMarkdownByPath(): Record<string, string> {
  try {
    const rootReadmeMarkdownByPath = import.meta.glob<string>("../../README.md", {
      eager: true,
      query: "?raw",
      import: "default",
    });

    const docsMarkdownByPath = import.meta.glob<string>([
      "../../docs/**/*.md",
      "!../../docs/archive/**/*.md",
      "!../../docs/release-smoke-runs/**/*.md",
    ], {
      eager: true,
      query: "?raw",
      import: "default",
    });

    return {
      ...rootReadmeMarkdownByPath,
      ...docsMarkdownByPath,
    };
  } catch {
    const rootReadmeMarkdownByPath = loadRawMarkdownFromTestGlob("../../README.md", import.meta.url);
    const docsMarkdownByPath = loadRawMarkdownFromTestGlob("../../docs/**/*.md", import.meta.url);

    return {
      ...rootReadmeMarkdownByPath,
      ...docsMarkdownByPath,
    };
  }
}

function buildBuiltinDocSources(): BuiltinDocSource[] {
  return Object.entries(readBundledMarkdownByPath())
    .map(([globPath, markdown]) => toBuiltinDocSource(globPath, markdown))
    .filter(isBuiltinDocSource)
    .sort((left, right) => left.path.localeCompare(right.path));
}

function toBuiltinWorkspaceDoc(source: BuiltinDocSource): BuiltinWorkspaceDoc {
  return {
    path: normalizeWorkspacePath(source.path),
    name: source.path.split("/").at(-1) ?? source.path,
    markdown: source.markdown,
    size: toByteLength(source.markdown),
  };
}

const BUILTIN_DOCS: readonly BuiltinWorkspaceDoc[] = buildBuiltinDocSources()
  .map((source) => toBuiltinWorkspaceDoc(source));

function mapBuiltinDocToFileEntry(doc: BuiltinWorkspaceDoc): WorkspaceFileEntry {
  return {
    path: doc.path,
    name: doc.name,
    size: doc.size,
    modifiedAt: BUILTIN_DOC_TIMESTAMP,
    mimeType: "text/markdown",
    kind: "text",
    sourceKind: "builtin-doc",
    readOnly: true,
  };
}

function mapBuiltinDocToReadResult(doc: BuiltinWorkspaceDoc): WorkspaceFileReadResult {
  return {
    ...mapBuiltinDocToFileEntry(doc),
    text: doc.markdown,
  };
}

export function listBuiltinWorkspaceDocs(): WorkspaceFileEntry[] {
  return BUILTIN_DOCS.map((doc) => mapBuiltinDocToFileEntry(doc));
}

export function getBuiltinWorkspaceDoc(path: string): WorkspaceFileReadResult | null {
  const normalizedPath = normalizeWorkspacePath(path);
  const found = BUILTIN_DOCS.find((doc) => doc.path === normalizedPath);
  if (!found) return null;

  return mapBuiltinDocToReadResult(found);
}

export function isBuiltinWorkspacePath(path: string): boolean {
  const normalizedPath = normalizeWorkspacePath(path);
  return BUILTIN_DOCS.some((doc) => doc.path === normalizedPath);
}
