/**
 * Built-in, read-only docs exposed through the files workspace.
 */

import projectReadmeMarkdown from "../../README.md?raw";
import docsReadmeMarkdown from "../../docs/README.md?raw";
import docsExtensionsMarkdown from "../../docs/extensions.md?raw";
import docsIntegrationsMarkdown from "../../docs/integrations-external-tools.md?raw";

import { normalizeWorkspacePath } from "./path.js";
import type { WorkspaceFileEntry, WorkspaceFileReadResult } from "./types.js";

interface BuiltinDocSource {
  path: string;
  markdown: string;
}

const BUILTIN_DOCS_PREFIX = "assistant-docs";
const BUILTIN_DOC_TIMESTAMP = Date.now();

const BUILTIN_DOC_SOURCES: readonly BuiltinDocSource[] = [
  {
    path: `${BUILTIN_DOCS_PREFIX}/README.md`,
    markdown: projectReadmeMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/README.md`,
    markdown: docsReadmeMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/extensions.md`,
    markdown: docsExtensionsMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/integrations-external-tools.md`,
    markdown: docsIntegrationsMarkdown,
  },
] as const;

interface BuiltinWorkspaceDoc {
  path: string;
  name: string;
  markdown: string;
  size: number;
}

function toByteLength(text: string): number {
  return new TextEncoder().encode(text).byteLength;
}

function toBuiltinWorkspaceDoc(source: BuiltinDocSource): BuiltinWorkspaceDoc {
  return {
    path: normalizeWorkspacePath(source.path),
    name: source.path.split("/").at(-1) ?? source.path,
    markdown: source.markdown,
    size: toByteLength(source.markdown),
  };
}

const BUILTIN_DOCS: readonly BuiltinWorkspaceDoc[] = BUILTIN_DOC_SOURCES
  .map((source) => toBuiltinWorkspaceDoc(source))
  .sort((left, right) => left.path.localeCompare(right.path));

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

