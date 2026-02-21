/**
 * Built-in, read-only docs exposed through the files workspace.
 *
 * Every non-archive doc is bundled so the assistant can answer user
 * questions accurately instead of hallucinating setup instructions.
 */

import projectReadmeMarkdown from "../../README.md?raw";
import docsReadmeMarkdown from "../../docs/README.md?raw";
import docsAgentSkillsInteropMarkdown from "../../docs/agent-skills-interop.md?raw";
import docsCompactionMarkdown from "../../docs/compaction.md?raw";
import docsContextManagementPolicyMarkdown from "../../docs/context-management-policy.md?raw";
import docsDeployVercelMarkdown from "../../docs/deploy-vercel.md?raw";
import docsExtensionsMarkdown from "../../docs/extensions.md?raw";
import docsFilesWorkspaceMarkdown from "../../docs/files-workspace.md?raw";
import docsInstallMarkdown from "../../docs/install.md?raw";
import docsIntegrationsMarkdown from "../../docs/integrations-external-tools.md?raw";
import docsManualFullBackupsMarkdown from "../../docs/manual-full-backups.md?raw";
import docsModelUpdatesMarkdown from "../../docs/model-updates.md?raw";
import docsPythonBridgeContractMarkdown from "../../docs/python-bridge-contract.md?raw";
import docsSecurityThreatModelMarkdown from "../../docs/security-threat-model.md?raw";
import docsTmuxBridgeContractMarkdown from "../../docs/tmux-bridge-contract.md?raw";
import docsReleaseNotesV070Markdown from "../../docs/release-notes/v0.7.0-pre.md?raw";
import docsReleaseNotesV080Markdown from "../../docs/release-notes/v0.8.0-pre.md?raw";

import { normalizeWorkspacePath } from "./path.js";
import type { WorkspaceFileEntry, WorkspaceFileReadResult } from "./types.js";

interface BuiltinDocSource {
  path: string;
  markdown: string;
}

const BUILTIN_DOCS_PREFIX = "assistant-docs";
const BUILTIN_DOC_TIMESTAMP = Date.now();

const BUILTIN_DOC_SOURCES: readonly BuiltinDocSource[] = [
  // Project root
  {
    path: `${BUILTIN_DOCS_PREFIX}/README.md`,
    markdown: projectReadmeMarkdown,
  },

  // Docs index
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/README.md`,
    markdown: docsReadmeMarkdown,
  },

  // Guides
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/install.md`,
    markdown: docsInstallMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/deploy-vercel.md`,
    markdown: docsDeployVercelMarkdown,
  },

  // Runtime features
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/extensions.md`,
    markdown: docsExtensionsMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/integrations-external-tools.md`,
    markdown: docsIntegrationsMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/agent-skills-interop.md`,
    markdown: docsAgentSkillsInteropMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/compaction.md`,
    markdown: docsCompactionMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/manual-full-backups.md`,
    markdown: docsManualFullBackupsMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/files-workspace.md`,
    markdown: docsFilesWorkspaceMarkdown,
  },

  // Architecture & policy
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/context-management-policy.md`,
    markdown: docsContextManagementPolicyMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/security-threat-model.md`,
    markdown: docsSecurityThreatModelMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/model-updates.md`,
    markdown: docsModelUpdatesMarkdown,
  },

  // Feature-flagged bridge contracts
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/tmux-bridge-contract.md`,
    markdown: docsTmuxBridgeContractMarkdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/python-bridge-contract.md`,
    markdown: docsPythonBridgeContractMarkdown,
  },

  // Release notes
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/release-notes/v0.7.0-pre.md`,
    markdown: docsReleaseNotesV070Markdown,
  },
  {
    path: `${BUILTIN_DOCS_PREFIX}/docs/release-notes/v0.8.0-pre.md`,
    markdown: docsReleaseNotesV080Markdown,
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

