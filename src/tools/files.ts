/**
 * Experimental files workspace tool.
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";

import { formatBytes } from "../files/mime.js";
import { normalizeWorkspacePath } from "../files/path.js";
import {
  getFilesWorkspace,
  type FilesWorkspaceAuditContext,
  type WorkspaceReadMode,
} from "../files/workspace.js";
import type {
  FilesDeleteDetails,
  FilesListDetails,
  FilesReadDetails,
  FilesToolDetails,
  FilesWriteDetails,
} from "./tool-details.js";

const schema = Type.Object({
  action: Type.Union([
    Type.Literal("list"),
    Type.Literal("read"),
    Type.Literal("write"),
    Type.Literal("delete"),
  ], {
    description: "Workspace action: list, read, write, or delete.",
  }),
  path: Type.Optional(Type.String({
    description: "Workspace-relative file path (required for read/write/delete). For list, optional folder prefix to filter results (e.g. \"notes/\").",
  })),
  content: Type.Optional(Type.String({
    description: "Content for write. Use plain text by default, or base64 when encoding=base64.",
  })),
  mode: Type.Optional(Type.Union([
    Type.Literal("auto"),
    Type.Literal("text"),
    Type.Literal("base64"),
  ], {
    description: "Read mode: auto (default), text, or base64.",
  })),
  encoding: Type.Optional(Type.Union([
    Type.Literal("text"),
    Type.Literal("base64"),
  ], {
    description: "Write encoding. Default: text.",
  })),
  mime_type: Type.Optional(Type.String({
    description: "Optional MIME type hint when writing files.",
  })),
  max_chars: Type.Optional(Type.Number({
    minimum: 128,
    maximum: 200000,
    description: "Maximum characters to return for read output (default: 20000).",
  })),
});

type Params = Static<typeof schema>;

const TOOL_AUDIT_CONTEXT: FilesWorkspaceAuditContext = {
  actor: "assistant",
  source: "tool:files",
};

function normalizeFolderPrefix(rawPath: string): string {
  const normalized = normalizeWorkspacePath(rawPath + "/placeholder").replace(/\/placeholder$/, "/");
  return normalized;
}

function filterByFolder<T extends { path: string }>(
  files: T[],
  rawFolder: string | undefined,
): T[] {
  const trimmed = rawFolder?.trim();
  if (!trimmed) return files;

  const prefix = normalizeFolderPrefix(trimmed);
  return files.filter((file) => file.path.startsWith(prefix));
}

function requirePath(path: string | undefined, action: Params["action"]): string {
  const trimmed = path?.trim();
  if (!trimmed) {
    throw new Error(`'path' is required for action='${action}'.`);
  }

  return trimmed;
}

function renderListMarkdown(args: {
  backendLabel: string;
  folderFilter?: string;
  totalCount: number;
  files: Array<{
    path: string;
    size: number;
    kind: string;
    mimeType: string;
    sourceKind: "workspace" | "builtin-doc";
    readOnly: boolean;
    workbookLabel?: string;
  }>;
}): string {
  const filterSuffix = args.folderFilter ? `, folder: ${args.folderFilter}` : "";
  if (args.files.length === 0) {
    const noFilesMsg = args.folderFilter
      ? `_No files in ${args.folderFilter}._`
      : "_No files yet._";
    return `Workspace files (${args.backendLabel}${filterSuffix}):\n\n${noFilesMsg}`;
  }

  const countNote = args.folderFilter && args.totalCount !== args.files.length
    ? ` (${args.files.length} of ${args.totalCount} total)`
    : "";
  const lines = [`Workspace files (${args.backendLabel}${filterSuffix})${countNote}:`, ""];
  for (const file of args.files) {
    const workbookSuffix = file.workbookLabel ? `, workbook: ${file.workbookLabel}` : "";
    const sourceSuffix = file.sourceKind === "builtin-doc" ? ", built-in doc" : "";
    const readOnlySuffix = file.readOnly ? ", read-only" : "";
    lines.push(`- ${file.path} (${formatBytes(file.size)}, ${file.kind}, ${file.mimeType}${sourceSuffix}${readOnlySuffix}${workbookSuffix})`);
  }

  return lines.join("\n");
}

function renderReadMarkdown(args: {
  path: string;
  size: number;
  mimeType: string;
  mode: "text" | "base64";
  content: string;
  truncated: boolean;
  sourceKind: "workspace" | "builtin-doc";
  readOnly: boolean;
  workbookLabel?: string;
}): string {
  const lines: string[] = [];
  const workbookSuffix = args.workbookLabel ? `, workbook: ${args.workbookLabel}` : "";
  const sourceSuffix = args.sourceKind === "builtin-doc" ? ", built-in doc" : "";
  const readOnlySuffix = args.readOnly ? ", read-only" : "";
  lines.push(`Read **${args.path}** (${formatBytes(args.size)}, ${args.mimeType}${sourceSuffix}${readOnlySuffix}${workbookSuffix})`);
  lines.push("");
  lines.push("```");
  lines.push(args.content);
  lines.push("```");

  if (args.truncated) {
    lines.push("");
    lines.push("⚠️ Output was truncated. Increase max_chars to read more.");
  }

  if (args.mode === "base64") {
    lines.push("");
    lines.push("(base64 output)");
  }

  return lines.join("\n");
}

function mapWorkbookTag(tag: {
  workbookId: string;
  workbookLabel: string;
  taggedAt: number;
} | undefined): {
  workbookId: string;
  workbookLabel: string;
  taggedAt: number;
} | undefined {
  if (!tag) return undefined;
  return {
    workbookId: tag.workbookId,
    workbookLabel: tag.workbookLabel,
    taggedAt: tag.taggedAt,
  };
}

export function createFilesTool(): AgentTool<typeof schema, FilesToolDetails> {
  return {
    name: "files",
    label: "Files",
    description:
      "Manage workspace files (list/read/write/delete). " +
      "Use this for artifacts like notes, CSV extracts, and generated documents.",
    parameters: schema,
    execute: async (_toolCallId: string, params: Params): Promise<AgentToolResult<FilesToolDetails>> => {
      const workspace = getFilesWorkspace();
      const backend = await workspace.getBackendStatus();

      if (params.action === "list") {
        const allFiles = await workspace.listFiles({
          audit: TOOL_AUDIT_CONTEXT,
        });

        const files = filterByFolder(allFiles, params.path);

        const details: FilesListDetails = {
          kind: "files_list",
          backend: backend.kind,
          count: files.length,
          files: files.map((file) => ({
            path: file.path,
            size: file.size,
            mimeType: file.mimeType,
            fileKind: file.kind,
            modifiedAt: file.modifiedAt,
            sourceKind: file.sourceKind,
            readOnly: file.readOnly,
            workbookTag: mapWorkbookTag(file.workbookTag),
          })),
        };

        return {
          content: [{
            type: "text",
            text: renderListMarkdown({
              backendLabel: backend.label,
              folderFilter: params.path?.trim(),
              totalCount: allFiles.length,
              files: files.map((file) => ({
                path: file.path,
                size: file.size,
                kind: file.kind,
                mimeType: file.mimeType,
                sourceKind: file.sourceKind,
                readOnly: file.readOnly,
                workbookLabel: file.workbookTag?.workbookLabel,
              })),
            }),
          }],
          details,
        };
      }

      if (params.action === "read") {
        const path = requirePath(params.path, "read");
        const mode: WorkspaceReadMode = params.mode ?? "auto";
        const maxChars = params.max_chars;

        const readResult = await workspace.readFile(path, {
          mode,
          maxChars,
          audit: TOOL_AUDIT_CONTEXT,
        });

        const outputMode: "text" | "base64" = readResult.text !== undefined ? "text" : "base64";
        const output = readResult.text ?? readResult.base64 ?? "";
        const details: FilesReadDetails = {
          kind: "files_read",
          backend: backend.kind,
          path: readResult.path,
          mode: outputMode,
          size: readResult.size,
          mimeType: readResult.mimeType,
          fileKind: readResult.kind,
          sourceKind: readResult.sourceKind,
          readOnly: readResult.readOnly,
          truncated: readResult.truncated === true,
          workbookTag: mapWorkbookTag(readResult.workbookTag),
        };

        return {
          content: [{
            type: "text",
            text: renderReadMarkdown({
              path: readResult.path,
              size: readResult.size,
              mimeType: readResult.mimeType,
              mode: outputMode,
              content: output,
              truncated: readResult.truncated === true,
              sourceKind: readResult.sourceKind,
              readOnly: readResult.readOnly,
              workbookLabel: readResult.workbookTag?.workbookLabel,
            }),
          }],
          details,
        };
      }

      if (params.action === "write") {
        const path = requirePath(params.path, "write");
        const normalizedPath = normalizeWorkspacePath(path);
        const content = params.content ?? "";
        const encoding = params.encoding ?? "text";

        if (encoding === "base64") {
          await workspace.writeBase64File(path, content, params.mime_type, {
            audit: TOOL_AUDIT_CONTEXT,
          });
        } else {
          await workspace.writeTextFile(path, content, params.mime_type, {
            audit: TOOL_AUDIT_CONTEXT,
          });
        }

        const filesAfterWrite = await workspace.listFiles();
        const writtenFile = filesAfterWrite.find((file) => file.path === normalizedPath);

        const details: FilesWriteDetails = {
          kind: "files_write",
          backend: backend.kind,
          path: normalizedPath,
          encoding,
          chars: content.length,
          workbookTag: mapWorkbookTag(writtenFile?.workbookTag),
        };

        return {
          content: [{
            type: "text",
            text: `Wrote **${normalizedPath}** (${content.length.toLocaleString()} chars, ${encoding}).`,
          }],
          details,
        };
      }

      const path = requirePath(params.path, "delete");
      const normalizedPath = normalizeWorkspacePath(path);
      const filesBeforeDelete = await workspace.listFiles();
      const deletedFile = filesBeforeDelete.find((file) => file.path === normalizedPath);

      await workspace.deleteFile(normalizedPath, {
        audit: TOOL_AUDIT_CONTEXT,
      });

      const details: FilesDeleteDetails = {
        kind: "files_delete",
        backend: backend.kind,
        path: normalizedPath,
        workbookTag: mapWorkbookTag(deletedFile?.workbookTag),
      };

      return {
        content: [{ type: "text", text: `Deleted **${normalizedPath}**.` }],
        details,
      };
    },
  };
}
