/**
 * Experimental files workspace tool.
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";

import { formatBytes } from "../files/mime.js";
import { getFilesWorkspace, type WorkspaceReadMode } from "../files/workspace.js";
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
    description: "Workspace-relative file path (required for read/write/delete).",
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

function requirePath(path: string | undefined, action: Params["action"]): string {
  const trimmed = path?.trim();
  if (!trimmed) {
    throw new Error(`'path' is required for action='${action}'.`);
  }

  return trimmed;
}

function renderListMarkdown(args: {
  backendLabel: string;
  files: Array<{
    path: string;
    size: number;
    kind: string;
    mimeType: string;
  }>;
}): string {
  if (args.files.length === 0) {
    return `Workspace files (${args.backendLabel}):\n\n_No files yet._`;
  }

  const lines = [`Workspace files (${args.backendLabel}):`, ""];
  for (const file of args.files) {
    lines.push(`- ${file.path} (${formatBytes(file.size)}, ${file.kind}, ${file.mimeType})`);
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
}): string {
  const lines: string[] = [];
  lines.push(`Read **${args.path}** (${formatBytes(args.size)}, ${args.mimeType})`);
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
        const files = await workspace.listFiles();
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
          })),
        };

        return {
          content: [{
            type: "text",
            text: renderListMarkdown({
              backendLabel: backend.label,
              files: files.map((file) => ({
                path: file.path,
                size: file.size,
                kind: file.kind,
                mimeType: file.mimeType,
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
          truncated: readResult.truncated === true,
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
            }),
          }],
          details,
        };
      }

      if (params.action === "write") {
        const path = requirePath(params.path, "write");
        const content = params.content ?? "";
        const encoding = params.encoding ?? "text";

        if (encoding === "base64") {
          await workspace.writeBase64File(path, content, params.mime_type);
        } else {
          await workspace.writeTextFile(path, content, params.mime_type);
        }

        const details: FilesWriteDetails = {
          kind: "files_write",
          backend: backend.kind,
          path,
          encoding,
          chars: content.length,
        };

        return {
          content: [{
            type: "text",
            text: `Wrote **${path}** (${content.length.toLocaleString()} chars, ${encoding}).`,
          }],
          details,
        };
      }

      const path = requirePath(params.path, "delete");
      await workspace.deleteFile(path);

      const details: FilesDeleteDetails = {
        kind: "files_delete",
        backend: backend.kind,
        path,
      };

      return {
        content: [{ type: "text", text: `Deleted **${path}**.` }],
        details,
      };
    },
  };
}
