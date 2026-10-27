/**
 * Consent gates for assistant actions that change local state outside the
 * workbook: Files workspace writes/deletes and extension lifecycle changes.
 *
 * Rule (see src/tools/DECISIONS.md, "Consent outside the workbook"):
 * - Confirm mode asks before every such change.
 * - Auto mode still asks before running new code (extension install/enable)
 *   and before irreversible loss: deleting or overwriting a file Pi did not
 *   create in this session, or uninstalling an extension.
 */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import { Type } from "typebox";
import { Value } from "typebox/value";

import { formatBytes } from "../../files/mime.js";
import { normalizeWorkspacePath } from "../../files/path.js";
import { getFilesWorkspace } from "../../files/workspace.js";
import { t } from "../../language/index.js";
import type {
  ExperimentalToolGateDependencies,
  FilesGateFileInfo,
  LocalChangeApprovalRequest,
} from "./types.js";

// The subset of each tool's parameters the gate reads. Extra fields are
// allowed; params that fail the schema go straight to the tool, which
// rejects them with its own validation error.
const FilesGateParamsSchema = Type.Object({
  action: Type.String(),
  path: Type.Optional(Type.String()),
});

const ExtensionsGateParamsSchema = Type.Object({
  action: Type.String(),
  extension_id: Type.Optional(Type.String()),
  name: Type.Optional(Type.String()),
  code: Type.Optional(Type.String()),
  enabled: Type.Optional(Type.Boolean()),
  replace_existing: Type.Optional(Type.Boolean()),
  capabilities: Type.Optional(Type.Array(Type.String())),
});

function nonEmpty(value: string | undefined): string | undefined {
  return value?.trim() || undefined;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) throw new Error("Aborted");
}

async function defaultInspectWorkspaceFile(path: string): Promise<FilesGateFileInfo | null> {
  const workspace = getFilesWorkspace();
  const entry = (await workspace.listFiles()).find((file) => file.path === path);
  if (!entry) return null;

  if (entry.locationKind === "native-directory") {
    const status = await workspace.getBackendStatus();
    return {
      size: entry.size,
      readOnly: entry.readOnly,
      ...(status.nativeDirectoryName !== undefined ? { nativeFolderName: status.nativeDirectoryName } : {}),
    };
  }

  return { size: entry.size, readOnly: entry.readOnly };
}

async function requireApproval(
  dependencies: ExperimentalToolGateDependencies,
  request: LocalChangeApprovalRequest,
  cancelledMessage: string,
  signal: AbortSignal | undefined,
): Promise<void> {
  throwIfAborted(signal);
  const requestApproval = dependencies.requestLocalChangeApproval;
  if (!requestApproval) {
    throw new Error("This change needs your approval, but the approval dialog is unavailable.");
  }
  const approved = await requestApproval(signal ? { ...request, signal } : request);
  throwIfAborted(signal);
  if (!approved) {
    throw new Error(cancelledMessage);
  }
}

function getMode(dependencies: ExperimentalToolGateDependencies): Promise<"yolo" | "safe"> {
  return dependencies.getExecutionMode?.() ?? Promise.resolve("safe");
}

export function wrapFilesToolWithChangeGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  const inspect = dependencies.inspectWorkspaceFile ?? defaultInspectWorkspaceFile;
  const assistantCreated = dependencies.assistantCreatedFiles ?? new Set<string>();

  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      throwIfAborted(signal);

      const parsed = Value.Check(FilesGateParamsSchema, params) ? params : null;
      const action = parsed?.action;
      const rawPath = nonEmpty(parsed?.path);
      if ((action !== "write" && action !== "delete") || rawPath === undefined) {
        return tool.execute(toolCallId, params, signal, onUpdate);
      }

      let path: string;
      try {
        path = normalizeWorkspacePath(rawPath);
      } catch {
        // Invalid paths fail inside the tool with its own message.
        return tool.execute(toolCallId, params, signal, onUpdate);
      }

      const existing = await inspect(path);
      // Built-in docs are read-only; the tool rejects the change itself.
      if (existing?.readOnly === true) {
        return tool.execute(toolCallId, params, signal, onUpdate);
      }

      const mode = await getMode(dependencies);
      const losesUserData = existing !== null && !assistantCreated.has(path);
      if (mode === "safe" || losesUserData) {
        const location = existing?.nativeFolderName !== undefined
          ? t("gate.files.locationFolder", { folder: existing.nativeFolderName })
          : t("gate.files.locationWorkspace");
        const size = existing ? formatBytes(existing.size) : "";
        const request: LocalChangeApprovalRequest = action === "delete"
          ? {
            toolName: tool.name,
            title: t("gate.files.deleteTitle"),
            message: t("gate.files.deleteMessage", { path, size, location }),
            confirmLabel: t("gate.files.deleteConfirm"),
            destructive: true,
          }
          : existing !== null
            ? {
              toolName: tool.name,
              title: t("gate.files.replaceTitle"),
              message: t("gate.files.replaceMessage", { path, size, location }),
              confirmLabel: t("gate.files.replaceConfirm"),
              destructive: true,
            }
            : {
              toolName: tool.name,
              title: t("gate.files.createTitle"),
              message: t("gate.files.createMessage", { path, location }),
              confirmLabel: t("gate.files.createConfirm"),
              destructive: false,
            };

        await requireApproval(dependencies, request, `Files ${action} of '${path}' cancelled by user.`, signal);
        throwIfAborted(signal);
      }

      const result = await tool.execute(toolCallId, params, signal, onUpdate);
      if (action === "write" && existing === null) assistantCreated.add(path);
      if (action === "delete") assistantCreated.delete(path);
      return result;
    },
  };
}

export function wrapExtensionsManagerToolWithChangeGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      throwIfAborted(signal);

      if (!Value.Check(ExtensionsGateParamsSchema, params)) {
        return tool.execute(toolCallId, params, signal, onUpdate);
      }
      const parsed = params;
      const action = parsed.action;
      const mode = await getMode(dependencies);
      const extensionId = nonEmpty(parsed.extension_id);
      const target = extensionId === undefined
        ? undefined
        : dependencies.listExtensions?.().find((entry) => entry.id === extensionId);
      const targetName = target?.name ?? extensionId ?? "?";
      const noneLabel = t("gate.extensions.none");

      if (action === "install_code") {
        const name = nonEmpty(parsed.name) ?? "?";
        const code = nonEmpty(parsed.code) ?? "";
        const requested = parsed.capabilities ?? [];
        const replacing = parsed.replace_existing !== false
          ? (dependencies.listExtensions?.() ?? []).filter((entry) => entry.name.toLowerCase() === name.toLowerCase())
          : [];

        const lines = [
          t("gate.extensions.installIntro", { name }),
          "",
          t("gate.extensions.extraPermissions", { list: requested.length > 0 ? requested.join(", ") : noneLabel }),
        ];
        if (replacing.length > 0) {
          lines.push(t("gate.extensions.replaces", { count: replacing.length }));
        }
        const codeLines = code.split("\n");
        const preview = codeLines.slice(0, 12).join("\n");
        lines.push("", t("gate.extensions.codeHeader", { lines: codeLines.length, chars: code.length }),
          codeLines.length > 12 ? `${preview}\n…` : preview);

        await requireApproval(dependencies, {
          toolName: tool.name,
          title: t("gate.extensions.installTitle"),
          message: lines.join("\n"),
          confirmLabel: t("gate.extensions.installConfirm"),
          destructive: requested.length > 0,
        }, `Extension install of "${name}" cancelled by user.`, signal);
      } else if (action === "set_enabled" && parsed.enabled === true) {
        const granted = target?.grantedCapabilities ?? [];
        await requireApproval(dependencies, {
          toolName: tool.name,
          title: t("gate.extensions.enableTitle"),
          message: t("gate.extensions.enableMessage", {
            name: targetName,
            list: granted.length > 0 ? granted.join(", ") : noneLabel,
          }),
          confirmLabel: t("gate.extensions.enableConfirm"),
          destructive: false,
        }, `Enabling extension "${targetName}" cancelled by user.`, signal);
      } else if (action === "uninstall") {
        await requireApproval(dependencies, {
          toolName: tool.name,
          title: t("gate.extensions.uninstallTitle"),
          message: t("gate.extensions.uninstallMessage", { name: targetName }),
          confirmLabel: t("gate.extensions.uninstallConfirm"),
          destructive: true,
        }, `Uninstalling extension "${targetName}" cancelled by user.`, signal);
      } else if (mode === "safe" && action === "set_enabled") {
        await requireApproval(dependencies, {
          toolName: tool.name,
          title: t("gate.extensions.disableTitle"),
          message: t("gate.extensions.disableMessage", { name: targetName }),
          confirmLabel: t("gate.extensions.disableConfirm"),
          destructive: false,
        }, `Disabling extension "${targetName}" cancelled by user.`, signal);
      }

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}
