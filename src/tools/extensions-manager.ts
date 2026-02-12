/**
 * extensions_manager — manage runtime extensions from chat.
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";

export interface ExtensionsManagerToolStatus {
  id: string;
  name: string;
  enabled: boolean;
  loaded: boolean;
  sourceLabel: string;
  trustLabel: string;
  effectiveCapabilities: string[];
  lastError: string | null;
}

export interface ExtensionsManagerToolRuntime {
  list(): ExtensionsManagerToolStatus[];
  installFromCode(name: string, code: string): Promise<string>;
  setExtensionEnabled(entryId: string, enabled: boolean): Promise<void>;
  reloadExtension(entryId: string): Promise<void>;
  uninstallExtension(entryId: string): Promise<void>;
}

export interface CreateExtensionsManagerToolOptions {
  getManager: () => ExtensionsManagerToolRuntime | null;
}

const schema = Type.Object({
  action: Type.Union([
    Type.Literal("list"),
    Type.Literal("install_code"),
    Type.Literal("set_enabled"),
    Type.Literal("reload"),
    Type.Literal("uninstall"),
  ], {
    description:
      "list installed extensions, install_code from JS module text, set_enabled true/false, reload, or uninstall by id",
  }),
  extension_id: Type.Optional(Type.String({
    description: "Target extension id for set_enabled/reload/uninstall.",
  })),
  name: Type.Optional(Type.String({
    description: "Extension display name for install_code.",
  })),
  code: Type.Optional(Type.String({
    description: "Single-file JavaScript ES module source for install_code.",
  })),
  enabled: Type.Optional(Type.Boolean({
    description: "Desired enabled state when action=set_enabled.",
  })),
  replace_existing: Type.Optional(Type.Boolean({
    description: "When install_code, uninstall existing extensions with the same name first (default: true).",
  })),
});

type Params = Static<typeof schema>;

function getRequiredManager(getManager: () => ExtensionsManagerToolRuntime | null): ExtensionsManagerToolRuntime {
  const manager = getManager();
  if (!manager) {
    throw new Error("Extension manager is not available in this runtime.");
  }

  return manager;
}

function requireNonEmptyString(value: string | undefined, field: string): string {
  if (typeof value !== "string") {
    throw new Error(`${field} is required.`);
  }

  const trimmed = value.trim();
  if (trimmed.length === 0) {
    throw new Error(`${field} cannot be empty.`);
  }

  return trimmed;
}

function requireBoolean(value: boolean | undefined, field: string): boolean {
  if (typeof value !== "boolean") {
    throw new Error(`${field} must be true or false.`);
  }

  return value;
}

function summarizeStatus(status: ExtensionsManagerToolStatus): string {
  const state = status.enabled
    ? status.loaded
      ? "enabled, loaded"
      : "enabled, not loaded"
    : "disabled";
  const capabilitySummary = status.effectiveCapabilities.length > 0
    ? status.effectiveCapabilities.join(", ")
    : "(none)";

  const parts = [
    `- ${status.name} (${status.id})`,
    `  - State: ${state}`,
    `  - Source: ${status.sourceLabel}`,
    `  - Trust: ${status.trustLabel}`,
    `  - Effective capabilities: ${capabilitySummary}`,
  ];

  if (status.lastError) {
    parts.push(`  - Last error: ${status.lastError}`);
  }

  return parts.join("\n");
}

export function createExtensionsManagerTool(
  options: CreateExtensionsManagerToolOptions,
): AgentTool<typeof schema, undefined> {
  return {
    name: "extensions_manager",
    label: "Extensions Manager",
    description:
      "List and manage runtime extensions (install from code, enable/disable, reload, uninstall). "
      + "Useful when a user asks you to create an extension directly in chat.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      const manager = getRequiredManager(options.getManager);

      if (params.action === "list") {
        const statuses = manager.list();
        if (statuses.length === 0) {
          return {
            content: [{ type: "text", text: "No extensions installed." }],
            details: undefined,
          };
        }

        const body = statuses.map((status) => summarizeStatus(status)).join("\n");
        return {
          content: [{ type: "text", text: `Installed extensions:\n${body}` }],
          details: undefined,
        };
      }

      if (params.action === "install_code") {
        const extensionName = requireNonEmptyString(params.name, "name");
        const extensionCode = requireNonEmptyString(params.code, "code");
        const replaceExisting = params.replace_existing ?? true;

        const existingWithName = manager.list().filter((status) => status.name.toLowerCase() === extensionName.toLowerCase());

        if (!replaceExisting && existingWithName.length > 0) {
          throw new Error(
            `Extension name "${extensionName}" already exists (${existingWithName.map((status) => status.id).join(", ")}). `
            + "Set replace_existing=true to replace it.",
          );
        }

        for (const existing of existingWithName) {
          await manager.uninstallExtension(existing.id);
        }

        const newId = await manager.installFromCode(extensionName, extensionCode);

        if (params.enabled === false) {
          await manager.setExtensionEnabled(newId, false);
        }

        const installed = manager.list().find((status) => status.id === newId) ?? null;

        const lines = [
          `Installed extension "${extensionName}" as ${newId}.`,
          replaceExisting && existingWithName.length > 0
            ? `Replaced ${existingWithName.length} existing extension${existingWithName.length === 1 ? "" : "s"} with the same name.`
            : "No existing extensions were replaced.",
        ];

        if (installed) {
          lines.push(
            `State: ${installed.enabled ? "enabled" : "disabled"} (${installed.loaded ? "loaded" : "not loaded"}).`,
            `Effective capabilities: ${installed.effectiveCapabilities.length > 0 ? installed.effectiveCapabilities.join(", ") : "(none)"}.`,
          );

          if (installed.lastError) {
            lines.push(`Last error: ${installed.lastError}`);
          }
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      }

      if (params.action === "set_enabled") {
        const extensionId = requireNonEmptyString(params.extension_id, "extension_id");
        const enabled = requireBoolean(params.enabled, "enabled");

        await manager.setExtensionEnabled(extensionId, enabled);

        const status = manager.list().find((entry) => entry.id === extensionId) ?? null;
        const lines = [
          `${enabled ? "Enabled" : "Disabled"} extension ${extensionId}.`,
        ];

        if (status?.lastError) {
          lines.push(`Last error: ${status.lastError}`);
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      }

      if (params.action === "reload") {
        const extensionId = requireNonEmptyString(params.extension_id, "extension_id");

        await manager.reloadExtension(extensionId);
        const status = manager.list().find((entry) => entry.id === extensionId) ?? null;

        const lines = [`Reloaded extension ${extensionId}.`];
        if (status?.lastError) {
          lines.push(`Last error: ${status.lastError}`);
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      }

      const extensionId = requireNonEmptyString(params.extension_id, "extension_id");
      await manager.uninstallExtension(extensionId);

      return {
        content: [{ type: "text", text: `Uninstalled extension ${extensionId}.` }],
        details: undefined,
      };
    },
  };
}
