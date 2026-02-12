/**
 * Extension trust + capability permissions.
 *
 * This module is storage/runtime-facing (no UI strings beyond short labels).
 */

import { classifyExtensionSource } from "../commands/extension-source-policy.js";
import { isRecord } from "../utils/type-guards.js";

export type StoredExtensionTrust = "builtin" | "local-module" | "inline-code" | "remote-url";

export type ExtensionSourceLike =
  | { kind: "module"; specifier: string }
  | { kind: "inline"; code: string };

export interface StoredExtensionPermissions {
  commandsRegister: boolean;
  toolsRegister: boolean;
  agentRead: boolean;
  agentEventsRead: boolean;
  uiOverlay: boolean;
  uiWidget: boolean;
  uiToast: boolean;
}

export const ALL_EXTENSION_CAPABILITIES = [
  "commands.register",
  "tools.register",
  "agent.read",
  "agent.events.read",
  "ui.overlay",
  "ui.widget",
  "ui.toast",
] as const;

export type ExtensionCapability = (typeof ALL_EXTENSION_CAPABILITIES)[number];

const TRUSTED_PERMISSIONS: StoredExtensionPermissions = {
  commandsRegister: true,
  toolsRegister: true,
  agentRead: true,
  agentEventsRead: true,
  uiOverlay: true,
  uiWidget: true,
  uiToast: true,
};

const RESTRICTED_UNTRUSTED_PERMISSIONS: StoredExtensionPermissions = {
  commandsRegister: true,
  toolsRegister: false,
  agentRead: false,
  agentEventsRead: false,
  uiOverlay: true,
  uiWidget: true,
  uiToast: true,
};

const TRUST_LABELS: Record<StoredExtensionTrust, string> = {
  builtin: "builtin",
  "local-module": "local module",
  "inline-code": "inline code",
  "remote-url": "remote URL",
};

const CAPABILITY_LABELS: Record<ExtensionCapability, string> = {
  "commands.register": "register commands",
  "tools.register": "register tools",
  "agent.read": "read agent state",
  "agent.events.read": "read agent events",
  "ui.overlay": "show overlays",
  "ui.widget": "show widgets",
  "ui.toast": "show toasts",
};

function clonePermissions(source: StoredExtensionPermissions): StoredExtensionPermissions {
  return {
    commandsRegister: source.commandsRegister,
    toolsRegister: source.toolsRegister,
    agentRead: source.agentRead,
    agentEventsRead: source.agentEventsRead,
    uiOverlay: source.uiOverlay,
    uiWidget: source.uiWidget,
    uiToast: source.uiToast,
  };
}

function normalizeBooleanOrFallback(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

export function deriveStoredExtensionTrust(entryId: string, source: ExtensionSourceLike): StoredExtensionTrust {
  if (source.kind === "inline") {
    return "inline-code";
  }

  const sourceKind = classifyExtensionSource(source.specifier);
  if (sourceKind === "remote-url") {
    return "remote-url";
  }

  if (sourceKind === "blob-url") {
    return "inline-code";
  }

  if (entryId === "builtin.snake" || entryId.startsWith("builtin.")) {
    return "builtin";
  }

  return "local-module";
}

export function getDefaultPermissionsForTrust(trust: StoredExtensionTrust): StoredExtensionPermissions {
  if (trust === "builtin" || trust === "local-module") {
    return clonePermissions(TRUSTED_PERMISSIONS);
  }

  return clonePermissions(RESTRICTED_UNTRUSTED_PERMISSIONS);
}

export function normalizeStoredExtensionPermissions(
  raw: unknown,
  trust: StoredExtensionTrust,
): StoredExtensionPermissions {
  const defaults = getDefaultPermissionsForTrust(trust);

  if (!isRecord(raw)) {
    return defaults;
  }

  return {
    commandsRegister: normalizeBooleanOrFallback(raw.commandsRegister, defaults.commandsRegister),
    toolsRegister: normalizeBooleanOrFallback(raw.toolsRegister, defaults.toolsRegister),
    agentRead: normalizeBooleanOrFallback(raw.agentRead, defaults.agentRead),
    agentEventsRead: normalizeBooleanOrFallback(raw.agentEventsRead, defaults.agentEventsRead),
    uiOverlay: normalizeBooleanOrFallback(raw.uiOverlay, defaults.uiOverlay),
    uiWidget: normalizeBooleanOrFallback(raw.uiWidget, defaults.uiWidget),
    uiToast: normalizeBooleanOrFallback(raw.uiToast, defaults.uiToast),
  };
}

export function isExtensionCapabilityAllowed(
  permissions: StoredExtensionPermissions,
  capability: ExtensionCapability,
): boolean {
  switch (capability) {
    case "commands.register":
      return permissions.commandsRegister;
    case "tools.register":
      return permissions.toolsRegister;
    case "agent.read":
      return permissions.agentRead;
    case "agent.events.read":
      return permissions.agentEventsRead;
    case "ui.overlay":
      return permissions.uiOverlay;
    case "ui.widget":
      return permissions.uiWidget;
    case "ui.toast":
      return permissions.uiToast;
  }
}

export function setExtensionCapabilityAllowed(
  permissions: StoredExtensionPermissions,
  capability: ExtensionCapability,
  allowed: boolean,
): StoredExtensionPermissions {
  switch (capability) {
    case "commands.register":
      return {
        ...permissions,
        commandsRegister: allowed,
      };
    case "tools.register":
      return {
        ...permissions,
        toolsRegister: allowed,
      };
    case "agent.read":
      return {
        ...permissions,
        agentRead: allowed,
      };
    case "agent.events.read":
      return {
        ...permissions,
        agentEventsRead: allowed,
      };
    case "ui.overlay":
      return {
        ...permissions,
        uiOverlay: allowed,
      };
    case "ui.widget":
      return {
        ...permissions,
        uiWidget: allowed,
      };
    case "ui.toast":
      return {
        ...permissions,
        uiToast: allowed,
      };
  }
}

export function describeStoredExtensionTrust(trust: StoredExtensionTrust): string {
  return TRUST_LABELS[trust];
}

export function describeExtensionCapability(capability: ExtensionCapability): string {
  return CAPABILITY_LABELS[capability];
}

export function listAllExtensionCapabilities(): ExtensionCapability[] {
  return [...ALL_EXTENSION_CAPABILITIES];
}

export function listGrantedExtensionCapabilities(
  permissions: StoredExtensionPermissions,
): ExtensionCapability[] {
  const capabilities: ExtensionCapability[] = [];

  if (permissions.commandsRegister) capabilities.push("commands.register");
  if (permissions.toolsRegister) capabilities.push("tools.register");
  if (permissions.agentRead) capabilities.push("agent.read");
  if (permissions.agentEventsRead) capabilities.push("agent.events.read");
  if (permissions.uiOverlay) capabilities.push("ui.overlay");
  if (permissions.uiWidget) capabilities.push("ui.widget");
  if (permissions.uiToast) capabilities.push("ui.toast");

  return capabilities;
}
