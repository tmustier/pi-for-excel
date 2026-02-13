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
  llmComplete: boolean;
  httpFetch: boolean;
  storageReadWrite: boolean;
  clipboardWrite: boolean;
  agentContextWrite: boolean;
  agentSteer: boolean;
  agentFollowUp: boolean;
  skillsRead: boolean;
  skillsWrite: boolean;
  downloadFile: boolean;
}

export const ALL_EXTENSION_CAPABILITIES = [
  "commands.register",
  "tools.register",
  "agent.read",
  "agent.events.read",
  "ui.overlay",
  "ui.widget",
  "ui.toast",
  "llm.complete",
  "http.fetch",
  "storage.readwrite",
  "clipboard.write",
  "agent.context.write",
  "agent.steer",
  "agent.followup",
  "skills.read",
  "skills.write",
  "download.file",
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
  llmComplete: true,
  httpFetch: true,
  storageReadWrite: true,
  clipboardWrite: true,
  agentContextWrite: false,
  agentSteer: false,
  agentFollowUp: false,
  skillsRead: true,
  skillsWrite: false,
  downloadFile: true,
};

const RESTRICTED_UNTRUSTED_PERMISSIONS: StoredExtensionPermissions = {
  commandsRegister: true,
  toolsRegister: false,
  agentRead: false,
  agentEventsRead: false,
  uiOverlay: true,
  uiWidget: true,
  uiToast: true,
  llmComplete: false,
  httpFetch: false,
  storageReadWrite: true,
  clipboardWrite: true,
  agentContextWrite: false,
  agentSteer: false,
  agentFollowUp: false,
  skillsRead: true,
  skillsWrite: false,
  downloadFile: true,
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
  "llm.complete": "call LLM completions",
  "http.fetch": "fetch external HTTP resources",
  "storage.readwrite": "read/write extension storage",
  "clipboard.write": "write clipboard text",
  "agent.context.write": "inject agent context",
  "agent.steer": "steer active agent runs",
  "agent.followup": "queue agent follow-up messages",
  "skills.read": "read skill catalog",
  "skills.write": "install/uninstall external skills",
  "download.file": "trigger file downloads",
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
    llmComplete: source.llmComplete,
    httpFetch: source.httpFetch,
    storageReadWrite: source.storageReadWrite,
    clipboardWrite: source.clipboardWrite,
    agentContextWrite: source.agentContextWrite,
    agentSteer: source.agentSteer,
    agentFollowUp: source.agentFollowUp,
    skillsRead: source.skillsRead,
    skillsWrite: source.skillsWrite,
    downloadFile: source.downloadFile,
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
    llmComplete: normalizeBooleanOrFallback(raw.llmComplete, defaults.llmComplete),
    httpFetch: normalizeBooleanOrFallback(raw.httpFetch, defaults.httpFetch),
    storageReadWrite: normalizeBooleanOrFallback(raw.storageReadWrite, defaults.storageReadWrite),
    clipboardWrite: normalizeBooleanOrFallback(raw.clipboardWrite, defaults.clipboardWrite),
    agentContextWrite: normalizeBooleanOrFallback(raw.agentContextWrite, defaults.agentContextWrite),
    agentSteer: normalizeBooleanOrFallback(raw.agentSteer, defaults.agentSteer),
    agentFollowUp: normalizeBooleanOrFallback(raw.agentFollowUp, defaults.agentFollowUp),
    skillsRead: normalizeBooleanOrFallback(raw.skillsRead, defaults.skillsRead),
    skillsWrite: normalizeBooleanOrFallback(raw.skillsWrite, defaults.skillsWrite),
    downloadFile: normalizeBooleanOrFallback(raw.downloadFile, defaults.downloadFile),
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
    case "llm.complete":
      return permissions.llmComplete;
    case "http.fetch":
      return permissions.httpFetch;
    case "storage.readwrite":
      return permissions.storageReadWrite;
    case "clipboard.write":
      return permissions.clipboardWrite;
    case "agent.context.write":
      return permissions.agentContextWrite;
    case "agent.steer":
      return permissions.agentSteer;
    case "agent.followup":
      return permissions.agentFollowUp;
    case "skills.read":
      return permissions.skillsRead;
    case "skills.write":
      return permissions.skillsWrite;
    case "download.file":
      return permissions.downloadFile;
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
    case "llm.complete":
      return {
        ...permissions,
        llmComplete: allowed,
      };
    case "http.fetch":
      return {
        ...permissions,
        httpFetch: allowed,
      };
    case "storage.readwrite":
      return {
        ...permissions,
        storageReadWrite: allowed,
      };
    case "clipboard.write":
      return {
        ...permissions,
        clipboardWrite: allowed,
      };
    case "agent.context.write":
      return {
        ...permissions,
        agentContextWrite: allowed,
      };
    case "agent.steer":
      return {
        ...permissions,
        agentSteer: allowed,
      };
    case "agent.followup":
      return {
        ...permissions,
        agentFollowUp: allowed,
      };
    case "skills.read":
      return {
        ...permissions,
        skillsRead: allowed,
      };
    case "skills.write":
      return {
        ...permissions,
        skillsWrite: allowed,
      };
    case "download.file":
      return {
        ...permissions,
        downloadFile: allowed,
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
  if (permissions.llmComplete) capabilities.push("llm.complete");
  if (permissions.httpFetch) capabilities.push("http.fetch");
  if (permissions.storageReadWrite) capabilities.push("storage.readwrite");
  if (permissions.clipboardWrite) capabilities.push("clipboard.write");
  if (permissions.agentContextWrite) capabilities.push("agent.context.write");
  if (permissions.agentSteer) capabilities.push("agent.steer");
  if (permissions.agentFollowUp) capabilities.push("agent.followup");
  if (permissions.skillsRead) capabilities.push("skills.read");
  if (permissions.skillsWrite) capabilities.push("skills.write");
  if (permissions.downloadFile) capabilities.push("download.file");

  return capabilities;
}
