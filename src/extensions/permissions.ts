function isExtensionsPermissionsPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Extension trust + capability permissions.
 *
 * This module is storage/runtime-facing (no UI strings beyond short labels).
 */

import { t } from "../language/index.js";
import { classifyExtensionSource } from "../commands/extension-source-policy.js";

export type StoredExtensionTrust = "builtin" | "local-module" | "inline-code" | "remote-url";

export type ExtensionSourceLike =
  | { kind: "module"; specifier: string }
  | { kind: "inline"; code: string };

export interface StoredExtensionPermissions {
  commandsRegister: boolean;
  toolsRegister: boolean;
  modelsRegister: boolean;
  agentRead: boolean;
  agentEventsRead: boolean;
  uiOverlay: boolean;
  uiWidget: boolean;
  uiToast: boolean;
  llmComplete: boolean;
  httpFetch: boolean;
  storageReadWrite: boolean;
  connectionsReadWrite: boolean;
  connectionsSecretsRead: boolean;
  clipboardWrite: boolean;
  agentContextWrite: boolean;
  agentSteer: boolean;
  agentFollowUp: boolean;
  skillsRead: boolean;
  skillsWrite: boolean;
  downloadFile: boolean;
}

interface CapabilityDescriptor {
  capability: string;
  permissionKey: keyof StoredExtensionPermissions;
  tKey: string;
}

function getCapabilityDescriptors(): readonly CapabilityDescriptor[] {
  return [
    { capability: "commands.register",           permissionKey: "commandsRegister",           tKey: "perm.commands.register" },
    { capability: "tools.register",              permissionKey: "toolsRegister",              tKey: "perm.tools.register" },
    { capability: "models.register",             permissionKey: "modelsRegister",             tKey: "perm.models.register" },
    { capability: "agent.read",                  permissionKey: "agentRead",                  tKey: "perm.agent.read" },
    { capability: "agent.events.read",           permissionKey: "agentEventsRead",            tKey: "perm.agent.events.read" },
    { capability: "ui.overlay",                  permissionKey: "uiOverlay",                  tKey: "perm.ui.overlay" },
    { capability: "ui.widget",                   permissionKey: "uiWidget",                   tKey: "perm.ui.widget" },
    { capability: "ui.toast",                    permissionKey: "uiToast",                    tKey: "perm.ui.toast" },
    { capability: "llm.complete",                permissionKey: "llmComplete",                tKey: "perm.llm.complete" },
    { capability: "http.fetch",                  permissionKey: "httpFetch",                  tKey: "perm.http.fetch" },
    { capability: "storage.readwrite",           permissionKey: "storageReadWrite",           tKey: "perm.storage.readwrite" },
    { capability: "connections.readwrite",       permissionKey: "connectionsReadWrite",       tKey: "perm.connections.readwrite" },
    { capability: "connections.secrets.read",    permissionKey: "connectionsSecretsRead",    tKey: "perm.connections.secrets.read" },
    { capability: "clipboard.write",             permissionKey: "clipboardWrite",             tKey: "perm.clipboard.write" },
    { capability: "agent.context.write",         permissionKey: "agentContextWrite",          tKey: "perm.agent.context.write" },
    { capability: "agent.steer",                 permissionKey: "agentSteer",                 tKey: "perm.agent.steer" },
    { capability: "agent.followup",              permissionKey: "agentFollowUp",              tKey: "perm.agent.followup" },
    { capability: "skills.read",                 permissionKey: "skillsRead",                 tKey: "perm.skills.read" },
    { capability: "skills.write",                permissionKey: "skillsWrite",                tKey: "perm.skills.write" },
    { capability: "download.file",               permissionKey: "downloadFile",               tKey: "perm.download.file" },
  ];
}

export type ExtensionCapability = (ReturnType<typeof getCapabilityDescriptors>)[number]["capability"];

export const ALL_EXTENSION_CAPABILITIES: ExtensionCapability[] = [
  "commands.register", "tools.register", "models.register", "agent.read", "agent.events.read",
  "ui.overlay", "ui.widget", "ui.toast", "llm.complete", "http.fetch",
  "storage.readwrite", "connections.readwrite", "connections.secrets.read",
  "clipboard.write", "agent.context.write", "agent.steer", "agent.followup",
  "skills.read", "skills.write", "download.file",
];

const TRUSTED_PERMISSIONS: StoredExtensionPermissions = {
  commandsRegister: true, toolsRegister: true, modelsRegister: true, agentRead: true, agentEventsRead: true,
  uiOverlay: true, uiWidget: true, uiToast: true, llmComplete: true, httpFetch: true,
  storageReadWrite: true, connectionsReadWrite: true, connectionsSecretsRead: false,
  clipboardWrite: true, agentContextWrite: false, agentSteer: false, agentFollowUp: false,
  skillsRead: true, skillsWrite: false, downloadFile: true,
};

const RESTRICTED_UNTRUSTED_PERMISSIONS: StoredExtensionPermissions = {
  commandsRegister: true, toolsRegister: false, modelsRegister: false, agentRead: false, agentEventsRead: false,
  uiOverlay: true, uiWidget: true, uiToast: true, llmComplete: false, httpFetch: false,
  storageReadWrite: true, connectionsReadWrite: false, connectionsSecretsRead: false,
  clipboardWrite: true, agentContextWrite: false, agentSteer: false, agentFollowUp: false,
  skillsRead: true, skillsWrite: false, downloadFile: true,
};

function getCapabilityDescriptor(capability: ExtensionCapability): CapabilityDescriptor {
  const descriptor = getCapabilityDescriptors().find((entry) => entry.capability === capability);
  if (!descriptor) {
    throw new Error(`Unknown extension capability: ${capability}`);
  }
  return descriptor;
}

function clonePermissions(source: StoredExtensionPermissions): StoredExtensionPermissions {
  return { ...source };
}

export function deriveStoredExtensionTrust(entryId: string, source: ExtensionSourceLike): StoredExtensionTrust {
  if (source.kind === "inline") return "inline-code";
  const sourceKind = classifyExtensionSource(source.specifier);
  if (sourceKind === "remote-url") return "remote-url";
  if (sourceKind === "blob-url") return "inline-code";
  if (entryId === "builtin.snake" || entryId.startsWith("builtin.")) return "builtin";
  return "local-module";
}

export function getDefaultPermissionsForTrust(trust: StoredExtensionTrust): StoredExtensionPermissions {
  if (trust === "builtin" || trust === "local-module") return clonePermissions(TRUSTED_PERMISSIONS);
  return clonePermissions(RESTRICTED_UNTRUSTED_PERMISSIONS);
}

export function normalizeStoredExtensionPermissions(raw: DynamicValue, trust: StoredExtensionTrust): StoredExtensionPermissions {
  const defaults = getDefaultPermissionsForTrust(trust);
  if (!isExtensionsPermissionsPayloadShape(raw)) return defaults;
  return {
    commandsRegister: normalizeBooleanOrFallback(raw.commandsRegister, defaults.commandsRegister),
    toolsRegister: normalizeBooleanOrFallback(raw.toolsRegister, defaults.toolsRegister),
    modelsRegister: normalizeBooleanOrFallback(raw.modelsRegister, defaults.modelsRegister),
    agentRead: normalizeBooleanOrFallback(raw.agentRead, defaults.agentRead),
    agentEventsRead: normalizeBooleanOrFallback(raw.agentEventsRead, defaults.agentEventsRead),
    uiOverlay: normalizeBooleanOrFallback(raw.uiOverlay, defaults.uiOverlay),
    uiWidget: normalizeBooleanOrFallback(raw.uiWidget, defaults.uiWidget),
    uiToast: normalizeBooleanOrFallback(raw.uiToast, defaults.uiToast),
    llmComplete: normalizeBooleanOrFallback(raw.llmComplete, defaults.llmComplete),
    httpFetch: normalizeBooleanOrFallback(raw.httpFetch, defaults.httpFetch),
    storageReadWrite: normalizeBooleanOrFallback(raw.storageReadWrite, defaults.storageReadWrite),
    connectionsReadWrite: normalizeBooleanOrFallback(raw.connectionsReadWrite, defaults.connectionsReadWrite),
    connectionsSecretsRead: normalizeBooleanOrFallback(raw.connectionsSecretsRead, defaults.connectionsSecretsRead),
    clipboardWrite: normalizeBooleanOrFallback(raw.clipboardWrite, defaults.clipboardWrite),
    agentContextWrite: normalizeBooleanOrFallback(raw.agentContextWrite, defaults.agentContextWrite),
    agentSteer: normalizeBooleanOrFallback(raw.agentSteer, defaults.agentSteer),
    agentFollowUp: normalizeBooleanOrFallback(raw.agentFollowUp, defaults.agentFollowUp),
    skillsRead: normalizeBooleanOrFallback(raw.skillsRead, defaults.skillsRead),
    skillsWrite: normalizeBooleanOrFallback(raw.skillsWrite, defaults.skillsWrite),
    downloadFile: normalizeBooleanOrFallback(raw.downloadFile, defaults.downloadFile),
  };
}

export function isExtensionCapabilityAllowed(permissions: StoredExtensionPermissions, capability: ExtensionCapability): boolean {
  const descriptor = getCapabilityDescriptor(capability);
  return permissions[descriptor.permissionKey];
}

export function setExtensionCapabilityAllowed(permissions: StoredExtensionPermissions, capability: ExtensionCapability, allowed: boolean): StoredExtensionPermissions {
  const descriptor = getCapabilityDescriptor(capability);
  return { ...permissions, [descriptor.permissionKey]: allowed };
}

export function describeStoredExtensionTrust(trust: StoredExtensionTrust): string {
  return t("perm.trust." + trust);
}

export function describeExtensionCapability(capability: ExtensionCapability): string {
  const descriptor = getCapabilityDescriptor(capability);
  return t(descriptor.tKey);
}

export function listAllExtensionCapabilities(): ExtensionCapability[] {
  return [...ALL_EXTENSION_CAPABILITIES];
}

export function listGrantedExtensionCapabilities(permissions: StoredExtensionPermissions): ExtensionCapability[] {
  return getCapabilityDescriptors()
    .filter((descriptor) => permissions[descriptor.permissionKey])
    .map((descriptor) => descriptor.capability);
}

function normalizeBooleanOrFallback(value: DynamicValue, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}
