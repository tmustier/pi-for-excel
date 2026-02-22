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

const EXTENSION_CAPABILITY_DESCRIPTORS = [
  {
    capability: "commands.register",
    permissionKey: "commandsRegister",
    label: "register commands",
  },
  {
    capability: "tools.register",
    permissionKey: "toolsRegister",
    label: "register tools",
  },
  {
    capability: "agent.read",
    permissionKey: "agentRead",
    label: "read agent state",
  },
  {
    capability: "agent.events.read",
    permissionKey: "agentEventsRead",
    label: "read agent events",
  },
  {
    capability: "ui.overlay",
    permissionKey: "uiOverlay",
    label: "show overlays",
  },
  {
    capability: "ui.widget",
    permissionKey: "uiWidget",
    label: "show widgets",
  },
  {
    capability: "ui.toast",
    permissionKey: "uiToast",
    label: "show toasts",
  },
  {
    capability: "llm.complete",
    permissionKey: "llmComplete",
    label: "call LLM completions",
  },
  {
    capability: "http.fetch",
    permissionKey: "httpFetch",
    label: "fetch external HTTP resources",
  },
  {
    capability: "storage.readwrite",
    permissionKey: "storageReadWrite",
    label: "read/write extension storage",
  },
  {
    capability: "connections.readwrite",
    permissionKey: "connectionsReadWrite",
    label: "manage connection definitions and secrets",
  },
  {
    capability: "connections.secrets.read",
    permissionKey: "connectionsSecretsRead",
    label: "read raw connection secret values",
  },
  {
    capability: "clipboard.write",
    permissionKey: "clipboardWrite",
    label: "write clipboard text",
  },
  {
    capability: "agent.context.write",
    permissionKey: "agentContextWrite",
    label: "inject agent context",
  },
  {
    capability: "agent.steer",
    permissionKey: "agentSteer",
    label: "steer active agent runs",
  },
  {
    capability: "agent.followup",
    permissionKey: "agentFollowUp",
    label: "queue agent follow-up messages",
  },
  {
    capability: "skills.read",
    permissionKey: "skillsRead",
    label: "read skill catalog",
  },
  {
    capability: "skills.write",
    permissionKey: "skillsWrite",
    label: "install/uninstall external skills",
  },
  {
    capability: "download.file",
    permissionKey: "downloadFile",
    label: "trigger file downloads",
  },
] as const satisfies ReadonlyArray<{
  capability: string;
  permissionKey: keyof StoredExtensionPermissions;
  label: string;
}>;

export type ExtensionCapability = (typeof EXTENSION_CAPABILITY_DESCRIPTORS)[number]["capability"];

export const ALL_EXTENSION_CAPABILITIES: ExtensionCapability[] = EXTENSION_CAPABILITY_DESCRIPTORS.map((descriptor) => {
  return descriptor.capability;
});

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
  connectionsReadWrite: true,
  connectionsSecretsRead: false,
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
  connectionsReadWrite: false,
  connectionsSecretsRead: false,
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

function getCapabilityDescriptor(capability: ExtensionCapability): (typeof EXTENSION_CAPABILITY_DESCRIPTORS)[number] {
  const descriptor = EXTENSION_CAPABILITY_DESCRIPTORS.find((entry) => entry.capability === capability);
  if (!descriptor) {
    throw new Error(`Unknown extension capability: ${capability}`);
  }

  return descriptor;
}

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
    connectionsReadWrite: source.connectionsReadWrite,
    connectionsSecretsRead: source.connectionsSecretsRead,
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

export function isExtensionCapabilityAllowed(
  permissions: StoredExtensionPermissions,
  capability: ExtensionCapability,
): boolean {
  const descriptor = getCapabilityDescriptor(capability);
  return permissions[descriptor.permissionKey];
}

export function setExtensionCapabilityAllowed(
  permissions: StoredExtensionPermissions,
  capability: ExtensionCapability,
  allowed: boolean,
): StoredExtensionPermissions {
  const descriptor = getCapabilityDescriptor(capability);
  return {
    ...permissions,
    [descriptor.permissionKey]: allowed,
  };
}

export function describeStoredExtensionTrust(trust: StoredExtensionTrust): string {
  return TRUST_LABELS[trust];
}

export function describeExtensionCapability(capability: ExtensionCapability): string {
  const descriptor = getCapabilityDescriptor(capability);
  return descriptor.label;
}

export function listAllExtensionCapabilities(): ExtensionCapability[] {
  return [...ALL_EXTENSION_CAPABILITIES];
}

export function listGrantedExtensionCapabilities(
  permissions: StoredExtensionPermissions,
): ExtensionCapability[] {
  return EXTENSION_CAPABILITY_DESCRIPTORS
    .filter((descriptor) => permissions[descriptor.permissionKey])
    .map((descriptor) => descriptor.capability);
}
