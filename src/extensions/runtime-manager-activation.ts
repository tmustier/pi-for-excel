import type { Agent } from "@mariozechner/pi-agent-core";

import type {
  CreateExtensionAPIOptions,
  HttpRequestOptions,
  HttpResponse,
  LlmCompletionRequest,
  LlmCompletionResult,
} from "../commands/extension-api.js";
import type { ConnectionManager } from "../connections/manager.js";
import type { ConnectionStatus } from "../connections/types.js";
import type { ExtensionCapability } from "./permissions.js";
import {
  deleteExtensionStorageValue,
  getExtensionStorageValue,
  listExtensionStorageKeys,
  setExtensionStorageValue,
} from "./storage-store.js";
import {
  installExternalExtensionSkill,
  listExtensionSkillSummaries,
  readExtensionSkill,
  uninstallExternalExtensionSkill,
} from "./skills-store.js";
import { createExtensionAgentMessage } from "./runtime-manager-helpers.js";
import type { SandboxActivationOptions } from "./sandbox-runtime.js";
import type { ExtensionSettingsStore, StoredExtensionEntry } from "./store.js";

type HostActivationBridge = Pick<
  CreateExtensionAPIOptions,
  | "getAgent"
  | "llmComplete"
  | "httpFetch"
  | "storageGet"
  | "storageSet"
  | "storageDelete"
  | "storageKeys"
  | "clipboardWriteText"
  | "injectAgentContext"
  | "steerAgent"
  | "followUpAgent"
  | "listSkills"
  | "readSkill"
  | "installSkill"
  | "uninstallSkill"
  | "downloadFile"
  | "registerConnection"
  | "unregisterConnection"
  | "listConnections"
  | "getConnection"
  | "setConnectionSecrets"
  | "clearConnectionSecrets"
  | "markConnectionValidated"
  | "markConnectionInvalid"
  | "markConnectionStatus"
  | "isCapabilityEnabled"
  | "formatCapabilityError"
  | "extensionOwnerId"
  | "widgetApiV2Enabled"
>;

type SandboxActivationBridge = Pick<
  SandboxActivationOptions,
  | "llmComplete"
  | "httpFetch"
  | "storageGet"
  | "storageSet"
  | "storageDelete"
  | "storageKeys"
  | "clipboardWriteText"
  | "injectAgentContext"
  | "steerAgent"
  | "followUpAgent"
  | "listSkills"
  | "readSkill"
  | "installSkill"
  | "uninstallSkill"
  | "downloadFile"
  | "registerConnection"
  | "unregisterConnection"
  | "listConnections"
  | "getConnection"
  | "setConnectionSecrets"
  | "clearConnectionSecrets"
  | "markConnectionValidated"
  | "markConnectionInvalid"
  | "markConnectionStatus"
  | "isCapabilityEnabled"
  | "formatCapabilityError"
  | "toast"
  | "widgetOwnerId"
  | "widgetApiV2Enabled"
>;

function qualifyConnectionIdForEntry(entryId: string, connectionId: string): string {
  const normalizedConnectionId = connectionId.trim().toLowerCase();
  if (normalizedConnectionId.length === 0) {
    throw new Error("Connection id cannot be empty.");
  }

  const ownerPrefix = `${entryId.toLowerCase()}.`;

  if (normalizedConnectionId.startsWith(ownerPrefix)) {
    return normalizedConnectionId;
  }

  return `${ownerPrefix}${normalizedConnectionId}`;
}

export interface RuntimeManagerActivationBridge {
  host: HostActivationBridge;
  sandbox: SandboxActivationBridge;
}

export interface BuildRuntimeManagerActivationBridgeOptions {
  entry: StoredExtensionEntry;
  settings: ExtensionSettingsStore;
  connectionManager: ConnectionManager;
  getRequiredActiveAgent: () => Agent;
  runExtensionLlmCompletion: (request: LlmCompletionRequest) => Promise<LlmCompletionResult>;
  runExtensionHttpFetch: (url: string, options?: HttpRequestOptions) => Promise<HttpResponse>;
  writeExtensionClipboard: (text: string) => Promise<void>;
  triggerExtensionDownload: (filename: string, content: string, mimeType?: string) => void;
  isCapabilityEnabled: (capability: ExtensionCapability) => boolean;
  formatCapabilityError: (capability: ExtensionCapability) => string;
  showToastMessage: (message: string) => void;
  widgetApiV2Enabled: boolean;
}

export function buildRuntimeManagerActivationBridge(
  options: BuildRuntimeManagerActivationBridgeOptions,
): RuntimeManagerActivationBridge {
  const {
    entry,
    settings,
    connectionManager,
    getRequiredActiveAgent,
    runExtensionLlmCompletion,
    runExtensionHttpFetch,
    writeExtensionClipboard,
    triggerExtensionDownload,
    isCapabilityEnabled,
    formatCapabilityError,
    showToastMessage,
    widgetApiV2Enabled,
  } = options;

  const buildExtensionMessage = (label: string, content: string) => {
    return createExtensionAgentMessage(entry.name, label, content);
  };

  const storageGet = (key: string) => getExtensionStorageValue(settings, entry.id, key);
  const storageSet = (key: string, value: unknown) => setExtensionStorageValue(settings, entry.id, key, value);
  const storageDelete = (key: string) => deleteExtensionStorageValue(settings, entry.id, key);
  const storageKeys = () => listExtensionStorageKeys(settings, entry.id);

  const injectAgentContext = (content: string): void => {
    const agent = getRequiredActiveAgent();
    agent.appendMessage(buildExtensionMessage("agent.injectContext content", content));
  };

  const steerAgent = (content: string): void => {
    const agent = getRequiredActiveAgent();
    agent.steer(buildExtensionMessage("agent.steer content", content));
  };

  const followUpAgent = (content: string): void => {
    const agent = getRequiredActiveAgent();
    agent.followUp(buildExtensionMessage("agent.followUp content", content));
  };

  const listSkills = () => listExtensionSkillSummaries();
  const readSkill = (name: string) => readExtensionSkill(name);
  const installSkill = (name: string, markdown: string) => installExternalExtensionSkill(name, markdown);
  const uninstallSkill = (name: string) => uninstallExternalExtensionSkill(name);

  const downloadFile = (filename: string, content: string, mimeType?: string): void => {
    triggerExtensionDownload(filename, content, mimeType);
  };

  const registerConnection = (definition: Parameters<ConnectionManager["registerDefinition"]>[1]) => {
    const normalizedDefinition = {
      ...definition,
      id: qualifyConnectionIdForEntry(entry.id, definition.id),
    };

    return connectionManager.registerDefinition(entry.id, normalizedDefinition);
  };

  const unregisterConnection = (connectionId: string): void => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);
    connectionManager.unregisterDefinition(entry.id, normalizedConnectionId);
  };

  const listConnections = async () => {
    const ownerPrefix = `${entry.id.toLowerCase()}.`;
    const snapshots = await connectionManager.listSnapshots();

    return snapshots
      .filter((snapshot) => snapshot.connectionId.startsWith(ownerPrefix))
      .map((snapshot) => ({
        connectionId: snapshot.connectionId,
        status: snapshot.status,
        lastValidatedAt: snapshot.lastValidatedAt,
        lastError: snapshot.lastError,
      }));
  };

  const getConnection = async (connectionId: string) => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);
    const snapshot = await connectionManager.getSnapshot(normalizedConnectionId);
    if (!snapshot) return null;

    return {
      connectionId: snapshot.connectionId,
      status: snapshot.status,
      lastValidatedAt: snapshot.lastValidatedAt,
      lastError: snapshot.lastError,
    };
  };

  const setConnectionSecrets = async (connectionId: string, secrets: Record<string, string>): Promise<void> => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);
    await connectionManager.setSecrets(entry.id, normalizedConnectionId, secrets);
  };

  const clearConnectionSecrets = async (connectionId: string): Promise<void> => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);
    await connectionManager.clearSecrets(entry.id, normalizedConnectionId);
  };

  const markConnectionValidated = async (connectionId: string): Promise<void> => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);
    await connectionManager.markValidated(entry.id, normalizedConnectionId);
  };

  const markConnectionInvalid = async (connectionId: string, reason: string): Promise<void> => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);
    await connectionManager.markInvalid(entry.id, normalizedConnectionId, reason);
  };

  const markConnectionStatus = async (
    connectionId: string,
    status: ConnectionStatus,
    reason?: string,
  ): Promise<void> => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);

    if (status === "connected") {
      await connectionManager.markValidated(entry.id, normalizedConnectionId);
      return;
    }

    if (status === "missing") {
      await connectionManager.clearSecrets(entry.id, normalizedConnectionId);
      return;
    }

    if (status === "invalid") {
      await connectionManager.markInvalid(entry.id, normalizedConnectionId, reason ?? "Connection marked invalid.");
      return;
    }

    await connectionManager.markRuntimeAuthFailure(normalizedConnectionId, {
      message: reason ?? "Connection reported runtime authentication failure.",
    });
  };

  const host: HostActivationBridge = {
    getAgent: getRequiredActiveAgent,
    llmComplete: runExtensionLlmCompletion,
    httpFetch: runExtensionHttpFetch,
    storageGet,
    storageSet,
    storageDelete,
    storageKeys,
    clipboardWriteText: writeExtensionClipboard,
    injectAgentContext,
    steerAgent,
    followUpAgent,
    listSkills,
    readSkill,
    installSkill,
    uninstallSkill,
    downloadFile,
    registerConnection,
    unregisterConnection,
    listConnections,
    getConnection,
    setConnectionSecrets,
    clearConnectionSecrets,
    markConnectionValidated,
    markConnectionInvalid,
    markConnectionStatus,
    isCapabilityEnabled,
    formatCapabilityError,
    extensionOwnerId: entry.id,
    widgetApiV2Enabled,
  };

  const sandbox: SandboxActivationBridge = {
    llmComplete: runExtensionLlmCompletion,
    httpFetch: runExtensionHttpFetch,
    storageGet,
    storageSet,
    storageDelete,
    storageKeys,
    clipboardWriteText: writeExtensionClipboard,
    injectAgentContext,
    steerAgent,
    followUpAgent,
    listSkills,
    readSkill,
    installSkill,
    uninstallSkill,
    downloadFile,
    registerConnection,
    unregisterConnection,
    listConnections,
    getConnection,
    setConnectionSecrets,
    clearConnectionSecrets,
    markConnectionValidated,
    markConnectionInvalid,
    markConnectionStatus,
    isCapabilityEnabled,
    formatCapabilityError,
    toast: showToastMessage,
    widgetOwnerId: entry.id,
    widgetApiV2Enabled,
  };

  return {
    host,
    sandbox,
  };
}
