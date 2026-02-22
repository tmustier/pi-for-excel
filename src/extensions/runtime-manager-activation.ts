import type { Agent } from "@mariozechner/pi-agent-core";

import type {
  CreateExtensionAPIOptions,
  HttpRequestOptions,
  HttpResponse,
  LlmCompletionRequest,
  LlmCompletionResult,
} from "../commands/extension-api.js";
import type { ConnectionManager } from "../connections/manager.js";
import type {
  ConnectionDefinition,
  ConnectionSnapshot,
  ConnectionStatus,
  ConnectionToolErrorDetails,
} from "../connections/types.js";
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
  | "getConnectionSecrets"
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
  | "getConnectionSecrets"
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

function mapStatusToConnectionErrorCode(status: ConnectionStatus): ConnectionToolErrorDetails["errorCode"] {
  if (status === "missing") return "missing_connection";
  if (status === "invalid") return "invalid_connection";
  if (status === "error") return "connection_auth_failed";
  return "invalid_connection";
}

function buildConnectionErrorMessage(details: ConnectionToolErrorDetails): string {
  if (details.errorCode === "missing_connection") {
    return `Connection "${details.connectionTitle}" is not configured. ${details.setupHint}.`;
  }

  if (details.errorCode === "invalid_connection") {
    const reasonSuffix = details.reason ? ` (${details.reason})` : "";
    return `Connection "${details.connectionTitle}" is invalid${reasonSuffix}. ${details.setupHint}.`;
  }

  const reasonSuffix = details.reason ? ` (${details.reason})` : "";
  return `Connection "${details.connectionTitle}" failed authentication${reasonSuffix}. ${details.setupHint}.`;
}

function createConnectionFetchError(details: ConnectionToolErrorDetails): Error {
  const error = new Error(buildConnectionErrorMessage(details));
  Reflect.set(error, "details", details);
  Reflect.set(error, "connectionId", details.connectionId);
  return error;
}

function buildConnectionErrorDetails(args: {
  snapshot: ConnectionSnapshot;
  errorCode: ConnectionToolErrorDetails["errorCode"];
  status?: ConnectionStatus;
  reason?: string;
}): ConnectionToolErrorDetails {
  return {
    kind: "connection_error",
    ok: false,
    errorCode: args.errorCode,
    connectionId: args.snapshot.connectionId,
    connectionTitle: args.snapshot.title,
    status: args.status ?? args.snapshot.status,
    setupHint: args.snapshot.setupHint,
    reason: args.reason,
  };
}

function renderHttpAuthValueTemplate(args: {
  valueTemplate: string;
  secrets: Record<string, string>;
}): string {
  const placeholderPattern = /\{([^{}]+)\}/g;
  let rendered = args.valueTemplate;
  let match = placeholderPattern.exec(args.valueTemplate);

  while (match) {
    const placeholderRaw = match[1];
    const placeholder = placeholderRaw.trim();
    if (placeholder.length === 0) {
      throw new Error("httpAuth.valueTemplate contains an empty placeholder.");
    }

    const value = args.secrets[placeholder];
    if (typeof value !== "string" || value.trim().length === 0) {
      throw new Error(
        `httpAuth.valueTemplate references secret field "${placeholder}" with no stored value.`,
      );
    }

    rendered = rendered.replaceAll(`{${placeholderRaw}}`, value);
    match = placeholderPattern.exec(args.valueTemplate);
  }

  return rendered;
}

function isAllowedHttpAuthHost(args: {
  definition: ConnectionDefinition;
  targetHost: string;
}): boolean {
  const httpAuth = args.definition.httpAuth;
  if (!httpAuth) {
    return false;
  }

  const normalizedTargetHost = args.targetHost.toLowerCase();
  return httpAuth.allowedHosts.some((host) => host.toLowerCase() === normalizedTargetHost);
}

function buildHttpAuthHeader(args: {
  definition: ConnectionDefinition;
  secrets: Record<string, string>;
}): { headerName: string; value: string } {
  const httpAuth = args.definition.httpAuth;
  if (!httpAuth) {
    throw new Error("Connection does not define httpAuth.");
  }

  return {
    headerName: httpAuth.headerName,
    value: renderHttpAuthValueTemplate({
      valueTemplate: httpAuth.valueTemplate,
      secrets: args.secrets,
    }),
  };
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

  const getConnectionSecrets = async (connectionId: string): Promise<Record<string, string> | null> => {
    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionId);
    return connectionManager.getSecretsForOwner(entry.id, normalizedConnectionId);
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

  const runConnectionAwareHttpFetch = async (url: string, options?: HttpRequestOptions): Promise<HttpResponse> => {
    const connectionName = options?.connection;

    if (typeof connectionName !== "string" || connectionName.trim().length === 0) {
      return runExtensionHttpFetch(url, options);
    }

    const normalizedConnectionId = qualifyConnectionIdForEntry(entry.id, connectionName);
    const snapshot = await connectionManager.getSnapshot(normalizedConnectionId);

    if (!snapshot) {
      throw createConnectionFetchError({
        kind: "connection_error",
        ok: false,
        errorCode: "invalid_connection",
        connectionId: normalizedConnectionId,
        connectionTitle: normalizedConnectionId,
        status: "invalid",
        setupHint: "Reload the extension, then open /tools → Connections.",
        reason: "Connection requirement is not registered in this session.",
      });
    }

    if (snapshot.status !== "connected") {
      throw createConnectionFetchError(buildConnectionErrorDetails({
        snapshot,
        errorCode: mapStatusToConnectionErrorCode(snapshot.status),
        reason: snapshot.lastError,
      }));
    }

    let definition: ConnectionDefinition | null = null;
    try {
      connectionManager.assertConnectionOwnedBy(entry.id, normalizedConnectionId);
      definition = connectionManager.getDefinition(normalizedConnectionId);
    } catch {
      definition = null;
    }

    if (!definition || !definition.httpAuth) {
      throw createConnectionFetchError(buildConnectionErrorDetails({
        snapshot,
        errorCode: "invalid_connection",
        reason: "Connection does not define host-managed httpAuth.",
      }));
    }

    let parsedUrl: URL;
    try {
      parsedUrl = new URL(url);
    } catch {
      throw new Error("Invalid URL.");
    }

    if (!isAllowedHttpAuthHost({
      definition,
      targetHost: parsedUrl.hostname,
    })) {
      throw createConnectionFetchError(buildConnectionErrorDetails({
        snapshot,
        errorCode: "invalid_connection",
        reason: `Host \"${parsedUrl.hostname}\" is not allowed for this connection.`,
      }));
    }

    const secrets = await connectionManager.getSecretsForOwner(entry.id, normalizedConnectionId);
    if (!secrets) {
      throw createConnectionFetchError(buildConnectionErrorDetails({
        snapshot,
        errorCode: "missing_connection",
        status: "missing",
        reason: "No credentials stored for this connection.",
      }));
    }

    let authHeader: { headerName: string; value: string };
    try {
      authHeader = buildHttpAuthHeader({
        definition,
        secrets,
      });
    } catch (error: unknown) {
      const reason = error instanceof Error ? error.message : String(error);
      throw createConnectionFetchError(buildConnectionErrorDetails({
        snapshot,
        errorCode: "invalid_connection",
        reason,
      }));
    }

    const mergedHeaders: Record<string, string> = {
      ...(options?.headers ?? {}),
      [authHeader.headerName]: authHeader.value,
    };

    const requestOptions: HttpRequestOptions = {
      method: options?.method,
      headers: mergedHeaders,
      body: options?.body,
      timeoutMs: options?.timeoutMs,
      connection: normalizedConnectionId,
    };

    const response = await runExtensionHttpFetch(url, requestOptions);

    if (response.status === 401 || response.status === 403) {
      const authFailureMessage = `HTTP ${response.status} ${response.statusText}`;

      try {
        await connectionManager.markRuntimeAuthFailure(normalizedConnectionId, {
          message: authFailureMessage,
        });
      } catch {
        // best-effort status update only
      }

      const latestSnapshot = await connectionManager.getSnapshot(normalizedConnectionId) ?? snapshot;
      throw createConnectionFetchError(buildConnectionErrorDetails({
        snapshot: latestSnapshot,
        status: "error",
        errorCode: "connection_auth_failed",
        reason: latestSnapshot.lastError ?? authFailureMessage,
      }));
    }

    return response;
  };

  const host: HostActivationBridge = {
    getAgent: getRequiredActiveAgent,
    llmComplete: runExtensionLlmCompletion,
    httpFetch: runConnectionAwareHttpFetch,
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
    getConnectionSecrets,
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
    httpFetch: runConnectionAwareHttpFetch,
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
    getConnectionSecrets,
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
