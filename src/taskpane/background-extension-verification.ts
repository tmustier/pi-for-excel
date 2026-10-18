import { prepareDurableReload } from "../durable/model-recovery.js";
import { entryMessages, readConversationHistory } from "../durable/conversation-view.js";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { isJsonValue, type JsonValue } from "@earendil-works/chord";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import { LegacySessionDoc } from "../durable/session-import.js";
import { InboxDoc } from "@earendil-works/pi-durable";
import { Type } from "typebox";
import { Value } from "typebox/value";
import { StringEnum } from "../tools/string-enum.js";
import type { DurableSessions } from "../durable/sessions.js";
import type { ConnectionManager } from "../connections/manager.js";
import {
  loadConnectionStoreDocumentForUpdate,
  saveConnectionStoreDocument,
} from "../connections/store.js";
import {
  ALL_EXTENSION_CAPABILITIES,
  type ExtensionCapability,
} from "../extensions/permissions.js";
import type { ExtensionRuntimeManager } from "../extensions/runtime-manager.js";
import { getExtensionStorageValue } from "../extensions/storage-store.js";
import {
  loadStoredExtensions,
  saveStoredExtensions,
} from "../extensions/store.js";
import {
  setExperimentalFeatureEnabled,
  type ExperimentalFeatureId,
} from "../experiments/flags.js";
import type { BrowserModelRuntime } from "../models/browser-model-runtime.js";
import { getAppStorage } from "../storage/local/app-storage.js";
import type { SessionRuntime } from "./session-runtime-manager.js";
import type { SettingsWriter } from "../storage/local/settings-store.js";
import type { DurableTaskpane } from "../durable/taskpane.js";

export const EXTENSION_VERIFICATION_COMMAND_TYPES = [
  "assertLastAssistantText",
  "getMessages",
  "getPendingMessages",
  "abort",
  "configure",
  "listSessions",
  "exportTranscript",
  "executeTool",
  "extensionInstallCode",
  "extensionList",
  "extensionSetCapability",
  "extensionSetEnabled",
  "extensionReload",
  "extensionUninstall",
  "extensionStorageGet",
  "connectionSetSecrets",
  "modelsRefresh",
  "modelsList",
  "setExperiment",
  "stageInlineExtensionUpgrade",
  "reloadTaskpane",
  "durableStatus",
  "durableSubmit",
  "durableFault",
  "durableResumeFault",
  "durableReconcile",
  "durableImport",
] as const;

export type ExtensionVerificationCommandType =
  (typeof EXTENSION_VERIFICATION_COMMAND_TYPES)[number];

interface JsonRecord {
  [key: string]: unknown;
}

export interface ExtensionVerificationOptions {
  durable: DurableTaskpane;
  getActiveRuntime: () => SessionRuntime | null;
  getSessions?: () => DurableSessions;
  /** Must dispatch through the installed Durable tool task, including mutation gates. */
  executeTool?: (runtime: SessionRuntime, name: string, args: { [key: string]: JsonValue }) => Promise<JsonValue>;
  extensionManager: ExtensionRuntimeManager;
  connectionManager: ConnectionManager;
  modelRuntime: BrowserModelRuntime;
  refreshModels: (allowNetwork: boolean) => Promise<{ errors: string[] }>;
}

const configureSchema = Type.Object({
  model: Type.Optional(Type.Object({ provider: Type.String(), modelId: Type.String() }, { additionalProperties: false })),
  thinkingLevel: Type.Optional(StringEnum(["off", "minimal", "low", "medium", "high", "xhigh", "max"])),
}, { additionalProperties: false });

const executeToolSchema = Type.Object({
  name: Type.String({ minLength: 1 }),
  args: Type.Record(Type.String(), Type.Unknown()),
}, { additionalProperties: false });

function isPayloadShape(value: unknown): value is JsonRecord {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function stringField(value: unknown, key: string): string | undefined {
  if (!isPayloadShape(value)) return undefined;
  const field = value[key];
  return typeof field === "string" && field.trim().length > 0 ? field.trim() : undefined;
}

function booleanField(value: unknown, key: string): boolean | undefined {
  if (!isPayloadShape(value)) return undefined;
  const field = value[key];
  return typeof field === "boolean" ? field : undefined;
}

function stringMapField(value: unknown, key: string): Record<string, string> | undefined {
  if (!isPayloadShape(value)) return undefined;
  const field = value[key];
  if (!isPayloadShape(field)) return undefined;

  const result: Record<string, string> = {};
  for (const [fieldName, fieldValue] of Object.entries(field)) {
    if (typeof fieldValue !== "string") return undefined;
    result[fieldName] = fieldValue;
  }
  return result;
}

export function isExtensionVerificationCommand(
  value: string,
): value is ExtensionVerificationCommandType {
  return EXTENSION_VERIFICATION_COMMAND_TYPES.some((command) => command === value);
}

async function readRuntimeMessages(runtime: SessionRuntime): Promise<AgentMessage[]> {
  const legacy = await runtime.view.harness.snapshot(LegacySessionDoc, runtime.view.conversation.id, context);
  return (await readConversationHistory(runtime.view.conversation))
    .flatMap((entry) => entryMessages(entry, legacy?.originalSessionJson));
}

function originalMessages(messages: readonly AgentMessage[]): AgentMessage[] {
  return messages.flatMap((message) => message.role === "archivedMessages" ? originalMessages(message.archivedMessages) : [message]);
}

async function latestAssistantText(runtime: SessionRuntime | null): Promise<string | null> {
  if (!runtime) return null;

  const messages = originalMessages(await readRuntimeMessages(runtime));
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (!message || message.role !== "assistant") continue;
    return message.content
      .filter((part) => part.type === "text")
      .map((part) => part.text)
      .join("");
  }

  return null;
}

async function assertLastAssistantText(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const expected = stringField(payload, "expected");
  if (!expected) throw new Error("assertLastAssistantText requires payload.expected");

  const actual = await latestAssistantText(options.getActiveRuntime());
  return {
    matches: actual === expected,
    expectedLength: expected.length,
    actualLength: actual?.length ?? 0,
  };
}

function extensionStatusSummary(options: ExtensionVerificationOptions): JsonRecord[] {
  return options.extensionManager.list().map((status) => ({
    id: status.id,
    name: status.name,
    enabled: status.enabled,
    loaded: status.loaded,
    sourceKind: status.source.kind,
    trust: status.trust,
    runtimeMode: status.runtimeMode,
    permissionsEnforced: status.permissionsEnforced,
    grantedCapabilities: status.grantedCapabilities,
    effectiveCapabilities: status.effectiveCapabilities,
    connectionIds: options.connectionManager.listRegisteredConnectionIds()
      .filter((connectionId) => connectionId.startsWith(`${status.id}.`)),
    modelProviderIds: status.modelProviderIds,
    lastError: status.lastError,
  }));
}

function parseExtensionCapability(value: string): ExtensionCapability {
  for (const capability of ALL_EXTENSION_CAPABILITIES) {
    if (value === capability) return capability;
  }
  throw new Error(`Unknown extension capability: ${value}`);
}

function parseExperimentalFeatureId(value: string): ExperimentalFeatureId {
  switch (value) {
    case "ui_dark_mode":
    case "remote_extension_urls":
    case "extension_permission_gates":
    case "extension_sandbox_runtime":
    case "extension_widget_v2":
      return value;
    default:
      throw new Error(`Unknown experimental feature: ${value}`);
  }
}

async function installExtensionCode(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const name = stringField(payload, "name");
  const code = stringField(payload, "code");
  if (!name || !code) throw new Error("extensionInstallCode requires payload.name and payload.code");

  const extensionId = await options.extensionManager.installFromCode(name, code);
  return { extensionId };
}

async function setExtensionCapability(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const extensionId = stringField(payload, "extensionId");
  const capabilityRaw = stringField(payload, "capability");
  const allowed = booleanField(payload, "allowed");
  if (!extensionId || !capabilityRaw || allowed === undefined) {
    throw new Error("extensionSetCapability requires extensionId, capability and boolean allowed");
  }

  const capability = parseExtensionCapability(capabilityRaw);
  await options.extensionManager.setExtensionCapability(extensionId, capability, allowed);
  return { extensionId, capability, allowed };
}

async function setExtensionEnabled(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const extensionId = stringField(payload, "extensionId");
  const enabled = booleanField(payload, "enabled");
  if (!extensionId || enabled === undefined) {
    throw new Error("extensionSetEnabled requires extensionId and boolean enabled");
  }

  await options.extensionManager.setExtensionEnabled(extensionId, enabled);
  return { extensionId, enabled };
}

async function reloadExtension(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const extensionId = stringField(payload, "extensionId");
  if (!extensionId) throw new Error("extensionReload requires extensionId");

  await options.extensionManager.reloadExtension(extensionId);
  return { extensionId, reloaded: true };
}

async function uninstallExtension(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const extensionId = stringField(payload, "extensionId");
  if (!extensionId) throw new Error("extensionUninstall requires extensionId");

  await options.extensionManager.uninstallExtension(extensionId);
  return { extensionId, uninstalled: true };
}

async function readExtensionStorage(payload: unknown): Promise<JsonRecord> {
  const extensionId = stringField(payload, "extensionId");
  const key = stringField(payload, "key");
  if (!extensionId || !key) throw new Error("extensionStorageGet requires extensionId and key");

  return {
    extensionId,
    key,
    value: await getExtensionStorageValue(getAppStorage().settings, extensionId, key),
  };
}

async function setConnectionSecrets(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const connectionId = stringField(payload, "connectionId");
  const secrets = stringMapField(payload, "secrets");
  if (!connectionId || !secrets) {
    throw new Error("connectionSetSecrets requires connectionId and string-valued secrets");
  }

  await options.connectionManager.updateSecretsFromHost(connectionId, secrets);
  const snapshot = await options.connectionManager.getSnapshot(connectionId);
  return {
    connectionId,
    status: snapshot?.status ?? null,
    savedFieldIds: Object.keys(secrets).sort(),
  };
}

async function refreshModels(
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<JsonRecord> {
  const allowNetwork = booleanField(payload, "allowNetwork") ?? true;
  const result = await options.refreshModels(allowNetwork);
  return { allowNetwork, errors: result.errors };
}

function listModels(payload: unknown, options: ExtensionVerificationOptions): JsonRecord {
  const provider = stringField(payload, "provider");
  if (!provider) throw new Error("modelsList requires payload.provider");

  return {
    provider,
    models: options.modelRuntime.models.getModels(provider).map((model) => ({
      id: model.id,
      name: model.name,
      api: model.api,
      provider: model.provider,
      baseUrl: model.baseUrl,
      contextWindow: model.contextWindow,
      maxTokens: model.maxTokens,
    })),
  };
}

function setExperiment(payload: unknown): JsonRecord {
  const featureRaw = stringField(payload, "feature");
  const enabled = booleanField(payload, "enabled");
  if (!featureRaw || enabled === undefined) {
    throw new Error("setExperiment requires feature and boolean enabled");
  }

  const feature = parseExperimentalFeatureId(featureRaw);
  setExperimentalFeatureEnabled(feature, enabled);
  return { feature, enabled };
}

export async function stageInlineExtensionUpgrade(
  payload: unknown,
  settings: SettingsWriter = getAppStorage().settings,
): Promise<JsonRecord> {
  const extensionId = stringField(payload, "extensionId");
  const code = stringField(payload, "code");
  const connectionId = stringField(payload, "connectionId");
  const secrets = stringMapField(payload, "secrets");
  if (!extensionId || !code || !connectionId || !secrets) {
    throw new Error(
      "stageInlineExtensionUpgrade requires extensionId, code, connectionId and string-valued secrets",
    );
  }
  if (!connectionId.startsWith(`${extensionId}.`)) {
    throw new Error("Staged connection must be owned by the staged extension");
  }

  const extensions = await loadStoredExtensions(settings);
  const entry = extensions.find((candidate) => candidate.id === extensionId);
  if (!entry) throw new Error("Extension not found");
  if (entry.source.kind !== "inline") throw new Error("Only inline extensions can be staged");

  const now = new Date().toISOString();
  entry.source = { kind: "inline", code };
  entry.updatedAt = now;

  const connections = await loadConnectionStoreDocumentForUpdate(settings);
  connections[connectionId] = {
    status: "connected",
    secrets,
    lastValidatedAt: now,
  };

  await saveStoredExtensions(settings, extensions);
  await saveConnectionStoreDocument(settings, connections);
  return {
    extensionId,
    connectionId,
    staged: true,
    savedFieldIds: Object.keys(secrets).sort(),
  };
}

function scheduleTaskpaneReload(): JsonRecord {
  window.setTimeout(() => {
    prepareDurableReload();
    window.location.reload();
  }, 100);
  return { scheduled: true };
}

function assertNever(value: never): never {
  throw new Error(`Unknown extension verification command: ${String(value)}`);
}

export async function executeExtensionVerificationCommand(
  type: ExtensionVerificationCommandType,
  payload: unknown,
  options: ExtensionVerificationOptions,
): Promise<unknown> {
  switch (type) {
    case "listSessions":
      return { sessions: await (options.getSessions?.() ?? options.durable.sessions).getAllMetadata() };
    case "exportTranscript": {
      const id = stringField(payload, "sessionId") ?? String(options.getActiveRuntime()?.view.conversation.id ?? "");
      if (!id) throw new Error("exportTranscript requires an active conversation or sessionId");
      const session = await (options.getSessions?.() ?? options.durable.sessions).loadSession(id);
      if (!session) throw new Error(`Session not found: ${id}`);
      return session;
    }
    case "getMessages":
    case "getPendingMessages":
    case "abort":
    case "configure":
    case "executeTool": {
      const runtime = options.getActiveRuntime();
      if (!runtime) throw new Error("No active conversation");
      const conversation = runtime.view.conversation;
      if (type === "getMessages") {
        return { messages: await readRuntimeMessages(runtime) };
      }
      if (type === "getPendingMessages") {
        return { pending: (await runtime.view.harness.snapshot(InboxDoc, conversation.id, context))?.items ?? [] };
      }
      if (type === "abort") {
        await conversation.abort(context);
        await runtime.view.refresh();
        return { aborted: true };
      }
      if (type === "configure") {
        if (!Value.Check(configureSchema, payload)) throw new Error("Invalid configure payload");
        if (runtime.view.current.busy || runtime.view.current.inbox.length) throw new Error("An idle active chat is required.");
        await conversation.configure(payload, context);
        await runtime.view.refresh();
        await runtime.refreshCapabilities();
        return { configured: true };
      }
      if (!Value.Check(executeToolSchema, payload) || !isJsonValue(payload.args)) throw new Error("executeTool requires name and JSON args");
      if (!runtime.tools.some((tool) => tool.name === payload.name)) throw new Error(`Tool not available: ${payload.name}`);
      // Never fall back to runtime.tools[].execute: that skips Durable mutation admission.
      if (!options.executeTool) throw new Error("Installed Durable tool execution is unavailable");
      return options.executeTool(runtime, payload.name, payload.args);
    }
    case "durableStatus":
    case "durableSubmit":
    case "durableFault":
    case "durableResumeFault":
    case "durableReconcile":
    case "durableImport":
      return options.durable.execute(type, JSON.stringify(payload ?? {}));
    case "assertLastAssistantText":
      return assertLastAssistantText(payload, options);
    case "extensionInstallCode":
      return await installExtensionCode(payload, options);
    case "extensionList":
      return { extensions: extensionStatusSummary(options) };
    case "extensionSetCapability":
      return await setExtensionCapability(payload, options);
    case "extensionSetEnabled":
      return await setExtensionEnabled(payload, options);
    case "extensionReload":
      return await reloadExtension(payload, options);
    case "extensionUninstall":
      return await uninstallExtension(payload, options);
    case "extensionStorageGet":
      return await readExtensionStorage(payload);
    case "connectionSetSecrets":
      return await setConnectionSecrets(payload, options);
    case "modelsRefresh":
      return await refreshModels(payload, options);
    case "modelsList":
      return listModels(payload, options);
    case "setExperiment":
      return setExperiment(payload);
    case "stageInlineExtensionUpgrade":
      return await stageInlineExtensionUpgrade(payload);
    case "reloadTaskpane":
      return scheduleTaskpaneReload();
    default:
      return assertNever(type);
  }
}
