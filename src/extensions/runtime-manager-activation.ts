import type { Agent } from "@mariozechner/pi-agent-core";

import type {
  CreateExtensionAPIOptions,
  HttpRequestOptions,
  HttpResponse,
  LlmCompletionRequest,
  LlmCompletionResult,
} from "../commands/extension-api.js";
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
  | "isCapabilityEnabled"
  | "formatCapabilityError"
  | "toast"
  | "widgetOwnerId"
  | "widgetApiV2Enabled"
>;

export interface RuntimeManagerActivationBridge {
  host: HostActivationBridge;
  sandbox: SandboxActivationBridge;
}

export interface BuildRuntimeManagerActivationBridgeOptions {
  entry: StoredExtensionEntry;
  settings: ExtensionSettingsStore;
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

  const listSkills = () => listExtensionSkillSummaries(settings);
  const readSkill = (name: string) => readExtensionSkill(settings, name);
  const installSkill = (name: string, markdown: string) => installExternalExtensionSkill(settings, name, markdown);
  const uninstallSkill = (name: string) => uninstallExternalExtensionSkill(settings, name);

  const downloadFile = (filename: string, content: string, mimeType?: string): void => {
    triggerExtensionDownload(filename, content, mimeType);
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
