/**
 * Runtime extension manager.
 *
 * Responsibilities:
 * - load persisted extension registry
 * - activate/deactivate extensions with failure isolation
 * - track extension-owned commands and tools for clean unload
 * - expose extension tool list so runtimes can refresh Agent toolsets
 */

import type { Agent, AgentEvent, AgentTool } from "@mariozechner/pi-agent-core";

import {
  createExtensionAPI,
  loadExtension,
  type ExtensionCommand,
  type HttpRequestOptions,
  type HttpResponse,
  type LlmCompletionRequest,
  type LlmCompletionResult,
  type LoadedExtensionHandle,
} from "../commands/extension-api.js";
import { activateExtensionInSandbox } from "./sandbox-runtime.js";
import {
  describeExtensionRuntimeMode,
  resolveExtensionRuntimeMode,
  type ExtensionRuntimeMode,
} from "./runtime-mode.js";
import { commandRegistry } from "../commands/types.js";
import {
  describeExtensionCapability,
  describeStoredExtensionTrust,
  deriveStoredExtensionTrust,
  getDefaultPermissionsForTrust,
  isExtensionCapabilityAllowed,
  listAllExtensionCapabilities,
  listGrantedExtensionCapabilities,
  setExtensionCapabilityAllowed,
  type ExtensionCapability,
  type StoredExtensionPermissions,
  type StoredExtensionTrust,
} from "./permissions.js";
import {
  loadStoredExtensions,
  saveStoredExtensions,
  type ExtensionSettingsStore,
  type StoredExtensionEntry,
  type StoredExtensionSource,
} from "./store.js";
import { isExperimentalFeatureEnabled } from "../experiments/flags.js";
import { clearExtensionWidgets } from "./internal/widget-surface.js";
import { showToast } from "../ui/toast.js";
import { getEnabledProxyBaseUrl, resolveOutboundRequestUrl } from "../tools/external-fetch.js";
import {
  clearExtensionStorage,
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
import {
  createExtensionAgentMessage,
  describeExtensionSource,
  extractAssistantText,
  getRuntimeManagerErrorMessage,
  isApiModel,
  isBlockedExtensionHostname,
  normalizeDownloadFilename,
  normalizeExtensionName,
  normalizeHttpOptions,
  normalizeInlineCode,
  normalizeRemoteUrl,
  parseLlmMessages,
  readLimitedResponseBody,
  resolveModelForCompletion,
} from "./runtime-manager-helpers.js";

type AnyAgentTool = AgentTool;

type ManagerListener = () => void;

interface LoadedExtensionState {
  entryId: string;
  runtimeMode: ExtensionRuntimeMode;
  commandNames: Set<string>;
  toolNames: Set<string>;
  eventUnsubscribers: Set<() => void>;
  handle: LoadedExtensionHandle | null;
  inlineBlobUrl: string | null;
}

export interface ExtensionRuntimeStatus {
  id: string;
  name: string;
  enabled: boolean;
  loaded: boolean;
  source: StoredExtensionSource;
  sourceLabel: string;
  trust: StoredExtensionTrust;
  trustLabel: string;
  runtimeMode: ExtensionRuntimeMode;
  runtimeLabel: string;
  permissions: StoredExtensionPermissions;
  grantedCapabilities: ExtensionCapability[];
  effectiveCapabilities: ExtensionCapability[];
  permissionsEnforced: boolean;
  commandNames: string[];
  toolNames: string[];
  lastError: string | null;
}

export interface ExtensionRuntimeManagerOptions {
  settings: ExtensionSettingsStore;
  getActiveAgent: () => Agent | null;
  refreshRuntimeTools: () => Promise<void>;
  reservedToolNames: ReadonlySet<string>;
  loadExtensionFromSource?: typeof loadExtension;
  activateInSandbox?: typeof activateExtensionInSandbox;
  showToastMessage?: typeof showToast;
}

export class ExtensionRuntimeManager {
  private readonly settings: ExtensionSettingsStore;
  private readonly getActiveAgent: () => Agent | null;
  private readonly refreshRuntimeTools: () => Promise<void>;
  private readonly reservedToolNames: ReadonlySet<string>;
  private readonly loadExtensionFromSource: typeof loadExtension;
  private readonly activateInSandbox: typeof activateExtensionInSandbox;
  private readonly showToastMessage: typeof showToast;

  private readonly listeners = new Set<ManagerListener>();
  private readonly activeStates = new Map<string, LoadedExtensionState>();
  private readonly extensionTools = new Map<string, AnyAgentTool>();
  private readonly toolOwners = new Map<string, string>();
  private readonly commandOwners = new Map<string, string>();
  private readonly lastErrors = new Map<string, string>();

  private entries: StoredExtensionEntry[] = [];
  private initialized = false;

  constructor(options: ExtensionRuntimeManagerOptions) {
    this.settings = options.settings;
    this.getActiveAgent = options.getActiveAgent;
    this.refreshRuntimeTools = options.refreshRuntimeTools;
    this.reservedToolNames = options.reservedToolNames;
    this.loadExtensionFromSource = options.loadExtensionFromSource ?? loadExtension;
    this.activateInSandbox = options.activateInSandbox ?? activateExtensionInSandbox;
    this.showToastMessage = options.showToastMessage ?? showToast;
  }

  subscribe(listener: ManagerListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  list(): ExtensionRuntimeStatus[] {
    const permissionsEnforced = isExperimentalFeatureEnabled("extension_permission_gates");
    const sandboxHostFallbackEnabled = isExperimentalFeatureEnabled("extension_sandbox_runtime");

    return this.entries.map((entry) => {
      const state = this.activeStates.get(entry.id);
      const grantedCapabilities = listGrantedExtensionCapabilities(entry.permissions);
      const effectiveCapabilities = permissionsEnforced
        ? grantedCapabilities
        : listAllExtensionCapabilities();
      const runtimeMode = state
        ? state.runtimeMode
        : resolveExtensionRuntimeMode(entry.trust, sandboxHostFallbackEnabled);

      return {
        id: entry.id,
        name: entry.name,
        enabled: entry.enabled,
        loaded: Boolean(state),
        source: entry.source,
        sourceLabel: describeExtensionSource(entry.source),
        trust: entry.trust,
        trustLabel: describeStoredExtensionTrust(entry.trust),
        runtimeMode,
        runtimeLabel: describeExtensionRuntimeMode(runtimeMode),
        permissions: entry.permissions,
        grantedCapabilities,
        effectiveCapabilities,
        permissionsEnforced,
        commandNames: state ? Array.from(state.commandNames).sort() : [],
        toolNames: state ? Array.from(state.toolNames).sort() : [],
        lastError: this.lastErrors.get(entry.id) ?? null,
      };
    });
  }

  getRegisteredTools(): AnyAgentTool[] {
    return Array.from(this.extensionTools.values());
  }

  async initialize(): Promise<void> {
    if (this.initialized) {
      return;
    }

    this.entries = await loadStoredExtensions(this.settings);

    for (const entry of this.entries) {
      if (!entry.enabled) {
        continue;
      }

      await this.tryActivateEntry(entry);
    }

    this.initialized = true;
    this.notify();
  }

  async reloadExtension(entryId: string): Promise<void> {
    const entry = this.getEntryById(entryId);
    if (!entry) {
      throw new Error("Extension not found");
    }

    await this.deactivateEntry(entry.id);

    if (entry.enabled) {
      await this.tryActivateEntry(entry);
    }

    this.notify();
  }

  async setExtensionEnabled(entryId: string, enabled: boolean): Promise<void> {
    const entry = this.getEntryById(entryId);
    if (!entry) {
      throw new Error("Extension not found");
    }

    if (entry.enabled === enabled) {
      return;
    }

    entry.enabled = enabled;
    entry.updatedAt = new Date().toISOString();
    await this.persistEntries();

    if (enabled) {
      await this.tryActivateEntry(entry);
    } else {
      await this.deactivateEntry(entry.id);
    }

    this.notify();
  }

  async setExtensionCapability(
    entryId: string,
    capability: ExtensionCapability,
    allowed: boolean,
  ): Promise<void> {
    const entry = this.getEntryById(entryId);
    if (!entry) {
      throw new Error("Extension not found");
    }

    const existing = isExtensionCapabilityAllowed(entry.permissions, capability);
    if (existing === allowed) {
      return;
    }

    entry.permissions = setExtensionCapabilityAllowed(entry.permissions, capability, allowed);
    entry.updatedAt = new Date().toISOString();
    await this.persistEntries();

    if (entry.enabled) {
      await this.reloadExtension(entry.id);
      return;
    }

    this.notify();
  }

  async uninstallExtension(entryId: string): Promise<void> {
    const entryIndex = this.entries.findIndex((entry) => entry.id === entryId);
    if (entryIndex < 0) {
      throw new Error("Extension not found");
    }

    await this.deactivateEntry(entryId);
    await clearExtensionStorage(this.settings, entryId);
    this.entries.splice(entryIndex, 1);
    this.lastErrors.delete(entryId);

    await this.persistEntries();
    this.notify();
  }

  async installFromUrl(name: string, url: string): Promise<string> {
    return this.installEntry({
      name: normalizeExtensionName(name),
      source: {
        kind: "module",
        specifier: normalizeRemoteUrl(url),
      },
    });
  }

  async installFromCode(name: string, code: string): Promise<string> {
    return this.installEntry({
      name: normalizeExtensionName(name),
      source: {
        kind: "inline",
        code: normalizeInlineCode(code),
      },
    });
  }

  async installFromModuleSpecifier(name: string, specifier: string): Promise<string> {
    const normalizedSpecifier = specifier.trim();
    if (normalizedSpecifier.length === 0) {
      throw new Error("Module specifier cannot be empty");
    }

    return this.installEntry({
      name: normalizeExtensionName(name),
      source: {
        kind: "module",
        specifier: normalizedSpecifier,
      },
    });
  }

  private notify(): void {
    for (const listener of this.listeners) {
      try {
        listener();
      } catch (error: unknown) {
        console.warn("[pi] Extension manager listener failed:", getRuntimeManagerErrorMessage(error));
      }
    }
  }

  private getEntryById(entryId: string): StoredExtensionEntry | null {
    return this.entries.find((entry) => entry.id === entryId) ?? null;
  }

  private async persistEntries(): Promise<void> {
    await saveStoredExtensions(this.settings, this.entries);
  }

  private getRequiredActiveAgent(): Agent {
    const activeAgent = this.getActiveAgent();
    if (!activeAgent) {
      throw new Error("No active runtime available for extension activation");
    }

    return activeAgent;
  }

  private async runExtensionLlmCompletion(
    entry: StoredExtensionEntry,
    request: LlmCompletionRequest,
  ): Promise<LlmCompletionResult> {
    const agent = this.getRequiredActiveAgent();

    if (!isApiModel(agent.state.model)) {
      throw new Error("Active model is unavailable for extension LLM completion.");
    }

    const model = resolveModelForCompletion({
      fallbackModel: agent.state.model,
      requestedModel: request.model,
    });

    const apiKey = agent.getApiKey ? await agent.getApiKey(model.provider) : undefined;
    if (!apiKey) {
      throw new Error(`No API key available for provider "${model.provider}".`);
    }

    if (!Array.isArray(request.messages)) {
      throw new Error("llm.complete requires a messages array.");
    }

    const stream = await agent.streamFn(
      model,
      {
        systemPrompt: request.systemPrompt,
        messages: parseLlmMessages(request.messages, model),
      },
      {
        apiKey,
        sessionId: agent.sessionId,
        maxTokens: request.maxTokens,
      },
    );

    const result = await stream.result();
    if (result.stopReason === "error" || result.stopReason === "aborted") {
      throw new Error(result.errorMessage ?? "LLM completion failed.");
    }

    return {
      content: extractAssistantText(result),
      model: `${result.provider}/${result.model}`,
      usage: {
        inputTokens: result.usage.input,
        outputTokens: result.usage.output,
      },
    };
  }

  private async runExtensionHttpFetch(url: string, options?: HttpRequestOptions): Promise<HttpResponse> {
    let parsedUrl: URL;

    try {
      parsedUrl = new URL(url);
    } catch {
      throw new Error("Invalid URL.");
    }

    if (parsedUrl.protocol !== "http:" && parsedUrl.protocol !== "https:") {
      throw new Error("Only http:// and https:// URLs are supported.");
    }

    if (isBlockedExtensionHostname(parsedUrl.hostname)) {
      throw new Error("Blocked target host: local and private-network addresses are not allowed.");
    }

    const normalizedOptions = normalizeHttpOptions(options);
    const proxyBaseUrl = await getEnabledProxyBaseUrl(this.settings);
    const resolved = resolveOutboundRequestUrl({
      targetUrl: parsedUrl.toString(),
      proxyBaseUrl,
    });

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), normalizedOptions.timeoutMs);

    try {
      const response = await fetch(resolved.requestUrl, {
        method: normalizedOptions.method,
        headers: normalizedOptions.headers,
        body: normalizedOptions.body,
        signal: controller.signal,
      });

      const responseBody = await readLimitedResponseBody(response);
      const headers: Record<string, string> = {};
      response.headers.forEach((value, key) => {
        headers[key] = value;
      });

      return {
        status: response.status,
        statusText: response.statusText,
        headers,
        body: responseBody,
      };
    } catch (error: unknown) {
      if (error instanceof DOMException && error.name === "AbortError") {
        throw new Error(`HTTP request timed out after ${normalizedOptions.timeoutMs}ms.`);
      }

      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  private async writeExtensionClipboard(text: string): Promise<void> {
    if (typeof navigator === "undefined" || !navigator.clipboard) {
      throw new Error("Clipboard API is unavailable.");
    }

    await navigator.clipboard.writeText(text);
  }

  private triggerExtensionDownload(filename: string, content: string, mimeType?: string): void {
    const normalizedFilename = normalizeDownloadFilename(filename);
    const blob = new Blob([content], {
      type: mimeType && mimeType.trim().length > 0 ? mimeType : "text/plain;charset=utf-8",
    });

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = normalizedFilename;
    anchor.style.display = "none";

    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();

    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 0);
  }

  private resolveRuntimeMode(entry: StoredExtensionEntry): ExtensionRuntimeMode {
    return resolveExtensionRuntimeMode(
      entry.trust,
      isExperimentalFeatureEnabled("extension_sandbox_runtime"),
    );
  }

  private async installEntry(input: {
    name: string;
    source: StoredExtensionSource;
  }): Promise<string> {
    const now = new Date().toISOString();
    const id = `ext.${crypto.randomUUID()}`;

    const trust = deriveStoredExtensionTrust(id, input.source);
    const entry: StoredExtensionEntry = {
      id,
      name: input.name,
      enabled: true,
      source: input.source,
      trust,
      permissions: getDefaultPermissionsForTrust(trust),
      createdAt: now,
      updatedAt: now,
    };

    this.entries.push(entry);
    await this.persistEntries();

    await this.tryActivateEntry(entry);
    this.notify();
    return id;
  }

  private async tryActivateEntry(entry: StoredExtensionEntry): Promise<void> {
    try {
      await this.activateEntry(entry);
      this.lastErrors.delete(entry.id);
    } catch (error: unknown) {
      const message = getRuntimeManagerErrorMessage(error);
      this.lastErrors.set(entry.id, message);
      console.warn(`[pi] Failed to load extension "${entry.name}": ${message}`);
    }
  }

  private async activateEntry(entry: StoredExtensionEntry): Promise<void> {
    await this.deactivateEntry(entry.id);

    const state: LoadedExtensionState = {
      entryId: entry.id,
      runtimeMode: this.resolveRuntimeMode(entry),
      commandNames: new Set<string>(),
      toolNames: new Set<string>(),
      eventUnsubscribers: new Set<() => void>(),
      handle: null,
      inlineBlobUrl: null,
    };

    let activationPhase = true;
    let toolsChangedDuringActivation = false;

    const refreshToolsForDynamicChange = (): void => {
      void this.refreshRuntimeTools().catch((error: unknown) => {
        console.warn(`[pi] Failed to refresh tools after extension tool update: ${getRuntimeManagerErrorMessage(error)}`);
      });
    };

    const registerCommand = (name: string, cmd: ExtensionCommand) => {
      const existing = commandRegistry.get(name);
      if (existing) {
        throw new Error(`Command /${name} is already registered (${existing.source})`);
      }

      commandRegistry.register({
        name,
        description: cmd.description,
        source: "extension",
        execute: cmd.handler,
      });

      this.commandOwners.set(name, entry.id);
      state.commandNames.add(name);
    };

    const registerTool = (tool: AnyAgentTool) => {
      if (this.reservedToolNames.has(tool.name)) {
        throw new Error(`Tool name "${tool.name}" conflicts with a built-in tool`);
      }

      const existingOwner = this.toolOwners.get(tool.name);
      if (existingOwner && existingOwner !== entry.id) {
        throw new Error(`Tool name "${tool.name}" is already registered by another extension`);
      }

      if (state.toolNames.has(tool.name)) {
        throw new Error(`Tool name "${tool.name}" is registered multiple times by this extension`);
      }

      this.toolOwners.set(tool.name, entry.id);
      this.extensionTools.set(tool.name, tool);
      state.toolNames.add(tool.name);

      if (activationPhase) {
        toolsChangedDuringActivation = true;
      } else {
        refreshToolsForDynamicChange();
      }
    };

    const unregisterTool = (toolName: string): void => {
      const normalizedName = toolName.trim();
      if (normalizedName.length === 0) {
        throw new Error("Tool name cannot be empty");
      }

      if (!state.toolNames.has(normalizedName)) {
        throw new Error(`Tool name "${normalizedName}" is not registered by this extension`);
      }

      const owner = this.toolOwners.get(normalizedName);
      if (owner !== entry.id) {
        throw new Error(`Tool name "${normalizedName}" is not owned by this extension`);
      }

      state.toolNames.delete(normalizedName);
      this.toolOwners.delete(normalizedName);
      this.extensionTools.delete(normalizedName);

      if (!activationPhase) {
        refreshToolsForDynamicChange();
      }
    };

    const subscribeAgentEvents = (handler: (ev: AgentEvent) => void): (() => void) => {
      const unsubscribe = this.getRequiredActiveAgent().subscribe(handler);
      state.eventUnsubscribers.add(unsubscribe);

      return () => {
        if (!state.eventUnsubscribers.has(unsubscribe)) {
          return;
        }
        state.eventUnsubscribers.delete(unsubscribe);
        unsubscribe();
      };
    };

    const isCapabilityEnabled = (capability: ExtensionCapability): boolean => {
      if (!isExperimentalFeatureEnabled("extension_permission_gates")) {
        return true;
      }

      return isExtensionCapabilityAllowed(entry.permissions, capability);
    };

    const formatCapabilityError = (capability: ExtensionCapability): string => {
      const capabilityLabel = describeExtensionCapability(capability);
      return (
        `Permission denied for extension "${entry.name}": cannot ${capabilityLabel}. `
        + "Disable /experimental extension-permissions or adjust extension permissions."
      );
    };

    const widgetApiV2Enabled = isExperimentalFeatureEnabled("extension_widget_v2");

    try {
      if (state.runtimeMode === "sandbox-iframe") {
        const source = entry.source.kind === "inline"
          ? {
            kind: "inline" as const,
            code: entry.source.code,
          }
          : {
            kind: "module" as const,
            specifier: entry.source.specifier,
          };

        state.handle = await this.activateInSandbox({
          instanceId: `${entry.id}.${crypto.randomUUID()}`,
          extensionName: entry.name,
          source,
          registerCommand,
          registerTool,
          unregisterTool,
          subscribeAgentEvents,
          llmComplete: (request) => this.runExtensionLlmCompletion(entry, request),
          httpFetch: (url, options) => this.runExtensionHttpFetch(url, options),
          storageGet: (key) => getExtensionStorageValue(this.settings, entry.id, key),
          storageSet: (key, value) => setExtensionStorageValue(this.settings, entry.id, key, value),
          storageDelete: (key) => deleteExtensionStorageValue(this.settings, entry.id, key),
          storageKeys: () => listExtensionStorageKeys(this.settings, entry.id),
          clipboardWriteText: (text) => this.writeExtensionClipboard(text),
          injectAgentContext: (content) => {
            const agent = this.getRequiredActiveAgent();
            agent.appendMessage(createExtensionAgentMessage(entry.name, "agent.injectContext content", content));
          },
          steerAgent: (content) => {
            const agent = this.getRequiredActiveAgent();
            agent.steer(createExtensionAgentMessage(entry.name, "agent.steer content", content));
          },
          followUpAgent: (content) => {
            const agent = this.getRequiredActiveAgent();
            agent.followUp(createExtensionAgentMessage(entry.name, "agent.followUp content", content));
          },
          listSkills: () => listExtensionSkillSummaries(this.settings),
          readSkill: (name) => readExtensionSkill(this.settings, name),
          installSkill: (name, markdown) => installExternalExtensionSkill(this.settings, name, markdown),
          uninstallSkill: (name) => uninstallExternalExtensionSkill(this.settings, name),
          downloadFile: (filename, content, mimeType) => this.triggerExtensionDownload(filename, content, mimeType),
          isCapabilityEnabled,
          formatCapabilityError,
          toast: this.showToastMessage,
          widgetOwnerId: entry.id,
          widgetApiV2Enabled,
        });
      } else {
        const api = createExtensionAPI({
          getAgent: () => this.getRequiredActiveAgent(),
          registerCommand,
          registerTool,
          unregisterTool,
          subscribeAgentEvents,
          llmComplete: (request) => this.runExtensionLlmCompletion(entry, request),
          httpFetch: (url, options) => this.runExtensionHttpFetch(url, options),
          storageGet: (key) => getExtensionStorageValue(this.settings, entry.id, key),
          storageSet: (key, value) => setExtensionStorageValue(this.settings, entry.id, key, value),
          storageDelete: (key) => deleteExtensionStorageValue(this.settings, entry.id, key),
          storageKeys: () => listExtensionStorageKeys(this.settings, entry.id),
          clipboardWriteText: (text) => this.writeExtensionClipboard(text),
          injectAgentContext: (content) => {
            const agent = this.getRequiredActiveAgent();
            agent.appendMessage(createExtensionAgentMessage(entry.name, "agent.injectContext content", content));
          },
          steerAgent: (content) => {
            const agent = this.getRequiredActiveAgent();
            agent.steer(createExtensionAgentMessage(entry.name, "agent.steer content", content));
          },
          followUpAgent: (content) => {
            const agent = this.getRequiredActiveAgent();
            agent.followUp(createExtensionAgentMessage(entry.name, "agent.followUp content", content));
          },
          listSkills: () => listExtensionSkillSummaries(this.settings),
          readSkill: (name) => readExtensionSkill(this.settings, name),
          installSkill: (name, markdown) => installExternalExtensionSkill(this.settings, name, markdown),
          uninstallSkill: (name) => uninstallExternalExtensionSkill(this.settings, name),
          downloadFile: (filename, content, mimeType) => this.triggerExtensionDownload(filename, content, mimeType),
          isCapabilityEnabled,
          formatCapabilityError,
          extensionOwnerId: entry.id,
          widgetApiV2Enabled,
        });

        let loadSource: string;
        if (entry.source.kind === "inline") {
          const blob = new Blob([entry.source.code], { type: "text/javascript" });
          state.inlineBlobUrl = URL.createObjectURL(blob);
          loadSource = state.inlineBlobUrl;
        } else {
          loadSource = entry.source.specifier;
        }

        state.handle = await this.loadExtensionFromSource(api, loadSource);
      }

      activationPhase = false;
      this.activeStates.set(entry.id, state);

      if (toolsChangedDuringActivation) {
        await this.refreshRuntimeTools();
      }
    } catch (error: unknown) {
      try {
        await this.cleanupState(state);
      } catch (cleanupError: unknown) {
        console.warn(
          `[pi] Extension cleanup after failed activation also failed: ${getRuntimeManagerErrorMessage(cleanupError)}`,
        );
      }

      throw error;
    }
  }

  private async deactivateEntry(entryId: string): Promise<void> {
    const state = this.activeStates.get(entryId);
    if (!state) {
      return;
    }

    this.activeStates.delete(entryId);
    await this.cleanupState(state);
  }

  private async cleanupState(state: LoadedExtensionState): Promise<void> {
    const failures: string[] = [];

    if (state.handle) {
      try {
        await state.handle.deactivate();
      } catch (error: unknown) {
        failures.push(getRuntimeManagerErrorMessage(error));
      }
    }

    try {
      clearExtensionWidgets(state.entryId);
    } catch (error: unknown) {
      failures.push(getRuntimeManagerErrorMessage(error));
    }

    for (const unsubscribe of state.eventUnsubscribers) {
      try {
        unsubscribe();
      } catch (error: unknown) {
        failures.push(getRuntimeManagerErrorMessage(error));
      }
    }
    state.eventUnsubscribers.clear();

    for (const commandName of state.commandNames) {
      const owner = this.commandOwners.get(commandName);
      if (owner === state.entryId) {
        commandRegistry.unregister(commandName);
        this.commandOwners.delete(commandName);
      }
    }
    state.commandNames.clear();

    let toolsChanged = false;
    for (const toolName of state.toolNames) {
      const owner = this.toolOwners.get(toolName);
      if (owner !== state.entryId) {
        continue;
      }

      this.toolOwners.delete(toolName);
      const deleted = this.extensionTools.delete(toolName);
      toolsChanged = toolsChanged || deleted;
    }
    state.toolNames.clear();

    if (state.inlineBlobUrl) {
      URL.revokeObjectURL(state.inlineBlobUrl);
      state.inlineBlobUrl = null;
    }

    if (toolsChanged) {
      try {
        await this.refreshRuntimeTools();
      } catch (error: unknown) {
        failures.push(getRuntimeManagerErrorMessage(error));
      }
    }

    if (failures.length > 0) {
      throw new Error(`Extension teardown failed:\n- ${failures.join("\n- ")}`);
    }
  }
}
