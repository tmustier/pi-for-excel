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
  type LoadedExtensionHandle,
} from "../commands/extension-api.js";
import { commandRegistry } from "../commands/types.js";
import {
  loadStoredExtensions,
  saveStoredExtensions,
  type ExtensionSettingsStore,
  type StoredExtensionEntry,
  type StoredExtensionSource,
} from "./store.js";

type AnyAgentTool = AgentTool;

type ManagerListener = () => void;

interface LoadedExtensionState {
  entryId: string;
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
  commandNames: string[];
  toolNames: string[];
  lastError: string | null;
}

export interface ExtensionRuntimeManagerOptions {
  settings: ExtensionSettingsStore;
  getActiveAgent: () => Agent | null;
  refreshRuntimeTools: () => Promise<void>;
  reservedToolNames: ReadonlySet<string>;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  return String(error);
}

function normalizeExtensionName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Extension name cannot be empty");
  }
  return trimmed;
}

function normalizeInlineCode(code: string): string {
  if (code.trim().length === 0) {
    throw new Error("Extension code cannot be empty");
  }
  return code;
}

function normalizeRemoteUrl(url: string): string {
  let parsed: URL;
  try {
    parsed = new URL(url.trim());
  } catch {
    throw new Error("Invalid URL");
  }

  if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
    throw new Error("Extension URL must use http:// or https://");
  }

  return parsed.toString();
}

function describeExtensionSource(source: StoredExtensionSource): string {
  if (source.kind === "module") {
    return source.specifier;
  }

  const lines = source.code.split("\n").length;
  return `inline code (${source.code.length} chars, ${lines} lines)`;
}

export class ExtensionRuntimeManager {
  private readonly settings: ExtensionSettingsStore;
  private readonly getActiveAgent: () => Agent | null;
  private readonly refreshRuntimeTools: () => Promise<void>;
  private readonly reservedToolNames: ReadonlySet<string>;

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
  }

  subscribe(listener: ManagerListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  list(): ExtensionRuntimeStatus[] {
    return this.entries.map((entry) => {
      const state = this.activeStates.get(entry.id);
      return {
        id: entry.id,
        name: entry.name,
        enabled: entry.enabled,
        loaded: Boolean(state),
        source: entry.source,
        sourceLabel: describeExtensionSource(entry.source),
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

  async uninstallExtension(entryId: string): Promise<void> {
    const entryIndex = this.entries.findIndex((entry) => entry.id === entryId);
    if (entryIndex < 0) {
      throw new Error("Extension not found");
    }

    await this.deactivateEntry(entryId);
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
        console.warn("[pi] Extension manager listener failed:", getErrorMessage(error));
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

  private async installEntry(input: {
    name: string;
    source: StoredExtensionSource;
  }): Promise<string> {
    const now = new Date().toISOString();
    const id = `ext.${crypto.randomUUID()}`;

    const entry: StoredExtensionEntry = {
      id,
      name: input.name,
      enabled: true,
      source: input.source,
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
      const message = getErrorMessage(error);
      this.lastErrors.set(entry.id, message);
      console.warn(`[pi] Failed to load extension "${entry.name}": ${message}`);
    }
  }

  private async activateEntry(entry: StoredExtensionEntry): Promise<void> {
    await this.deactivateEntry(entry.id);

    const state: LoadedExtensionState = {
      entryId: entry.id,
      commandNames: new Set<string>(),
      toolNames: new Set<string>(),
      eventUnsubscribers: new Set<() => void>(),
      handle: null,
      inlineBlobUrl: null,
    };

    let toolsChanged = false;

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
      toolsChanged = true;
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

    const api = createExtensionAPI({
      getAgent: () => this.getRequiredActiveAgent(),
      registerCommand,
      registerTool,
      subscribeAgentEvents,
    });

    let loadSource: string;
    if (entry.source.kind === "inline") {
      const blob = new Blob([entry.source.code], { type: "text/javascript" });
      state.inlineBlobUrl = URL.createObjectURL(blob);
      loadSource = state.inlineBlobUrl;
    } else {
      loadSource = entry.source.specifier;
    }

    try {
      state.handle = await loadExtension(api, loadSource);
      this.activeStates.set(entry.id, state);

      if (toolsChanged) {
        await this.refreshRuntimeTools();
      }
    } catch (error: unknown) {
      try {
        await this.cleanupState(state);
      } catch (cleanupError: unknown) {
        console.warn(
          `[pi] Extension cleanup after failed activation also failed: ${getErrorMessage(cleanupError)}`,
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
        failures.push(getErrorMessage(error));
      }
    }

    for (const unsubscribe of state.eventUnsubscribers) {
      try {
        unsubscribe();
      } catch (error: unknown) {
        failures.push(getErrorMessage(error));
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
        failures.push(getErrorMessage(error));
      }
    }

    if (failures.length > 0) {
      throw new Error(`Extension teardown failed:\n- ${failures.join("\n- ")}`);
    }
  }
}
