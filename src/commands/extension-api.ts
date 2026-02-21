/**
 * Extension API for Pi for Excel.
 *
 * Extensions are ES modules that export an `activate(api)` function.
 * They run in the same webview sandbox — no Node.js APIs.
 *
 * Extensions can:
 * - Register slash commands + tools (including runtime unregister)
 * - Call host-mediated LLM/HTTP/storage/clipboard/skills/download APIs
 * - Steer/inject/follow-up the active agent via capability-gated bridges
 * - Show overlay/widget UIs and subscribe to agent events
 *
 * Example extension:
 * ```ts
 * export function activate(api: ExcelExtensionAPI) {
 *   api.registerCommand("snake", {
 *     description: "Play Snake!",
 *     handler: () => {
 *       // ...
 *     },
 *   });
 * }
 * ```
 */

import type { AgentEvent, AgentTool } from "@mariozechner/pi-agent-core";

import {
  ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY,
  classifyExtensionSource,
} from "./extension-source-policy.js";
import {
  importExtensionModule,
  readRemoteExtensionOptInFromStorage,
} from "./extension-module-import.js";
import {
  collectActivationCleanups,
  createLoadedExtensionHandle,
  getExtensionActivator,
  getExtensionDeactivator,
  type ExtensionActivator,
  type LoadedExtensionHandle,
} from "./extension-loader.js";
import type {
  CreateExtensionAPIOptions,
  ExcelExtensionAPI,
  ExtensionCommand,
  ExtensionConnectionDefinition,
  ExtensionToolDefinition,
  HttpRequestOptions,
  HttpResponse,
  LlmCompletionRequest,
  LlmCompletionResult,
  SkillSummary,
  WidgetUpsertSpec,
} from "./extension-api-types.js";
import { commandRegistry } from "./types.js";
import { isExperimentalFeatureEnabled } from "../experiments/flags.js";

export type { LoadedExtensionHandle } from "./extension-loader.js";
export type {
  ClipboardAPI,
  CreateExtensionAPIOptions,
  DownloadAPI,
  ExcelExtensionAPI,
  ExtensionAgentAPI,
  ExtensionCleanup,
  ExtensionCommand,
  ExtensionConnectionDefinition,
  ExtensionConnectionsAPI,
  ExtensionToolDefinition,
  HttpAPI,
  HttpRequestOptions,
  HttpResponse,
  LlmAPI,
  LlmCompletionMessage,
  LlmCompletionRequest,
  LlmCompletionResult,
  OverlayAPI,
  SkillSummary,
  SkillsAPI,
  StorageAPI,
  WidgetAPI,
  WidgetPlacement,
  WidgetUpsertSpec,
} from "./extension-api-types.js";
import type { ExtensionCapability } from "../extensions/permissions.js";
import {
  clearExtensionWidgets,
  removeExtensionWidget,
  upsertExtensionWidget,
} from "../extensions/internal/widget-surface.js";
import { EXTENSION_OVERLAY_ID } from "../ui/overlay-ids.js";
import { createOverlayDialogManager } from "../ui/overlay-dialog.js";

function normalizeIdentifier(kind: "command" | "tool", value: string): string {
  const trimmed = value.trim();
  if (trimmed.length === 0) {
    throw new Error(`Extension ${kind} name cannot be empty`);
  }
  return trimmed;
}

function normalizeConnectionIdentifier(value: string): string {
  const trimmed = value.trim().toLowerCase();
  if (trimmed.length === 0) {
    throw new Error("Connection id cannot be empty");
  }

  return trimmed;
}

function qualifyOwnedConnectionId(ownerId: string, connectionId: string): string {
  const normalizedConnectionId = normalizeConnectionIdentifier(connectionId);
  const ownerPrefix = `${ownerId.toLowerCase()}.`;

  if (normalizedConnectionId.startsWith(ownerPrefix)) {
    return normalizedConnectionId;
  }

  return `${ownerPrefix}${normalizedConnectionId}`;
}

function normalizeToolConnectionRequirements(
  rawValue: unknown,
  ownerId: string,
): string[] | undefined {
  const normalizedIds: string[] = [];

  if (typeof rawValue === "string") {
    normalizedIds.push(qualifyOwnedConnectionId(ownerId, rawValue));
  } else if (Array.isArray(rawValue)) {
    for (const item of rawValue) {
      if (typeof item !== "string") {
        throw new Error("requiresConnection entries must be strings.");
      }

      normalizedIds.push(qualifyOwnedConnectionId(ownerId, item));
    }
  } else if (rawValue !== undefined) {
    throw new Error("requiresConnection must be a string or array of strings.");
  }

  if (normalizedIds.length === 0) {
    return undefined;
  }

  return Array.from(new Set(normalizedIds));
}

function normalizeConnectionDefinitionForOwner(
  ownerId: string,
  definition: ExtensionConnectionDefinition,
): ExtensionConnectionDefinition {
  return {
    ...definition,
    id: qualifyOwnedConnectionId(ownerId, definition.id),
  };
}

function assertValidToolDefinition(name: string, tool: ExtensionToolDefinition): void {
  const execute: unknown = Reflect.get(tool, "execute");
  if (typeof execute === "function") {
    return;
  }

  const handler: unknown = Reflect.get(tool, "handler");
  if (typeof handler === "function") {
    throw new Error(
      `Extension tool "${name}" is invalid: use execute(params, signal?, onUpdate?) instead of handler.`,
    );
  }

  throw new Error(
    `Extension tool "${name}" is invalid: execute must be a function.`,
  );
}

function defaultRegisterCommand(name: string, cmd: ExtensionCommand): void {
  commandRegistry.register({
    name,
    description: cmd.description,
    source: "extension",
    execute: cmd.handler,
    busyAllowed: cmd.busyAllowed ?? true,
  });
}

function defaultToast(message: string): void {
  let toast = document.getElementById("pi-toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "pi-toast";
    toast.className = "pi-toast";
    document.body.appendChild(toast);
  }

  toast.textContent = message;
  toast.classList.add("visible");
  const toastEl = toast;
  setTimeout(() => toastEl.classList.remove("visible"), 2000);
}

function getDefaultCapabilityErrorMessage(capability: ExtensionCapability): string {
  return `Extension is not allowed to use capability "${capability}".`;
}

const LEGACY_WIDGET_ID = "__legacy__";
const WIDGET_V2_DISABLED_ERROR = "Widget API v2 is disabled. Enable /experimental on extension-widget-v2.";

function getWidgetOwnerId(options: CreateExtensionAPIOptions): string {
  if (typeof options.extensionOwnerId !== "string") {
    return "extension.unknown";
  }

  const normalized = options.extensionOwnerId.trim();
  return normalized.length > 0 ? normalized : "extension.unknown";
}

function resolveWidgetApiV2Enabled(options: CreateExtensionAPIOptions): boolean {
  if (typeof options.widgetApiV2Enabled === "boolean") {
    return options.widgetApiV2Enabled;
  }

  return isExperimentalFeatureEnabled("extension_widget_v2");
}

function showLegacyWidget(el: HTMLElement): void {
  let slot = document.getElementById("pi-widget-slot");
  if (!slot) {
    // Fallback: insert before .pi-input-area inside the sidebar
    const inputArea = document.querySelector(".pi-input-area");
    if (inputArea) {
      slot = document.createElement("div");
      slot.id = "pi-widget-slot";
      slot.className = "pi-widget-slot";
      const parent = inputArea.parentElement;
      if (!parent) {
        console.warn("[pi] No widget slot parent found");
        return;
      }
      parent.insertBefore(slot, inputArea);
    } else {
      console.warn("[pi] No widget slot or input area found");
      return;
    }
  }

  slot.replaceChildren(el);
  slot.style.display = "flex";
}

function dismissLegacyWidget(): void {
  const slot = document.getElementById("pi-widget-slot");
  if (slot) {
    slot.style.display = "none";
    slot.replaceChildren();
  }
}

function normalizeWidgetId(id: string): string {
  const normalized = id.trim();
  if (normalized.length === 0) {
    throw new Error("Widget id cannot be empty.");
  }

  return normalized;
}

/** Create the extension API for a given host context. */
export function createExtensionAPI(options: CreateExtensionAPIOptions): ExcelExtensionAPI {
  const registerCommand = options.registerCommand ?? defaultRegisterCommand;
  const registerTool = options.registerTool;
  const unregisterTool = options.unregisterTool;
  const registerConnection = options.registerConnection;
  const unregisterConnection = options.unregisterConnection;
  const listConnections = options.listConnections;
  const getConnection = options.getConnection;
  const setConnectionSecrets = options.setConnectionSecrets;
  const clearConnectionSecrets = options.clearConnectionSecrets;
  const markConnectionValidated = options.markConnectionValidated;
  const markConnectionInvalid = options.markConnectionInvalid;
  const markConnectionStatus = options.markConnectionStatus;
  const subscribeAgentEvents = options.subscribeAgentEvents
    ?? ((handler: (ev: AgentEvent) => void) => options.getAgent().subscribe(handler));
  const llmComplete = options.llmComplete;
  const httpFetch = options.httpFetch;
  const storageGet = options.storageGet;
  const storageSet = options.storageSet;
  const storageDelete = options.storageDelete;
  const storageKeys = options.storageKeys;
  const clipboardWriteText = options.clipboardWriteText;
  const injectAgentContext = options.injectAgentContext;
  const steerAgent = options.steerAgent;
  const followUpAgent = options.followUpAgent;
  const listSkills = options.listSkills;
  const readSkill = options.readSkill;
  const installSkill = options.installSkill;
  const uninstallSkill = options.uninstallSkill;
  const downloadFile = options.downloadFile;
  const toast = options.toast ?? defaultToast;
  const isCapabilityEnabled = options.isCapabilityEnabled;
  const formatCapabilityError = options.formatCapabilityError ?? getDefaultCapabilityErrorMessage;
  const widgetOwnerId = getWidgetOwnerId(options);
  const connectionOwnerId = widgetOwnerId;
  const widgetApiV2Enabled = resolveWidgetApiV2Enabled(options);

  const assertCapability = (capability: ExtensionCapability): void => {
    if (!isCapabilityEnabled) {
      return;
    }

    if (!isCapabilityEnabled(capability)) {
      throw new Error(formatCapabilityError(capability));
    }
  };

  const overlayDialogManager = createOverlayDialogManager({
    overlayId: EXTENSION_OVERLAY_ID,
    cardClassName: "",
    zIndex: 250,
  });

  return {
    registerCommand(name: string, cmd: ExtensionCommand) {
      assertCapability("commands.register");
      registerCommand(normalizeIdentifier("command", name), cmd);
    },

    registerTool(name: string, tool: ExtensionToolDefinition) {
      assertCapability("tools.register");

      if (!registerTool) {
        throw new Error("Extension host does not support registerTool()");
      }

      const normalizedName = normalizeIdentifier("tool", name);
      assertValidToolDefinition(normalizedName, tool);

      const wrappedTool: AgentTool = {
        name: normalizedName,
        label: tool.label ?? normalizedName,
        description: tool.description,
        parameters: tool.parameters,
        execute: async (_toolCallId, params, signal, onUpdate) => {
          return tool.execute(params, signal, onUpdate);
        },
      };

      const requiresConnection = normalizeToolConnectionRequirements(
        Reflect.get(tool, "requiresConnection"),
        connectionOwnerId,
      );

      if (requiresConnection && requiresConnection.length > 0) {
        Reflect.set(wrappedTool, "requiresConnection", requiresConnection);
      }

      registerTool(wrappedTool);
    },

    unregisterTool(name: string) {
      assertCapability("tools.register");

      if (!unregisterTool) {
        throw new Error("Extension host does not support unregisterTool()");
      }

      unregisterTool(normalizeIdentifier("tool", name));
    },

    connections: {
      register(definition: ExtensionConnectionDefinition): string {
        assertCapability("connections.readwrite");

        if (!registerConnection) {
          throw new Error("Extension host does not support connections.register()");
        }

        const normalizedDefinition = normalizeConnectionDefinitionForOwner(connectionOwnerId, definition);
        return registerConnection(normalizedDefinition);
      },

      unregister(connectionId: string): void {
        assertCapability("connections.readwrite");

        if (!unregisterConnection) {
          throw new Error("Extension host does not support connections.unregister()");
        }

        unregisterConnection(qualifyOwnedConnectionId(connectionOwnerId, connectionId));
      },

      async list() {
        assertCapability("connections.readwrite");

        if (!listConnections) {
          throw new Error("Extension host does not support connections.list()");
        }

        return listConnections();
      },

      async get(connectionId: string) {
        assertCapability("connections.readwrite");

        if (!getConnection) {
          throw new Error("Extension host does not support connections.get()");
        }

        return getConnection(qualifyOwnedConnectionId(connectionOwnerId, connectionId));
      },

      async setSecrets(connectionId: string, secrets: Record<string, string>): Promise<void> {
        assertCapability("connections.readwrite");

        if (!setConnectionSecrets) {
          throw new Error("Extension host does not support connections.setSecrets()");
        }

        await setConnectionSecrets(
          qualifyOwnedConnectionId(connectionOwnerId, connectionId),
          secrets,
        );
      },

      async clearSecrets(connectionId: string): Promise<void> {
        assertCapability("connections.readwrite");

        if (!clearConnectionSecrets) {
          throw new Error("Extension host does not support connections.clearSecrets()");
        }

        await clearConnectionSecrets(qualifyOwnedConnectionId(connectionOwnerId, connectionId));
      },

      async markValidated(connectionId: string): Promise<void> {
        assertCapability("connections.readwrite");

        if (!markConnectionValidated) {
          throw new Error("Extension host does not support connections.markValidated()");
        }

        await markConnectionValidated(qualifyOwnedConnectionId(connectionOwnerId, connectionId));
      },

      async markInvalid(connectionId: string, reason: string): Promise<void> {
        assertCapability("connections.readwrite");

        if (!markConnectionInvalid) {
          throw new Error("Extension host does not support connections.markInvalid()");
        }

        await markConnectionInvalid(
          qualifyOwnedConnectionId(connectionOwnerId, connectionId),
          reason,
        );
      },

      async markStatus(connectionId: string, status: "connected" | "missing" | "invalid" | "error", reason?: string): Promise<void> {
        assertCapability("connections.readwrite");

        const normalizedConnectionId = qualifyOwnedConnectionId(connectionOwnerId, connectionId);

        if (markConnectionStatus) {
          await markConnectionStatus(normalizedConnectionId, status, reason);
          return;
        }

        if (status === "connected") {
          if (!markConnectionValidated) {
            throw new Error("Extension host does not support setting connection status to connected.");
          }
          await markConnectionValidated(normalizedConnectionId);
          return;
        }

        if (status === "invalid") {
          if (!markConnectionInvalid) {
            throw new Error("Extension host does not support setting connection status to invalid.");
          }
          await markConnectionInvalid(normalizedConnectionId, reason ?? "Connection marked invalid.");
          return;
        }

        if (status === "missing") {
          if (!clearConnectionSecrets) {
            throw new Error("Extension host does not support setting connection status to missing.");
          }
          await clearConnectionSecrets(normalizedConnectionId);
          return;
        }

        throw new Error("Extension host does not support setting connection status to error.");
      },
    },

    agent: {
      get raw() {
        assertCapability("agent.read");
        assertCapability("agent.events.read");
        return options.getAgent();
      },

      injectContext(content: string) {
        assertCapability("agent.context.write");

        if (!injectAgentContext) {
          throw new Error("Extension host does not support agent.injectContext()");
        }

        injectAgentContext(content);
      },

      steer(content: string) {
        assertCapability("agent.steer");

        if (!steerAgent) {
          throw new Error("Extension host does not support agent.steer()");
        }

        steerAgent(content);
      },

      followUp(content: string) {
        assertCapability("agent.followup");

        if (!followUpAgent) {
          throw new Error("Extension host does not support agent.followUp()");
        }

        followUpAgent(content);
      },
    },

    llm: {
      async complete(request: LlmCompletionRequest): Promise<LlmCompletionResult> {
        assertCapability("llm.complete");

        if (!llmComplete) {
          throw new Error("Extension host does not support llm.complete()");
        }

        return llmComplete(request);
      },
    },

    http: {
      async fetch(url: string, requestOptions?: HttpRequestOptions): Promise<HttpResponse> {
        assertCapability("http.fetch");

        if (!httpFetch) {
          throw new Error("Extension host does not support http.fetch()");
        }

        return httpFetch(url, requestOptions);
      },
    },

    storage: {
      async get(key: string): Promise<unknown> {
        assertCapability("storage.readwrite");
        if (!storageGet) {
          throw new Error("Extension host does not support storage.get()");
        }

        return storageGet(key);
      },

      async set(key: string, value: unknown): Promise<void> {
        assertCapability("storage.readwrite");
        if (!storageSet) {
          throw new Error("Extension host does not support storage.set()");
        }

        await storageSet(key, value);
      },

      async delete(key: string): Promise<void> {
        assertCapability("storage.readwrite");
        if (!storageDelete) {
          throw new Error("Extension host does not support storage.delete()");
        }

        await storageDelete(key);
      },

      async keys(): Promise<string[]> {
        assertCapability("storage.readwrite");
        if (!storageKeys) {
          throw new Error("Extension host does not support storage.keys()");
        }

        return storageKeys();
      },
    },

    clipboard: {
      async writeText(text: string): Promise<void> {
        assertCapability("clipboard.write");
        if (!clipboardWriteText) {
          throw new Error("Extension host does not support clipboard.writeText()");
        }

        await clipboardWriteText(text);
      },
    },

    skills: {
      async list(): Promise<SkillSummary[]> {
        assertCapability("skills.read");
        if (!listSkills) {
          throw new Error("Extension host does not support skills.list()");
        }

        return listSkills();
      },

      async read(name: string): Promise<string> {
        assertCapability("skills.read");
        if (!readSkill) {
          throw new Error("Extension host does not support skills.read()");
        }

        return readSkill(name);
      },

      async install(name: string, markdown: string): Promise<void> {
        assertCapability("skills.write");
        if (!installSkill) {
          throw new Error("Extension host does not support skills.install()");
        }

        await installSkill(name, markdown);
      },

      async uninstall(name: string): Promise<void> {
        assertCapability("skills.write");
        if (!uninstallSkill) {
          throw new Error("Extension host does not support skills.uninstall()");
        }

        await uninstallSkill(name);
      },
    },

    download: {
      download(filename: string, content: string, mimeType?: string): void {
        assertCapability("download.file");
        if (!downloadFile) {
          throw new Error("Extension host does not support download.download()");
        }

        downloadFile(filename, content, mimeType);
      },
    },

    overlay: {
      show(el: HTMLElement) {
        assertCapability("ui.overlay");

        const dialog = overlayDialogManager.ensure();
        dialog.card.replaceChildren(el);

        if (!dialog.overlay.isConnected) {
          dialog.mount();
        }
      },

      dismiss() {
        overlayDialogManager.dismiss();
      },
    },

    widget: {
      show(el: HTMLElement) {
        assertCapability("ui.widget");

        if (!widgetApiV2Enabled) {
          showLegacyWidget(el);
          return;
        }

        upsertExtensionWidget({
          ownerId: widgetOwnerId,
          id: LEGACY_WIDGET_ID,
          element: el,
          placement: "above-input",
          order: 0,
        });
      },

      dismiss() {
        if (!widgetApiV2Enabled) {
          dismissLegacyWidget();
          return;
        }

        removeExtensionWidget(widgetOwnerId, LEGACY_WIDGET_ID);
      },

      upsert(spec: WidgetUpsertSpec) {
        assertCapability("ui.widget");
        if (!widgetApiV2Enabled) {
          throw new Error(WIDGET_V2_DISABLED_ERROR);
        }

        const widgetId = normalizeWidgetId(spec.id);
        upsertExtensionWidget({
          ownerId: widgetOwnerId,
          id: widgetId,
          element: spec.el,
          title: spec.title,
          placement: spec.placement,
          order: spec.order,
          collapsible: spec.collapsible,
          collapsed: spec.collapsed,
          minHeightPx: spec.minHeightPx,
          maxHeightPx: spec.maxHeightPx,
        });
      },

      remove(id: string) {
        assertCapability("ui.widget");
        if (!widgetApiV2Enabled) {
          throw new Error(WIDGET_V2_DISABLED_ERROR);
        }

        const widgetId = normalizeWidgetId(id);
        removeExtensionWidget(widgetOwnerId, widgetId);
      },

      clear() {
        assertCapability("ui.widget");
        if (!widgetApiV2Enabled) {
          throw new Error(WIDGET_V2_DISABLED_ERROR);
        }

        clearExtensionWidgets(widgetOwnerId);
      },
    },

    toast(message: string) {
      assertCapability("ui.toast");
      toast(message);
    },

    onAgentEvent(handler: (ev: AgentEvent) => void) {
      assertCapability("agent.events.read");
      return subscribeAgentEvents(handler);
    },
  };
}

/**
 * Load and activate an extension from an inline function or module specifier.
 *
 * Security default: remote http(s) module URLs are blocked unless the user
 * explicitly opts in by setting localStorage key
 * `pi.allowRemoteExtensionUrls` to `1` or `true`.
 */
export async function loadExtension(
  api: ExcelExtensionAPI,
  source: string | ExtensionActivator<ExcelExtensionAPI>,
): Promise<LoadedExtensionHandle> {
  if (typeof source === "function") {
    const cleanups = collectActivationCleanups(await source(api));
    return createLoadedExtensionHandle(cleanups, null);
  }

  const specifier = source.trim();
  const sourceKind = classifyExtensionSource(specifier);

  if (sourceKind === "unsupported") {
    throw new Error(
      `Unsupported extension source "${specifier}". Only local module specifiers (./, ../, /), blob: URLs, and inline function activators are allowed by default.`,
    );
  }

  if (sourceKind === "remote-url" && !readRemoteExtensionOptInFromStorage(globalThis.localStorage)) {
    throw new Error(
      "Remote extension URL imports are disabled by default. "
      + `Set localStorage['${ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY}']='1' to opt in (unsafe).`,
    );
  }

  if (sourceKind === "remote-url") {
    console.warn(`[pi] WARNING: loading remote extension URL due to explicit opt-in: ${specifier}`);
  }

  const importedModule = await importExtensionModule(specifier, sourceKind);
  const activate = getExtensionActivator(importedModule);
  if (!activate) {
    throw new Error(`Extension module "${specifier}" must export an activate(api) function`);
  }

  const cleanups = collectActivationCleanups(await activate(api));
  const deactivate = getExtensionDeactivator(importedModule);
  return createLoadedExtensionHandle(cleanups, deactivate);
}
