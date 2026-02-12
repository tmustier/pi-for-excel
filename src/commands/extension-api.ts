/**
 * Extension API for Pi for Excel.
 *
 * Extensions are ES modules that export an `activate(api)` function.
 * They run in the same webview sandbox — no Node.js APIs.
 *
 * Extensions can:
 * - Register slash commands
 * - Register custom tools
 * - Show overlay UIs (via the overlay API)
 * - Subscribe to agent events
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

import type {
  Agent,
  AgentEvent,
  AgentTool,
  AgentToolResult,
  AgentToolUpdateCallback,
} from "@mariozechner/pi-agent-core";
import type { Static, TSchema } from "@sinclair/typebox";

import {
  ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY,
  classifyExtensionSource,
  isRemoteExtensionOptIn,
} from "./extension-source-policy.js";
import { commandRegistry } from "./types.js";
import type { ExtensionCapability } from "../extensions/permissions.js";
import { isRecord } from "../utils/type-guards.js";

export interface ExtensionCommand {
  description: string;
  handler: (args: string) => void | Promise<void>;
}

export type ExtensionCleanup = () => void | Promise<void>;

export interface ExtensionToolDefinition<TParameters extends TSchema = TSchema, TDetails = unknown> {
  description: string;
  parameters: TParameters;
  label?: string;
  execute: (
    params: Static<TParameters>,
    signal?: AbortSignal,
    onUpdate?: AgentToolUpdateCallback<TDetails>,
  ) => Promise<AgentToolResult<TDetails>> | AgentToolResult<TDetails>;
}

export interface OverlayAPI {
  /** Show an HTML element as a full-screen overlay */
  show(el: HTMLElement): void;
  /** Remove the overlay */
  dismiss(): void;
}

export interface WidgetAPI {
  /** Show an HTML element as an inline widget above the input area */
  show(el: HTMLElement): void;
  /** Remove the widget */
  dismiss(): void;
}

export interface ExcelExtensionAPI {
  /** Register a slash command */
  registerCommand(name: string, cmd: ExtensionCommand): void;
  /** Register a custom tool callable by the agent */
  registerTool(name: string, tool: ExtensionToolDefinition): void;
  /** Access the active agent */
  readonly agent: Agent;
  /** Show/dismiss full-screen overlay UI */
  overlay: OverlayAPI;
  /** Show/dismiss inline widget above input (messages still visible above) */
  widget: WidgetAPI;
  /** Show a toast notification */
  toast(message: string): void;
  /** Subscribe to agent events */
  onAgentEvent(handler: (ev: AgentEvent) => void): () => void;
}

export interface CreateExtensionAPIOptions {
  getAgent: () => Agent;
  registerCommand?: (name: string, cmd: ExtensionCommand) => void;
  registerTool?: (tool: AgentTool) => void;
  subscribeAgentEvents?: (handler: (ev: AgentEvent) => void) => () => void;
  toast?: (message: string) => void;
  isCapabilityEnabled?: (capability: ExtensionCapability) => boolean;
  formatCapabilityError?: (capability: ExtensionCapability) => string;
}

function normalizeIdentifier(kind: "command" | "tool", value: string): string {
  const trimmed = value.trim();
  if (trimmed.length === 0) {
    throw new Error(`Extension ${kind} name cannot be empty`);
  }
  return trimmed;
}

function defaultRegisterCommand(name: string, cmd: ExtensionCommand): void {
  commandRegistry.register({
    name,
    description: cmd.description,
    source: "extension",
    execute: cmd.handler,
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

const BUNDLED_LOCAL_EXTENSION_IMPORTERS = import.meta.glob("../extensions/*.{ts,js}");

function getLocalExtensionImportCandidates(specifier: string): string[] {
  const normalized = specifier.trim();
  const candidates = new Set<string>([normalized]);

  if (normalized.endsWith(".js")) {
    candidates.add(`${normalized.slice(0, -3)}.ts`);
  } else if (normalized.endsWith(".ts")) {
    candidates.add(`${normalized.slice(0, -3)}.js`);
  } else {
    candidates.add(`${normalized}.ts`);
    candidates.add(`${normalized}.js`);
  }

  return Array.from(candidates);
}

async function importExtensionModule(
  specifier: string,
  sourceKind: ReturnType<typeof classifyExtensionSource>,
): Promise<unknown> {
  if (sourceKind === "local-module") {
    for (const candidate of getLocalExtensionImportCandidates(specifier)) {
      const importer = BUNDLED_LOCAL_EXTENSION_IMPORTERS[candidate];
      if (!importer) {
        continue;
      }

      return importer();
    }

    if (import.meta.env.DEV) {
      return import(/* @vite-ignore */ specifier);
    }

    throw new Error(
      `Local extension module "${specifier}" was not bundled. `
      + "Use a bundled module under src/extensions, paste code, or a remote URL (with explicit opt-in).",
    );
  }

  return import(/* @vite-ignore */ specifier);
}

function getDefaultCapabilityErrorMessage(capability: ExtensionCapability): string {
  return `Extension is not allowed to use capability "${capability}".`;
}

/** Create the extension API for a given host context. */
export function createExtensionAPI(options: CreateExtensionAPIOptions): ExcelExtensionAPI {
  const registerCommand = options.registerCommand ?? defaultRegisterCommand;
  const registerTool = options.registerTool;
  const subscribeAgentEvents = options.subscribeAgentEvents
    ?? ((handler: (ev: AgentEvent) => void) => options.getAgent().subscribe(handler));
  const toast = options.toast ?? defaultToast;
  const isCapabilityEnabled = options.isCapabilityEnabled;
  const formatCapabilityError = options.formatCapabilityError ?? getDefaultCapabilityErrorMessage;

  const assertCapability = (capability: ExtensionCapability): void => {
    if (!isCapabilityEnabled) {
      return;
    }

    if (!isCapabilityEnabled(capability)) {
      throw new Error(formatCapabilityError(capability));
    }
  };

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
      const wrappedTool: AgentTool = {
        name: normalizedName,
        label: tool.label ?? normalizedName,
        description: tool.description,
        parameters: tool.parameters,
        execute: async (_toolCallId, params, signal, onUpdate) => {
          return tool.execute(params, signal, onUpdate);
        },
      };

      registerTool(wrappedTool);
    },

    get agent() {
      assertCapability("agent.read");
      // Raw Agent includes event subscription; require both until a narrowed agent facade exists.
      assertCapability("agent.events.read");
      return options.getAgent();
    },

    overlay: {
      show(el: HTMLElement) {
        assertCapability("ui.overlay");

        let container = document.getElementById("pi-ext-overlay");
        if (!container) {
          container = document.createElement("div");
          container.id = "pi-ext-overlay";
          container.className = "pi-welcome-overlay";
          container.style.zIndex = "250";
          document.body.appendChild(container);
        }

        container.replaceChildren(el);
        container.style.display = "flex";

        // ESC to dismiss
        const handler = (e: KeyboardEvent) => {
          if (e.key === "Escape") {
            this.dismiss();
            document.removeEventListener("keydown", handler);
          }
        };
        document.addEventListener("keydown", handler);
      },

      dismiss() {
        const container = document.getElementById("pi-ext-overlay");
        if (container) {
          container.style.display = "none";
          container.replaceChildren();
        }
      },
    },

    widget: {
      show(el: HTMLElement) {
        assertCapability("ui.widget");

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
        slot.style.display = "block";
      },

      dismiss() {
        const slot = document.getElementById("pi-widget-slot");
        if (slot) {
          slot.style.display = "none";
          slot.replaceChildren();
        }
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

type ExtensionActivateResult = void | ExtensionCleanup | readonly ExtensionCleanup[];
type ExtensionActivator = (api: ExcelExtensionAPI) => ExtensionActivateResult | Promise<ExtensionActivateResult>;
type ExtensionDeactivator = () => void | Promise<void>;

export interface LoadedExtensionHandle {
  deactivate: () => Promise<void>;
}

function isExtensionActivator(value: unknown): value is ExtensionActivator {
  return typeof value === "function";
}

function isExtensionDeactivator(value: unknown): value is ExtensionDeactivator {
  return typeof value === "function";
}

function isExtensionCleanup(value: unknown): value is ExtensionCleanup {
  return typeof value === "function";
}

function getRemoteExtensionOptInFromStorage(): boolean {
  if (typeof localStorage === "undefined") return false;

  try {
    const raw = localStorage.getItem(ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY);
    return isRemoteExtensionOptIn(raw);
  } catch {
    return false;
  }
}

function getExtensionActivator(mod: unknown): ExtensionActivator | null {
  if (!isRecord(mod)) return null;

  const activate = mod.activate;
  if (isExtensionActivator(activate)) {
    return activate;
  }

  const fallback = mod.default;
  if (isExtensionActivator(fallback)) {
    return fallback;
  }

  return null;
}

function getExtensionDeactivator(mod: unknown): ExtensionDeactivator | null {
  if (!isRecord(mod)) return null;

  const deactivate = mod.deactivate;
  return isExtensionDeactivator(deactivate) ? deactivate : null;
}

function collectActivationCleanups(result: unknown): ExtensionCleanup[] {
  if (typeof result === "undefined") {
    return [];
  }

  if (isExtensionCleanup(result)) {
    return [result];
  }

  if (!Array.isArray(result)) {
    throw new Error("activate(api) must return void, a cleanup function, or an array of cleanup functions");
  }

  const cleanups: ExtensionCleanup[] = [];
  for (const value of result) {
    const item: unknown = value;
    if (!isExtensionCleanup(item)) {
      throw new Error("activate(api) returned an invalid cleanup entry; expected a function");
    }
    cleanups.push(item);
  }

  return cleanups;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  return String(error);
}

function createLoadedExtensionHandle(
  cleanups: readonly ExtensionCleanup[],
  moduleDeactivate: ExtensionDeactivator | null,
): LoadedExtensionHandle {
  let deactivated = false;

  return {
    deactivate: async () => {
      if (deactivated) {
        return;
      }
      deactivated = true;

      const failures: string[] = [];

      for (let i = cleanups.length - 1; i >= 0; i -= 1) {
        const cleanup = cleanups[i];
        try {
          await cleanup();
        } catch (error: unknown) {
          failures.push(getErrorMessage(error));
        }
      }

      if (moduleDeactivate) {
        try {
          await moduleDeactivate();
        } catch (error: unknown) {
          failures.push(getErrorMessage(error));
        }
      }

      if (failures.length > 0) {
        throw new Error(`Extension cleanup failed:\n- ${failures.join("\n- ")}`);
      }
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
  source: string | ((api: ExcelExtensionAPI) => ExtensionActivateResult | Promise<ExtensionActivateResult>),
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

  if (sourceKind === "remote-url" && !getRemoteExtensionOptInFromStorage()) {
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
