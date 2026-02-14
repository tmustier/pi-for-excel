/**
 * Sandboxed extension runtime host (iframe + postMessage RPC).
 *
 * Scope: default runtime for untrusted extension execution (with rollback switch).
 * This bridge intentionally starts small and only forwards a constrained,
 * sanitized UI projection format (no raw HTML injection).
 */

import type {
  AgentEvent,
  AgentTool,
  AgentToolResult,
} from "@mariozechner/pi-agent-core";
import { Kind, Type, type TSchema } from "@sinclair/typebox";

import type {
  ExtensionCommand,
  HttpRequestOptions,
  HttpResponse,
  LlmCompletionRequest,
  LlmCompletionResult,
  LoadedExtensionHandle,
  SkillSummary,
  WidgetPlacement,
} from "../commands/extension-api.js";
import type { ExtensionCapability } from "./permissions.js";
import {
  clearExtensionWidgets,
  removeExtensionWidget,
} from "./internal/widget-surface.js";
import { normalizeSandboxUiNode } from "./sandbox-ui.js";
import {
  SANDBOX_CHANNEL,
  SANDBOX_REQUEST_TIMEOUT_MS,
  isSandboxEnvelope,
  serializeForSandboxInlineScript,
  type SandboxRequestEnvelope,
  type SandboxResponseEnvelope,
  type SandboxEventEnvelope,
} from "./sandbox/protocol.js";
import {
  createTextOnlyUiNode,
  dismissOverlay,
  dismissWidget,
  showOverlayNode,
  showWidgetNode,
  upsertSandboxWidgetNode,
} from "./sandbox/surfaces.js";
import { isRecord } from "../utils/type-guards.js";

const LEGACY_WIDGET_ID = "__legacy__";

interface SandboxInlineSource {
  kind: "inline";
  code: string;
}

interface SandboxModuleSource {
  kind: "module";
  specifier: string;
}

export type SandboxExtensionSource = SandboxInlineSource | SandboxModuleSource;

export interface SandboxActivationOptions {
  instanceId: string;
  extensionName: string;
  source: SandboxExtensionSource;
  registerCommand: (name: string, cmd: ExtensionCommand) => void;
  registerTool: (tool: AgentTool<TSchema, unknown>) => void;
  unregisterTool: (name: string) => void;
  subscribeAgentEvents: (handler: (event: AgentEvent) => void) => () => void;
  llmComplete: (request: LlmCompletionRequest) => Promise<LlmCompletionResult>;
  httpFetch: (url: string, options?: HttpRequestOptions) => Promise<HttpResponse>;
  storageGet: (key: string) => Promise<unknown>;
  storageSet: (key: string, value: unknown) => Promise<void>;
  storageDelete: (key: string) => Promise<void>;
  storageKeys: () => Promise<string[]>;
  clipboardWriteText: (text: string) => Promise<void>;
  injectAgentContext: (content: string) => void;
  steerAgent: (content: string) => void;
  followUpAgent: (content: string) => void;
  listSkills: () => Promise<SkillSummary[]>;
  readSkill: (name: string) => Promise<string>;
  installSkill: (name: string, markdown: string) => Promise<void>;
  uninstallSkill: (name: string) => Promise<void>;
  downloadFile: (filename: string, content: string, mimeType?: string) => void;
  isCapabilityEnabled: (capability: ExtensionCapability) => boolean;
  formatCapabilityError: (capability: ExtensionCapability) => string;
  toast: (message: string) => void;
  widgetOwnerId?: string;
  widgetApiV2Enabled?: boolean;
}

interface SandboxPendingRequest {
  resolve: (value: unknown) => void;
  reject: (reason: unknown) => void;
  timeoutId: ReturnType<typeof setTimeout>;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }

  return String(error);
}

function sanitizeText(value: unknown): string {
  if (typeof value !== "string") {
    return "";
  }

  return value;
}

function asNonEmptyString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`${field} must be a non-empty string.`);
  }

  return value.trim();
}

function asRecord(value: unknown, field: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`${field} must be an object.`);
  }

  return value;
}

function asFiniteNumberOrNull(value: unknown): number | null {
  if (typeof value !== "number" || Number.isNaN(value) || !Number.isFinite(value)) {
    return null;
  }

  return value;
}

function asFiniteNumberOrNullOrUndefined(value: unknown): number | null | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (value === null) {
    return null;
  }

  if (typeof value !== "number" || Number.isNaN(value) || !Number.isFinite(value)) {
    return undefined;
  }

  return value;
}

function asWidgetPlacementOrUndefined(value: unknown): WidgetPlacement | undefined {
  if (value === "above-input" || value === "below-input") {
    return value;
  }

  return undefined;
}

function asBooleanOrUndefined(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

function isTypeBoxSchema(value: unknown): value is TSchema {
  return isRecord(value) && Kind in value;
}

export function normalizeSandboxToolParameters(raw: unknown): TSchema {
  if (isTypeBoxSchema(raw)) {
    return raw;
  }

  if (!isRecord(raw)) {
    throw new Error("register_tool parameters must be an object schema.");
  }

  return Type.Unsafe<unknown>(raw);
}

function normalizeToolResult(raw: unknown): AgentToolResult<unknown> {
  const content: Array<{ type: "text"; text: string }> = [];

  if (isRecord(raw) && Array.isArray(raw.content)) {
    for (const item of raw.content) {
      if (!isRecord(item)) {
        continue;
      }

      if (item.type !== "text") {
        continue;
      }

      if (typeof item.text !== "string") {
        continue;
      }

      content.push({
        type: "text",
        text: item.text,
      });
    }
  }

  if (content.length === 0) {
    const fallbackText = isRecord(raw) && Array.isArray(raw.content)
      ? "Sandbox tool returned non-text content; showing serialized payload instead."
      : "Sandbox tool returned an invalid payload; showing serialized payload instead.";

    content.push({
      type: "text",
      text: `${fallbackText}\n\n\`\`\`json\n${JSON.stringify(raw, null, 2)}\n\`\`\``,
    });
  }

  const details = isRecord(raw) && Object.prototype.hasOwnProperty.call(raw, "details")
    ? raw.details
    : undefined;

  return {
    content,
    details,
  };
}

class SandboxRuntimeHost {
  private readonly options: SandboxActivationOptions;
  private readonly iframe: HTMLIFrameElement;
  private readonly pendingRequests = new Map<string, SandboxPendingRequest>();
  private readonly eventSubscriptions = new Map<string, () => void>();
  private readonly overlayActionIds = new Set<string>();
  private readonly widgetActionIdsByWidgetId = new Map<string, Set<string>>();

  private readonly onWindowMessage = (event: MessageEvent<unknown>) => {
    this.handleWindowMessage(event);
  };

  private readonly readyPromise: Promise<void>;
  private resolveReady: (() => void) | null = null;
  private rejectReady: ((reason: unknown) => void) | null = null;

  private nextRequestId = 1;
  private disposed = false;
  private readonly widgetOwnerId: string;
  private readonly widgetApiV2Enabled: boolean;

  constructor(options: SandboxActivationOptions) {
    this.options = options;
    this.widgetOwnerId = typeof options.widgetOwnerId === "string" && options.widgetOwnerId.trim().length > 0
      ? options.widgetOwnerId.trim()
      : options.instanceId;
    this.widgetApiV2Enabled = options.widgetApiV2Enabled === true;

    this.readyPromise = new Promise<void>((resolve, reject) => {
      this.resolveReady = resolve;
      this.rejectReady = reject;
    });

    this.iframe = document.createElement("iframe");
    this.iframe.setAttribute("sandbox", "allow-scripts");
    this.iframe.setAttribute("aria-hidden", "true");
    this.iframe.tabIndex = -1;
    this.iframe.style.display = "none";
    this.iframe.srcdoc = this.buildSandboxSrcdoc();

    window.addEventListener("message", this.onWindowMessage);
    document.body.appendChild(this.iframe);
  }

  async waitUntilReady(): Promise<void> {
    const timeout = setTimeout(() => {
      if (this.resolveReady || this.rejectReady) {
        this.rejectReady?.(new Error("Sandbox extension bootstrap timed out."));
        this.resolveReady = null;
        this.rejectReady = null;
      }
    }, SANDBOX_REQUEST_TIMEOUT_MS);

    try {
      await this.readyPromise;
    } finally {
      clearTimeout(timeout);
    }
  }

  async dispose(gracefulDeactivate: boolean): Promise<void> {
    if (this.disposed) {
      return;
    }
    this.disposed = true;

    if (gracefulDeactivate) {
      try {
        await this.callSandbox("deactivate", {}, { allowWhenDisposed: true });
      } catch (error: unknown) {
        console.warn(`[pi] Sandbox deactivate failed: ${getErrorMessage(error)}`);
      }
    }

    for (const unsubscribe of this.eventSubscriptions.values()) {
      try {
        unsubscribe();
      } catch {
        // ignore cleanup failures during shutdown
      }
    }
    this.eventSubscriptions.clear();

    for (const pending of this.pendingRequests.values()) {
      clearTimeout(pending.timeoutId);
      pending.reject(new Error("Sandbox runtime disposed before response."));
    }
    this.pendingRequests.clear();

    this.overlayActionIds.clear();
    this.widgetActionIdsByWidgetId.clear();

    dismissOverlay();
    if (this.widgetApiV2Enabled) {
      clearExtensionWidgets(this.widgetOwnerId);
    } else {
      dismissWidget();
    }

    window.removeEventListener("message", this.onWindowMessage);
    this.iframe.remove();
  }

  private buildSandboxSrcdoc(): string {
    const sourceConfig = this.options.source.kind === "inline"
      ? { kind: "inline", code: this.options.source.code }
      : { kind: "module", specifier: this.options.source.specifier };

    const config = {
      channel: SANDBOX_CHANNEL,
      instanceId: this.options.instanceId,
      extensionName: this.options.extensionName,
      source: sourceConfig,
      widgetApiV2Enabled: this.widgetApiV2Enabled,
    };

    const serializedConfig = serializeForSandboxInlineScript(config);

    return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
  </head>
  <body>
    <script type="module">
      const config = ${serializedConfig};

      const pendingHostRequests = new Map();
      const commandHandlers = new Map();
      const toolHandlers = new Map();
      const agentEventHandlers = new Map();
      const uiActionHandlers = new Map();
      const overlayActionIds = new Set();
      const widgetActionIds = new Set();
      const widgetActionIdsByWidgetId = new Map();
      const LEGACY_WIDGET_ID = "__legacy__";
      const ALLOWED_UI_TAGS = new Set([
        "div",
        "span",
        "p",
        "strong",
        "em",
        "code",
        "pre",
        "ul",
        "ol",
        "li",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "button",
      ]);

      let nextRequestId = 1;
      let moduleDeactivate = null;
      let cleanups = [];
      let activating = true;
      const activationOps = [];

      function getErrorMessage(error) {
        if (error instanceof Error && error.message.trim().length > 0) {
          return error.message;
        }
        return String(error);
      }

      function sendToHost(envelope) {
        parent.postMessage(envelope, "*");
      }

      function sendEvent(eventName, data) {
        sendToHost({
          channel: config.channel,
          instanceId: config.instanceId,
          direction: "sandbox_to_host",
          kind: "event",
          event: eventName,
          data,
        });
      }

      function respondToHost(requestId, ok, payload) {
        const message = {
          channel: config.channel,
          instanceId: config.instanceId,
          direction: "sandbox_to_host",
          kind: "response",
          requestId,
          ok,
        };

        if (ok) {
          message.result = payload;
        } else {
          message.error = typeof payload === "string" ? payload : "Unknown sandbox error";
        }

        sendToHost(message);
      }

      function requestHost(method, params) {
        const requestId = 'host-' + String(nextRequestId++);

        return new Promise((resolve, reject) => {
          pendingHostRequests.set(requestId, { resolve, reject });

          sendToHost({
            channel: config.channel,
            instanceId: config.instanceId,
            direction: "sandbox_to_host",
            kind: "request",
            requestId,
            method,
            params,
          });
        });
      }

      function collectActivationCleanups(result) {
        if (typeof result === "undefined") {
          return [];
        }

        if (typeof result === "function") {
          return [result];
        }

        if (!Array.isArray(result)) {
          throw new Error("activate(api) must return void, a cleanup function, or an array of cleanup functions");
        }

        const list = [];
        for (const item of result) {
          if (typeof item !== "function") {
            throw new Error("activate(api) returned an invalid cleanup entry; expected a function");
          }

          list.push(item);
        }

        return list;
      }

      function clearSurfaceActions(surfaceActions) {
        for (const actionId of surfaceActions) {
          uiActionHandlers.delete(actionId);
        }

        surfaceActions.clear();
      }

      function clearWidgetActions(widgetId) {
        const actionIds = widgetActionIdsByWidgetId.get(widgetId);
        if (!actionIds) {
          return;
        }

        clearSurfaceActions(actionIds);
        widgetActionIdsByWidgetId.delete(widgetId);
      }

      function clearAllWidgetActions() {
        for (const widgetId of widgetActionIdsByWidgetId.keys()) {
          clearWidgetActions(widgetId);
        }
      }

      function getWidgetSurfaceActions(widgetId) {
        if (!config.widgetApiV2Enabled) {
          return widgetActionIds;
        }

        const existing = widgetActionIdsByWidgetId.get(widgetId);
        if (existing) {
          clearSurfaceActions(existing);
        }

        const next = new Set();
        widgetActionIdsByWidgetId.set(widgetId, next);
        return next;
      }

      function sanitizeActionToken(value) {
        if (typeof value !== "string") {
          return "";
        }

        const cleaned = value.trim().replace(/[^A-Za-z0-9:_-]/g, "");
        if (cleaned.length === 0) {
          return "";
        }

        return cleaned.slice(0, 48);
      }

      function registerUiAction(surface, sourceAction, element, surfaceActions) {
        const actionToken = sanitizeActionToken(sourceAction);
        if (!actionToken) {
          return null;
        }

        const baseActionId = surface + ":" + actionToken;
        let actionId = baseActionId;
        let suffix = 1;

        while (uiActionHandlers.has(actionId)) {
          actionId = baseActionId + "-" + String(suffix);
          suffix += 1;
        }

        uiActionHandlers.set(actionId, () => {
          const click = new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
          });
          element.dispatchEvent(click);
        });

        surfaceActions.add(actionId);
        return actionId;
      }

      function normalizeUiTag(tagName) {
        const lowered = typeof tagName === "string"
          ? tagName.toLowerCase()
          : "div";

        if (!ALLOWED_UI_TAGS.has(lowered)) {
          return "div";
        }

        return lowered;
      }

      function projectUiNode(node, surface, surfaceActions, depth = 0) {
        if (!node || depth > 12) {
          return null;
        }

        if (node.nodeType === Node.TEXT_NODE) {
          const text = typeof node.textContent === "string" ? node.textContent : "";
          return {
            kind: "text",
            text,
          };
        }

        if (node.nodeType !== Node.ELEMENT_NODE) {
          return null;
        }

        const element = node;
        const projection = {
          kind: "element",
          tag: normalizeUiTag(element.tagName),
          children: [],
        };

        const className = typeof element.className === "string"
          ? element.className.trim()
          : "";
        if (className.length > 0) {
          projection.className = className;
        }

        const declaredAction = element.getAttribute("data-pi-action");
        if (declaredAction) {
          const actionId = registerUiAction(surface, declaredAction, element, surfaceActions);
          if (actionId) {
            projection.actionId = actionId;
          }
        }

        for (const child of element.childNodes) {
          const projectedChild = projectUiNode(child, surface, surfaceActions, depth + 1);
          if (!projectedChild) {
            continue;
          }

          projection.children.push(projectedChild);
        }

        return projection;
      }

      function projectSurfaceUi(surface, element, widgetId) {
        const surfaceActions = surface === "overlay"
          ? overlayActionIds
          : getWidgetSurfaceActions(widgetId || LEGACY_WIDGET_ID);

        clearSurfaceActions(surfaceActions);

        const projected = projectUiNode(element, surface, surfaceActions, 0);
        if (projected) {
          return projected;
        }

        return {
          kind: "text",
          text: element && typeof element.textContent === "string"
            ? element.textContent
            : "",
        };
      }

      async function runDeactivate() {
        const failures = [];

        for (let i = cleanups.length - 1; i >= 0; i -= 1) {
          const cleanup = cleanups[i];
          try {
            await cleanup();
          } catch (error) {
            failures.push(getErrorMessage(error));
          }
        }

        if (typeof moduleDeactivate === "function") {
          try {
            await moduleDeactivate();
          } catch (error) {
            failures.push(getErrorMessage(error));
          }
        }

        for (const subscriptionId of agentEventHandlers.keys()) {
          requestHost("unsubscribe_agent_events", { subscriptionId })
            .catch(() => {
              // ignore
            });
        }

        agentEventHandlers.clear();
        clearSurfaceActions(overlayActionIds);
        clearSurfaceActions(widgetActionIds);
        clearAllWidgetActions();

        if (failures.length > 0) {
          throw new Error('Extension cleanup failed:\n- ' + failures.join('\n- '));
        }
      }

      async function handleHostRequest(message) {
        const { method, requestId, params } = message;

        try {
          if (method === "invoke_command") {
            const payload = typeof params === "object" && params !== null ? params : {};
            const commandId = typeof payload.commandId === "string" ? payload.commandId : "";
            const args = typeof payload.args === "string" ? payload.args : "";

            const handler = commandHandlers.get(commandId);
            if (typeof handler !== "function") {
              throw new Error('Unknown sandbox command id: ' + commandId);
            }

            await handler(args);
            respondToHost(requestId, true, null);
            return;
          }

          if (method === "invoke_tool") {
            const payload = typeof params === "object" && params !== null ? params : {};
            const toolId = typeof payload.toolId === "string" ? payload.toolId : "";

            const execute = toolHandlers.get(toolId);
            if (typeof execute !== "function") {
              throw new Error('Unknown sandbox tool id: ' + toolId);
            }

            const result = await execute(payload.params);
            respondToHost(requestId, true, result);
            return;
          }

          if (method === "ui_action") {
            const payload = typeof params === "object" && params !== null ? params : {};
            const actionId = typeof payload.actionId === "string" ? payload.actionId : "";

            const handler = uiActionHandlers.get(actionId);
            if (typeof handler !== "function") {
              throw new Error('Unknown sandbox UI action id: ' + actionId);
            }

            await handler();
            respondToHost(requestId, true, null);
            return;
          }

          if (method === "deactivate") {
            await runDeactivate();
            respondToHost(requestId, true, null);
            return;
          }

          throw new Error('Unsupported host request method: ' + method);
        } catch (error) {
          respondToHost(requestId, false, getErrorMessage(error));
        }
      }

      function handleHostEvent(message) {
        if (message.event !== "agent_event") {
          return;
        }

        const data = typeof message.data === "object" && message.data !== null ? message.data : {};
        const subscriptionId = typeof data.subscriptionId === "string" ? data.subscriptionId : "";

        if (!subscriptionId) {
          return;
        }

        const handler = agentEventHandlers.get(subscriptionId);
        if (typeof handler !== "function") {
          return;
        }

        try {
          handler(data.event);
        } catch (error) {
          console.warn('[pi] sandbox agent event handler failed:', getErrorMessage(error));
        }
      }

      window.addEventListener("message", (event) => {
        const message = event.data;
        if (!message || typeof message !== "object") {
          return;
        }

        if (message.channel !== config.channel || message.instanceId !== config.instanceId) {
          return;
        }

        if (message.direction !== "host_to_sandbox") {
          return;
        }

        if (message.kind === "response") {
          const pending = pendingHostRequests.get(message.requestId);
          if (!pending) {
            return;
          }

          pendingHostRequests.delete(message.requestId);
          if (message.ok) {
            pending.resolve(message.result);
          } else {
            pending.reject(new Error(typeof message.error === "string" ? message.error : "Sandbox host request failed."));
          }

          return;
        }

        if (message.kind === "request") {
          void handleHostRequest(message);
          return;
        }

        if (message.kind === "event") {
          handleHostEvent(message);
        }
      });

      function queueActivationOp(promise) {
        if (activating) {
          activationOps.push(promise);
          return;
        }

        promise.catch((error) => {
          console.warn('[pi] sandbox host operation failed:', getErrorMessage(error));
        });
      }

      function makeApi() {
        return {
          registerCommand(name, cmd) {
            const normalizedName = typeof name === "string" ? name.trim() : "";
            if (!normalizedName) {
              throw new Error('Extension command name cannot be empty');
            }

            if (!cmd || typeof cmd !== "object") {
              throw new Error('registerCommand requires a command definition');
            }

            if (typeof cmd.handler !== "function") {
              throw new Error('registerCommand handler must be a function');
            }

            const description = typeof cmd.description === "string" ? cmd.description : "";
            const commandId = 'cmd-' + String(nextRequestId++);
            commandHandlers.set(commandId, cmd.handler);

            queueActivationOp(requestHost("register_command", {
              commandId,
              name: normalizedName,
              description,
            }));
          },

          registerTool(name, tool) {
            const normalizedName = typeof name === "string" ? name.trim() : "";
            if (!normalizedName) {
              throw new Error('Extension tool name cannot be empty');
            }

            if (!tool || typeof tool !== "object") {
              throw new Error('registerTool requires a tool definition');
            }

            if (typeof tool.execute !== "function") {
              throw new Error('registerTool execute must be a function');
            }

            const toolId = 'tool-' + String(nextRequestId++);
            toolHandlers.set(toolId, (params) => tool.execute(params));

            queueActivationOp(requestHost("register_tool", {
              toolId,
              name: normalizedName,
              label: typeof tool.label === "string" ? tool.label : normalizedName,
              description: typeof tool.description === "string" ? tool.description : "",
              parameters: tool.parameters,
            }));
          },

          unregisterTool(name) {
            const normalizedName = typeof name === "string" ? name.trim() : "";
            if (!normalizedName) {
              throw new Error('Extension tool name cannot be empty');
            }

            queueActivationOp(requestHost("unregister_tool", {
              name: normalizedName,
            }));
          },

          agent: {
            get raw() {
              throw new Error('api.agent is not available in sandbox runtime. Use onAgentEvent() and explicit APIs.');
            },

            injectContext(content) {
              queueActivationOp(requestHost("agent_inject_context", {
                content: typeof content === "string" ? content : String(content),
              }));
            },

            steer(content) {
              queueActivationOp(requestHost("agent_steer", {
                content: typeof content === "string" ? content : String(content),
              }));
            },

            followUp(content) {
              queueActivationOp(requestHost("agent_follow_up", {
                content: typeof content === "string" ? content : String(content),
              }));
            },
          },

          llm: {
            complete(request) {
              return requestHost("llm_complete", {
                request,
              });
            },
          },

          http: {
            fetch(url, options) {
              return requestHost("http_fetch", {
                url,
                options,
              });
            },
          },

          storage: {
            get(key) {
              return requestHost("storage_get", { key });
            },
            set(key, value) {
              return requestHost("storage_set", { key, value });
            },
            delete(key) {
              return requestHost("storage_delete", { key });
            },
            keys() {
              return requestHost("storage_keys", {});
            },
          },

          clipboard: {
            writeText(text) {
              return requestHost("clipboard_write_text", {
                text: typeof text === "string" ? text : String(text),
              });
            },
          },

          skills: {
            list() {
              return requestHost("skills_list", {});
            },
            read(name) {
              return requestHost("skills_read", { name });
            },
            install(name, markdown) {
              return requestHost("skills_install", { name, markdown });
            },
            uninstall(name) {
              return requestHost("skills_uninstall", { name });
            },
          },

          download: {
            download(filename, content, mimeType) {
              queueActivationOp(requestHost("download_file", {
                filename,
                content,
                mimeType,
              }));
            },
          },

          overlay: {
            show(el) {
              const tree = projectSurfaceUi("overlay", el);
              queueActivationOp(requestHost("overlay_show", { tree }));
            },
            dismiss() {
              clearSurfaceActions(overlayActionIds);
              queueActivationOp(requestHost("overlay_dismiss", {}));
            },
          },

          widget: {
            show(el) {
              if (config.widgetApiV2Enabled) {
                const tree = projectSurfaceUi("widget", el, LEGACY_WIDGET_ID);
                queueActivationOp(requestHost("widget_upsert", {
                  widgetId: LEGACY_WIDGET_ID,
                  tree,
                  placement: "above-input",
                  order: 0,
                }));
                return;
              }

              const tree = projectSurfaceUi("widget", el, LEGACY_WIDGET_ID);
              queueActivationOp(requestHost("widget_show", { tree }));
            },
            dismiss() {
              if (config.widgetApiV2Enabled) {
                clearWidgetActions(LEGACY_WIDGET_ID);
                queueActivationOp(requestHost("widget_remove", { widgetId: LEGACY_WIDGET_ID }));
                return;
              }

              clearSurfaceActions(widgetActionIds);
              queueActivationOp(requestHost("widget_dismiss", {}));
            },
            upsert(spec) {
              if (!config.widgetApiV2Enabled) {
                throw new Error('Widget API v2 is disabled. Enable /experimental on extension-widget-v2.');
              }

              const payload = spec && typeof spec === "object" ? spec : null;
              if (!payload) {
                throw new Error("widget.upsert requires a widget spec object");
              }

              const widgetId = typeof payload.id === "string" ? payload.id.trim() : "";
              if (!widgetId) {
                throw new Error("widget.upsert requires a non-empty id");
              }

              if (!(payload.el instanceof HTMLElement)) {
                throw new Error("widget.upsert requires an HTMLElement in spec.el");
              }

              const tree = projectSurfaceUi("widget", payload.el, widgetId);

              queueActivationOp(requestHost("widget_upsert", {
                widgetId,
                tree,
                title: typeof payload.title === "string" ? payload.title : undefined,
                placement: payload.placement === "above-input" || payload.placement === "below-input"
                  ? payload.placement
                  : undefined,
                order: typeof payload.order === "number" ? payload.order : undefined,
                collapsible: typeof payload.collapsible === "boolean" ? payload.collapsible : undefined,
                collapsed: typeof payload.collapsed === "boolean" ? payload.collapsed : undefined,
                minHeightPx: typeof payload.minHeightPx === "number"
                  ? payload.minHeightPx
                  : payload.minHeightPx === null
                    ? null
                    : undefined,
                maxHeightPx: typeof payload.maxHeightPx === "number"
                  ? payload.maxHeightPx
                  : payload.maxHeightPx === null
                    ? null
                    : undefined,
              }));
            },
            remove(id) {
              if (!config.widgetApiV2Enabled) {
                throw new Error('Widget API v2 is disabled. Enable /experimental on extension-widget-v2.');
              }

              const widgetId = typeof id === "string" ? id.trim() : "";
              if (!widgetId) {
                throw new Error("widget.remove requires a non-empty id");
              }

              clearWidgetActions(widgetId);
              queueActivationOp(requestHost("widget_remove", { widgetId }));
            },
            clear() {
              if (!config.widgetApiV2Enabled) {
                throw new Error('Widget API v2 is disabled. Enable /experimental on extension-widget-v2.');
              }

              clearAllWidgetActions();
              queueActivationOp(requestHost("widget_clear", {}));
            },
          },

          toast(message) {
            queueActivationOp(requestHost("toast", {
              message: typeof message === "string" ? message : String(message),
            }));
          },

          onAgentEvent(handler) {
            if (typeof handler !== "function") {
              throw new Error('onAgentEvent requires a function handler');
            }

            const subscriptionId = 'ev-' + String(nextRequestId++);
            agentEventHandlers.set(subscriptionId, handler);

            queueActivationOp(requestHost("subscribe_agent_events", { subscriptionId }));

            return () => {
              agentEventHandlers.delete(subscriptionId);
              requestHost("unsubscribe_agent_events", { subscriptionId })
                .catch(() => {
                  // ignore unsubscribe failures
                });
            };
          },
        };
      }

      async function importExtensionModule() {
        if (config.source.kind === "inline") {
          const blob = new Blob([config.source.code], { type: 'text/javascript' });
          const blobUrl = URL.createObjectURL(blob);

          try {
            return await import(blobUrl);
          } finally {
            URL.revokeObjectURL(blobUrl);
          }
        }

        return import(config.source.specifier);
      }

      async function activateExtension() {
        const importedModule = await importExtensionModule();

        const activate = typeof importedModule.activate === "function"
          ? importedModule.activate
          : typeof importedModule.default === "function"
            ? importedModule.default
            : null;

        if (!activate) {
          throw new Error('Extension module "' + config.extensionName + '" must export an activate(api) function');
        }

        moduleDeactivate = typeof importedModule.deactivate === "function"
          ? importedModule.deactivate
          : null;

        const activationResult = await activate(makeApi());
        cleanups = collectActivationCleanups(activationResult);

        activating = false;
        await Promise.all(activationOps);

        sendEvent("ready", null);
      }

      activateExtension().catch((error) => {
        sendEvent("error", { message: getErrorMessage(error) });
      });
    </script>
  </body>
</html>`;
  }

  private async callSandbox(
    method: string,
    params: unknown,
    options?: {
      allowWhenDisposed?: boolean;
    },
  ): Promise<unknown> {
    if (this.disposed && !options?.allowWhenDisposed) {
      throw new Error("Sandbox runtime is already disposed.");
    }

    const targetWindow = this.iframe.contentWindow;
    if (!targetWindow) {
      throw new Error("Sandbox frame is not ready.");
    }

    const requestId = `req-${this.nextRequestId}`;
    this.nextRequestId += 1;

    return new Promise<unknown>((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        this.pendingRequests.delete(requestId);
        reject(new Error(`Sandbox request timed out: ${method}`));
      }, SANDBOX_REQUEST_TIMEOUT_MS);

      this.pendingRequests.set(requestId, {
        resolve,
        reject,
        timeoutId,
      });

      const envelope: SandboxRequestEnvelope = {
        channel: SANDBOX_CHANNEL,
        instanceId: this.options.instanceId,
        direction: "host_to_sandbox",
        kind: "request",
        requestId,
        method,
        params,
      };

      targetWindow.postMessage(envelope, "*");
    });
  }

  private sendResponse(
    requestId: string,
    ok: boolean,
    payload: unknown,
  ): void {
    const targetWindow = this.iframe.contentWindow;
    if (!targetWindow) {
      return;
    }

    const envelope: SandboxResponseEnvelope = {
      channel: SANDBOX_CHANNEL,
      instanceId: this.options.instanceId,
      direction: "host_to_sandbox",
      kind: "response",
      requestId,
      ok,
    };

    if (ok) {
      envelope.result = payload;
    } else {
      envelope.error = typeof payload === "string" ? payload : "Unknown host error";
    }

    targetWindow.postMessage(envelope, "*");
  }

  private sendEvent(eventName: string, data: unknown): void {
    const targetWindow = this.iframe.contentWindow;
    if (!targetWindow) {
      return;
    }

    const envelope: SandboxEventEnvelope = {
      channel: SANDBOX_CHANNEL,
      instanceId: this.options.instanceId,
      direction: "host_to_sandbox",
      kind: "event",
      event: eventName,
      data,
    };

    targetWindow.postMessage(envelope, "*");
  }

  private replaceActionIds(target: Set<string>, next: Set<string>): void {
    target.clear();
    for (const actionId of next) {
      target.add(actionId);
    }
  }

  private replaceWidgetActionIds(widgetId: string, next: Set<string>): void {
    this.widgetActionIdsByWidgetId.set(widgetId, new Set(next));
  }

  private clearWidgetActionIds(widgetId: string): void {
    this.widgetActionIdsByWidgetId.delete(widgetId);
  }

  private clearAllWidgetActionIds(): void {
    this.widgetActionIdsByWidgetId.clear();
  }

  private dispatchSandboxUiAction(actionId: string): void {
    void this.callSandbox("ui_action", { actionId })
      .catch((error: unknown) => {
        console.warn(`[pi] Sandbox UI action failed: ${getErrorMessage(error)}`);
      });
  }

  private handleWindowMessage(event: MessageEvent<unknown>): void {
    if (event.source !== this.iframe.contentWindow) {
      return;
    }

    const envelope = event.data;
    if (!isSandboxEnvelope(envelope)) {
      return;
    }

    if (envelope.instanceId !== this.options.instanceId) {
      return;
    }

    if (envelope.direction !== "sandbox_to_host") {
      return;
    }

    if (envelope.kind === "response") {
      const pending = this.pendingRequests.get(envelope.requestId);
      if (!pending) {
        return;
      }

      this.pendingRequests.delete(envelope.requestId);
      clearTimeout(pending.timeoutId);

      if (envelope.ok) {
        pending.resolve(envelope.result);
      } else {
        pending.reject(new Error(typeof envelope.error === "string" ? envelope.error : "Sandbox request failed."));
      }

      return;
    }

    if (envelope.kind === "event") {
      if (envelope.event === "ready") {
        this.resolveReady?.();
        this.resolveReady = null;
        this.rejectReady = null;
        return;
      }

      if (envelope.event === "error") {
        const payload = isRecord(envelope.data) ? envelope.data : null;
        const message = payload && typeof payload.message === "string"
          ? payload.message
          : "Sandbox bootstrap failed.";

        this.rejectReady?.(new Error(message));
        this.resolveReady = null;
        this.rejectReady = null;
      }

      return;
    }

    void this.handleSandboxRequest(envelope);
  }

  private assertCapability(capability: ExtensionCapability): void {
    if (this.options.isCapabilityEnabled(capability)) {
      return;
    }

    throw new Error(this.options.formatCapabilityError(capability));
  }

  private async handleSandboxRequest(envelope: SandboxRequestEnvelope): Promise<void> {
    const { method, requestId, params } = envelope;

    try {
      switch (method) {
        case "register_command": {
          this.assertCapability("commands.register");

          const payload = asRecord(params, "register_command params");
          const commandId = asNonEmptyString(payload.commandId, "commandId");
          const name = asNonEmptyString(payload.name, "name");
          const description = typeof payload.description === "string" ? payload.description : "";

          this.options.registerCommand(name, {
            description,
            handler: async (args: string) => {
              await this.callSandbox("invoke_command", {
                commandId,
                args,
              });
            },
          });

          this.sendResponse(requestId, true, null);
          return;
        }

        case "register_tool": {
          this.assertCapability("tools.register");

          const payload = asRecord(params, "register_tool params");
          const toolId = asNonEmptyString(payload.toolId, "toolId");
          const name = asNonEmptyString(payload.name, "name");
          const label = typeof payload.label === "string" && payload.label.trim().length > 0
            ? payload.label.trim()
            : name;
          const description = typeof payload.description === "string" ? payload.description : "";
          const parametersRaw = payload.parameters;
          const parameters = normalizeSandboxToolParameters(parametersRaw);

          const tool: AgentTool<TSchema, unknown> = {
            name,
            label,
            description,
            parameters,
            execute: async (
              _toolCallId: string,
              toolParams: unknown,
            ): Promise<AgentToolResult<unknown>> => {
              const result = await this.callSandbox("invoke_tool", {
                toolId,
                params: toolParams,
              });

              return normalizeToolResult(result);
            },
          };

          this.options.registerTool(tool);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "unregister_tool": {
          this.assertCapability("tools.register");

          const payload = asRecord(params, "unregister_tool params");
          const name = asNonEmptyString(payload.name, "name");
          this.options.unregisterTool(name);

          this.sendResponse(requestId, true, null);
          return;
        }

        case "llm_complete": {
          this.assertCapability("llm.complete");

          const payload = asRecord(params, "llm_complete params");
          const request = payload.request;

          if (!isRecord(request)) {
            throw new Error("llm_complete request must be an object.");
          }

          const messagesRaw = request.messages;
          if (!Array.isArray(messagesRaw)) {
            throw new Error("llm_complete request.messages must be an array.");
          }

          const messages: LlmCompletionRequest["messages"] = [];
          for (const value of messagesRaw) {
            if (!isRecord(value)) {
              throw new Error("llm_complete messages entries must be objects.");
            }

            const role = value.role;
            const content = value.content;
            if ((role !== "user" && role !== "assistant") || typeof content !== "string") {
              throw new Error("llm_complete messages entries must contain role + string content.");
            }

            messages.push({ role, content });
          }

          const completionRequest: LlmCompletionRequest = {
            model: typeof request.model === "string" ? request.model : undefined,
            systemPrompt: typeof request.systemPrompt === "string" ? request.systemPrompt : undefined,
            messages,
            maxTokens: typeof request.maxTokens === "number" ? request.maxTokens : undefined,
          };

          const result = await this.options.llmComplete(completionRequest);
          this.sendResponse(requestId, true, result);
          return;
        }

        case "http_fetch": {
          this.assertCapability("http.fetch");

          const payload = asRecord(params, "http_fetch params");
          const url = asNonEmptyString(payload.url, "url");
          const optionsRaw = payload.options;
          let options: HttpRequestOptions | undefined;

          if (isRecord(optionsRaw)) {
            const methodRaw = optionsRaw.method;
            const method = methodRaw === "GET"
              || methodRaw === "POST"
              || methodRaw === "PUT"
              || methodRaw === "PATCH"
              || methodRaw === "DELETE"
              || methodRaw === "HEAD"
              ? methodRaw
              : undefined;

            const headersRaw = optionsRaw.headers;
            let headers: Record<string, string> | undefined;
            if (isRecord(headersRaw)) {
              headers = {};
              for (const [key, value] of Object.entries(headersRaw)) {
                if (typeof value === "string") {
                  headers[key] = value;
                }
              }
            }

            options = {
              method,
              headers,
              body: typeof optionsRaw.body === "string" ? optionsRaw.body : undefined,
              timeoutMs: typeof optionsRaw.timeoutMs === "number" ? optionsRaw.timeoutMs : undefined,
            };
          }

          const response = await this.options.httpFetch(url, options);
          this.sendResponse(requestId, true, response);
          return;
        }

        case "storage_get": {
          this.assertCapability("storage.readwrite");

          const payload = asRecord(params, "storage_get params");
          const key = asNonEmptyString(payload.key, "key");
          const value = await this.options.storageGet(key);
          this.sendResponse(requestId, true, value);
          return;
        }

        case "storage_set": {
          this.assertCapability("storage.readwrite");

          const payload = asRecord(params, "storage_set params");
          const key = asNonEmptyString(payload.key, "key");
          await this.options.storageSet(key, payload.value);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "storage_delete": {
          this.assertCapability("storage.readwrite");

          const payload = asRecord(params, "storage_delete params");
          const key = asNonEmptyString(payload.key, "key");
          await this.options.storageDelete(key);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "storage_keys": {
          this.assertCapability("storage.readwrite");
          const keys = await this.options.storageKeys();
          this.sendResponse(requestId, true, keys);
          return;
        }

        case "clipboard_write_text": {
          this.assertCapability("clipboard.write");

          const payload = asRecord(params, "clipboard_write_text params");
          const text = sanitizeText(payload.text);
          await this.options.clipboardWriteText(text);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "agent_inject_context": {
          this.assertCapability("agent.context.write");

          const payload = asRecord(params, "agent_inject_context params");
          const content = asNonEmptyString(payload.content, "content");
          this.options.injectAgentContext(content);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "agent_steer": {
          this.assertCapability("agent.steer");

          const payload = asRecord(params, "agent_steer params");
          const content = asNonEmptyString(payload.content, "content");
          this.options.steerAgent(content);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "agent_follow_up": {
          this.assertCapability("agent.followup");

          const payload = asRecord(params, "agent_follow_up params");
          const content = asNonEmptyString(payload.content, "content");
          this.options.followUpAgent(content);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "skills_list": {
          this.assertCapability("skills.read");
          const skills = await this.options.listSkills();
          this.sendResponse(requestId, true, skills);
          return;
        }

        case "skills_read": {
          this.assertCapability("skills.read");

          const payload = asRecord(params, "skills_read params");
          const name = asNonEmptyString(payload.name, "name");
          const markdown = await this.options.readSkill(name);
          this.sendResponse(requestId, true, markdown);
          return;
        }

        case "skills_install": {
          this.assertCapability("skills.write");

          const payload = asRecord(params, "skills_install params");
          const name = asNonEmptyString(payload.name, "name");
          const markdown = asNonEmptyString(payload.markdown, "markdown");
          await this.options.installSkill(name, markdown);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "skills_uninstall": {
          this.assertCapability("skills.write");

          const payload = asRecord(params, "skills_uninstall params");
          const name = asNonEmptyString(payload.name, "name");
          await this.options.uninstallSkill(name);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "download_file": {
          this.assertCapability("download.file");

          const payload = asRecord(params, "download_file params");
          const filename = asNonEmptyString(payload.filename, "filename");
          if (typeof payload.content !== "string") {
            throw new Error("download_file content must be a string.");
          }

          const content = payload.content;
          const mimeType = typeof payload.mimeType === "string" ? payload.mimeType : undefined;

          this.options.downloadFile(filename, content, mimeType);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "toast": {
          this.assertCapability("ui.toast");

          const payload = asRecord(params, "toast params");
          const message = sanitizeText(payload.message);
          this.options.toast(message);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "overlay_show": {
          this.assertCapability("ui.overlay");

          const payload = asRecord(params, "overlay_show params");
          const tree = normalizeSandboxUiNode(payload.tree);
          const actionIds = showOverlayNode(tree, (actionId) => {
            if (!this.overlayActionIds.has(actionId)) {
              return;
            }

            this.dispatchSandboxUiAction(actionId);
          });

          this.replaceActionIds(this.overlayActionIds, actionIds);

          this.sendResponse(requestId, true, null);
          return;
        }

        case "overlay_show_text": {
          this.assertCapability("ui.overlay");

          const payload = asRecord(params, "overlay_show_text params");
          const fallbackNode = createTextOnlyUiNode(sanitizeText(payload.text));
          const actionIds = showOverlayNode(fallbackNode, () => {
            // legacy text-only path has no actions
          });

          this.replaceActionIds(this.overlayActionIds, actionIds);

          this.sendResponse(requestId, true, null);
          return;
        }

        case "overlay_dismiss": {
          this.assertCapability("ui.overlay");
          this.overlayActionIds.clear();
          dismissOverlay();
          this.sendResponse(requestId, true, null);
          return;
        }

        case "widget_show": {
          this.assertCapability("ui.widget");

          const payload = asRecord(params, "widget_show params");
          const tree = normalizeSandboxUiNode(payload.tree);

          if (this.widgetApiV2Enabled) {
            const actionIds = upsertSandboxWidgetNode({
              ownerId: this.widgetOwnerId,
              widgetId: LEGACY_WIDGET_ID,
              node: tree,
              onAction: (actionId) => {
                const knownActionIds = this.widgetActionIdsByWidgetId.get(LEGACY_WIDGET_ID);
                if (!knownActionIds || !knownActionIds.has(actionId)) {
                  return;
                }

                this.dispatchSandboxUiAction(actionId);
              },
              placement: "above-input",
              order: 0,
            });

            this.replaceWidgetActionIds(LEGACY_WIDGET_ID, actionIds);
          } else {
            const actionIds = showWidgetNode(tree, (actionId) => {
              const knownActionIds = this.widgetActionIdsByWidgetId.get(LEGACY_WIDGET_ID);
              if (!knownActionIds || !knownActionIds.has(actionId)) {
                return;
              }

              this.dispatchSandboxUiAction(actionId);
            });

            this.replaceWidgetActionIds(LEGACY_WIDGET_ID, actionIds);
          }

          this.sendResponse(requestId, true, null);
          return;
        }

        case "widget_show_text": {
          this.assertCapability("ui.widget");

          const payload = asRecord(params, "widget_show_text params");
          const fallbackNode = createTextOnlyUiNode(sanitizeText(payload.text));

          if (this.widgetApiV2Enabled) {
            const actionIds = upsertSandboxWidgetNode({
              ownerId: this.widgetOwnerId,
              widgetId: LEGACY_WIDGET_ID,
              node: fallbackNode,
              onAction: () => {
                // legacy text-only path has no actions
              },
              placement: "above-input",
              order: 0,
            });

            this.replaceWidgetActionIds(LEGACY_WIDGET_ID, actionIds);
          } else {
            const actionIds = showWidgetNode(fallbackNode, () => {
              // legacy text-only path has no actions
            });

            this.replaceWidgetActionIds(LEGACY_WIDGET_ID, actionIds);
          }

          this.sendResponse(requestId, true, null);
          return;
        }

        case "widget_dismiss": {
          this.assertCapability("ui.widget");
          this.clearWidgetActionIds(LEGACY_WIDGET_ID);

          if (this.widgetApiV2Enabled) {
            removeExtensionWidget(this.widgetOwnerId, LEGACY_WIDGET_ID);
          } else {
            dismissWidget();
          }

          this.sendResponse(requestId, true, null);
          return;
        }

        case "widget_upsert": {
          this.assertCapability("ui.widget");
          if (!this.widgetApiV2Enabled) {
            throw new Error("Widget API v2 is disabled. Enable /experimental on extension-widget-v2.");
          }

          const payload = asRecord(params, "widget_upsert params");
          const widgetId = asNonEmptyString(payload.widgetId, "widgetId");
          const tree = normalizeSandboxUiNode(payload.tree);
          const title = typeof payload.title === "string" ? payload.title : undefined;
          const placement = asWidgetPlacementOrUndefined(payload.placement);
          const order = asFiniteNumberOrNull(payload.order);
          const minHeightPx = asFiniteNumberOrNullOrUndefined(payload.minHeightPx);
          const maxHeightPx = asFiniteNumberOrNullOrUndefined(payload.maxHeightPx);
          const collapsible = asBooleanOrUndefined(payload.collapsible);
          const collapsed = asBooleanOrUndefined(payload.collapsed);

          const actionIds = upsertSandboxWidgetNode({
            ownerId: this.widgetOwnerId,
            widgetId,
            node: tree,
            onAction: (actionId) => {
              const knownActionIds = this.widgetActionIdsByWidgetId.get(widgetId);
              if (!knownActionIds || !knownActionIds.has(actionId)) {
                return;
              }

              this.dispatchSandboxUiAction(actionId);
            },
            title,
            placement,
            order: order ?? undefined,
            collapsible,
            collapsed,
            minHeightPx,
            maxHeightPx,
          });

          this.replaceWidgetActionIds(widgetId, actionIds);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "widget_remove": {
          this.assertCapability("ui.widget");
          if (!this.widgetApiV2Enabled) {
            throw new Error("Widget API v2 is disabled. Enable /experimental on extension-widget-v2.");
          }

          const payload = asRecord(params, "widget_remove params");
          const widgetId = asNonEmptyString(payload.widgetId, "widgetId");
          this.clearWidgetActionIds(widgetId);
          removeExtensionWidget(this.widgetOwnerId, widgetId);

          this.sendResponse(requestId, true, null);
          return;
        }

        case "widget_clear": {
          this.assertCapability("ui.widget");
          if (!this.widgetApiV2Enabled) {
            throw new Error("Widget API v2 is disabled. Enable /experimental on extension-widget-v2.");
          }

          this.clearAllWidgetActionIds();
          clearExtensionWidgets(this.widgetOwnerId);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "subscribe_agent_events": {
          this.assertCapability("agent.events.read");

          const payload = asRecord(params, "subscribe_agent_events params");
          const subscriptionId = asNonEmptyString(payload.subscriptionId, "subscriptionId");

          if (!this.eventSubscriptions.has(subscriptionId)) {
            const unsubscribe = this.options.subscribeAgentEvents((agentEvent) => {
              this.sendEvent("agent_event", {
                subscriptionId,
                event: agentEvent,
              });
            });

            this.eventSubscriptions.set(subscriptionId, unsubscribe);
          }

          this.sendResponse(requestId, true, null);
          return;
        }

        case "unsubscribe_agent_events": {
          const payload = asRecord(params, "unsubscribe_agent_events params");
          const subscriptionId = asNonEmptyString(payload.subscriptionId, "subscriptionId");

          const unsubscribe = this.eventSubscriptions.get(subscriptionId);
          if (unsubscribe) {
            this.eventSubscriptions.delete(subscriptionId);
            unsubscribe();
          }

          this.sendResponse(requestId, true, null);
          return;
        }

        default:
          throw new Error(`Unsupported sandbox request method: ${method}`);
      }
    } catch (error: unknown) {
      this.sendResponse(requestId, false, getErrorMessage(error));
    }
  }
}

export async function activateExtensionInSandbox(
  options: SandboxActivationOptions,
): Promise<LoadedExtensionHandle> {
  if (typeof document === "undefined" || typeof window === "undefined") {
    throw new Error("Sandbox runtime is only available in a browser environment.");
  }

  const host = new SandboxRuntimeHost(options);

  try {
    await host.waitUntilReady();
  } catch (error: unknown) {
    await host.dispose(false);
    throw error;
  }

  let deactivated = false;

  return {
    deactivate: async () => {
      if (deactivated) {
        return;
      }

      deactivated = true;
      await host.dispose(true);
    },
  };
}
