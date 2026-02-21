/**
 * Sandboxed extension runtime host (iframe + postMessage RPC).
 *
 * Scope: default runtime for untrusted extension execution (with rollback switch).
 * This bridge intentionally starts small and only forwards a constrained,
 * sanitized UI projection format (no raw HTML injection).
 */

import type { AgentEvent, AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import type { TSchema } from "@sinclair/typebox";

import type {
  ExtensionCommand,
  ExtensionConnectionDefinition,
  HttpRequestOptions,
  HttpResponse,
  LlmCompletionRequest,
  LlmCompletionResult,
  LoadedExtensionHandle,
  SkillSummary,
} from "../commands/extension-api.js";
import type { ConnectionState, ConnectionStatus } from "../connections/types.js";
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
  type SandboxRequestEnvelope,
  type SandboxResponseEnvelope,
  type SandboxEventEnvelope,
} from "./sandbox/protocol.js";
import { buildSandboxSrcdoc } from "./sandbox/srcdoc.js";
import {
  createTextOnlyUiNode,
  dismissOverlay,
  dismissWidget,
  showOverlayNode,
  showWidgetNode,
  upsertSandboxWidgetNode,
} from "./sandbox/surfaces.js";
import {
  asBooleanOrUndefined,
  asFiniteNumberOrNull,
  asFiniteNumberOrNullOrUndefined,
  asNonEmptyString,
  asRecord,
  asWidgetPlacementOrUndefined,
  getErrorMessage,
  normalizeSandboxToolParameters,
  normalizeSandboxToolResult,
  parseSandboxHttpRequestOptions,
  parseSandboxLlmCompletionRequest,
  sanitizeText,
} from "./sandbox/runtime-helpers.js";
import { isRecord } from "../utils/type-guards.js";

export { normalizeSandboxToolParameters } from "./sandbox/runtime-helpers.js";

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
  registerConnection: (definition: ExtensionConnectionDefinition) => string;
  unregisterConnection: (connectionId: string) => void;
  listConnections: () => Promise<ConnectionState[]>;
  getConnection: (connectionId: string) => Promise<ConnectionState | null>;
  setConnectionSecrets: (connectionId: string, secrets: Record<string, string>) => Promise<void>;
  clearConnectionSecrets: (connectionId: string) => Promise<void>;
  markConnectionValidated: (connectionId: string) => Promise<void>;
  markConnectionInvalid: (connectionId: string, reason: string) => Promise<void>;
  markConnectionStatus: (connectionId: string, status: ConnectionStatus, reason?: string) => Promise<void>;
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

function parseConnectionStatus(value: unknown): ConnectionStatus {
  if (value === "connected" || value === "missing" || value === "invalid" || value === "error") {
    return value;
  }

  throw new Error("Invalid connection status.");
}

function parseConnectionSecrets(value: unknown): Record<string, string> {
  const payload = asRecord(value, "connection secrets");
  const secrets: Record<string, string> = {};

  for (const [fieldId, rawSecret] of Object.entries(payload)) {
    if (typeof rawSecret !== "string") {
      throw new Error(`Connection secret field \"${fieldId}\" must be a string.`);
    }

    secrets[fieldId] = rawSecret;
  }

  return secrets;
}

function parseConnectionDefinition(value: unknown): ExtensionConnectionDefinition {
  const payload = asRecord(value, "connection definition");
  const title = asNonEmptyString(payload.title, "title");
  const id = asNonEmptyString(payload.id, "id");
  const capability = asNonEmptyString(payload.capability, "capability");

  const authKindRaw = payload.authKind;
  if (
    authKindRaw !== "api_key"
    && authKindRaw !== "bearer_token"
    && authKindRaw !== "oauth"
    && authKindRaw !== "custom"
  ) {
    throw new Error("connection definition authKind must be one of: api_key, bearer_token, oauth, custom.");
  }

  if (!Array.isArray(payload.secretFields)) {
    throw new Error("connection definition secretFields must be an array.");
  }

  const secretFields = payload.secretFields.map((fieldRaw, index) => {
    const field = asRecord(fieldRaw, `secretFields[${index}]`);
    const fieldId = asNonEmptyString(field.id, `secretFields[${index}].id`);
    const label = asNonEmptyString(field.label, `secretFields[${index}].label`);
    const required = field.required;
    if (typeof required !== "boolean") {
      throw new Error(`secretFields[${index}].required must be a boolean.`);
    }

    const maskInUi = field.maskInUi;
    if (maskInUi !== undefined && typeof maskInUi !== "boolean") {
      throw new Error(`secretFields[${index}].maskInUi must be a boolean when provided.`);
    }

    return {
      id: fieldId,
      label,
      required,
      maskInUi,
    };
  });

  const setupHintRaw = payload.setupHint;
  const setupHint = typeof setupHintRaw === "string" && setupHintRaw.trim().length > 0
    ? setupHintRaw.trim()
    : undefined;

  return {
    id,
    title,
    capability,
    authKind: authKindRaw,
    secretFields,
    setupHint,
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
    this.iframe.srcdoc = buildSandboxSrcdoc({
      instanceId: this.options.instanceId,
      extensionName: this.options.extensionName,
      source: this.options.source,
      widgetApiV2Enabled: this.widgetApiV2Enabled,
    });

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
          const busyAllowed = typeof payload.busyAllowed === "boolean" ? payload.busyAllowed : true;

          this.options.registerCommand(name, {
            description,
            busyAllowed,
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

          const requiresConnectionRaw = payload.requiresConnection;
          let requiresConnection: string[] | undefined;

          if (requiresConnectionRaw !== undefined) {
            if (!Array.isArray(requiresConnectionRaw)) {
              throw new Error("register_tool requiresConnection must be an array of strings.");
            }

            const normalizedRequirements: string[] = [];
            for (const requirement of requiresConnectionRaw) {
              if (typeof requirement !== "string") {
                throw new Error("register_tool requiresConnection entries must be strings.");
              }

              const trimmed = requirement.trim().toLowerCase();
              if (trimmed.length === 0) continue;

              const ownerPrefix = `${this.widgetOwnerId.toLowerCase()}.`;
              const qualified = trimmed.startsWith(ownerPrefix)
                ? trimmed
                : `${ownerPrefix}${trimmed}`;

              normalizedRequirements.push(qualified);
            }

            requiresConnection = normalizedRequirements.length > 0
              ? Array.from(new Set(normalizedRequirements))
              : undefined;
          }

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

              return normalizeSandboxToolResult(result);
            },
          };

          if (requiresConnection && requiresConnection.length > 0) {
            Reflect.set(tool, "requiresConnection", requiresConnection);
          }

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
          const completionRequest = parseSandboxLlmCompletionRequest(payload.request);
          const result = await this.options.llmComplete(completionRequest);
          this.sendResponse(requestId, true, result);
          return;
        }

        case "http_fetch": {
          this.assertCapability("http.fetch");

          const payload = asRecord(params, "http_fetch params");
          const url = asNonEmptyString(payload.url, "url");
          const options = parseSandboxHttpRequestOptions(payload.options);
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

        case "connections_register": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_register params");
          const definition = parseConnectionDefinition(payload.definition);
          const connectionId = this.options.registerConnection(definition);
          this.sendResponse(requestId, true, { connectionId });
          return;
        }

        case "connections_unregister": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_unregister params");
          const connectionId = asNonEmptyString(payload.connectionId, "connectionId");
          this.options.unregisterConnection(connectionId);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "connections_list": {
          this.assertCapability("connections.readwrite");
          const states = await this.options.listConnections();
          this.sendResponse(requestId, true, states);
          return;
        }

        case "connections_get": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_get params");
          const connectionId = asNonEmptyString(payload.connectionId, "connectionId");
          const state = await this.options.getConnection(connectionId);
          this.sendResponse(requestId, true, state);
          return;
        }

        case "connections_set_secrets": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_set_secrets params");
          const connectionId = asNonEmptyString(payload.connectionId, "connectionId");
          const secrets = parseConnectionSecrets(payload.secrets);
          await this.options.setConnectionSecrets(connectionId, secrets);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "connections_clear_secrets": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_clear_secrets params");
          const connectionId = asNonEmptyString(payload.connectionId, "connectionId");
          await this.options.clearConnectionSecrets(connectionId);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "connections_mark_validated": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_mark_validated params");
          const connectionId = asNonEmptyString(payload.connectionId, "connectionId");
          await this.options.markConnectionValidated(connectionId);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "connections_mark_invalid": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_mark_invalid params");
          const connectionId = asNonEmptyString(payload.connectionId, "connectionId");
          const reason = asNonEmptyString(payload.reason, "reason");
          await this.options.markConnectionInvalid(connectionId, reason);
          this.sendResponse(requestId, true, null);
          return;
        }

        case "connections_mark_status": {
          this.assertCapability("connections.readwrite");

          const payload = asRecord(params, "connections_mark_status params");
          const connectionId = asNonEmptyString(payload.connectionId, "connectionId");
          const status = parseConnectionStatus(payload.status);
          const reason = typeof payload.reason === "string" && payload.reason.trim().length > 0
            ? payload.reason
            : undefined;

          await this.options.markConnectionStatus(connectionId, status, reason);
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
