import {
  SANDBOX_CHANNEL,
  serializeForSandboxInlineScript,
} from "./protocol.js";

interface SandboxSrcdocInlineSource {
  kind: "inline";
  code: string;
}

interface SandboxSrcdocModuleSource {
  kind: "module";
  specifier: string;
}

export type SandboxSrcdocSource = SandboxSrcdocInlineSource | SandboxSrcdocModuleSource;

export interface BuildSandboxSrcdocOptions {
  instanceId: string;
  extensionName: string;
  source: SandboxSrcdocSource;
  widgetApiV2Enabled: boolean;
}

export function buildSandboxSrcdoc(options: BuildSandboxSrcdocOptions): string {
    const sourceConfig = options.source.kind === "inline"
      ? { kind: "inline", code: options.source.code }
      : { kind: "module", specifier: options.source.specifier };

    const config = {
      channel: SANDBOX_CHANNEL,
      instanceId: options.instanceId,
      extensionName: options.extensionName,
      source: sourceConfig,
      widgetApiV2Enabled: options.widgetApiV2Enabled,
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
            const busyAllowed = typeof cmd.busyAllowed === "boolean" ? cmd.busyAllowed : true;
            const commandId = 'cmd-' + String(nextRequestId++);
            commandHandlers.set(commandId, cmd.handler);

            queueActivationOp(requestHost("register_command", {
              commandId,
              name: normalizedName,
              description,
              busyAllowed,
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
