/**
 * Pi for Excel — Pop-out dialog.
 *
 * Hosts the chat UI in a modeless dialog and proxies all actions
 * to the taskpane via Office dialog messaging.
 */

import "./boot.js";

import { html, render } from "lit";
import type { AgentEvent, AgentMessage, AgentState, ThinkingLevel } from "@mariozechner/pi-agent-core";
import type { Model } from "@mariozechner/pi-ai";
import { getModel } from "@mariozechner/pi-ai";
import { ChatPanel, ApiKeyPromptDialog } from "@mariozechner/pi-web-ui";

import { installFetchInterceptor } from "./auth/cors-proxy.js";
import { restoreCredentials } from "./auth/restore.js";
import { initAppStorage } from "./storage/init-app-storage.js";

// ============================================================================
// Globals
// ============================================================================

declare const Office: any;

const appEl = document.getElementById("app")!;

// ============================================================================
// Dialog bridge types
// ============================================================================

type ToolMeta = { name: string; label?: string; description?: string };

type SerializedAgentState = Omit<AgentState, "pendingToolCalls" | "tools"> & {
  pendingToolCalls: string[];
  tools: ToolMeta[];
};

// ============================================================================
// Remote agent proxy (dialog → taskpane)
// ============================================================================

class RemoteAgent {
  state: AgentState;
  private listeners = new Set<(ev: AgentEvent) => void>();

  constructor() {
    this.state = {
      systemPrompt: "",
      model: getModel("anthropic", "claude-opus-4-5"),
      thinkingLevel: "off",
      tools: [],
      messages: [],
      isStreaming: false,
      streamMessage: null,
      pendingToolCalls: new Set(),
      error: undefined,
    };
  }

  subscribe(fn: (ev: AgentEvent) => void) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  prompt(message: string | AgentMessage) {
    sendToParent({ type: "pi-dialog-prompt", message });
  }

  abort() {
    sendToParent({ type: "pi-dialog-abort" });
  }

  setModel(model: Model<any>) {
    this.state.model = model;
    sendToParent({ type: "pi-dialog-set-model", model });
    this.emit({ type: "agent_start" });
  }

  setThinkingLevel(level: ThinkingLevel) {
    this.state.thinkingLevel = level;
    sendToParent({ type: "pi-dialog-set-thinking", level });
    this.emit({ type: "agent_start" });
  }

  setTools(tools: any[]) {
    this.state.tools = tools;
  }

  applyRemoteState(state: SerializedAgentState, event?: AgentEvent | null) {
    this.state = {
      ...state,
      pendingToolCalls: new Set(state.pendingToolCalls || []),
      tools: (state.tools || []) as any,
    };

    if (event) {
      this.emit(event);
    } else {
      this.emit({ type: "agent_end", messages: this.state.messages });
    }
  }

  private emit(event: AgentEvent) {
    for (const listener of this.listeners) {
      listener(event);
    }
  }
}

function sendToParent(payload: any) {
  if (!Office?.context?.ui?.messageParent) return;
  Office.context.ui.messageParent(JSON.stringify(payload));
}

// ============================================================================
// Initialization
// ============================================================================

Office.onReady(async () => {
  installFetchInterceptor();
  const { providerKeys } = initAppStorage();
  await restoreCredentials(providerKeys);

  const agent = new RemoteAgent();

  const chatPanel = new ChatPanel();
  await chatPanel.setAgent(agent as any, {
    onApiKeyRequired: async (provider: string) => {
      return await ApiKeyPromptDialog.prompt(provider);
    },
  });

  appEl.innerHTML = "";
  render(
    html`<div class="w-full h-full flex flex-col overflow-hidden">${chatPanel}</div>`,
    appEl,
  );

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (args: any) => {
      let data: any;
      try {
        data = JSON.parse(args.message);
      } catch {
        return;
      }

      if (data.type === "pi-dialog-state") {
        agent.applyRemoteState(data.state, data.event);
      } else if (data.type === "pi-dialog-close") {
        window.close();
      }
    },
  );

  sendToParent({ type: "pi-dialog-ready" });
});
