/**
 * Built-in slash commands for Pi for Excel.
 */

import type {
  AssistantMessage,
  StopReason,
  Usage,
  UserMessage,
} from "@mariozechner/pi-ai";
import type { Agent, AgentMessage } from "@mariozechner/pi-agent-core";
import {
  ApiKeysTab,
  ModelSelector,
  ProxyTab,
  SettingsDialog,
} from "@mariozechner/pi-web-ui";

import { commandRegistry, type SlashCommand } from "./types.js";
import { showToast } from "../ui/toast.js";
import { getErrorMessage } from "../utils/errors.js";
import { extractTextBlocks, summarizeContentForTranscript } from "../utils/content.js";
import type { PiSidebar } from "../ui/pi-sidebar.js";
import { showProviderPicker, showResumeDialog, showShortcutsDialog } from "./builtins/overlays.js";

type TranscriptEntry = {
  role: AgentMessage["role"];
  text: string;
  usage?: Usage;
  stopReason?: StopReason;
};

const ZERO_USAGE: Usage = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: {
    input: 0,
    output: 0,
    cacheRead: 0,
    cacheWrite: 0,
    total: 0,
  },
};

/** Register all built-in commands. Call once after agent is created. */
export function registerBuiltins(agent: Agent): void {
  const builtins: SlashCommand[] = [
    {
      name: "model",
      description: "Change the AI model",
      source: "builtin",
      execute: () => {
        openModelSelector(agent);
      },
    },
    {
      name: "default-models",
      description: "Cycle models with Ctrl+P",
      source: "builtin",
      execute: () => {
        // TODO: implement scoped models dialog
        // For now, open model selector as a placeholder
        openModelSelector(agent);
      },
    },
    {
      name: "settings",
      description: "Settings (API keys + CORS proxy)",
      source: "builtin",
      execute: () => {
        SettingsDialog.open([new ApiKeysTab(), new ProxyTab()]);
      },
    },
    {
      name: "login",
      description: "Add or change provider API keys",
      source: "builtin",
      execute: async () => {
        await showProviderPicker();
      },
    },
    {
      name: "copy",
      description: "Copy last agent message to clipboard",
      source: "builtin",
      execute: () => {
        const msgs = agent.state.messages;
        const text = getLastAssistantText(msgs);
        if (text) {
          navigator.clipboard.writeText(text).then(() => {
            showToast("Copied to clipboard");
          });
          return;
        }
        showToast("No agent message to copy");
      },
    },
    {
      name: "export",
      description: "Export session transcript (JSON to clipboard or download)",
      source: "builtin",
      execute: (args: string) => {
        const msgs = agent.state.messages;
        if (msgs.length === 0) {
          showToast("No messages to export");
          return;
        }

        const transcript: TranscriptEntry[] = msgs.map((m) => {
          const text = summarizeContentForTranscript(m.content);
          if (m.role === "assistant") {
            return {
              role: m.role,
              text,
              usage: m.usage,
              stopReason: m.stopReason,
            };
          }
          return { role: m.role, text };
        });

        const exportData = {
          exported: new Date().toISOString(),
          model: agent.state.model
            ? {
              id: agent.state.model.id,
              name: agent.state.model.name,
              provider: agent.state.model.provider,
            }
            : null,
          thinkingLevel: agent.state.thinkingLevel,
          messageCount: msgs.length,
          transcript,
          // Also include raw messages for full fidelity debugging
          raw: msgs,
        };

        const json = JSON.stringify(exportData, null, 2);

        if (args.trim() === "clipboard" || !args.trim()) {
          navigator.clipboard.writeText(json).then(() => {
            showToast(
              `Transcript copied (${msgs.length} messages, ${(json.length / 1024).toFixed(0)}KB)`,
            );
          });
        } else {
          // Download as file
          const blob = new Blob([json], { type: "application/json" });
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = `pi-session-${new Date().toISOString().slice(0, 10)}.json`;
          a.click();
          URL.revokeObjectURL(url);
          showToast(`Downloaded transcript (${msgs.length} messages)`);
        }
      },
    },
    {
      name: "name",
      description: "Name the current chat session",
      source: "builtin",
      execute: (args: string) => {
        if (!args.trim()) {
          showToast("Usage: /name My Session Name");
          return;
        }
        document.dispatchEvent(
          new CustomEvent("pi:session-rename", { detail: { title: args.trim() } }),
        );
        showToast(`Session named: ${args.trim()}`);
      },
    },
    {
      name: "share-session",
      description: "Share session as a link",
      source: "builtin",
      execute: () => {
        showToast("Session sharing coming soon");
      },
    },
    {
      name: "shortcuts",
      description: "Show keyboard shortcuts",
      source: "builtin",
      execute: () => {
        showShortcutsDialog();
      },
    },
    {
      name: "new",
      description: "Start a new chat session",
      source: "builtin",
      execute: () => {
        // Signal new session (resets ID) then clear messages
        document.dispatchEvent(new CustomEvent("pi:session-new"));
        agent.clearMessages();

        // Force sidebar + status bar to re-render
        const sidebar = document.querySelector<PiSidebar>("pi-sidebar");
        sidebar?.requestUpdate();

        document.dispatchEvent(new CustomEvent("pi:status-update"));
        showToast("New session started");
      },
    },
    {
      name: "resume",
      description: "Resume a previous session",
      source: "builtin",
      execute: async () => {
        await showResumeDialog(agent);
      },
    },
    {
      name: "compact",
      description: "Summarize conversation to free context",
      source: "builtin",
      execute: async () => {
        const msgs = agent.state.messages;
        if (msgs.length < 4) {
          showToast("Too few messages to compact");
          return;
        }
        showToast("Compacting…");

        try {
          const { completeSimple } = await import("@mariozechner/pi-ai");

          // Serialize conversation for summarization
          const convo = conversationToText(msgs);

          const result = await completeSimple(agent.state.model, {
            systemPrompt:
              "You are a conversation summarizer. Summarize the following conversation concisely, preserving key decisions, facts, and context. Output ONLY the summary, no preamble.",
            messages: [
              {
                role: "user",
                content: [
                  {
                    type: "text",
                    text: `Summarize this conversation:\n\n${convo}`,
                  },
                ],
                timestamp: Date.now(),
              },
            ],
          });

          const summary = extractTextBlocks(result.content) || "Summary unavailable";

          const now = Date.now();
          const model = agent.state.model;

          const marker: UserMessage = {
            role: "user",
            content: [{ type: "text", text: "[This conversation was compacted]" }],
            timestamp: now,
          };

          const summaryMessage: AssistantMessage = {
            role: "assistant",
            content: [
              {
                type: "text",
                text: `**Session Summary (compacted)**\n\n${summary}`,
              },
            ],
            api: model.api,
            provider: model.provider,
            model: model.id,
            usage: ZERO_USAGE,
            stopReason: "stop",
            timestamp: now,
          };

          agent.replaceMessages([marker, summaryMessage]);

          const iface = document.querySelector<PiSidebar>("pi-sidebar");
          iface?.requestUpdate();

          showToast(`Compacted ${msgs.length} messages → summary`);
        } catch (e: unknown) {
          showToast(`Compact failed: ${getErrorMessage(e)}`);
        }
      },
    },
  ];

  for (const cmd of builtins) {
    commandRegistry.register(cmd);
  }
}

// ── Helpers ────────────────────────────────────────────────

function getLastAssistantText(messages: AgentMessage[]): string | null {
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg.role === "assistant") {
      const text = extractTextBlocks(msg.content).trim();
      return text || null;
    }
  }
  return null;
}

function conversationToText(messages: AgentMessage[]): string {
  return messages
    .map((m) => {
      const role = m.role === "user" ? "User" : m.role === "assistant" ? "Assistant" : "Tool";
      const text = extractTextBlocks(m.content);
      return `${role}: ${text}`;
    })
    .join("\n\n");
}

function openModelSelector(agent: Agent): void {
  ModelSelector.open(agent.state.model, (model) => {
    agent.setModel(model);
    // Header update is handled by the agent subscriber in taskpane.ts
    document.dispatchEvent(new CustomEvent("pi:model-changed"));
  });
}

