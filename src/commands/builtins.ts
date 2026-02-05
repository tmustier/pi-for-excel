/**
 * Built-in slash commands for Pi for Excel.
 */

import { commandRegistry, type SlashCommand } from "./types.js";
import type { Agent } from "@mariozechner/pi-agent-core";
import { ModelSelector } from "@mariozechner/pi-web-ui";

/** Register all built-in commands. Call once after agent is created. */
export function registerBuiltins(agent: Agent): void {
  const builtins: SlashCommand[] = [
    {
      name: "model",
      description: "Change the AI model",
      source: "builtin",
      execute: () => {
        ModelSelector.open(agent.state.model, (model) => {
          agent.setModel(model);
          // Header update is handled by the agent subscriber in taskpane.ts
          document.dispatchEvent(new CustomEvent("pi:model-changed"));
        });
      },
    },
    {
      name: "default-models",
      description: "Cycle models with Ctrl+P",
      source: "builtin",
      execute: () => {
        // TODO: implement scoped models dialog
        // For now, open model selector as a placeholder
        ModelSelector.open(agent.state.model, (model) => {
          agent.setModel(model);
          document.dispatchEvent(new CustomEvent("pi:model-changed"));
        });
      },
    },
    {
      name: "settings",
      description: "Open settings",
      source: "builtin",
      execute: () => {
        // Open the API key dialog for the current provider
        import("@mariozechner/pi-web-ui").then(({ ApiKeyPromptDialog }) => {
          const provider = agent.state.model?.provider || "anthropic";
          ApiKeyPromptDialog.prompt(provider);
        });
      },
    },
    {
      name: "copy",
      description: "Copy last agent message to clipboard",
      source: "builtin",
      execute: () => {
        const msgs = agent.state.messages;
        // Find last assistant message
        for (let i = msgs.length - 1; i >= 0; i--) {
          const msg = msgs[i] as any;
          if (msg.role === "assistant") {
            const text = msg.content
              ?.filter((b: any) => b.type === "text")
              .map((b: any) => b.text)
              .join("\n") || "";
            if (text) {
              navigator.clipboard.writeText(text).then(() => {
                showToast("Copied to clipboard");
              });
            }
            return;
          }
        }
        showToast("No agent message to copy");
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
        // Session naming would be handled through SessionsStore
        // For now, store in a simple way
        document.title = args.trim();
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
        agent.clearMessages();
        document.dispatchEvent(new CustomEvent("pi:session-new"));
        showToast("New session started");
      },
    },
  ];

  for (const cmd of builtins) {
    commandRegistry.register(cmd);
  }
}

// ── Helpers ────────────────────────────────────────────────

function showToast(message: string): void {
  let toast = document.getElementById("pi-toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "pi-toast";
    toast.className = "pi-toast";
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.classList.add("visible");
  setTimeout(() => toast!.classList.remove("visible"), 2000);
}

function showShortcutsDialog(): void {
  const shortcuts = [
    ["Enter", "Send message"],
    ["Shift+Tab", "Cycle thinking level"],
    ["Esc", "Abort agent / dismiss menu"],
    ["Enter (streaming)", "Steer — redirect agent"],
    ["⌥Enter", "Queue follow-up message"],
    ["/", "Open command menu"],
    ["↑↓", "Navigate command menu"],
  ];

  let overlay = document.getElementById("pi-shortcuts-overlay");
  if (overlay) { overlay.remove(); return; }

  overlay = document.createElement("div");
  overlay.id = "pi-shortcuts-overlay";
  overlay.className = "pi-welcome-overlay";
  overlay.innerHTML = `
    <div class="pi-welcome-card" style="text-align: left;">
      <h2 style="font-size: 16px; font-weight: 600; margin: 0 0 12px; font-family: var(--font-sans);">Keyboard Shortcuts</h2>
      <div style="display: flex; flex-direction: column; gap: 6px;">
        ${shortcuts.map(([key, desc]) => `
          <div style="display: flex; justify-content: space-between; align-items: center; gap: 12px;">
            <kbd style="font-family: var(--font-mono); font-size: 11px; padding: 2px 6px; background: oklch(0 0 0 / 0.05); border-radius: 4px; white-space: nowrap;">${key}</kbd>
            <span style="font-size: 12.5px; color: var(--muted-foreground); font-family: var(--font-sans);">${desc}</span>
          </div>
        `).join("")}
      </div>
      <button onclick="this.closest('.pi-welcome-overlay').remove()" style="margin-top: 16px; width: 100%; padding: 8px; border-radius: 8px; border: 1px solid oklch(0 0 0 / 0.08); background: oklch(0 0 0 / 0.03); cursor: pointer; font-family: var(--font-sans); font-size: 13px;">Close</button>
    </div>
  `;
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay!.remove();
  });
  document.body.appendChild(overlay);
}
