/**
 * Extension API for Pi for Excel.
 *
 * Extensions are ES modules that export an `activate(api)` function.
 * They run in the same webview sandbox — no Node.js APIs.
 *
 * Extensions can:
 * - Register slash commands
 * - Add custom tools for the agent
 * - Show overlay UIs (via the overlay API)
 * - Subscribe to agent events
 *
 * Example extension:
 * ```ts
 * export function activate(api: ExcelExtensionAPI) {
 *   api.registerCommand("snake", {
 *     description: "Play Snake!",
 *     handler: (args) => {
 *       api.overlay.show(createSnakeGame(api.overlay));
 *     },
 *   });
 * }
 * ```
 */

import { commandRegistry, type CommandSource } from "./types.js";
import type { Agent } from "@mariozechner/pi-agent-core";

export interface ExtensionCommand {
  description: string;
  handler: (args: string) => void | Promise<void>;
}

export interface OverlayAPI {
  /** Show an HTML element as a full-screen overlay */
  show(el: HTMLElement): void;
  /** Remove the overlay */
  dismiss(): void;
}

export interface ExcelExtensionAPI {
  /** Register a slash command */
  registerCommand(name: string, cmd: ExtensionCommand): void;
  /** Access the agent */
  agent: Agent;
  /** Show/dismiss overlay UI */
  overlay: OverlayAPI;
  /** Show a toast notification */
  toast(message: string): void;
  /** Subscribe to agent events */
  onAgentEvent(handler: (ev: any) => void): () => void;
}

/** Create the extension API for a given agent instance */
export function createExtensionAPI(agent: Agent): ExcelExtensionAPI {
  return {
    registerCommand(name: string, cmd: ExtensionCommand) {
      commandRegistry.register({
        name,
        description: cmd.description,
        source: "extension" as CommandSource,
        execute: cmd.handler,
      });
    },

    agent,

    overlay: {
      show(el: HTMLElement) {
        let container = document.getElementById("pi-ext-overlay");
        if (!container) {
          container = document.createElement("div");
          container.id = "pi-ext-overlay";
          container.className = "pi-welcome-overlay";
          container.style.zIndex = "250";
          document.body.appendChild(container);
        }
        container.innerHTML = "";
        container.appendChild(el);
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
          container.innerHTML = "";
        }
      },
    },

    toast(message: string) {
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
    },

    onAgentEvent(handler: (ev: any) => void) {
      return agent.subscribe(handler);
    },
  };
}

/**
 * Load and activate an extension from a URL or inline function.
 */
export async function loadExtension(
  api: ExcelExtensionAPI,
  source: string | ((api: ExcelExtensionAPI) => void | Promise<void>),
): Promise<void> {
  if (typeof source === "function") {
    await source(api);
  } else {
    // Dynamic import from URL
    const mod = await import(/* @vite-ignore */ source);
    if (typeof mod.activate === "function") {
      await mod.activate(api);
    } else if (typeof mod.default === "function") {
      await mod.default(api);
    }
  }
}
