/**
 * Pi for Excel — main entry point.
 *
 * Initializes Office.js, mounts the ChatPanel sidebar,
 * wires up tools and context injection.
 */

// MUST be first — Lit fix + CSS (theme.css loaded after pi-web-ui/app.css)
import "./boot.js";

import { html, render } from "lit";
import { Agent } from "@mariozechner/pi-agent-core";
import { getModel } from "@mariozechner/pi-ai";
import {
  ChatPanel,
  AppStorage,
  IndexedDBStorageBackend,
  ProviderKeysStore,
  CustomProvidersStore,
  SessionsStore,
  SettingsStore,
  setAppStorage,
  ApiKeyPromptDialog,
} from "@mariozechner/pi-web-ui";

import { installFetchInterceptor } from "./auth/cors-proxy.js";
import { restoreCredentials } from "./auth/restore.js";
import { createAllTools } from "./tools/index.js";
import { buildSystemPrompt } from "./prompt/system-prompt.js";
import { getBlueprint } from "./context/blueprint.js";
import { readSelectionContext } from "./context/selection.js";
import { ChangeTracker } from "./context/change-tracker.js";

// UI components — extracted for easy swapping
import { renderHeader, headerStyles } from "./ui/header.js";
import { renderLoading, renderError, loadingStyles } from "./ui/loading.js";

// ============================================================================
// Globals
// ============================================================================

declare const Office: any;

const headerRoot = document.getElementById("header-root")!;
const appEl = document.getElementById("app")!;
const loadingRoot = document.getElementById("loading-root")!;
const errorRoot = document.getElementById("error-root")!;

const changeTracker = new ChangeTracker();

// ============================================================================
// Inject component styles + render initial UI
// ============================================================================

const styleSheet = document.createElement("style");
styleSheet.textContent = headerStyles + loadingStyles;
document.head.appendChild(styleSheet);

// Render header and loading state immediately
function updateHeader(opts: { status?: "ready" | "working" | "error"; modelAlias?: string } = {}) {
  render(renderHeader({
    status: opts.status || "ready",
    modelAlias: opts.modelAlias,
    onModelClick: () => {
      // Trigger the model selector by clicking the (now hidden) model button in toolbar
      const modelBtn = document.querySelector("message-editor .px-2.pb-2 > .flex.gap-2:last-child > button:first-child") as HTMLElement;
      if (modelBtn) modelBtn.click();
    },
  }), headerRoot);
}
updateHeader();
render(renderLoading(), loadingRoot);

// ============================================================================
// Bootstrap
// ============================================================================

// Install CORS proxy before any fetch calls
installFetchInterceptor();

let initialized = false;

// Wait for Office.js
Office.onReady(async (info: { host: any; platform: any }) => {
  console.log(`[pi] Office.js ready: host=${info.host}, platform=${info.platform}`);

  try {
    initialized = true;
    await init();
  } catch (e: any) {
    showError(`Failed to initialize: ${e.message}`);
    console.error("[pi] Init error:", e);
  }
});

// Fallback if not in Excel (dev/testing)
setTimeout(() => {
  if (!initialized) {
    console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
    initialized = true;
    init().catch((e) => {
      showError(`Failed to initialize: ${e.message}`);
      console.error("[pi] Init error:", e);
    });
  }
}, 3000);

// ============================================================================
// Initialization
// ============================================================================

async function init(): Promise<void> {
  // 1. Set up storage
  const settings = new SettingsStore();
  const providerKeys = new ProviderKeysStore();
  const sessions = new SessionsStore();
  const customProviders = new CustomProvidersStore();

  const backend = new IndexedDBStorageBackend({
    dbName: "pi-for-excel",
    version: 1,
    stores: [
      settings.getConfig(),
      providerKeys.getConfig(),
      sessions.getConfig(),
      SessionsStore.getMetadataConfig(),
      customProviders.getConfig(),
    ],
  });

  settings.setBackend(backend);
  providerKeys.setBackend(backend);
  sessions.setBackend(backend);
  customProviders.setBackend(backend);

  const storage = new AppStorage(settings, providerKeys, sessions, customProviders, backend);
  setAppStorage(storage);

  // 2. Restore auth credentials
  await restoreCredentials(providerKeys);

  // 3. Build initial workbook blueprint
  let blueprint: string | undefined;
  try {
    blueprint = await getBlueprint();
    console.log("[pi] Workbook blueprint built");
  } catch {
    console.warn("[pi] Could not build blueprint (not in Excel?)");
  }

  // 4. Start change tracker
  changeTracker.start().catch(() => {});

  // 5. Create agent
  const systemPrompt = buildSystemPrompt(blueprint);

  const agent = new Agent({
    initialState: {
      systemPrompt,
      model: getModel("anthropic", "claude-opus-4-5"),
      thinkingLevel: "off",
      messages: [],
      tools: [],
    },
    transformContext: async (context) => {
      return await injectContext(context);
    },
  });

  // 6. Create and mount ChatPanel
  const chatPanel = new ChatPanel();

  await chatPanel.setAgent(agent, {
    onApiKeyRequired: async (provider: string) => {
      return await ApiKeyPromptDialog.prompt(provider);
    },
    toolsFactory: () => createAllTools(),
  });

  // 7. Clear loading, mount ChatPanel with empty state
  appEl.innerHTML = "";
  render(
    html`
      <div class="w-full h-full flex flex-col overflow-hidden"
           style="background: var(--background); color: var(--foreground); position: relative;">
        <!-- Empty state — shown when no messages -->
        <div id="empty-state">
          <div class="empty-logo">π</div>
          <div class="empty-tagline">
            Your AI assistant for Excel.<br/>Ask anything about your spreadsheet.
          </div>
          <div class="empty-hints">
            ${["Summarize this sheet", "Add a VLOOKUP formula", "Format as a table"].map(
              (hint) => html`
                <button
                  class="empty-hint"
                  @click=${() => {
                    const iface = document.querySelector("agent-interface") as any;
                    if (iface?.sendMessage) iface.sendMessage(hint);
                    document.getElementById("empty-state")?.classList.add("hidden");
                  }}
                >${hint}</button>
              `,
            )}
          </div>
        </div>
        ${chatPanel}
      </div>
    `,
    appEl,
  );

  // Update header with model name + status; hide empty state when messages arrive
  const emptyState = document.getElementById("empty-state");
  const getModelAlias = () => {
    const m = agent.state.model;
    return m ? (m.name || m.id) : undefined;
  };
  updateHeader({ modelAlias: getModelAlias() });

  agent.subscribe((ev) => {
    // Header status
    if (ev.type === "message_start") {
      updateHeader({ status: "working", modelAlias: getModelAlias() });
    } else if (ev.type === "message_end") {
      updateHeader({ status: "ready", modelAlias: getModelAlias() });
    }

    // Empty state
    if (ev.type === "message_start" || ev.type === "message_end") {
      if (agent.state.messages.length > 0 && emptyState) {
        emptyState.classList.add("hidden");
      }
    }
  });

  // ── Keyboard shortcuts ──────────────────────────────────────────
  const THINKING_LEVELS = ["off", "low", "medium", "high"] as const;
  const THINKING_COLORS: Record<string, string> = {
    off: "#a0a0a0",       // grey
    low: "#4488cc",       // blue
    medium: "#22998a",    // teal
    high: "#875f87",      // purple
  };

  // We capture on the document to intercept before MessageEditor's handler
  document.addEventListener("keydown", (e) => {
    const textarea = document.querySelector("message-editor textarea") as HTMLTextAreaElement | null;
    const isInEditor = textarea && (e.target === textarea || textarea.contains(e.target as Node));
    const isStreaming = agent.state.isStreaming;

    // ESC — abort (MessageEditor already handles this, but we add it globally too)
    if (e.key === "Escape" && isStreaming) {
      e.preventDefault();
      agent.abort();
      return;
    }

    // Shift+Tab — cycle thinking level
    if (e.shiftKey && e.key === "Tab") {
      e.preventDefault();
      const current = agent.state.thinkingLevel;
      const idx = THINKING_LEVELS.indexOf(current as any);
      const next = THINKING_LEVELS[(idx + 1) % THINKING_LEVELS.length];
      agent.setThinkingLevel(next);
      const iface = document.querySelector("agent-interface") as any;
      if (iface) iface.requestUpdate();
      updateStatusBar(agent);
      flashThinkingLevel(next, THINKING_COLORS[next] || "#a0a0a0");
      return;
    }

    // Enter/Alt+Enter in textarea while streaming — steer or follow-up
    if (isInEditor && e.key === "Enter" && !e.shiftKey && isStreaming) {
      const text = textarea!.value.trim();
      if (!text) return;

      e.preventDefault();
      e.stopImmediatePropagation(); // prevent MessageEditor's handler

      const msg = { role: "user" as const, content: [{ type: "text" as const, text }], timestamp: Date.now() };

      if (e.altKey) {
        // Alt+Enter → follow-up (queued for after agent finishes current turn)
        agent.followUp(msg);
      } else {
        // Enter → steer (interrupts current turn)
        agent.steer(msg);
      }

      // Clear the textarea
      textarea!.value = "";
      textarea!.dispatchEvent(new Event("input", { bubbles: true }));
      return;
    }

    // Alt+Enter when NOT streaming — also queue follow-up for after next response
    if (isInEditor && e.key === "Enter" && e.altKey && !isStreaming) {
      const text = textarea!.value.trim();
      if (!text) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      const msg = { role: "user" as const, content: [{ type: "text" as const, text }], timestamp: Date.now() };
      agent.followUp(msg);
      textarea!.value = "";
      textarea!.dispatchEvent(new Event("input", { bubbles: true }));
      return;
    }
  }, true); // capture phase — fires before MessageEditor

  // Custom status bar — shows context % and thinking level
  injectStatusBar(agent);

  // Auto-resize textarea — override pi-web-ui's inline max-height: 200px
  const patchTextarea = () => {
    const ta = document.querySelector("message-editor textarea") as HTMLTextAreaElement | null;
    if (ta) {
      ta.style.maxHeight = "50vh";
      ta.style.height = "auto";
      // Classic auto-grow: reset height then set to scrollHeight
      const autoGrow = () => {
        ta.style.height = "auto";
        ta.style.height = ta.scrollHeight + "px";
      };
      ta.addEventListener("input", autoGrow);
      // Also observe value changes (e.g. clearing after send)
      new MutationObserver(autoGrow).observe(ta, { attributes: true, attributeFilter: ["value"] });
    } else {
      requestAnimationFrame(patchTextarea);
    }
  };
  requestAnimationFrame(patchTextarea);

  // Dynamic placeholder — changes during streaming to hint at steer/follow-up
  agent.subscribe((ev) => {
    if (ev.type === "message_start" || ev.type === "message_end") {
      const textarea = document.querySelector("message-editor textarea") as HTMLTextAreaElement | null;
      if (textarea) {
        textarea.placeholder = agent.state.isStreaming
          ? "Steer (Enter) · Follow-up (⌥Enter)…"
          : "Type a message…";
      }
    }
  });

  // Make status bar thinking indicator clickable
  document.addEventListener("click", (e) => {
    const target = (e.target as HTMLElement).closest?.(".pi-status-thinking");
    if (target) {
      const current = agent.state.thinkingLevel;
      const idx = THINKING_LEVELS.indexOf(current as any);
      const next = THINKING_LEVELS[(idx + 1) % THINKING_LEVELS.length];
      agent.setThinkingLevel(next);
      const iface = document.querySelector("agent-interface") as any;
      if (iface) iface.requestUpdate();
      updateStatusBar(agent);
      flashThinkingLevel(next, THINKING_COLORS[next] || "#a0a0a0");
    }
  });

  console.log("[pi] ChatPanel mounted");
}

// ============================================================================
// Context injection — runs before every LLM call
// ============================================================================

async function injectContext(context: any): Promise<any> {
  const injections: string[] = [];

  // 1. Selection context (what the user is looking at)
  try {
    const sel = await readSelectionContext();
    if (sel) {
      injections.push(sel.text);
    }
  } catch {
    // Not in Excel or selection read failed — skip
  }

  // 2. User changes since last message
  const changes = changeTracker.flush();
  if (changes) {
    injections.push(changes);
  }

  // If nothing to inject, return context unchanged
  if (injections.length === 0) return context;

  // Inject as a system message before the last user message
  const injection = injections.join("\n\n");
  const injectionMessage = {
    role: "user" as const,
    content: [{ type: "text" as const, text: `[Auto-context]\n${injection}` }],
  };

  // Find the last user message and insert before it
  const messages = [...context.messages];
  let lastUserIdx = -1;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === "user") {
      lastUserIdx = i;
      break;
    }
  }

  if (lastUserIdx >= 0) {
    messages.splice(lastUserIdx, 0, injectionMessage);
  } else {
    messages.push(injectionMessage);
  }

  return { ...context, messages };
}

// ============================================================================
// Status bar — context % + thinking level
// ============================================================================

function injectStatusBar(agent: Agent): void {
  // Find the stats area rendered by AgentInterface (below the editor)
  // We'll replace it with our own content via a MutationObserver
  const bar = document.createElement("div");
  bar.id = "pi-status-bar";
  bar.className = "pi-status-bar";

  // Initial render
  updateStatusBar(agent, bar);

  // Update on agent events
  agent.subscribe(() => updateStatusBar(agent, bar));

  // Insert after the message editor area
  const tryInsert = () => {
    const editorWrap = document.querySelector("agent-interface .shrink-0 .max-w-3xl");
    if (editorWrap && !document.getElementById("pi-status-bar")) {
      editorWrap.appendChild(bar);
    } else {
      requestAnimationFrame(tryInsert);
    }
  };
  requestAnimationFrame(tryInsert);
}

function updateStatusBar(agent: Agent, bar?: HTMLElement): void {
  const el = bar || document.getElementById("pi-status-bar");
  if (!el) return;

  const state = agent.state;

  // Compute token usage from messages
  let totalTokens = 0;
  for (const msg of state.messages) {
    const usage = (msg as any).usage;
    if (usage) {
      totalTokens += (usage.input || 0) + (usage.output || 0);
    }
  }

  const contextWindow = state.model?.contextWindow || 200000;
  const pct = Math.min(100, Math.round((totalTokens / contextWindow) * 100));
  const ctxLabel = contextWindow >= 1_000_000
    ? `${(contextWindow / 1_000_000).toFixed(0)}M`
    : `${Math.round(contextWindow / 1000)}k`;

  // Thinking level display
  const thinkingLabels: Record<string, string> = {
    off: "off", minimal: "min", low: "low", medium: "med", high: "high", xhigh: "max",
  };
  const thinkingLevel = thinkingLabels[state.thinkingLevel] || state.thinkingLevel;

  const brainSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 18V5"/><path d="M15 13a4.17 4.17 0 0 1-3-4 4.17 4.17 0 0 1-3 4"/><path d="M17.598 6.5A3 3 0 1 0 12 5a3 3 0 1 0-5.598 1.5"/><path d="M17.997 5.125a4 4 0 0 1 2.526 5.77"/><path d="M18 18a4 4 0 0 0 2-7.464"/><path d="M19.967 17.483A4 4 0 1 1 12 18a4 4 0 1 1-7.967-.517"/><path d="M6 18a4 4 0 0 1-2-7.464"/><path d="M6.003 5.125a4 4 0 0 0-2.526 5.77"/></svg>`;

  el.innerHTML = `
    <span class="pi-status-ctx">${pct}% / ${ctxLabel}</span>
    <span class="pi-status-thinking" title="Shift+Tab to cycle">${brainSvg} ${thinkingLevel}</span>
  `;
}

// ============================================================================
// Thinking level flash — visual feedback on change
// ============================================================================

function flashThinkingLevel(level: string, color: string): void {
  // Flash the status bar thinking indicator
  const el = document.querySelector(".pi-status-thinking") as HTMLElement;
  if (!el) return;

  // Apply color + pulse animation
  el.style.color = color;
  el.style.background = `${color}18`; // 10% opacity fill
  el.style.boxShadow = `0 0 8px ${color}40`;
  el.style.transition = "none";

  // Also flash a thin bar at the bottom of the input area
  let flashBar = document.getElementById("pi-thinking-flash");
  if (!flashBar) {
    flashBar = document.createElement("div");
    flashBar.id = "pi-thinking-flash";
    flashBar.style.cssText = `
      position: fixed; bottom: 0; left: 0; right: 0; height: 2px;
      pointer-events: none; z-index: 100;
      transition: opacity 0.6s ease-out;
    `;
    document.body.appendChild(flashBar);
  }
  flashBar.style.background = `linear-gradient(90deg, transparent, ${color}, transparent)`;
  flashBar.style.opacity = "1";

  // Fade out
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      el.style.transition = "color 0.8s ease, background 0.8s ease, box-shadow 0.8s ease";
      el.style.color = "";
      el.style.background = "";
      el.style.boxShadow = "";
      flashBar!.style.opacity = "0";
    });
  });
}

// ============================================================================
// Error display
// ============================================================================

function showError(message: string): void {
  render(renderError(message), errorRoot);
}
