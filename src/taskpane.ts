/**
 * Pi for Excel — main entry point.
 *
 * Initializes Office.js, mounts the PiSidebar,
 * wires up tools and context injection.
 */

// MUST be first — Lit fix + CSS (theme.css loaded after pi-web-ui/app.css)
import "./boot.js";

import { html, render } from "lit";
import { Agent, type AgentMessage, type ThinkingLevel } from "@mariozechner/pi-agent-core";
import { getModel, getModels, supportsXhigh, type Api, type Model } from "@mariozechner/pi-ai";
import {
  ApiKeyPromptDialog,
  ModelSelector,
  type ProviderKeysStore,
  getAppStorage,
} from "@mariozechner/pi-web-ui";

import { installFetchInterceptor } from "./auth/cors-proxy.js";
import { createOfficeStreamFn } from "./auth/stream-proxy.js";
import { restoreCredentials } from "./auth/restore.js";
import { createAllTools } from "./tools/index.js";
import { buildSystemPrompt } from "./prompt/system-prompt.js";
import { getBlueprint } from "./context/blueprint.js";
import { readSelectionContext } from "./context/selection.js";
import { ChangeTracker } from "./context/change-tracker.js";
import { initAppStorage } from "./storage/init-app-storage.js";
import { getErrorMessage } from "./utils/errors.js";
import { extractTextFromContent } from "./utils/content.js";

// UI components
import { headerStyles } from "./ui/header.js";
import { renderLoading, renderError, loadingStyles } from "./ui/loading.js";
import { showToast } from "./ui/toast.js";
import { PiSidebar } from "./ui/pi-sidebar.js";

// Slash commands + extensions
import { registerBuiltins } from "./commands/builtins.js";
import { commandRegistry } from "./commands/types.js";
import { wireCommandMenu, handleCommandMenuKey, isCommandMenuVisible, hideCommandMenu } from "./commands/command-menu.js";
import { createExtensionAPI, loadExtension } from "./commands/extension-api.js";

import { modelRecencyScore, parseMajorMinor } from "./models/model-ordering.js";
import {
  installModelSelectorPatch,
  setActiveProviders,
} from "./taskpane/model-selector-patch.js";
import { setupSessionPersistence } from "./taskpane/sessions.js";
import { createQueueDisplay } from "./taskpane/queue-display.js";

// ============================================================================
// ModelSelector patch
// ============================================================================

installModelSelectorPatch();


// ============================================================================
// Globals
// ============================================================================

function getRequiredElement<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) {
    throw new Error(`[pi] Missing required element #${id}`);
  }
  return el as T;
}

const appEl = getRequiredElement<HTMLElement>("app");
const loadingRoot = getRequiredElement<HTMLElement>("loading-root");
const errorRoot = getRequiredElement<HTMLElement>("error-root");

const changeTracker = new ChangeTracker();


// ============================================================================
// Inject component styles + render initial UI
// ============================================================================

const styleSheet = document.createElement("style");
styleSheet.textContent = headerStyles + loadingStyles;
document.head.appendChild(styleSheet);

let _agent: Agent | null = null;
let _sidebar: PiSidebar | null = null;

function openModelSelector(): void {
  const agent = _agent;
  if (!agent) return;
  ModelSelector.open(agent.state.model, (model) => {
    agent.setModel(model);
    updateStatusBar(agent);
    requestAnimationFrame(() => _sidebar?.requestUpdate());
  });
}

function showErrorBanner(message: string): void {
  render(renderError(message), errorRoot);
}

function clearErrorBanner(): void {
  render(html``, errorRoot);
}

function isLikelyCorsErrorMessage(msg: string): boolean {
  const m = msg.toLowerCase();

  // Browser/network errors
  if (m.includes("failed to fetch")) return true;
  if (m.includes("load failed")) return true; // WebKit/Safari
  if (m.includes("networkerror")) return true;

  // Explicit CORS wording
  if (m.includes("cors") || m.includes("cross-origin")) return true;

  // Anthropic sometimes returns a JSON 401 with a CORS-specific message when direct browser access is disabled.
  if (m.includes("cors requests are not allowed")) return true;

  return false;
}

render(renderLoading(), loadingRoot);


// ============================================================================
// Bootstrap
// ============================================================================

installFetchInterceptor();

let initialized = false;

Office.onReady(async (info) => {
  console.log(`[pi] Office.js ready: host=${info.host}, platform=${info.platform}`);
  try {
    initialized = true;
    await init();
  } catch (e: unknown) {
    showError(`Failed to initialize: ${getErrorMessage(e)}`);
    console.error("[pi] Init error:", e);
  }
});

setTimeout(() => {
  if (!initialized) {
    console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
    initialized = true;
    init().catch((e: unknown) => {
      showError(`Failed to initialize: ${getErrorMessage(e)}`);
      console.error("[pi] Init error:", e);
    });
  }
}, 3000);


// ============================================================================
// Initialization
// ============================================================================

async function init(): Promise<void> {
  // 1. Storage
  const { providerKeys, sessions } = initAppStorage();

  // 2. Restore auth
  await restoreCredentials(providerKeys);

  // 2b. Welcome/login if no providers
  const configuredProviders = await providerKeys.list();
  if (configuredProviders.length === 0) {
    await showWelcomeLogin(providerKeys);
  }

  // 3. Workbook blueprint
  let blueprint: string | undefined;
  try {
    blueprint = await getBlueprint();
    console.log("[pi] Workbook blueprint built");
  } catch {
    console.warn("[pi] Could not build blueprint (not in Excel?)");
  }

  // 4. Change tracker
  changeTracker.start().catch(() => {});

  // 5. Create agent
  const systemPrompt = buildSystemPrompt(blueprint);
  const availableProviders = await providerKeys.list();
  setActiveProviders(new Set(availableProviders));
  const defaultModel = pickDefaultModel(availableProviders);

  const streamFn = createOfficeStreamFn(async () => {
    try {
      const storage = getAppStorage();
      const enabled = await storage.settings.get("proxy.enabled");
      if (!enabled) return undefined;
      const url = await storage.settings.get("proxy.url");
      return typeof url === "string" && url.trim().length > 0 ? url.trim().replace(/\/+$/, "") : undefined;
    } catch {
      return undefined;
    }
  });

  const agent = _agent = new Agent({
    initialState: {
      systemPrompt,
      model: defaultModel,
      thinkingLevel: "off",
      messages: [],
      tools: createAllTools({ changeTracker }),
    },
    transformContext: async (context) => await injectContext(context),
    streamFn,
  });

  // 6. Set up API key resolution
  agent.getApiKey = async (provider: string) => {
    const key = await getAppStorage().providerKeys.get(provider);
    if (key) return key;

    // Prompt for key
    const success = await ApiKeyPromptDialog.prompt(provider);
    const updated = await providerKeys.list();
    setActiveProviders(new Set(updated));
    if (success) {
      clearErrorBanner();
      return (await getAppStorage().providerKeys.get(provider)) ?? undefined;
    } else {
      showErrorBanner(`API key required for ${provider}.`);
      return undefined;
    }
  };

  // ── Abort tracking (hoisted — used by onAbort + error handler below) ──
  let _userAborted = false;

  // 7. Create and mount PiSidebar
  const sidebar = _sidebar = new PiSidebar();
  sidebar.agent = agent;
  sidebar.emptyHints = ["Summarize this sheet", "Add a VLOOKUP formula", "Format as a table"];
  sidebar.onSend = (text) => {
    clearErrorBanner();
    agent.prompt(text).catch((e: unknown) => {
      const msg = e instanceof Error ? e.message : String(e);
      if (isLikelyCorsErrorMessage(msg)) {
        showErrorBanner("Network error (likely CORS). Start the local HTTPS proxy (npm run proxy:https) and enable it in /settings → Proxy.");
      } else {
        showErrorBanner(`LLM error: ${msg}`);
      }
    });
  };
  sidebar.onAbort = () => {
    _userAborted = true;
    agent.abort();
  };

  appEl.innerHTML = "";
  appEl.appendChild(sidebar);

  // 8. Error tracking
  agent.subscribe((ev) => {
    if (ev.type === "message_start" && ev.message.role === "user") {
      clearErrorBanner();
    }
    if (ev.type === "agent_end") {
      if (agent.state.error) {
        const isAbort = _userAborted ||
          /abort/i.test(agent.state.error) ||
          /cancel/i.test(agent.state.error);
        if (!isAbort) {
          const err = agent.state.error;
          if (isLikelyCorsErrorMessage(err)) {
            showErrorBanner("Network error (likely CORS). Start the local HTTPS proxy (npm run proxy:https) and enable it in /settings → Proxy.");
          } else {
            showErrorBanner(`LLM error: ${err}`);
          }
        }
      } else {
        clearErrorBanner();
      }
      _userAborted = false;
    }
  });

  // ── Session persistence ──
  await setupSessionPersistence({ agent, sidebar, sessions });

  // ── Register slash commands + extensions ──
  registerBuiltins(agent);
  const extensionAPI = createExtensionAPI(agent);
  const { activate: activateSnake } = await import("./extensions/snake.js");
  await loadExtension(extensionAPI, activateSnake);

  document.addEventListener("pi:providers-changed", async () => {
    const updated = await providerKeys.list();
    setActiveProviders(new Set(updated));
  });

  // ── Queue display ──
  const queueDisplay = createQueueDisplay({ agent, sidebar });

  // ── Keyboard shortcuts ──
  const THINKING_COLORS: Record<ThinkingLevel, string> = {
    off: "#a0a0a0",
    minimal: "#767676",
    low: "#4488cc",
    medium: "#22998a",
    high: "#875f87",
    xhigh: "#8b008b",
  };

  function getThinkingLevels(): ThinkingLevel[] {
    const model = agent.state.model;
    if (!model || !model.reasoning) return ["off"];

    const provider = model.provider;
    if (provider === "openai" || provider === "openai-codex") {
      const levels: ThinkingLevel[] = ["off", "minimal", "low", "medium", "high"];
      if (supportsXhigh(model)) levels.push("xhigh");
      return levels;
    }

    if (provider === "anthropic") {
      const levels: ThinkingLevel[] = ["off", "low", "medium", "high"];
      if (supportsXhigh(model)) levels.push("xhigh");
      return levels;
    }

    return ["off", "low", "medium", "high"];
  }

  document.addEventListener("keydown", (e) => {
    // Command menu takes priority
    if (isCommandMenuVisible()) {
      if (handleCommandMenuKey(e)) return;
    }

    const textarea = sidebar.getTextarea();
    const isInEditor = Boolean(textarea && (e.target === textarea || textarea.contains(e.target as Node)));
    const isStreaming = agent.state.isStreaming;

    // ESC — dismiss command menu
    if (e.key === "Escape" && isCommandMenuVisible()) {
      e.preventDefault();
      hideCommandMenu();
      return;
    }

    // ESC — abort
    if (e.key === "Escape" && isStreaming) {
      e.preventDefault();
      _userAborted = true;
      agent.abort();
      return;
    }

    // Shift+Tab — cycle thinking level
    if (e.shiftKey && e.key === "Tab") {
      e.preventDefault();
      const levels = getThinkingLevels();
      const current = agent.state.thinkingLevel;
      const idx = levels.indexOf(current);
      const next = levels[(idx >= 0 ? idx + 1 : 0) % levels.length];
      agent.setThinkingLevel(next);
      updateStatusBar(agent);
      flashThinkingLevel(next, THINKING_COLORS[next] || "#a0a0a0");
      return;
    }

    // Ctrl+O — toggle thinking/tool visibility
    if ((e.ctrlKey || e.metaKey) && e.key === "o") {
      e.preventDefault();
      const collapsed = document.body.classList.toggle("pi-hide-internals");
      showToast(collapsed ? "Details hidden (⌃O)" : "Details shown (⌃O)", 1500);
      return;
    }

    // Slash command execution
    if (isInEditor && textarea && e.key === "Enter" && !e.shiftKey && textarea.value.startsWith("/") && !isStreaming) {
      const val = textarea.value.trim();
      const spaceIdx = val.indexOf(" ");
      const cmdName = spaceIdx > 0 ? val.slice(1, spaceIdx) : val.slice(1);
      const args = spaceIdx > 0 ? val.slice(spaceIdx + 1) : "";
      const cmd = commandRegistry.get(cmdName);
      if (cmd) {
        e.preventDefault();
        e.stopImmediatePropagation();
        hideCommandMenu();
        const input = sidebar.getInput();
        if (input) input.clear();
        cmd.execute(args);
        return;
      }
    }

    // Enter/Alt+Enter while streaming — steer or follow-up
    if (isInEditor && textarea && e.key === "Enter" && !e.shiftKey && isStreaming) {
      const text = textarea.value.trim();
      if (!text) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      const msg = { role: "user" as const, content: [{ type: "text" as const, text }], timestamp: Date.now() };
      if (e.altKey) {
        agent.followUp(msg);
        queueDisplay.add("follow-up", text);
      } else {
        agent.steer(msg);
        queueDisplay.add("steer", text);
      }
      const input = sidebar.getInput();
      if (input) input.clear();
      return;
    }
  }, true);

  // ── Status bar ──
  injectStatusBar(agent);

  // ── Wire command menu to textarea ──
  const wireTextarea = () => {
    const ta = sidebar.getTextarea();
    if (ta) {
      wireCommandMenu(ta);
    } else {
      requestAnimationFrame(wireTextarea);
    }
  };
  requestAnimationFrame(wireTextarea);

  // ── Status bar click handlers ──
  document.addEventListener("click", (e) => {
    const el = e.target as HTMLElement;

    // Model picker
    if (el.closest?.(".pi-status-model")) {
      openModelSelector();
      return;
    }

    // Thinking level toggle
    if (el.closest?.(".pi-status-thinking")) {
      const levels = getThinkingLevels();
      const current = agent.state.thinkingLevel;
      const idx = levels.indexOf(current);
      const next = levels[(idx >= 0 ? idx + 1 : 0) % levels.length];
      agent.setThinkingLevel(next);
      updateStatusBar(agent);
      flashThinkingLevel(next, THINKING_COLORS[next] || "#a0a0a0");
    }
  });

  console.log("[pi] PiSidebar mounted");
}


// ============================================================================
// Context injection
// ============================================================================

function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T | null> {
  let timeoutId: number | undefined;
  const timeoutPromise = new Promise<null>((resolve) => {
    timeoutId = window.setTimeout(() => resolve(null), timeoutMs);
  });
  return Promise.race([promise, timeoutPromise]).finally(() => {
    if (timeoutId !== undefined) window.clearTimeout(timeoutId);
  }) as Promise<T | null>;
}

async function injectContext(messages: AgentMessage[]): Promise<AgentMessage[]> {
  const injections: string[] = [];
  try {
    const sel = await withTimeout(readSelectionContext().catch(() => null), 1500);
    if (sel) injections.push(sel.text);
  } catch {}
  const changes = changeTracker.flush();
  if (changes) injections.push(changes);
  if (injections.length === 0) return messages;

  const injection = injections.join("\n\n");
  const injectionMessage: AgentMessage = {
    role: "user",
    content: [{ type: "text", text: `[Auto-context]\n${injection}` }],
    timestamp: Date.now(),
  };
  const nextMessages = [...messages];
  let lastUserIdx = -1;
  for (let i = nextMessages.length - 1; i >= 0; i--) {
    if (nextMessages[i].role === "user") { lastUserIdx = i; break; }
  }
  if (lastUserIdx >= 0) {
    nextMessages.splice(lastUserIdx, 0, injectionMessage);
  } else {
    nextMessages.push(injectionMessage);
  }
  return nextMessages;
}


// ============================================================================
// Status bar
// ============================================================================

function injectStatusBar(agent: Agent): void {
  agent.subscribe(() => updateStatusBar(agent));
  document.addEventListener("pi:status-update", () => updateStatusBar(agent));
  // Initial render after sidebar mounts
  requestAnimationFrame(() => updateStatusBar(agent));
}

function updateStatusBar(agent: Agent): void {
  const el = document.getElementById("pi-status-bar");
  if (!el) return;

  const state = agent.state;

  // Model alias
  const model = state.model;
  const modelAlias = model ? (model.name || model.id) : "Select model";

  // Context usage
  let totalTokens = 0;
  for (const msg of state.messages) {
    if (msg.role !== "assistant") continue;
    totalTokens += msg.usage.input + msg.usage.output;
  }

  const contextWindow = state.model?.contextWindow || 200000;
  const pct = Math.min(100, Math.round((totalTokens / contextWindow) * 100));
  const ctxLabel = contextWindow >= 1_000_000
    ? `${(contextWindow / 1_000_000).toFixed(0)}M`
    : `${Math.round(contextWindow / 1000)}k`;

  // Thinking level
  const thinkingLabels: Record<string, string> = {
    off: "off", minimal: "min", low: "low", medium: "med", high: "high", xhigh: "max",
  };
  const thinkingLevel = thinkingLabels[state.thinkingLevel] || state.thinkingLevel;

  const chevronSvg = `<svg width="8" height="8" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>`;
  const brainSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 18V5"/><path d="M15 13a4.17 4.17 0 0 1-3-4 4.17 4.17 0 0 1-3 4"/><path d="M17.598 6.5A3 3 0 1 0 12 5a3 3 0 1 0-5.598 1.5"/><path d="M17.997 5.125a4 4 0 0 1 2.526 5.77"/><path d="M18 18a4 4 0 0 0 2-7.464"/><path d="M19.967 17.483A4 4 0 1 1 12 18a4 4 0 1 1-7.967-.517"/><path d="M6 18a4 4 0 0 1-2-7.464"/><path d="M6.003 5.125a4 4 0 0 0-2.526 5.77"/></svg>`;

  el.innerHTML = `
    <button class="pi-status-model" title="Click to change model">
      <span class="pi-status-model__mark">π</span>
      <span class="pi-status-model__name">${modelAlias}</span>
      ${chevronSvg}
    </button>
    <span class="pi-status-ctx" title="Context window usage">${pct}% / ${ctxLabel}</span>
    <span class="pi-status-thinking" title="Click or ⇧Tab to cycle thinking depth">${brainSvg} ${thinkingLevel}</span>
  `;
}


// ============================================================================
// Thinking level flash
// ============================================================================

function flashThinkingLevel(level: string, color: string): void {
  const labels: Record<string, string> = {
    off: "Off",
    minimal: "Min",
    low: "Low",
    medium: "Medium",
    high: "High",
    xhigh: "Max",
  };
  showToast(`Thinking: ${labels[level] || level} (next turn)`, 1500);

  const el = document.querySelector(".pi-status-thinking") as HTMLElement;
  if (!el) return;

  el.style.color = color;
  el.style.background = `${color}18`;
  el.style.boxShadow = `0 0 8px ${color}40`;
  el.style.transition = "none";

  let flashBar = document.getElementById("pi-thinking-flash");
  if (!flashBar) {
    flashBar = document.createElement("div");
    flashBar.id = "pi-thinking-flash";
    flashBar.style.cssText = `
      position: fixed; bottom: 0; left: 0; right: 0; height: 2px;
      pointer-events: none; z-index: 100; transition: opacity 0.6s ease-out;
    `;
    document.body.appendChild(flashBar);
  }
  flashBar.style.background = `linear-gradient(90deg, transparent, ${color}, transparent)`;
  flashBar.style.opacity = "1";

  const bar = flashBar;

  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      el.style.transition = "color 0.8s ease, background 0.8s ease, box-shadow 0.8s ease";
      el.style.color = "";
      el.style.background = "";
      el.style.boxShadow = "";
      bar.style.opacity = "0";
    });
  });
}


// ============================================================================
// Welcome / Login
// ============================================================================

async function showWelcomeLogin(providerKeys: ProviderKeysStore): Promise<void> {
  const { ALL_PROVIDERS, buildProviderRow } = await import("./ui/provider-login.js");

  return new Promise<void>((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "pi-welcome-overlay";
    overlay.innerHTML = `
      <div class="pi-welcome-card" style="text-align: left;">
        <div class="pi-welcome-logo" style="text-align: center;">π</div>
        <h2 class="pi-welcome-title" style="text-align: center;">Pi for Excel</h2>
        <p class="pi-welcome-subtitle" style="text-align: center;">Connect a provider to get started</p>
        <div class="pi-welcome-providers"></div>
      </div>
    `;
    const providerList = overlay.querySelector<HTMLDivElement>(".pi-welcome-providers");
    if (!providerList) {
      throw new Error("Welcome provider list not found");
    }
    const expandedRef = { current: null as HTMLElement | null };
    for (const provider of ALL_PROVIDERS) {
      const row = buildProviderRow(provider, {
        isActive: false,
        expandedRef,
        onConnected: async (_row, _id, label) => {
          const updated = await providerKeys.list();
          setActiveProviders(new Set(updated));
          document.dispatchEvent(new CustomEvent("pi:providers-changed"));
          showToast(`${label} connected`);
          overlay.remove();
          resolve();
        },
      });
      providerList.appendChild(row);
    }
    document.body.appendChild(overlay);
  });
}


// ============================================================================
// Default model selection
// ============================================================================

type DefaultProvider = "openai-codex" | "openai" | "google";

type DefaultModelRule = { provider: DefaultProvider; match: RegExp };

const DEFAULT_MODEL_RULES: DefaultModelRule[] = [
  // Prefer latest GPT-5.x Codex on ChatGPT subscription (openai-codex)
  { provider: "openai-codex", match: /^gpt-5\.(\d+)-codex$/ },
  { provider: "openai-codex", match: /^gpt-5\./ },

  // API key OpenAI provider (if user connected OpenAI instead of openai-codex)
  { provider: "openai", match: /^gpt-5\.(\d+)-codex$/ },
  { provider: "openai", match: /^gpt-5\./ },

  // Gemini defaults: Pro-ish first, then any Gemini
  { provider: "google", match: /^gemini-.*-pro/i },
  { provider: "google", match: /^gemini-/i },
];

function pickLatestMatchingModel(
  provider: DefaultProvider,
  match: RegExp,
): Model<Api> | null {
  const models: Model<Api>[] = getModels(provider);
  const candidates = models.filter((m) => match.test(m.id));
  candidates.sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id));
  return candidates[0] ?? null;
}

function pickDefaultModel(availableProviders: string[]): Model<Api> {
  // Anthropic special-case:
  // Prefer Opus, except if there's a *newer-version* Sonnet, use that first.
  if (availableProviders.includes("anthropic")) {
    const models: Model<Api>[] = getModels("anthropic");
    const opus = models
      .filter((m) => m.id.startsWith("claude-opus-"))
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];
    const sonnet = models
      .filter((m) => m.id.startsWith("claude-sonnet-"))
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];

    if (opus && sonnet) {
      return parseMajorMinor(sonnet.id) > parseMajorMinor(opus.id) ? sonnet : opus;
    }

    if (opus) return opus;
    if (sonnet) return sonnet;
  }

  // Other providers: pattern-based rules
  for (const rule of DEFAULT_MODEL_RULES) {
    if (!availableProviders.includes(rule.provider)) continue;
    const m = pickLatestMatchingModel(rule.provider, rule.match);
    if (m) return m;
  }

  // Absolute fallback: keep this resilient across pi-ai version bumps
  return getModel("anthropic", "claude-opus-4-5");
}


// ============================================================================
// Error display
// ============================================================================

function showError(message: string): void {
  render(renderError(message), errorRoot);
}
