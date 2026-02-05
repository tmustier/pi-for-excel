/**
 * Pi for Excel — main entry point.
 *
 * Initializes Office.js, mounts the PiSidebar,
 * wires up tools and context injection.
 */

// MUST be first — Lit fix + CSS (theme.css loaded after pi-web-ui/app.css)
import "./boot.js";

import { html, render } from "lit";
import { Agent } from "@mariozechner/pi-agent-core";
import { getModel, getModels, supportsXhigh } from "@mariozechner/pi-ai";
import {
  ApiKeyPromptDialog,
  ModelSelector,
  type ProviderKeysStore,
  getAppStorage,
} from "@mariozechner/pi-web-ui";

import { installFetchInterceptor } from "./auth/cors-proxy.js";
import { restoreCredentials } from "./auth/restore.js";
import { createAllTools } from "./tools/index.js";
import { buildSystemPrompt } from "./prompt/system-prompt.js";
import { getBlueprint } from "./context/blueprint.js";
import { readSelectionContext } from "./context/selection.js";
import { ChangeTracker } from "./context/change-tracker.js";
import { initAppStorage } from "./storage/init-app-storage.js";

// UI components
import { renderHeader, headerStyles } from "./ui/header.js";
import { renderLoading, renderError, loadingStyles } from "./ui/loading.js";
import { showToast } from "./ui/toast.js";
import { PiSidebar } from "./ui/pi-sidebar.js";

// Slash commands + extensions
import { registerBuiltins } from "./commands/builtins.js";
import { commandRegistry } from "./commands/types.js";
import { wireCommandMenu, handleCommandMenuKey, isCommandMenuVisible, hideCommandMenu } from "./commands/command-menu.js";
import { createExtensionAPI, loadExtension } from "./commands/extension-api.js";

// ============================================================================
// Patch ModelSelector to only show models from providers with API keys
// ============================================================================

let _activeProviders: Set<string> | null = null;

export function setActiveProviders(providers: Set<string>) {
  _activeProviders = providers;
}

const PROVIDER_ORDER = new Map<string, number>([
  ["anthropic", 1],
  ["openai-codex", 2],
  ["openai", 3],
  ["google", 4],
  ["github-copilot", 5],
]);

function familyPriority(provider: string, id: string): number {
  if (provider === "anthropic") {
    if (id.startsWith("claude-opus-")) return 0;
    if (id.startsWith("claude-sonnet-")) return 1;
    if (id.startsWith("claude-haiku-")) return 2;
    return 9;
  }

  if (provider === "openai-codex" || provider === "openai") {
    if (id.includes("codex")) return 0;
    if (id.startsWith("gpt-")) return 1;
    if (id.startsWith("o")) return 2;
    return 9;
  }

  if (provider === "google") {
    // Prefer Pro-ish variants first, then Flash-ish, then any Gemini.
    if (/^gemini-.*-pro/i.test(id)) return 0;
    if (/^gemini-.*-flash/i.test(id)) return 1;
    if (id.includes("gemini")) return 2;
    return 9;
  }

  return 9;
}

function parseMajorMinor(id: string): number {
  // Extract a comparable major/minor number from common model ID formats.
  // Important: don't misinterpret 8-digit date suffixes (e.g. 20250514) as "minor".
  // Examples:
  // - claude-opus-4-5           -> 45
  // - claude-opus-4-6           -> 46
  // - claude-opus-4-20250514    -> 40 (major only; date handled separately)
  // - gpt-5.3-codex             -> 53
  // - gemini-2.5-pro            -> 25
  // - gemini-3-pro-preview      -> 30

  const pack = (major: number, minor: number | null): number => {
    if (minor === null) return major * 10;
    // minor < 10 => major*10 + minor (4.6 -> 46)
    if (minor < 10) return major * 10 + minor;
    // allow 2-digit minors (e.g. 4.12 -> 412)
    return major * 100 + minor;
  };

  // Claude-style: -4-6 (but NOT -4-20250514)
  const hyphenVer = id.match(/-(\d+)-(\d{1,2})(?:-|$)/);
  if (hyphenVer) {
    return pack(parseInt(hyphenVer[1], 10), parseInt(hyphenVer[2], 10));
  }

  // OpenAI/Gemini-style: 5.3 / 2.5
  const dotVer = id.match(/(\d+)\.(\d{1,2})/);
  if (dotVer) {
    return pack(parseInt(dotVer[1], 10), parseInt(dotVer[2], 10));
  }

  // Fallback: first major number after hyphen
  const majorMatch = id.match(/-(\d+)(?:-|$)/);
  if (majorMatch) {
    return pack(parseInt(majorMatch[1], 10), null);
  }

  return 0;
}

function modelRecencyScore(id: string): number {
  // Prefer higher major/minor first, then higher date suffix.
  const majorMinor = parseMajorMinor(id);

  let date = 0;
  const dateMatch = id.match(/(\d{8})/);
  if (dateMatch) date = parseInt(dateMatch[1], 10);

  // date is at most 8 digits → multiplier must exceed that range
  return majorMinor * 100_000_000 + date;
}

const _origGetFilteredModels = (ModelSelector.prototype as any).getFilteredModels;
(ModelSelector.prototype as any).getFilteredModels = function () {
  const all: Array<{ provider: string; id: string; model: any }> = _origGetFilteredModels.call(this);
  let filtered = all;
  if (_activeProviders && _activeProviders.size > 0) {
    filtered = all.filter((m: any) => _activeProviders!.has(m.provider));
  }

  const currentModel = this.currentModel;

  const isCurrent = (x: any) =>
    currentModel && x.model.id === currentModel.id && x.model.provider === currentModel.provider;

  const keyOf = (x: any) => `${x.provider}:${x.id}`;

  const compareModels = (a: any, b: any) => {
    const aProv = PROVIDER_ORDER.get(a.provider) ?? 999;
    const bProv = PROVIDER_ORDER.get(b.provider) ?? 999;
    if (aProv !== bProv) return aProv - bProv;

    const aFam = familyPriority(a.provider, a.id);
    const bFam = familyPriority(b.provider, b.id);
    if (aFam !== bFam) return aFam - bFam;

    const aRec = modelRecencyScore(a.id);
    const bRec = modelRecencyScore(b.id);
    if (aRec !== bRec) return bRec - aRec;

    return a.id.localeCompare(b.id);
  };

  // "Latest for each" behavior:
  // - keep current model at the very top
  // - then show "featured" models (latest per provider, pattern-based)
  //   - Anthropic: latest Sonnet if its version > latest Opus, then latest Opus
  //   - OpenAI Codex: latest gpt-5.x-codex, then latest gpt-5.x
  //   - Google: latest gemini-*-pro*
  // - then show the remaining models, sorted deterministically

  const byProvider = new Map<string, any[]>();
  for (const m of filtered) {
    const list = byProvider.get(m.provider) || [];
    list.push(m);
    byProvider.set(m.provider, list);
  }

  const providers = Array.from(byProvider.keys()).sort((a, b) => {
    const aProv = PROVIDER_ORDER.get(a) ?? 999;
    const bProv = PROVIDER_ORDER.get(b) ?? 999;
    if (aProv !== bProv) return aProv - bProv;
    return a.localeCompare(b);
  });

  const pickBest = (models: any[], filter?: (m: any) => boolean) => {
    const list = filter ? models.filter(filter) : models;
    if (!list.length) return null;
    return list
      .slice()
      .sort((a, b) => {
        const aFam = familyPriority(a.provider, a.id);
        const bFam = familyPriority(b.provider, b.id);
        if (aFam !== bFam) return aFam - bFam;
        const aRec = modelRecencyScore(a.id);
        const bRec = modelRecencyScore(b.id);
        if (aRec !== bRec) return bRec - aRec;
        return a.id.localeCompare(b.id);
      })[0];
  };

  const pickBestByRecency = (models: any[], filter: (m: any) => boolean) => {
    const list = models.filter(filter);
    if (!list.length) return null;
    return list
      .slice()
      .sort((a, b) => {
        const aRec = modelRecencyScore(a.id);
        const bRec = modelRecencyScore(b.id);
        if (aRec !== bRec) return bRec - aRec;
        return a.id.localeCompare(b.id);
      })[0];
  };

  const featured: any[] = [];
  for (const provider of providers) {
    const models = byProvider.get(provider) || [];
    if (!models.length) continue;

    // Provider-specific "latest" rules
    if (provider === "anthropic") {
      const bestOpus = pickBestByRecency(models, (m) => String(m.id).startsWith("claude-opus-"));
      const bestSonnet = pickBestByRecency(models, (m) => String(m.id).startsWith("claude-sonnet-"));

      if (bestOpus && bestSonnet) {
        const opusVer = parseMajorMinor(bestOpus.id);
        const sonnetVer = parseMajorMinor(bestSonnet.id);
        if (sonnetVer > opusVer) {
          featured.push(bestSonnet);
          featured.push(bestOpus);
          continue;
        }
        featured.push(bestOpus);
        continue;
      }

      if (bestOpus) {
        featured.push(bestOpus);
        continue;
      }

      if (bestSonnet) {
        featured.push(bestSonnet);
        continue;
      }

      const best = pickBest(models);
      if (best) featured.push(best);
      continue;
    }

    if (provider === "openai-codex") {
      const bestCodex = pickBestByRecency(models, (m) => /^gpt-5\.(\d+)-codex$/.test(String(m.id)));
      const bestGpt5 = pickBestByRecency(models, (m) => /^gpt-5\./.test(String(m.id)) && !/codex/.test(String(m.id)));

      if (bestCodex) featured.push(bestCodex);
      if (bestGpt5) featured.push(bestGpt5);
      if (bestCodex || bestGpt5) continue;

      const best = pickBest(models);
      if (best) featured.push(best);
      continue;
    }

    if (provider === "google") {
      const bestPro = pickBestByRecency(models, (m) => /^gemini-.*-pro/i.test(String(m.id)));
      if (bestPro) {
        featured.push(bestPro);
        continue;
      }

      const best = pickBest(models);
      if (best) featured.push(best);
      continue;
    }

    // Generic fallback
    const best = pickBest(models);
    if (best) featured.push(best);
  }

  const out: any[] = [];
  const used = new Set<string>();
  const push = (m: any) => {
    const k = keyOf(m);
    if (used.has(k)) return;
    used.add(k);
    out.push(m);
  };

  // Current model first (if it's in the filtered list)
  for (const m of filtered) {
    if (isCurrent(m)) push(m);
  }

  // Then latest-for-each-provider
  for (const m of featured) {
    push(m);
  }

  // Then the remaining models
  const remaining = filtered.filter((m: any) => !used.has(keyOf(m)));
  remaining.sort(compareModels);
  for (const m of remaining) push(m);

  return out;
};


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

let _agent: Agent | null = null;
let _sidebar: PiSidebar | null = null;
let _headerState: { status: "ready" | "working" | "error"; modelAlias?: string } = {
  status: "ready",
};

function getAgentModelAlias(agent: Agent | null): string | undefined {
  const m = agent?.state.model;
  return m ? (m.name || m.id) : undefined;
}

function setModelAndSync(agent: Agent, model: any): void {
  agent.setModel(model);
  updateHeader({ modelAlias: getAgentModelAlias(agent) });
  updateStatusBar(agent);
  // Ensure sidebar reacts to model capability changes (thinking levels, etc.)
  requestAnimationFrame(() => _sidebar?.requestUpdate());
}

function updateHeader(opts: { status?: "ready" | "working" | "error"; modelAlias?: string } = {}) {
  _headerState = { ..._headerState, ...opts };
  render(renderHeader({
    status: _headerState.status,
    modelAlias: _headerState.modelAlias,
    onModelClick: () => {
      if (!_agent) return;
      ModelSelector.open(_agent.state.model, (model) => {
        setModelAndSync(_agent!, model);
      });
    },
  }), headerRoot);
}

function showErrorBanner(message: string): void {
  render(renderError(message), errorRoot);
}

function clearErrorBanner(): void {
  render(html``, errorRoot);
}

updateHeader();
render(renderLoading(), loadingRoot);


// ============================================================================
// Bootstrap
// ============================================================================

installFetchInterceptor();

let initialized = false;

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

  const agent = _agent = new Agent({
    initialState: {
      systemPrompt,
      model: defaultModel,
      thinkingLevel: "off",
      messages: [],
      tools: createAllTools({ changeTracker }),
    },
    transformContext: async (context) => await injectContext(context),
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

  // 7. Create and mount PiSidebar
  const sidebar = _sidebar = new PiSidebar();
  sidebar.agent = agent;
  sidebar.emptyHints = ["Summarize this sheet", "Add a VLOOKUP formula", "Format as a table"];
  sidebar.onSend = (text) => {
    clearErrorBanner();
    agent.prompt(text).catch((e) => {
      showErrorBanner(`LLM error: ${e.message}`);
    });
  };
  sidebar.onAbort = () => {
    _userAborted = true;
    agent.abort();
  };

  appEl.innerHTML = "";
  appEl.appendChild(sidebar);

  // 8. Header + status tracking
  const getModelAlias = () => {
    const m = agent.state.model;
    return m ? (m.name || m.id) : undefined;
  };
  updateHeader({ modelAlias: getModelAlias() });

  agent.subscribe((ev) => {
    // Header status
    if (ev.type === "turn_start") {
      updateHeader({ status: "working", modelAlias: getModelAlias() });
    } else if (ev.type === "turn_end" || ev.type === "agent_end") {
      updateHeader({
        status: agent.state.error ? "error" : "ready",
        modelAlias: getModelAlias(),
      });
    }

    // Error banner
    if (ev.type === "message_start" && ev.message.role === "user") {
      clearErrorBanner();
    }
    if (ev.type === "agent_end") {
      if (agent.state.error) {
        const isAbort = _userAborted ||
          /abort/i.test(agent.state.error) ||
          /cancel/i.test(agent.state.error);
        if (isAbort) {
          clearErrorBanner();
          updateHeader({ status: "ready", modelAlias: getModelAlias() });
        } else {
          showErrorBanner(`LLM error: ${agent.state.error}`);
        }
      } else {
        clearErrorBanner();
      }
      _userAborted = false;
    }

  });

  // ── Session persistence ──
  let _sessionId: string = crypto.randomUUID();
  let _sessionTitle = "";
  let _sessionCreatedAt = new Date().toISOString();
  let _firstAssistantSeen = false;

  async function saveSession() {
    if (!_firstAssistantSeen) return;
    try {
      const now = new Date().toISOString();
      const messages = agent.state.messages;
      if (!_sessionTitle && messages.length > 0) {
        const firstUser = messages.find((m) => m.role === "user");
        if (firstUser) {
          const content = firstUser.content;
          const text = typeof content === "string"
            ? content
            : Array.isArray(content)
              ? content.filter((b: any) => b.type === "text").map((b: any) => b.text).join("")
              : "";
          _sessionTitle = text.slice(0, 80) || "Untitled";
        }
      }
      let preview = "";
      for (const m of messages) {
        if (m.role !== "user" && m.role !== "assistant") continue;
        const content = m.content;
        const text = typeof content === "string"
          ? content
          : Array.isArray(content)
            ? content.filter((b: any) => b.type === "text").map((b: any) => b.text).join("")
            : "";
        preview += text + "\n";
        if (preview.length > 2048) { preview = preview.slice(0, 2048); break; }
      }
      let inputTokens = 0, outputTokens = 0, totalCost = 0;
      for (const m of messages) {
        const u = (m as any).usage;
        if (u) {
          inputTokens += u.inputTokens || 0;
          outputTokens += u.outputTokens || 0;
          totalCost += u.totalCost || 0;
        }
      }
      await sessions.saveSession(_sessionId, agent.state, {
        id: _sessionId,
        title: _sessionTitle,
        createdAt: _sessionCreatedAt,
        lastModified: now,
        messageCount: messages.length,
        usage: {
          input: inputTokens,
          output: outputTokens,
          cacheRead: 0,
          cacheWrite: 0,
          totalTokens: inputTokens + outputTokens,
          cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: totalCost },
        },
        thinkingLevel: agent.state.thinkingLevel || "off",
        preview,
      }, _sessionTitle);
    } catch (err) {
      console.warn("[pi] Session save failed:", err);
    }
  }

  function startNewSession() {
    _sessionId = crypto.randomUUID();
    _sessionTitle = "";
    _sessionCreatedAt = new Date().toISOString();
    _firstAssistantSeen = false;
  }

  agent.subscribe((ev) => {
    if (ev.type === "message_end") {
      if (ev.message.role === "assistant") _firstAssistantSeen = true;
      if (_firstAssistantSeen) saveSession();
    }
  });

  // Auto-restore latest session
  try {
    const latestId = await sessions.getLatestSessionId();
    if (latestId) {
      const sessionData = await sessions.loadSession(latestId);
      if (sessionData && sessionData.messages.length > 0) {
        _sessionId = sessionData.id;
        _sessionTitle = sessionData.title || "";
        _sessionCreatedAt = sessionData.createdAt;
        _firstAssistantSeen = true;
        agent.replaceMessages(sessionData.messages);
        if (sessionData.model) {
          setModelAndSync(agent, sessionData.model);
        }
        if (sessionData.thinkingLevel) {
          agent.setThinkingLevel(sessionData.thinkingLevel);
          updateStatusBar(agent);
        }
        // Force sidebar to re-render with restored messages
        requestAnimationFrame(() => sidebar.requestUpdate());
        console.log(`[pi] Restored session: ${_sessionTitle || latestId}`);
      }
    }
  } catch (err) {
    console.warn("[pi] Session restore failed:", err);
  }

  document.addEventListener("pi:session-new", () => startNewSession());
  document.addEventListener("pi:session-rename", ((e: CustomEvent) => {
    _sessionTitle = e.detail?.title || _sessionTitle;
    saveSession();
  }) as EventListener);
  document.addEventListener("pi:session-resumed", ((e: CustomEvent) => {
    _sessionId = e.detail?.id || _sessionId;
    _sessionTitle = e.detail?.title || "";
    _sessionCreatedAt = e.detail?.createdAt || new Date().toISOString();
    _firstAssistantSeen = true;
  }) as EventListener);

  // ── Register slash commands + extensions ──
  registerBuiltins(agent);
  const extensionAPI = createExtensionAPI(agent);
  const { activate: activateSnake } = await import("./extensions/snake.js");
  await loadExtension(extensionAPI, activateSnake);

  document.addEventListener("pi:providers-changed", async () => {
    const updated = await providerKeys.list();
    setActiveProviders(new Set(updated));
  });

  // ── Abort tracking ──
  let _userAborted = false;

  // ── Queue display ──
  type QueuedItem = { type: "steer" | "follow-up"; text: string };
  const _queuedMessages: QueuedItem[] = [];

  function addQueuedMessage(type: QueuedItem["type"], text: string) {
    _queuedMessages.push({ type, text });
    updateQueueDisplay();
  }

  function clearQueue() {
    _queuedMessages.length = 0;
    updateQueueDisplay();
  }

  function updateQueueDisplay() {
    let container = document.getElementById("pi-queue-display");
    if (_queuedMessages.length === 0) {
      container?.remove();
      return;
    }
    if (!container) {
      container = document.createElement("div");
      container.id = "pi-queue-display";
      container.className = "pi-queue";
      document.body.appendChild(container);
    }
    // Position above the input area
    const inputArea = sidebar.querySelector(".pi-input-area") as HTMLElement | null;
    const inputTop = inputArea ? inputArea.getBoundingClientRect().top : window.innerHeight - 80;
    container.style.bottom = `${window.innerHeight - inputTop}px`;

    container.innerHTML = _queuedMessages.map(({ type, text }) => {
      const label = type === "steer" ? "Steering" : "Follow-up";
      const cls = type === "steer" ? "pi-queue__label--steer" : "pi-queue__label--followup";
      const truncated = text.length > 50 ? text.slice(0, 47) + "…" : text;
      return `<div class="pi-queue__item">
        <span class="pi-queue__label ${cls}">${label}</span>
        <span class="pi-queue__text">${truncated}</span>
      </div>`;
    }).join("");
  }

  agent.subscribe((ev) => {
    if (_queuedMessages.length === 0) return;
    if (ev.type === "message_start" && ev.message.role === "user") {
      const content = ev.message.content;
      const msgText = typeof content === "string"
        ? content
        : Array.isArray(content)
          ? content.filter((b: any) => b.type === "text").map((b: any) => b.text).join("")
          : "";
      const idx = _queuedMessages.findIndex((q) => q.text === msgText);
      if (idx !== -1) {
        _queuedMessages.splice(idx, 1);
        updateQueueDisplay();
      }
    }
    if (ev.type === "agent_end" && _queuedMessages.length > 0) clearQueue();
  });

  // ── Keyboard shortcuts ──
  const THINKING_COLORS: Record<string, string> = {
    off: "#a0a0a0", minimal: "#767676", low: "#4488cc",
    medium: "#22998a", high: "#875f87", xhigh: "#8b008b",
  };

  function getThinkingLevels(): string[] {
    const model = agent.state.model;
    if (!model || !model.reasoning) return ["off"];
    const provider = model.provider;
    if (provider === "openai" || provider === "openai-codex") {
      const levels = ["off", "minimal", "low", "medium", "high"];
      if (supportsXhigh(model)) levels.push("xhigh");
      return levels;
    }
    if (provider === "anthropic") return ["off", "low", "medium", "high"];
    return ["off", "low", "medium", "high"];
  }

  document.addEventListener("keydown", (e) => {
    // Command menu takes priority
    if (isCommandMenuVisible()) {
      if (handleCommandMenuKey(e)) return;
    }

    const textarea = sidebar.getTextarea();
    const isInEditor = textarea && (e.target === textarea || textarea.contains(e.target as Node));
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
      const next = levels[(idx + 1) % levels.length];
      agent.setThinkingLevel(next as any);
      updateStatusBar(agent);
      flashThinkingLevel(next, THINKING_COLORS[next] || "#a0a0a0");
      return;
    }

    // Slash command execution
    if (isInEditor && e.key === "Enter" && !e.shiftKey && textarea!.value.startsWith("/") && !isStreaming) {
      const val = textarea!.value.trim();
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
    if (isInEditor && e.key === "Enter" && !e.shiftKey && isStreaming) {
      const text = textarea!.value.trim();
      if (!text) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      const msg = { role: "user" as const, content: [{ type: "text" as const, text }], timestamp: Date.now() };
      if (e.altKey) {
        agent.followUp(msg);
        addQueuedMessage("follow-up", text);
      } else {
        agent.steer(msg);
        addQueuedMessage("steer", text);
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

  // ── Thinking indicator click ──
  document.addEventListener("click", (e) => {
    const target = (e.target as HTMLElement).closest?.(".pi-status-thinking");
    if (target) {
      const levels = getThinkingLevels();
      const current = agent.state.thinkingLevel;
      const idx = levels.indexOf(current);
      const next = levels[(idx + 1) % levels.length];
      agent.setThinkingLevel(next as any);
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

async function injectContext(messages: any[]): Promise<any[]> {
  const injections: string[] = [];
  try {
    const sel = await withTimeout(readSelectionContext().catch(() => null), 1500);
    if (sel) injections.push(sel.text);
  } catch {}
  const changes = changeTracker.flush();
  if (changes) injections.push(changes);
  if (injections.length === 0) return messages;

  const injection = injections.join("\n\n");
  const injectionMessage = {
    role: "user" as const,
    content: [{ type: "text" as const, text: `[Auto-context]\n${injection}` }],
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
  let totalTokens = 0;
  for (const msg of state.messages) {
    const usage = (msg as any).usage;
    if (usage) totalTokens += (usage.input || 0) + (usage.output || 0);
  }

  const contextWindow = state.model?.contextWindow || 200000;
  const pct = Math.min(100, Math.round((totalTokens / contextWindow) * 100));
  const ctxLabel = contextWindow >= 1_000_000
    ? `${(contextWindow / 1_000_000).toFixed(0)}M`
    : `${Math.round(contextWindow / 1000)}k`;

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
// Thinking level flash
// ============================================================================

function flashThinkingLevel(level: string, color: string): void {
  const labels: Record<string, string> = { off: "Off", low: "Low", medium: "Medium", high: "High" };
  showToast(`Thinking: ${labels[level] || level} (⇧Tab to toggle)`, 1500);

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
// Welcome / Login
// ============================================================================

async function showWelcomeLogin(providerKeys: InstanceType<typeof ProviderKeysStore>): Promise<void> {
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
    const providerList = overlay.querySelector(".pi-welcome-providers")!;
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

type DefaultModelRule = { provider: string; match: RegExp };

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

function pickLatestMatchingModel(provider: string, match: RegExp) {
  try {
    const models = (getModels as any)(provider) as Array<{ id: string }>;
    const candidates = models.filter((m) => match.test(m.id));
    candidates.sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id));
    return candidates[0] || null;
  } catch {
    return null;
  }
}

function pickDefaultModel(availableProviders: string[]) {
  // Anthropic special-case:
  // Prefer Opus, except if there's a *newer-version* Sonnet, use that first.
  if (availableProviders.includes("anthropic")) {
    try {
      const models = (getModels as any)("anthropic") as Array<{ id: string }>;
      const opus = models.filter((m) => String(m.id).startsWith("claude-opus-"))
        .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];
      const sonnet = models.filter((m) => String(m.id).startsWith("claude-sonnet-"))
        .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];

      if (opus && sonnet) {
        if (parseMajorMinor(sonnet.id) > parseMajorMinor(opus.id)) {
          return (getModel as any)("anthropic", sonnet.id);
        }
        return (getModel as any)("anthropic", opus.id);
      }

      if (opus) return (getModel as any)("anthropic", opus.id);
      if (sonnet) return (getModel as any)("anthropic", sonnet.id);
    } catch {}
  }

  // Other providers: pattern-based rules
  for (const rule of DEFAULT_MODEL_RULES) {
    if (!availableProviders.includes(rule.provider)) continue;
    const m = pickLatestMatchingModel(rule.provider, rule.match);
    if (m) {
      try { return (getModel as any)(rule.provider, m.id); } catch {}
    }
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
