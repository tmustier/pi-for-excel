/**
 * Pi for Excel — main entry point.
 *
 * Initializes Office.js, mounts the ChatPanel sidebar,
 * wires up tools and context injection.
 */

// MUST be first — Lit fix + CSS
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
import { getBlueprint, invalidateBlueprint } from "./context/blueprint.js";
import { readSelectionContext } from "./context/selection.js";
import { ChangeTracker } from "./context/change-tracker.js";

// ============================================================================
// Globals
// ============================================================================

declare const Office: any;

const appEl = document.getElementById("app")!;
const loadingEl = document.getElementById("loading")!;
const errorEl = document.getElementById("error")!;

const changeTracker = new ChangeTracker();

// ============================================================================
// Bootstrap
// ============================================================================

// Install CORS proxy before any fetch calls
installFetchInterceptor();

// Wait for Office.js
Office.onReady(async (info: { host: any; platform: any }) => {
  console.log(`[pi] Office.js ready: host=${info.host}, platform=${info.platform}`);

  try {
    await init();
  } catch (e: any) {
    showError(`Failed to initialize: ${e.message}`);
    console.error("[pi] Init error:", e);
  }
});

// Fallback if not in Excel (dev/testing)
setTimeout(() => {
  if (loadingEl.style.display !== "none") {
    console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
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
      model: getModel("anthropic", "claude-sonnet-4-5-20250929"),
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

  // 7. Render
  loadingEl.style.display = "none";

  render(
    html`
      <div class="w-full h-full flex flex-col overflow-hidden" style="background: var(--background); color: var(--foreground);">
        ${chatPanel}
      </div>
    `,
    appEl,
  );

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
// Error display
// ============================================================================

function showError(message: string): void {
  errorEl.style.display = "block";
  errorEl.textContent = message;
  loadingEl.style.display = "none";
}
