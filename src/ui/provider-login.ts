/**
 * Shared provider login row builder — used by both welcome screen and /login command.
 *
 * Renders an inline expandable row with:
 * - OAuth button (for providers that support it)
 * - "or enter API key" divider
 * - API key input + Save button
 */

import { getAppStorage, isCorsError } from "@mariozechner/pi-web-ui";
import { getErrorMessage } from "../utils/errors.js";

export interface ProviderDef {
  id: string;
  label: string;
  oauth?: string;
  desc?: string;
}

export const ALL_PROVIDERS: ProviderDef[] = [
  // OAuth providers first (free with subscription)
  { id: "anthropic",           label: "Anthropic",        oauth: "anthropic",          desc: "Claude Pro/Max" },
  { id: "openai-codex",       label: "ChatGPT Plus/Pro", oauth: "openai-codex",       desc: "Codex Subscription" },
  { id: "github-copilot",     label: "GitHub Copilot",   oauth: "github-copilot" },
  { id: "google-gemini-cli",  label: "Gemini CLI",       oauth: "google-gemini-cli",  desc: "Google Cloud Code Assist" },
  { id: "google-antigravity", label: "Antigravity",      oauth: "google-antigravity", desc: "Gemini 3, Claude, GPT-OSS" },
  // API key providers
  { id: "openai",             label: "OpenAI",           desc: "API key" },
  { id: "google",             label: "Google Gemini",    desc: "API key" },
  { id: "deepseek",           label: "DeepSeek" },
  { id: "amazon-bedrock",     label: "Amazon Bedrock" },
  { id: "mistral",            label: "Mistral" },
  { id: "groq",               label: "Groq" },
  { id: "xai",                label: "xAI / Grok" },
];

export interface ProviderRowCallbacks {
  onConnected: (row: HTMLElement, id: string, label: string) => void;
}

/**
 * Build a provider login row with inline OAuth + API key.
 * Manages expand/collapse via the shared expandedRef.
 */
export function buildProviderRow(
  provider: ProviderDef,
  opts: {
    isActive: boolean;
    expandedRef: { current: HTMLElement | null };
    onConnected: (row: HTMLElement, id: string, label: string) => void;
  }
): HTMLElement {
  const { id, label, oauth, desc } = provider;
  const { isActive, expandedRef, onConnected } = opts;
  const storage = getAppStorage();

  const row = document.createElement("div");
  row.className = "pi-login-row";
  row.innerHTML = `
    <button class="pi-welcome-provider" style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
      <span style="display: flex; flex-direction: column; align-items: flex-start; gap: 1px;">
        <span style="font-size: 13px;">${label}</span>
        ${desc ? `<span style="font-size: 10px; color: var(--muted-foreground); font-family: var(--font-sans);">${desc}</span>` : ""}
      </span>
      <span class="pi-login-status" style="font-size: 11px; color: ${isActive ? "var(--pi-green)" : "var(--muted-foreground)"}; font-family: var(--font-mono);">
        ${isActive ? "✓ connected" : "set up →"}
      </span>
    </button>
    <div class="pi-login-detail" style="display: none; padding: 8px 14px 12px; border: 1px solid oklch(0 0 0 / 0.05); border-top: none; border-radius: 0 0 10px 10px; margin-top: -1px; background: oklch(1 0 0 / 0.3);">
      ${oauth ? `
        <button class="pi-login-oauth" style="
          width: 100%; padding: 9px 14px; margin-bottom: 8px;
          background: var(--pi-green); color: white; border: none;
          border-radius: 9px; font-family: var(--font-sans);
          font-size: 13px; font-weight: 500; cursor: pointer;
          transition: background 0.15s;
        ">Login with ${label}</button>
        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
          <div style="flex: 1; height: 1px; background: oklch(0 0 0 / 0.08);"></div>
          <span style="font-size: 11px; color: var(--muted-foreground); font-family: var(--font-sans);">or enter API key</span>
          <div style="flex: 1; height: 1px; background: oklch(0 0 0 / 0.08);"></div>
        </div>
      ` : ""}
      <div style="display: flex; gap: 6px;">
        <input class="pi-login-key" type="password" placeholder="Enter API key"
          style="flex: 1; padding: 7px 10px; border: 1px solid oklch(0 0 0 / 0.10);
          border-radius: 8px; font-family: var(--font-mono); font-size: 12px;
          background: oklch(1 0 0 / 0.6); outline: none;"
        />
        <button class="pi-login-save" style="
          padding: 7px 12px; background: var(--pi-green); color: white;
          border: none; border-radius: 8px; font-family: var(--font-sans);
          font-size: 12px; font-weight: 500; cursor: pointer;
          transition: background 0.15s; white-space: nowrap;
        ">Save</button>
      </div>
      <p class="pi-login-error" style="display: none; font-size: 11px; color: oklch(0.55 0.22 25); margin: 6px 0 0; font-family: var(--font-sans);"></p>
    </div>
  `;

  const headerBtn = row.querySelector<HTMLButtonElement>(".pi-welcome-provider");
  if (!headerBtn) {
    throw new Error("Provider row header button not found");
  }
  const detail = row.querySelector(".pi-login-detail") as HTMLElement;
  const keyInput = row.querySelector(".pi-login-key") as HTMLInputElement;
  const saveBtn = row.querySelector(".pi-login-save") as HTMLButtonElement;
  const errorEl = row.querySelector(".pi-login-error") as HTMLElement;
  const oauthBtn = row.querySelector(".pi-login-oauth") as HTMLButtonElement | null;

  // Toggle expand
  headerBtn.addEventListener("click", () => {
    if (expandedRef.current === detail) {
      detail.style.display = "none";
      expandedRef.current = null;
    } else {
      if (expandedRef.current) expandedRef.current.style.display = "none";
      detail.style.display = "block";
      expandedRef.current = detail;
      keyInput.focus();
    }
  });

  // OAuth login
  if (oauthBtn) {
    oauthBtn.addEventListener("click", async (e) => {
      e.stopPropagation();
      oauthBtn.textContent = "Opening login…";
      oauthBtn.style.opacity = "0.7";
      try {
        const { getOAuthProvider } = await import("@mariozechner/pi-ai");
        if (!oauth) {
          throw new Error("OAuth provider id missing");
        }
        const oauthProvider = getOAuthProvider(oauth);
        if (oauthProvider) {
          const cred = await oauthProvider.login({
            onAuth: (info) => { window.open(info.url, "_blank"); },
            onPrompt: async (prompt) => window.prompt(prompt.message, prompt.placeholder || "") || "",
            onProgress: (msg) => { oauthBtn.textContent = msg; },
          });
          const apiKey = oauthProvider.getApiKey(cred);
          await storage.providerKeys.set(id, apiKey);
          localStorage.setItem(`oauth_${id}`, JSON.stringify(cred));
          markConnected(row);
          onConnected(row, id, label);
          detail.style.display = "none";
          expandedRef.current = null;
        }
      } catch (err: unknown) {
        if (isCorsError(err)) {
          errorEl.textContent = "Login was blocked by browser CORS. Start the local proxy (npm run proxy) and enable it in /settings → Proxy.";
        } else {
          errorEl.textContent = getErrorMessage(err) || "Login failed";
        }
        errorEl.style.display = "block";
      } finally {
        oauthBtn.textContent = `Login with ${label}`;
        oauthBtn.style.opacity = "1";
      }
    });
  }

  // API key save
  saveBtn.addEventListener("click", async () => {
    const key = keyInput.value.trim();
    if (!key) return;
    saveBtn.textContent = "Testing…";
    saveBtn.style.opacity = "0.7";
    errorEl.style.display = "none";
    try {
      await storage.providerKeys.set(id, key);
      markConnected(row);
      onConnected(row, id, label);
      detail.style.display = "none";
      expandedRef.current = null;
    } catch (err: unknown) {
      const msg = getErrorMessage(err);
      errorEl.textContent = msg ? `Failed to save key: ${msg}` : "Failed to save key";
      errorEl.style.display = "block";
    } finally {
      saveBtn.textContent = "Save";
      saveBtn.style.opacity = "1";
    }
  });

  // Enter key in input
  keyInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") saveBtn.click();
  });

  return row;
}

function markConnected(row: HTMLElement) {
  const status = row.querySelector(".pi-login-status") as HTMLElement;
  if (status) {
    status.textContent = "✓ connected";
    status.style.color = "var(--pi-green)";
  }
}
