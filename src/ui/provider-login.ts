/**
 * Shared provider login row builder — used by both welcome screen and /login command.
 *
 * Renders an inline expandable row with:
 * - OAuth button (for providers that support it)
 * - "or enter API key" divider
 * - API key input + Save button
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";
import { isCorsError } from "@mariozechner/pi-web-ui/dist/utils/proxy-utils.js";
import { getOAuthProvider } from "../auth/oauth-provider-registry.js";
import { clearOAuthCredentials, saveOAuthCredentials } from "../auth/oauth-storage.js";
import {
  DEFAULT_LOCAL_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
} from "../auth/proxy-validation.js";
import { PROVIDER_PROMPT_OVERLAY_ID } from "./overlay-ids.js";
import { closeOverlayById, createOverlayDialog } from "./overlay-dialog.js";
import { getErrorMessage } from "../utils/errors.js";

export interface ProviderDef {
  id: string;
  label: string;
  oauth?: string;
  desc?: string;
}

export const ALL_PROVIDERS: ProviderDef[] = [
  // OAuth providers first (subscription / account-based flows)
  // Only list flows that are supported in-browser (PKCE/manual paste, no local callback server).
  { id: "anthropic",          label: "Anthropic",                oauth: "anthropic",          desc: "Claude Pro/Max" },
  { id: "openai-codex",       label: "OpenAI (ChatGPT)",         oauth: "openai-codex",       desc: "Plus/Pro subscription" },
  { id: "google-gemini-cli",  label: "Google Code Assist",       oauth: "google-gemini-cli",  desc: "Gemini via Google account" },
  { id: "google-antigravity", label: "Google Antigravity",       oauth: "google-antigravity", desc: "Gemini/Claude/GPT-OSS" },
  { id: "github-copilot",     label: "GitHub Copilot",           oauth: "github-copilot" },

  // API key providers
  { id: "openai",             label: "OpenAI (API)",             desc: "API key" },
  { id: "google",             label: "Google Gemini (API)",      desc: "API key" },
  { id: "deepseek",           label: "DeepSeek" },
  { id: "amazon-bedrock",     label: "Amazon Bedrock" },
  { id: "mistral",            label: "Mistral" },
  { id: "groq",               label: "Groq" },
  { id: "xai",                label: "xAI / Grok" },
];

export interface ProviderRowCallbacks {
  onConnected: (row: HTMLElement, id: string, label: string) => void;
  onDisconnected?: (row: HTMLElement, id: string, label: string) => void;
}

class PromptCancelledError extends Error {
  constructor() {
    super("Prompt cancelled");
  }
}

function normalizeAnthropicAuthorizationInput(input: string): string {
  const value = input.trim();
  if (!value) return value;

  // Accept full redirect URL (or any URL with code/state query params)
  try {
    const url = new URL(value);
    const code = url.searchParams.get("code");
    const state = url.searchParams.get("state");
    if (code) return state ? `${code}#${state}` : code;
  } catch {
    // ignore
  }

  // Accept query-string style pastes (code=...&state=...)
  if (value.includes("code=")) {
    try {
      const params = new URLSearchParams(value.startsWith("?") ? value.slice(1) : value);
      const code = params.get("code");
      const state = params.get("state");
      if (code) return state ? `${code}#${state}` : code;
    } catch {
      // ignore
    }
  }

  // Accept whitespace-separated values (code state)
  if (!value.includes("#")) {
    const parts = value.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return `${parts[0]}#${parts[1]}`;
  }

  return value;
}

function looksLikeOAuthRedirectInput(value: string): boolean {
  const lower = value.toLowerCase();
  return (
    value.includes("#")
    || value.includes("code=")
    || lower.startsWith("http://localhost:1455/")
    || lower.startsWith("http://localhost:8085/")
    || lower.startsWith("http://localhost:51121/")
    || lower.startsWith("https://auth.openai.com/")
    || lower.startsWith("https://accounts.google.com/")
    || lower.includes("oauth2callback")
    || lower.includes("oauth-callback")
  );
}

function normalizeApiKeyForProvider(
  providerId: string,
  raw: string,
): { ok: true; key: string } | { ok: false; error: string } {
  let key = raw.trim();
  if (!key) return { ok: false, error: "API key is empty" };

  // Common copy/paste format: "Bearer <token>"
  if (/^bearer\s+/i.test(key)) {
    key = key.replace(/^bearer\s+/i, "").trim();
  }

  if (providerId === "anthropic") {
    // Prevent saving Anthropic OAuth *authorization code* (code#state) as an API key.
    // OAuth access tokens are sk-ant-oat*, API keys are sk-ant-api*.
    const looksLikeAuthCode = key.includes("#") && !key.includes("sk-ant-");
    if (looksLikeAuthCode) {
      return {
        ok: false,
        error:
          "That looks like an OAuth authorization code (code#state). Use “Login with Anthropic” and paste it when prompted (don’t Save it as an API key).",
      };
    }
  }

  if (providerId === "openai-codex" && looksLikeOAuthRedirectInput(key)) {
    return {
      ok: false,
      error:
        "That looks like an OAuth redirect URL/code. Use “Login with OpenAI (ChatGPT)” and paste it in the login prompt (don’t Save it as an API key).",
    };
  }

  if ((providerId === "google-gemini-cli" || providerId === "google-antigravity") && looksLikeOAuthRedirectInput(key)) {
    return {
      ok: false,
      error:
        "That looks like an OAuth redirect URL/code. Use “Login with Google …” and paste it in the login prompt (don’t Save it as an API key).",
    };
  }

  if (providerId === "google" && looksLikeOAuthRedirectInput(key)) {
    return {
      ok: false,
      error:
        "That looks like an OAuth redirect URL/code. Use Google API key auth here, or use the dedicated Google OAuth login rows.",
    };
  }

  return { ok: true, key };
}

function promptForText(opts: {
  title: string;
  message: string;
  placeholder?: string;
  helperText?: string;
  submitLabel?: string;
}): Promise<string> {
  return new Promise((resolve, reject) => {
    closeOverlayById(PROVIDER_PROMPT_OVERLAY_ID);

    const dialog = createOverlayDialog({
      overlayId: PROVIDER_PROMPT_OVERLAY_ID,
      cardClassName: "pi-welcome-card pi-prompt-card",
      restoreFocusOnClose: false,
    });

    const titleEl = document.createElement("h2");
    titleEl.className = "pi-prompt-title";
    titleEl.textContent = opts.title;

    const messageEl = document.createElement("p");
    messageEl.className = "pi-prompt-message";
    messageEl.textContent = opts.message;

    const helperEl = document.createElement("p");
    helperEl.className = "pi-prompt-helper";
    helperEl.hidden = true;

    const input = document.createElement("input");
    input.className = "pi-prompt-input";
    input.type = "text";
    input.autocomplete = "off";
    input.setAttribute("aria-label", opts.title);

    const actions = document.createElement("div");
    actions.className = "pi-prompt-actions";

    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "pi-prompt-cancel";
    cancelBtn.textContent = "Cancel";

    const okBtn = document.createElement("button");
    okBtn.type = "button";
    okBtn.className = "pi-prompt-ok";
    okBtn.textContent = opts.submitLabel ?? "Continue";

    actions.append(cancelBtn, okBtn);
    dialog.card.append(titleEl, messageEl, helperEl, input, actions);

    if (opts.helperText) {
      helperEl.textContent = opts.helperText;
      helperEl.hidden = false;
    }

    if (opts.placeholder) {
      input.placeholder = opts.placeholder;
    }

    let settled = false;

    const submit = (): void => {
      if (settled) {
        return;
      }

      settled = true;
      const value = input.value.trim();
      dialog.close();
      resolve(value);
    };

    const cancel = (): void => {
      if (settled) {
        return;
      }

      settled = true;
      dialog.close();
      reject(new PromptCancelledError());
    };

    const onInputKeyDown = (event: KeyboardEvent): void => {
      if (event.key !== "Enter") {
        return;
      }

      event.preventDefault();
      submit();
    };

    cancelBtn.addEventListener("click", cancel);
    okBtn.addEventListener("click", submit);
    input.addEventListener("keydown", onInputKeyDown);

    dialog.addCleanup(() => {
      cancelBtn.removeEventListener("click", cancel);
      okBtn.removeEventListener("click", submit);
      input.removeEventListener("keydown", onInputKeyDown);

      if (!settled) {
        settled = true;
        reject(new PromptCancelledError());
      }
    });

    dialog.mount();
    requestAnimationFrame(() => input.focus());
  });
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
  } & ProviderRowCallbacks,
): HTMLElement {
  const { id, label, oauth, desc } = provider;
  const { isActive, expandedRef, onConnected, onDisconnected } = opts;
  const storage = getAppStorage();

  const keyPlaceholder = id === "anthropic"
    ? "sk-ant-api… or sk-ant-oat…"
    : id === "openai-codex"
      ? "ChatGPT OAuth access token"
      : id === "google-gemini-cli" || id === "google-antigravity"
        ? "Google OAuth credential JSON"
        : "Enter API key";

  const row = document.createElement("div");
  row.className = "pi-login-row";
  row.innerHTML = `
    <button class="pi-welcome-provider pi-login-trigger">
      <span class="pi-login-meta">
        <span class="pi-login-label">${label}</span>
        ${desc ? `<span class="pi-login-desc">${desc}</span>` : ""}
      </span>
      <span class="pi-login-status ${isActive ? "is-connected" : ""}">
        ${isActive ? "✓ connected" : "set up →"}
      </span>
    </button>
    <div class="pi-login-detail" hidden>
      <button class="pi-login-disconnect" type="button" ${isActive ? "" : "hidden"}>Disconnect ${label}</button>
      ${oauth ? `
        <button class="pi-login-oauth">Login with ${label}</button>
        <div class="pi-login-divider">
          <div class="pi-login-divider__line"></div>
          <span class="pi-login-divider__text">or enter API key</span>
          <div class="pi-login-divider__line"></div>
        </div>
      ` : ""}
      <div class="pi-login-key-row">
        <input class="pi-login-key" type="password" placeholder="${keyPlaceholder}" aria-label="API key for ${label}" autocomplete="off" spellcheck="false" />
        <button class="pi-login-save">Save</button>
      </div>
      <p class="pi-login-error" hidden></p>
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
  const statusEl = row.querySelector<HTMLElement>(".pi-login-status");
  const oauthBtn = row.querySelector<HTMLButtonElement>(".pi-login-oauth");
  const disconnectBtn = row.querySelector<HTMLButtonElement>(".pi-login-disconnect");

  const setConnectedState = (connected: boolean): void => {
    if (statusEl) {
      statusEl.textContent = connected ? "✓ connected" : "set up →";
      statusEl.classList.toggle("is-connected", connected);
    }

    if (disconnectBtn) {
      disconnectBtn.hidden = !connected;
    }
  };

  setConnectedState(isActive);

  // Toggle expand
  headerBtn.addEventListener("click", () => {
    if (expandedRef.current === detail) {
      detail.hidden = true;
      expandedRef.current = null;
    } else {
      if (expandedRef.current) expandedRef.current.hidden = true;
      detail.hidden = false;
      expandedRef.current = detail;
      keyInput.focus();
    }
  });

  // OAuth login
  if (oauthBtn) {
    oauthBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      oauthBtn.textContent = "Opening login…";
      oauthBtn.style.opacity = "0.7";
      void (async () => {
        errorEl.hidden = true;
        try {
          if (!oauth) {
            throw new Error("OAuth provider id missing");
          }

          const oauthProvider = getOAuthProvider(oauth);
          if (!oauthProvider) {
            throw new Error(`OAuth provider not supported: ${oauth}`);
          }

          const cred = await oauthProvider.login({
            onAuth: (info) => {
              // Prevent the OAuth page from gaining a handle to the add-in window.
              const w = window.open(info.url, "_blank", "noopener,noreferrer");
              if (w) w.opener = null;
            },
            onPrompt: async (prompt) => {
              const helperText = id === "anthropic"
                ? "After completing login, copy the authorization string from the browser. You can paste the full URL, or a CODE#STATE value."
                : id === "openai-codex"
                  ? "After completing login, copy the final browser URL (usually localhost) from the address bar, then paste it here."
                  : id === "google-gemini-cli" || id === "google-antigravity"
                    ? "After sign-in, copy the final browser URL (localhost callback) from the address bar, then paste it here."
                    : undefined;

              const value = await promptForText({
                title: `Login with ${label}`,
                message: prompt.message,
                placeholder: prompt.placeholder || "",
                helperText,
                submitLabel: "Continue",
              });

              if (id === "anthropic") {
                return normalizeAnthropicAuthorizationInput(value);
              }

              return value;
            },
            onProgress: (msg) => { oauthBtn.textContent = msg; },
          });

          const apiKey = oauthProvider.getApiKey(cred);
          await storage.providerKeys.set(id, apiKey);
          await saveOAuthCredentials(storage.settings, id, cred);
          setConnectedState(true);
          onConnected(row, id, label);
          detail.hidden = true;
          expandedRef.current = null;
        } catch (err: unknown) {
          if (err instanceof PromptCancelledError) {
            // User cancelled the prompt; leave UI unchanged.
            return;
          }

          const msg = getErrorMessage(err);
          const isLikelyCors =
            isCorsError(err) ||
            (typeof msg === "string" && /load failed|failed to fetch|cors|cross-origin|networkerror/i.test(msg));

          if (isLikelyCors) {
            errorEl.textContent =
              "Login was blocked by browser CORS. Enable /settings → Proxy with " +
              `${DEFAULT_LOCAL_PROXY_URL} (local HTTPS proxy helper running), then retry. ` +
              `Guide: ${PROXY_HELPER_DOCS_URL}`;
          } else {
            errorEl.textContent = msg || "Login failed";
          }
          errorEl.hidden = false;
        } finally {
          oauthBtn.textContent = `Login with ${label}`;
          oauthBtn.style.opacity = "1";
        }
      })();
    });
  }

  // Credential disconnect
  if (disconnectBtn) {
    disconnectBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      void (async () => {
        disconnectBtn.textContent = "Disconnecting…";
        disconnectBtn.disabled = true;
        disconnectBtn.style.opacity = "0.7";
        errorEl.hidden = true;

        try {
          await storage.providerKeys.delete(id);
          await clearOAuthCredentials(storage.settings, id);

          setConnectedState(false);
          keyInput.value = "";
          onDisconnected?.(row, id, label);
        } catch (err: unknown) {
          const msg = getErrorMessage(err);
          errorEl.textContent = msg ? `Failed to disconnect: ${msg}` : "Failed to disconnect";
          errorEl.hidden = false;
        } finally {
          disconnectBtn.textContent = `Disconnect ${label}`;
          disconnectBtn.disabled = false;
          disconnectBtn.style.opacity = "1";
        }
      })();
    });
  }

  // API key save
  saveBtn.addEventListener("click", () => { void (async () => {
    const rawKey = keyInput.value.trim();
    if (!rawKey) return;

    const normalized = normalizeApiKeyForProvider(id, rawKey);
    if (!normalized.ok) {
      errorEl.textContent = normalized.error;
      errorEl.hidden = false;
      return;
    }

    const key = normalized.key;
    saveBtn.textContent = "Testing…";
    saveBtn.style.opacity = "0.7";
    errorEl.hidden = true;
    try {
      await storage.providerKeys.set(id, key);
      setConnectedState(true);
      onConnected(row, id, label);
      detail.hidden = true;
      expandedRef.current = null;
    } catch (err: unknown) {
      const msg = getErrorMessage(err);
      errorEl.textContent = msg ? `Failed to save key: ${msg}` : "Failed to save key";
      errorEl.hidden = false;
    } finally {
      saveBtn.textContent = "Save";
      saveBtn.style.opacity = "1";
    }
  })(); });

  // Enter key in input
  keyInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") saveBtn.click();
  });

  return row;
}
