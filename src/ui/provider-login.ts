/**
 * Shared provider login row builder — used by both welcome screen and /login command.
 *
 * Renders an inline expandable row with:
 * - OAuth button (for providers that support it)
 * - "or enter API key" divider
 * - API key input + Save button
 */

import { getAppStorage } from "../storage/local/app-storage.js";
import { isCorsError } from "../auth/cors-error.js";
import { pollOAuthCallbackCapture } from "../auth/oauth-callback-capture.js";
import { getOAuthProvider } from "../auth/oauth-provider-registry.js";
import { clearOAuthCredentials, saveOAuthCredentials } from "../auth/oauth-storage.js";
import {
  DEFAULT_PROXY_IS_REMOTE,
  DEFAULT_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  probeProxyReachability,
  resolveConfiguredProxyUrl,
} from "../auth/proxy-validation.js";
import { PROVIDER_PROMPT_OVERLAY_ID, PROXY_GATE_OVERLAY_ID } from "./overlay-ids.js";
import { closeOverlayById, createOverlayDialog } from "./overlay-dialog.js";
import { getErrorMessage } from "../utils/errors.js";
import { escapeAttr, escapeHtml, setSafeInnerHTML } from "../utils/html.js";
import { t } from "../language/index.js";
import { filterProvidersByAllowlist, resolveAllowedProviderIds } from "./provider-allowlist.js";

/**
 * Quick reachability check against the configured proxy URL.
 * Returns true if the proxy is enabled and responding.
 */
async function isProxyReachable(): Promise<boolean> {
  try {
    const storage = getAppStorage();
    const enabled = await storage.settings.get("proxy.enabled");
    if (!enabled) return false;

    const raw = await storage.settings.get("proxy.url");
    const proxyUrl = resolveConfiguredProxyUrl(raw);
    return probeProxyReachability(proxyUrl, 1500);
  } catch {
    return false;
  }
}

/**
 * Show a blocking dialog explaining the proxy is needed, with
 * a copy-able terminal command and retry / cancel buttons.
 *
 * Resolves `true` if the user retried and proxy is now reachable.
 * Resolves `false` if the user cancelled.
 */
function showProxyGateDialog(): Promise<boolean> {
  return new Promise((resolve) => {
    closeOverlayById(PROXY_GATE_OVERLAY_ID);

    const dialog = createOverlayDialog({
      overlayId: PROXY_GATE_OVERLAY_ID,
      cardClassName: "pi-welcome-card pi-prompt-card",
      restoreFocusOnClose: false,
    });

    const title = document.createElement("h2");
    title.className = "pi-prompt-title";
    title.textContent = t("provider.proxy_gate.title");

    const message = document.createElement("p");
    message.className = "pi-prompt-message";
    message.style.lineHeight = "1.5";
    message.textContent = DEFAULT_PROXY_IS_REMOTE
      ? t("provider.proxy_gate.message_remote")
      : t("provider.proxy_gate.message");

    const codeRow = document.createElement("div");
    codeRow.style.cssText = "display:flex;align-items:center;gap:8px;margin:12px 0;";

    const codeEl = document.createElement("code");
    codeEl.style.cssText =
      "flex:1;padding:8px 10px;border-radius:6px;" +
      "background:var(--pi-code-bg, #1e1e1e);color:var(--pi-code-fg, #d4d4d4);" +
      "font-size:13px;font-family:var(--pi-monospace, monospace);user-select:all;";
    codeEl.textContent = "npx pi-for-excel-proxy";

    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.textContent = t("provider.proxy_gate.copy");
    copyBtn.style.cssText = "padding:6px 12px;border-radius:6px;font-size:13px;cursor:pointer;";
    copyBtn.addEventListener("click", () => {
      void navigator.clipboard.writeText("npx pi-for-excel-proxy").then(() => {
        copyBtn.textContent = t("provider.proxy_gate.copied");
        setTimeout(() => { copyBtn.textContent = t("provider.proxy_gate.copy"); }, 1500);
      });
    });

    codeRow.append(codeEl, copyBtn);

    const hint = document.createElement("p");
    hint.className = "pi-prompt-helper";
    hint.style.lineHeight = "1.5";
    if (DEFAULT_PROXY_IS_REMOTE) {
      codeRow.style.display = "none";
      hint.textContent = t("provider.proxy_gate.hint_remote", { url: DEFAULT_PROXY_URL });
    } else {
      setSafeInnerHTML(
        hint,
        `${escapeHtml(t("provider.proxy_gate.hint_html"))} ` +
          `<a href="${escapeAttr(PROXY_HELPER_DOCS_URL)}" target="_blank" rel="noopener noreferrer">${escapeHtml(t("provider.proxy_gate.guide"))}</a>`,
        "proxy gate helper link markup with escaped localized text",
      );
    }

    const actions = document.createElement("div");
    actions.className = "pi-prompt-actions";

    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "pi-prompt-cancel";
    cancelBtn.textContent = t("provider.prompt.cancel");

    const retryBtn = document.createElement("button");
    retryBtn.type = "button";
    retryBtn.className = "pi-prompt-ok";
    retryBtn.textContent = t("provider.proxy_gate.retry");

    actions.append(cancelBtn, retryBtn);
    dialog.card.append(title, message, codeRow, hint, actions);

    let settled = false;

    const doCancel = (): void => {
      if (settled) return;
      settled = true;
      dialog.close();
      resolve(false);
    };

    const doRetry = (): void => {
      if (settled) return;
      retryBtn.textContent = t("provider.proxy_gate.checking");
      retryBtn.style.opacity = "0.7";

      void (async () => {
        // Auto-enable the proxy setting so the fetch interceptor will route through it.
        try {
          const storage = getAppStorage();
          const url = await storage.settings.get("proxy.url");
          const proxyUrl = resolveConfiguredProxyUrl(url);
          const ok = await probeProxyReachability(proxyUrl, 1500);

          if (ok) {
            await storage.settings.set("proxy.enabled", true);
            settled = true;
            dialog.close();
            resolve(true);
            return;
          }
        } catch {
          // fall through
        }

        retryBtn.textContent = t("provider.proxy_gate.retry");
        retryBtn.style.opacity = "1";
        if (DEFAULT_PROXY_IS_REMOTE) {
          hint.textContent = t("provider.proxy_gate.not_detected_remote", { url: DEFAULT_PROXY_URL });
        } else {
          setSafeInnerHTML(
            hint,
            `${escapeHtml(t("provider.proxy_gate.not_detected_html"))} ` +
              `<a href="${escapeAttr(PROXY_HELPER_DOCS_URL)}" target="_blank" rel="noopener noreferrer">${escapeHtml(t("provider.proxy_gate.guide"))}</a>`,
            "proxy retry helper link markup with escaped localized text",
          );
        }
      })();
    };

    cancelBtn.addEventListener("click", doCancel);
    retryBtn.addEventListener("click", doRetry);

    dialog.addCleanup(() => {
      cancelBtn.removeEventListener("click", doCancel);
      retryBtn.removeEventListener("click", doRetry);
      if (!settled) { settled = true; resolve(false); }
    });

    dialog.mount();
  });
}

/**
 * OAuth providers whose token exchange / API calls are CORS-blocked in Office
 * webviews and therefore require the local proxy.
 */
const OAUTH_IDS_NEEDING_PROXY = new Set([
  "anthropic",
  "openai-codex",
  "google-gemini-cli",
  "google-antigravity",
  "github-copilot",
]);

const OAUTH_CALLBACK_CAPTURE_IDS = new Set([
  "anthropic",
  "openai-codex",
  "google-gemini-cli",
  "google-antigravity",
]);

export interface ProviderDef {
  id: string;
  label: string;
  oauth?: string;
  desc?: string;
}

export const ALL_PROVIDERS: ProviderDef[] = [
  // OAuth providers first (subscription / account-based flows)
  // Only list flows that are supported in-browser (PKCE with proxy-assisted or manual callback handling).
  // desc holds a locale key (resolved via t() at render time in buildProviderRow).
  { id: "anthropic",          label: /* brand */ "Anthropic",                oauth: "anthropic",          desc: "provider.desc.claude" },
  { id: "openai-codex",       label: /* brand */ "OpenAI (ChatGPT)",         oauth: "openai-codex",       desc: "provider.desc.openai_sub" },
  { id: "google-gemini-cli",  label: /* brand */ "Google Code Assist",       oauth: "google-gemini-cli",  desc: "provider.desc.gemini_account" },
  { id: "google-antigravity", label: /* brand */ "Google Antigravity",       oauth: "google-antigravity", desc: "provider.desc.antigravity" },
  { id: "github-copilot",     label: /* brand */ "GitHub Copilot",           oauth: "github-copilot" },

  // API key providers
  { id: "openai",             label: /* brand */ "OpenAI (API)",             desc: "provider.desc.api_key" },
  { id: "google",             label: /* brand */ "Google Gemini (API)",      desc: "provider.desc.api_key" },
  { id: "deepseek",           label: /* brand */ "DeepSeek" },
  { id: "amazon-bedrock",     label: /* brand */ "Amazon Bedrock" },
  { id: "mistral",            label: /* brand */ "Mistral" },
  { id: "groq",               label: /* brand */ "Groq" },
  { id: "xai",                label: /* brand */ "xAI / Grok" },
];

/**
 * Providers to show in connect UIs. Equals ALL_PROVIDERS unless the build
 * sets VITE_PI_ALLOWED_PROVIDERS (org deployments — see docs/central-proxy.md).
 * UI-level filter only; enforcement lives at the proxy/network layer.
 */
export const VISIBLE_PROVIDERS: ProviderDef[] = filterProvidersByAllowlist(
  ALL_PROVIDERS,
  resolveAllowedProviderIds(
    typeof import.meta.env === "undefined" ? undefined : import.meta.env.VITE_PI_ALLOWED_PROVIDERS,
  ),
);

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
    const [code, state] = parts;
    if (code !== undefined && state !== undefined) return `${code}#${state}`;
  }

  return value;
}

function looksLikeOAuthRedirectInput(value: string): boolean {
  const lower = value.toLowerCase();
  return (
    value.includes("#")
    || value.includes("code=")
    || lower.startsWith("http://localhost:1455/")
    || lower.startsWith("http://localhost:53692/")
    || lower.startsWith("http://localhost:8085/")
    || lower.startsWith("http://localhost:51121/")
    || lower.startsWith("https://auth.openai.com/")
    || lower.startsWith("https://accounts.google.com/")
    || lower.includes("oauth2callback")
    || lower.includes("oauth-callback")
  );
}

function getOAuthStateFromAuthUrl(authUrl: string | null): string | undefined {
  if (!authUrl) return undefined;

  try {
    const state = new URL(authUrl).searchParams.get("state");
    return state && state.length > 0 ? state : undefined;
  } catch {
    return undefined;
  }
}

function normalizeApiKeyForProvider(
  providerId: string,
  raw: string,
): { ok: true; key: string } | { ok: false; error: string } {
  let key = raw.trim();
  if (!key) return { ok: false, error: t("provider.error.empty_key") };

  // Common copy/paste format: "Bearer <token>"
  if (/^bearer\s+/i.test(key)) {
    key = key.replace(/^bearer\s+/i, "").trim();
  }

  if (providerId === "anthropic") {
    // Prevent saving Anthropic OAuth *authorization code* (code#state) as an API key.
    // OAuth access tokens are sk-ant-oat*, API keys are sk-ant-api*.
    const looksLikeAuthCode = key.includes("#") && !key.includes("sk-ant-");
    if (looksLikeAuthCode) {
      return { ok: false, error: t("provider.error.oauth_code_as_key") };
    }
  }

  if (providerId === "openai-codex" && looksLikeOAuthRedirectInput(key)) {
    return { ok: false, error: t("provider.error.oauth_url_codex") };
  }

  if ((providerId === "google-gemini-cli" || providerId === "google-antigravity") && looksLikeOAuthRedirectInput(key)) {
    return { ok: false, error: t("provider.error.oauth_url_google") };
  }

  if (providerId === "google" && looksLikeOAuthRedirectInput(key)) {
    return { ok: false, error: t("provider.error.oauth_url_google_api") };
  }

  return { ok: true, key };
}

/**
 * Show a non-blocking dialog with a device-code login code (e.g. GitHub
 * Copilot). The login flow keeps polling in the background, so the dialog
 * only informs the user; the caller closes it when login settles.
 */
function showDeviceCodeDialog(info: { userCode: string; verificationUri: string }): () => void {
  closeOverlayById(PROVIDER_PROMPT_OVERLAY_ID);

  const dialog = createOverlayDialog({
    overlayId: PROVIDER_PROMPT_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-prompt-card",
    restoreFocusOnClose: false,
  });

  const title = document.createElement("h2");
  title.className = "pi-prompt-title";
  title.textContent = t("provider.device_code.title");

  const message = document.createElement("p");
  message.className = "pi-prompt-message";
  message.textContent = t("provider.device_code.message");

  const codeRow = document.createElement("div");
  codeRow.style.cssText = "display:flex;align-items:center;gap:8px;margin:12px 0;";

  const codeEl = document.createElement("code");
  codeEl.style.cssText =
    "flex:1;padding:8px 10px;border-radius:6px;text-align:center;" +
    "background:var(--pi-code-bg, #1e1e1e);color:var(--pi-code-fg, #d4d4d4);" +
    "font-size:16px;letter-spacing:2px;font-family:var(--pi-monospace, monospace);user-select:all;";
  codeEl.textContent = info.userCode;

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.textContent = t("provider.proxy_gate.copy");
  copyBtn.style.cssText = "padding:6px 12px;border-radius:6px;font-size:13px;cursor:pointer;";
  copyBtn.addEventListener("click", () => {
    void navigator.clipboard.writeText(info.userCode).then(() => {
      copyBtn.textContent = t("provider.proxy_gate.copied");
      setTimeout(() => { copyBtn.textContent = t("provider.proxy_gate.copy"); }, 1500);
    });
  });

  codeRow.append(codeEl, copyBtn);

  const helper = document.createElement("p");
  helper.className = "pi-prompt-helper";
  helper.textContent = t("provider.device_code.helper", { uri: info.verificationUri });

  dialog.card.append(title, message, codeRow, helper);
  dialog.mount();

  return () => dialog.close();
}

/**
 * Show a select dialog for OAuth flows that need the user to pick between
 * options (pi-ai `OAuthLoginCallbacks.onSelect`). Resolves the selected
 * option id, or `undefined` if the user cancels.
 */
function promptForSelect(opts: {
  title: string;
  message: string;
  options: { id: string; label: string }[];
}): Promise<string | undefined> {
  return new Promise((resolve) => {
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

    const optionList = document.createElement("div");
    optionList.style.cssText = "display:flex;flex-direction:column;gap:8px;margin:12px 0;";

    let settled = false;

    const settle = (value: string | undefined): void => {
      if (settled) return;
      settled = true;
      dialog.close();
      resolve(value);
    };

    for (const option of opts.options) {
      const optionBtn = document.createElement("button");
      optionBtn.type = "button";
      optionBtn.className = "pi-prompt-ok";
      optionBtn.textContent = option.label;
      optionBtn.addEventListener("click", () => settle(option.id));
      optionList.append(optionBtn);
    }

    const actions = document.createElement("div");
    actions.className = "pi-prompt-actions";

    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "pi-prompt-cancel";
    cancelBtn.textContent = t("provider.prompt.cancel");
    cancelBtn.addEventListener("click", () => settle(undefined));
    actions.append(cancelBtn);

    dialog.card.append(titleEl, messageEl, optionList, actions);

    dialog.addCleanup(() => {
      if (!settled) {
        settled = true;
        resolve(undefined);
      }
    });

    dialog.mount();
  });
}

async function copyTextToClipboard(text: string): Promise<void> {
  const clipboard = navigator.clipboard;
  if (clipboard?.writeText) {
    try {
      await clipboard.writeText(text);
      return;
    } catch {
      // Fall through to legacy copy for WebViews/insecure dev origins.
    }
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.setAttribute("readonly", "");
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  textarea.style.top = "0";
  document.body.append(textarea);
  textarea.select();
  const copied = document.execCommand("copy");
  textarea.remove();

  if (!copied) {
    throw new Error("Clipboard copy failed");
  }
}

function promptForText(opts: {
  title: string;
  message: string;
  placeholder?: string;
  helperText?: string;
  submitLabel?: string;
  externalUrl?: string;
  autoCapture?: {
    providerId: string;
    state: string;
  };
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

    const captureStatusEl = document.createElement("p");
    captureStatusEl.className = "pi-prompt-helper";
    captureStatusEl.hidden = true;

    const input = document.createElement("input");
    input.className = "pi-prompt-input";
    input.type = "text";
    input.autocomplete = "off";
    input.setAttribute("aria-label", opts.title);

    const externalUrlEl = document.createElement("div");
    externalUrlEl.className = "pi-prompt-external-url";
    externalUrlEl.hidden = true;

    const externalUrlFallbackEl = document.createElement("div");
    externalUrlFallbackEl.className = "pi-prompt-external-url__manual";
    externalUrlFallbackEl.hidden = true;

    const actions = document.createElement("div");
    actions.className = "pi-prompt-actions";

    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "pi-prompt-cancel";
    cancelBtn.textContent = t("provider.prompt.cancel");

    const okBtn = document.createElement("button");
    okBtn.type = "button";
    okBtn.className = "pi-prompt-ok";
    okBtn.textContent = opts.submitLabel ?? t("provider.prompt.continue");

    actions.append(cancelBtn, okBtn);
    dialog.card.append(
      titleEl,
      messageEl,
      helperEl,
      externalUrlEl,
      externalUrlFallbackEl,
      captureStatusEl,
      input,
      actions,
    );

    if (opts.helperText) {
      helperEl.textContent = opts.helperText;
      helperEl.hidden = false;
    }

    if (opts.placeholder) {
      input.placeholder = opts.placeholder;
    }

    if (opts.externalUrl) {
      const openLink = document.createElement("a");
      openLink.className = "pi-prompt-external-url__open";
      openLink.href = opts.externalUrl;
      openLink.target = "_blank";
      openLink.rel = "noopener noreferrer";
      openLink.textContent = t("provider.prompt.openLoginPage");

      const manualLabel = document.createElement("label");
      manualLabel.className = "pi-prompt-external-url__manual-label";
      manualLabel.textContent = t("provider.prompt.loginUrlFallback");

      const manualInput = document.createElement("textarea");
      manualInput.className = "pi-prompt-external-url__manual-value";
      manualInput.readOnly = true;
      manualInput.rows = 3;
      manualInput.value = opts.externalUrl;
      manualInput.setAttribute("aria-label", t("provider.prompt.loginUrlFallback"));
      const selectManualUrl = (): void => {
        manualInput.focus();
        manualInput.select();
      };
      manualInput.addEventListener("focus", selectManualUrl);
      manualInput.addEventListener("click", selectManualUrl);

      const copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.className = "pi-prompt-external-url__copy";
      copyBtn.textContent = t("provider.prompt.copyLoginLink");
      copyBtn.addEventListener("click", () => {
        void copyTextToClipboard(opts.externalUrl ?? "")
          .then(() => {
            copyBtn.textContent = t("provider.prompt.copied");
            setTimeout(() => {
              copyBtn.textContent = t("provider.prompt.copyLoginLink");
            }, 1500);
          })
          .catch(() => {
            copyBtn.textContent = t("provider.prompt.copyFailedShort");
            selectManualUrl();
            setTimeout(() => {
              copyBtn.textContent = t("provider.prompt.copyLoginLink");
            }, 2000);
          });
      });

      externalUrlFallbackEl.append(manualLabel, manualInput);
      externalUrlEl.append(openLink, copyBtn);
      externalUrlEl.hidden = false;
      externalUrlFallbackEl.hidden = false;
    }

    let settled = false;
    let captureAbortController: AbortController | undefined;

    const submit = (): void => {
      if (settled) {
        return;
      }

      captureAbortController?.abort();
      settled = true;
      const value = input.value.trim();
      dialog.close();
      resolve(value);
    };

    const cancel = (): void => {
      if (settled) {
        return;
      }

      captureAbortController?.abort();
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
      captureAbortController?.abort();

      if (!settled) {
        settled = true;
        reject(new PromptCancelledError());
      }
    });

    dialog.mount();

    if (opts.autoCapture) {
      captureStatusEl.textContent = t("provider.oauth.capture_waiting");
      captureStatusEl.hidden = false;
      captureAbortController = new AbortController();

      void pollOAuthCallbackCapture({
        providerId: opts.autoCapture.providerId,
        state: opts.autoCapture.state,
        signal: captureAbortController.signal,
      }).then(
        (capture) => {
          if (settled) return;
          if (!capture) {
            captureStatusEl.textContent = t("provider.oauth.capture_fallback");
            return;
          }

          captureStatusEl.textContent = t("provider.oauth.capture_captured");
          input.value = capture.url;
          submit();
        },
        (error: DynamicValue) => {
          if (settled) return;
          if (error instanceof Error && error.name === "AbortError") return;
          captureStatusEl.textContent = t("provider.oauth.capture_fallback");
        },
      );
    }

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
    ? t("provider.placeholder.anthropic")
    : id === "openai-codex"
      ? t("provider.placeholder.chatgpt")
      : id === "google-gemini-cli" || id === "google-antigravity"
        ? t("provider.placeholder.google_oauth")
        : t("provider.placeholder.api_key");

  const row = document.createElement("div");
  row.className = "pi-login-row";
  const labelText = escapeHtml(label);
  const descriptionMarkup = desc ? `<span class="pi-login-desc">${escapeHtml(t(desc))}</span>` : "";
  const oauthMarkup = oauth
    ? `
        <button class="pi-login-oauth">${escapeHtml(t("provider.login_with", { label }))}</button>
        <div class="pi-login-divider">
          <div class="pi-login-divider__line"></div>
          <span class="pi-login-divider__text">${escapeHtml(t("provider.or_api_key"))}</span>
          <div class="pi-login-divider__line"></div>
        </div>
      `
    : "";
  setSafeInnerHTML(
    row,
    `
    <button class="pi-welcome-provider pi-login-trigger">
      <span class="pi-login-meta">
        <span class="pi-login-label">${labelText}</span>
        ${descriptionMarkup}
      </span>
      <span class="pi-login-status ${isActive ? "is-connected" : ""}">
        ${escapeHtml(isActive ? t("provider.connected") : t("provider.set_up"))}
      </span>
    </button>
    <div class="pi-login-detail" hidden>
      <button class="pi-login-disconnect" type="button" ${isActive ? "" : "hidden"}>${escapeHtml(t("provider.disconnect", { label }))}</button>
      ${oauthMarkup}
      <div class="pi-login-key-row">
        <input class="pi-login-key" type="password" placeholder="${escapeAttr(keyPlaceholder)}" aria-label="${escapeAttr(t("provider.keyAria", { label }))}" autocomplete="off" spellcheck="false" />
        <button class="pi-login-save">${escapeHtml(t("provider.save"))}</button>
      </div>
      <p class="pi-login-error" hidden></p>
    </div>
  `,
    "provider login row template with escaped provider and localized text",
  );

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
      statusEl.textContent = connected ? t("provider.connected") : t("provider.set_up");
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
      oauthBtn.textContent = t("provider.opening_login");
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

          // In production, OAuth providers need the local CORS proxy.
          // Check reachability before sending the user through the browser login.
          if (!import.meta.env.DEV && OAUTH_IDS_NEEDING_PROXY.has(id)) {
            const reachable = await isProxyReachable();
            if (!reachable) {
              const userRetried = await showProxyGateDialog();
              if (!userRetried) {
                // User cancelled — reset button and bail.
                oauthBtn.textContent = t("provider.login_with", { label });
                oauthBtn.style.opacity = "1";
                return;
              }
            }
          }

          const deviceCodeDialogRef: { close: (() => void) | null } = { close: null };
          const authUrlRef: { current: string | null } = { current: null };

          let cred;
          try {
            cred = await oauthProvider.login({
              onAuth: (info) => {
                authUrlRef.current = info.url;
                // Prevent the OAuth page from gaining a handle to the add-in window.
                const w = window.open(info.url, "_blank", "noopener,noreferrer");
                if (w) w.opener = null;
              },
              onDeviceCode: (info) => {
                // Device-code flows (e.g. GitHub Copilot): open the verification
                // page and show the user code to enter there.
                const w = window.open(info.verificationUri, "_blank", "noopener,noreferrer");
                if (w) w.opener = null;
                deviceCodeDialogRef.close = showDeviceCodeDialog(info);
              },
              onSelect: (prompt) =>
                promptForSelect({
                  title: t("provider.login_with", { label }),
                  message: prompt.message,
                  options: prompt.options,
                }),
              onPrompt: async (prompt) => {
                const helperText = id === "anthropic"
                  ? t("provider.oauth.helper.anthropic")
                  : id === "openai-codex"
                    ? t("provider.oauth.helper.openai")
                    : id === "google-gemini-cli" || id === "google-antigravity"
                      ? t("provider.oauth.helper.google")
                      : undefined;

                const oauthCallbackProviderId = oauth && OAUTH_CALLBACK_CAPTURE_IDS.has(oauth)
                  ? oauth
                  : undefined;
                const oauthCallbackState = oauthCallbackProviderId
                  ? getOAuthStateFromAuthUrl(authUrlRef.current)
                  : undefined;

                const value = await promptForText({
                  title: t("provider.login_with", { label }),
                  message: prompt.message,
                  placeholder: prompt.placeholder || "",
                  ...(helperText !== undefined ? { helperText } : {}),
                  submitLabel: t("provider.prompt.continue"),
                  ...(authUrlRef.current !== null ? { externalUrl: authUrlRef.current } : {}),
                  ...(oauthCallbackProviderId && oauthCallbackState
                    ? { autoCapture: { providerId: oauthCallbackProviderId, state: oauthCallbackState } }
                    : {}),
                });

                if (id === "anthropic") {
                  return normalizeAnthropicAuthorizationInput(value);
                }

                return value;
              },
              onProgress: (msg) => { oauthBtn.textContent = msg; },
            });
          } finally {
            deviceCodeDialogRef.close?.();
          }

          const apiKey = oauthProvider.getApiKey(cred);
          await storage.providerKeys.set(id, apiKey);
          await saveOAuthCredentials(storage.settings, id, cred);
          setConnectedState(true);
          onConnected(row, id, label);
          detail.hidden = true;
          expandedRef.current = null;
        } catch (err) {
          if (err instanceof PromptCancelledError) {
            // User cancelled the prompt; leave UI unchanged.
            return;
          }

          const msg = getErrorMessage(err);
          const isLikelyCors =
            isCorsError(err) ||
            (typeof msg === "string" && /load failed|failed to fetch|cors|cross-origin|networkerror/i.test(msg));

          if (isLikelyCors) {
            if (DEFAULT_PROXY_IS_REMOTE) {
              errorEl.textContent = t("provider.cors_error_remote", { url: DEFAULT_PROXY_URL });
            } else {
              setSafeInnerHTML(
                errorEl,
                `${escapeHtml(t("provider.cors_error"))} <code style="padding:2px 5px;border-radius:4px;` +
                  "background:var(--pi-code-bg, #1e1e1e);color:var(--pi-code-fg, #d4d4d4)\">" +
                  `npx pi-for-excel-proxy</code>${escapeHtml(t("provider.cors_error.retry"))} ` +
                  `<a href="${escapeAttr(PROXY_HELPER_DOCS_URL)}" target="_blank" rel="noopener noreferrer">${escapeHtml(t("provider.proxy_gate.guide"))}</a>`,
                "provider CORS helper error markup with escaped localized text",
              );
            }
          } else {
            errorEl.textContent = msg || t("provider.login_failed");
          }
          errorEl.hidden = false;
        } finally {
          oauthBtn.textContent = t("provider.login_with", { label });
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
        disconnectBtn.textContent = t("provider.disconnecting");
        disconnectBtn.disabled = true;
        disconnectBtn.style.opacity = "0.7";
        errorEl.hidden = true;

        try {
          await storage.providerKeys.delete(id);
          await clearOAuthCredentials(storage.settings, id);

          setConnectedState(false);
          keyInput.value = "";
          onDisconnected?.(row, id, label);
        } catch (err) {
          const msg = getErrorMessage(err);
          errorEl.textContent = msg ? t("provider.disconnect_failed_msg", { msg }) : t("provider.disconnect_failed");
          errorEl.hidden = false;
        } finally {
          disconnectBtn.textContent = t("provider.disconnect", { label });
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
    saveBtn.textContent = t("provider.testing");
    saveBtn.style.opacity = "0.7";
    errorEl.hidden = true;
    try {
      await storage.providerKeys.set(id, key);
      setConnectedState(true);
      onConnected(row, id, label);
      detail.hidden = true;
      expandedRef.current = null;
    } catch (err) {
      const msg = getErrorMessage(err);
      errorEl.textContent = msg ? t("provider.save_failed_msg", { msg }) : t("provider.save_failed");
      errorEl.hidden = false;
    } finally {
      saveBtn.textContent = t("provider.save");
      saveBtn.style.opacity = "1";
    }
  })(); });

  // Enter key in input
  keyInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") saveBtn.click();
  });

  return row;
}
