import { getAppStorage } from "../storage/local/app-storage.js";
import { checkProviderApiKey } from "../auth/api-key-check.js";
import { showToast } from "./toast.js";
import { isMacPlatform, isWindowsPlatform } from "../utils/platform.js";

function terminalAppName(): string {
  if (isMacPlatform()) return t("platform.terminal.mac");
  if (isWindowsPlatform()) return t("platform.terminal.windows");
  return t("platform.terminal.other");
}
import { isCorsError, isModuleLoadError } from "../auth/cors-error.js";
import { checkProxyOnce, isProxyBlockedByBrowser } from "../taskpane/proxy-status.js";
import { pollOAuthCallbackCapture } from "../auth/oauth-callback-capture.js";
import { getOAuthProvider } from "../auth/oauth-provider-registry.js";
import { withOAuthProviderLock } from "../auth/oauth-refresh.js";
import { clearOAuthCredentials, saveOAuthCredentials } from "../auth/oauth-storage.js";
import {
  DEFAULT_PROXY_IS_REMOTE,
  DEFAULT_PROXY_URL,
  PROXY_HELPER_COMMAND,
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

function showProxyGateDialog(signal: AbortSignal): Promise<boolean> {
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
      : t("provider.proxy_gate.message", { terminal: terminalAppName() });

    const codeRow = document.createElement("div");
    codeRow.style.cssText = "display:flex;align-items:center;gap:8px;margin:12px 0;";

    const codeEl = document.createElement("code");
    codeEl.style.cssText =
      "flex:1;padding:8px 10px;border-radius:6px;" +
      "background:var(--pi-code-bg, #1e1e1e);color:var(--pi-code-fg, #d4d4d4);" +
      "font-size:13px;font-family:var(--pi-monospace, monospace);user-select:all;";
    codeEl.textContent = PROXY_HELPER_COMMAND;

    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.textContent = t("provider.proxy_gate.copy");
    copyBtn.style.cssText = "padding:6px 12px;border-radius:6px;font-size:13px;cursor:pointer;";
    copyBtn.addEventListener("click", () => {
      void navigator.clipboard.writeText(PROXY_HELPER_COMMAND).then(() => {
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
        `${escapeHtml(t("provider.proxy_gate.hint_html", { terminal: terminalAppName() }))} ` +
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

    let phase: "idle" | "probing" | "committing" | "settled" = "idle";

    const doCancel = (): void => {
      if (phase === "settled") return;
      phase = "settled";
      dialog.close();
      resolve(false);
    };

    const doRetry = (): void => {
      if (phase !== "idle") return;
      phase = "probing";
      retryBtn.disabled = true;
      retryBtn.textContent = t("provider.proxy_gate.checking");
      retryBtn.style.opacity = "0.7";

      void (async () => {
        try {
          const storage = getAppStorage();
          const url = await storage.settings.get("proxy.url");
          if (signal.aborted || phase !== "probing") return;
          const proxyUrl = resolveConfiguredProxyUrl(url);
          const ok = await probeProxyReachability(proxyUrl, 1500);

          if (signal.aborted || phase !== "probing") return;
          if (ok) {
            phase = "committing";
            await storage.settings.set("proxy.enabled", true);
            if (signal.aborted) return;
            phase = "settled";
            dialog.close();
            resolve(true);
            return;
          }
        } catch {
        }

        if (signal.aborted || phase === "settled") return;
        phase = "idle";
        retryBtn.disabled = false;
        retryBtn.textContent = t("provider.proxy_gate.retry");
        retryBtn.style.opacity = "1";
        if (DEFAULT_PROXY_IS_REMOTE) {
          hint.textContent = t("provider.proxy_gate.not_detected_remote", { url: DEFAULT_PROXY_URL });
        } else {
          setSafeInnerHTML(
            hint,
            `${escapeHtml(t("provider.proxy_gate.not_detected_html", { terminal: terminalAppName() }))} ` +
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
      signal.removeEventListener("abort", doCancel);
      if (phase !== "settled") { phase = "settled"; resolve(false); }
    });

    signal.addEventListener("abort", doCancel, { once: true });
    dialog.mount();
    if (signal.aborted) doCancel();
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
  brand: string;
  oauth?: string;
  desc?: string;
}

export const ALL_PROVIDERS: ProviderDef[] = [
  { id: "openai", brand: "OpenAI (API)", desc: "provider.desc.api_key" },
  { id: "google", brand: "Google Gemini (API)", desc: "provider.desc.api_key" },
  { id: "deepseek", brand: "DeepSeek" },
  { id: "mistral", brand: "Mistral" },
  { id: "groq", brand: "Groq" },
  { id: "xai", brand: "xAI / Grok" },
  { id: "anthropic",          brand: "Anthropic",                oauth: "anthropic",          desc: "provider.desc.claude" },
  { id: "openai-codex",       brand: "OpenAI (ChatGPT)",         oauth: "openai-codex",       desc: "provider.desc.openai_sub" },
  { id: "google-gemini-cli",  brand: "Google Code Assist",       oauth: "google-gemini-cli",  desc: "provider.desc.gemini_account" },
  { id: "google-antigravity", brand: "Google Antigravity",       oauth: "google-antigravity", desc: "provider.desc.antigravity" },
  { id: "github-copilot",     brand: "GitHub Copilot",           oauth: "github-copilot" },

  { id: "amazon-bedrock",     brand: "Amazon Bedrock" },
];

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

  try {
    const url = new URL(value);
    const code = url.searchParams.get("code");
    const state = url.searchParams.get("state");
    if (code) return state ? `${code}#${state}` : code;
  } catch {
  }

  if (value.includes("code=")) {
    try {
      const params = new URLSearchParams(value.startsWith("?") ? value.slice(1) : value);
      const code = params.get("code");
      const state = params.get("state");
      if (code) return state ? `${code}#${state}` : code;
    } catch {
    }
  }

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

  if (/^bearer\s+/i.test(key)) {
    key = key.replace(/^bearer\s+/i, "").trim();
  }

  if (providerId === "anthropic") {
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

function showDeviceCodeDialog(
  info: { userCode: string; verificationUri: string },
  onClose: () => void,
): () => void {
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
  let completed = false;
  dialog.addCleanup(() => { if (!completed) onClose(); });
  dialog.mount();

  return () => {
    completed = true;
    dialog.close();
  };
}

function promptForSelect(opts: {
  signal: AbortSignal;
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

    const onAbort = (): void => settle(undefined);
    opts.signal.addEventListener("abort", onAbort, { once: true });
    dialog.addCleanup(() => {
      opts.signal.removeEventListener("abort", onAbort);
      if (!settled) {
        settled = true;
        resolve(undefined);
      }
    });

    dialog.mount();
    if (opts.signal.aborted) onAbort();
  });
}

async function copyTextToClipboard(text: string): Promise<void> {
  const clipboard = navigator.clipboard;
  if (clipboard?.writeText) {
    try {
      await clipboard.writeText(text);
      return;
    } catch {
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
  signal: AbortSignal;
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

    opts.signal.addEventListener("abort", cancel, { once: true });
    cancelBtn.addEventListener("click", cancel);
    okBtn.addEventListener("click", submit);
    input.addEventListener("keydown", onInputKeyDown);

    dialog.addCleanup(() => {
      opts.signal.removeEventListener("abort", cancel);
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
    if (opts.signal.aborted) {
      cancel();
      return;
    }

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
        (error: unknown) => {
          if (settled) return;
          if (error instanceof Error && error.name === "AbortError") return;
          captureStatusEl.textContent = t("provider.oauth.capture_fallback");
        },
      );
    }

    requestAnimationFrame(() => input.focus());
  });
}

export function buildProviderRow(
  provider: ProviderDef,
  opts: {
    isActive: boolean;
    signal: AbortSignal;
    expandedRef: { current: HTMLElement | null };
  } & ProviderRowCallbacks,
): HTMLElement {
  const { id, brand: label, oauth, desc } = provider;
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
  const requirement = t(id === "amazon-bedrock"
    ? "provider.requirement.aws"
    : oauth ? "provider.requirement.helper" : "provider.requirement.key");
  const descriptionMarkup = `<span class="pi-login-desc">${escapeHtml([desc ? t(desc) : "", requirement].filter(Boolean).join(" · "))}</span>`;
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

  type Operation = { kind: "logging-in" | "saving-key" | "disconnecting"; controller: AbortController };
  type RowState = { kind: "idle" } | Operation | { kind: "disposed" };
  let state: RowState = { kind: "idle" };

  const transition = (next: RowState): void => {
    state = next;
    const busy = next.kind !== "idle";
    saveBtn.disabled = busy;
    saveBtn.textContent = t(next.kind === "saving-key" ? "provider.testing" : "provider.save");
    saveBtn.style.opacity = next.kind === "saving-key" ? "0.7" : "1";
    keyInput.disabled = busy;
    if (oauthBtn) {
      oauthBtn.disabled = busy;
      oauthBtn.textContent = next.kind === "logging-in" ? t("provider.opening_login") : t("provider.login_with", { label });
      oauthBtn.style.opacity = next.kind === "logging-in" ? "0.7" : "1";
    }
    if (disconnectBtn) {
      disconnectBtn.disabled = next.kind === "disconnecting" || next.kind === "disposed";
      disconnectBtn.textContent = next.kind === "disconnecting" ? t("provider.disconnecting") : t("provider.disconnect", { label });
      disconnectBtn.style.opacity = next.kind === "disconnecting" ? "0.7" : "1";
    }
    headerBtn.disabled = next.kind === "disposed";
  };
  const begin = (kind: Operation["kind"]): Operation | undefined => {
    if (state.kind === "disposed" || state.kind === "disconnecting") return undefined;
    if (state.kind !== "idle") {
      if (kind !== "disconnecting") return undefined;
      state.controller.abort();
    }
    const operation = { kind, controller: new AbortController() };
    transition(operation);
    return operation;
  };
  const finish = (operation: Operation): void => {
    if (state === operation) transition({ kind: "idle" });
  };
  const collapseOwnDetail = (): void => {
    detail.hidden = true;
    if (expandedRef.current === detail) expandedRef.current = null;
  };
  const dispose = (): void => {
    if (state.kind !== "idle" && state.kind !== "disposed") state.controller.abort();
    transition({ kind: "disposed" });
    if (expandedRef.current === detail) expandedRef.current = null;
  };
  opts.signal.addEventListener("abort", dispose, { once: true });
  if (opts.signal.aborted) dispose();
  setConnectedState(isActive);

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

  if (oauthBtn) {
    oauthBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const operation = begin("logging-in");
      if (!operation) return;
      const loginAbort = operation.controller;
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

          if (!import.meta.env.DEV && OAUTH_IDS_NEEDING_PROXY.has(id)) {
            const reachable = await isProxyReachable();
            if (!reachable) {
              loginAbort.signal.throwIfAborted();
              await checkProxyOnce(getAppStorage().settings);
              loginAbort.signal.throwIfAborted();
              if (isProxyBlockedByBrowser()) {
                errorEl.textContent = t("proxy.blocked_by_browser");
                errorEl.hidden = false;
                return;
              }
              const userRetried = await showProxyGateDialog(loginAbort.signal);
              if (!userRetried) {
                return;
              }
            }
          }

          const deviceCodeDialogRef: { close: (() => void) | null } = { close: null };
          const authUrlRef: { current: string | null } = { current: null };

          loginAbort.signal.throwIfAborted();
          let cred;
          try {
            cred = await oauthProvider.login({
              onAuth: (info) => {
                loginAbort.signal.throwIfAborted();
                authUrlRef.current = info.url;
                const w = window.open(info.url, "_blank", "noopener,noreferrer");
                if (w) w.opener = null;
              },
              onDeviceCode: (info) => {
                loginAbort.signal.throwIfAborted();
                const w = window.open(info.verificationUri, "_blank", "noopener,noreferrer");
                if (w) w.opener = null;
                deviceCodeDialogRef.close = showDeviceCodeDialog(info, () => loginAbort.abort());
                loginAbort.signal.addEventListener("abort", () => deviceCodeDialogRef.close?.(), { once: true });
              },
              onSelect: (prompt) =>
                promptForSelect({
                  signal: loginAbort.signal,
                  title: t("provider.login_with", { label }),
                  message: prompt.message,
                  options: prompt.options,
                }),
              onPrompt: async (prompt) => {
                loginAbort.signal.throwIfAborted();
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
                  signal: loginAbort.signal,
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
              onProgress: (msg) => { if (!loginAbort.signal.aborted) oauthBtn.textContent = msg; },
              signal: loginAbort.signal,
            });
          } catch (error) {
            if (loginAbort.signal.aborted) {
              return;
            }
            throw error;
          } finally {
            deviceCodeDialogRef.close?.();
          }

          loginAbort.signal.throwIfAborted();
          const apiKey = oauthProvider.getApiKey(cred);
          await withOAuthProviderLock(oauth, async () => {
            loginAbort.signal.throwIfAborted();
            await saveOAuthCredentials(storage.settings, oauth, cred);
            await storage.providerKeys.set(id, apiKey);
          });
          loginAbort.signal.throwIfAborted();
          setConnectedState(true);
          onConnected(row, id, label);
          collapseOwnDetail();
        } catch (err) {
          if (loginAbort.signal.aborted || err instanceof PromptCancelledError) {
            return;
          }

          const msg = getErrorMessage(err);
          const flowUnavailable = err instanceof Error && isModuleLoadError(err);
          const isLikelyCors = !flowUnavailable && (
            isCorsError(err) ||
            (typeof msg === "string" && /load failed|failed to fetch|cors|cross-origin|networkerror/i.test(msg)));

          if (isLikelyCors) {
            if (isProxyBlockedByBrowser()) {
              errorEl.textContent = t("proxy.blocked_by_browser");
            } else if (DEFAULT_PROXY_IS_REMOTE) {
              errorEl.textContent = t("provider.cors_error_remote", { url: DEFAULT_PROXY_URL });
            } else {
              setSafeInnerHTML(
                errorEl,
                `${escapeHtml(t("provider.cors_error", { terminal: terminalAppName() }))} <code style="padding:2px 5px;border-radius:4px;` +
                  "background:var(--pi-code-bg, #1e1e1e);color:var(--pi-code-fg, #d4d4d4)\">" +
                  `${escapeHtml(PROXY_HELPER_COMMAND)}</code>${escapeHtml(t("provider.cors_error.retry"))} ` +
                  `<a href="${escapeAttr(PROXY_HELPER_DOCS_URL)}" target="_blank" rel="noopener noreferrer">${escapeHtml(t("provider.proxy_gate.guide"))}</a>`,
                "provider CORS helper error markup with escaped localized text",
              );
            }
          } else if (flowUnavailable) {
            errorEl.textContent = t("provider.oauth.flow_unavailable", { label });
          } else {
            errorEl.textContent = msg || t("provider.login_failed");
          }
          errorEl.hidden = false;
        } finally {
          finish(operation);
        }
      })();
    });
  }

  if (disconnectBtn) {
    disconnectBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const operation = begin("disconnecting");
      if (!operation) return;
      void (async () => {
        errorEl.hidden = true;

        try {
          await withOAuthProviderLock(oauth ?? id, async () => {
            operation.controller.signal.throwIfAborted();
            await clearOAuthCredentials(storage.settings, oauth ?? id);
            await storage.providerKeys.delete(id);
          });
          operation.controller.signal.throwIfAborted();
          setConnectedState(false);
          keyInput.value = "";
          onDisconnected?.(row, id, label);
        } catch (err) {
          if (operation.controller.signal.aborted) return;
          const msg = getErrorMessage(err);
          errorEl.textContent = msg ? t("provider.disconnect_failed_msg", { msg }) : t("provider.disconnect_failed");
          errorEl.hidden = false;
        } finally {
          finish(operation);
        }
      })();
    });
  }

  saveBtn.addEventListener("click", () => { void (async () => {
    const rawKey = keyInput.value.trim();
    if (!rawKey) return;

    const normalized = normalizeApiKeyForProvider(id, rawKey);
    if (!normalized.ok) {
      errorEl.textContent = normalized.error;
      errorEl.hidden = false;
      return;
    }

    const operation = begin("saving-key");
    if (!operation) return;
    const key = normalized.key;
    errorEl.hidden = true;
    try {
      const check = await checkProviderApiKey(id, key, fetch, operation.controller.signal);
      operation.controller.signal.throwIfAborted();
      if (check.status === "rejected") {
        errorEl.textContent = t("provider.key_rejected", { label, status: String(check.httpStatus) });
        errorEl.hidden = false;
        return;
      }

      await withOAuthProviderLock(oauth ?? id, async () => {
        operation.controller.signal.throwIfAborted();
        await clearOAuthCredentials(storage.settings, oauth ?? id);
        await storage.providerKeys.set(id, key);
      });
      operation.controller.signal.throwIfAborted();
      setConnectedState(true);
      onConnected(row, id, label);
      if (check.status === "unverified") {
        showToast(t("provider.key_unverified", { label }), 4000);
      }
      collapseOwnDetail();
    } catch (err) {
      if (operation.controller.signal.aborted) return;
      const msg = getErrorMessage(err);
      errorEl.textContent = msg ? t("provider.save_failed_msg", { msg }) : t("provider.save_failed");
      errorEl.hidden = false;
    } finally {
      finish(operation);
    }
  })(); });

  keyInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") saveBtn.click();
  });

  return row;
}
