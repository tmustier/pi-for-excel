/**
 * Inline setup card for web search failures.
 *
 * Rendered inside the chat stream after a failed `web_search` tool card.
 * Detects the failure mode (missing key, proxy down, or both) and shows
 * contextual setup steps with inline actions.
 */

import { probeProxyReachability } from "../auth/proxy-validation.js";
import { DEFAULT_LOCAL_PROXY_URL } from "../auth/proxy-validation.js";
import { getEnabledProxyBaseUrl } from "../tools/external-fetch.js";
import {
  checkApiKeyFormat,
  getApiKeyForProvider,
  isApiKeyRequired,
  loadWebSearchProviderConfig,
  saveWebSearchApiKey,
  WEB_SEARCH_PROVIDER_INFO,
  WEB_SEARCH_PROVIDERS,
  type WebSearchConfigStore,
  type WebSearchProvider,
  type WebSearchProviderConfig,
} from "../tools/web-search-config.js";
import { validateWebSearchApiKey } from "../tools/web-search.js";
import type { WebSearchDetails } from "../tools/tool-details.js";
import { AlertTriangle, Check, Copy, Search, lucide } from "./lucide-icons.js";
import { showToast } from "./toast.js";

/* ── Types ──────────────────────────────────────────────────── */

type SetupMode =
  | { type: "needs_key" }
  | { type: "needs_proxy" }
  | { type: "needs_both" }
  | { type: "wrong_provider"; availableProvider: WebSearchProvider }
  | { type: "generic_error" };

interface SetupContext {
  mode: SetupMode;
  provider: WebSearchProvider;
  providerConfig: WebSearchProviderConfig;
  proxyBaseUrl: string | undefined;
}

/* ── Constants ──────────────────────────────────────────────── */

const PROXY_COMMAND = "npx pi-for-excel-proxy";

/* ── Helpers ────────────────────────────────────────────────── */

function selectElementText(element: HTMLElement): void {
  const selection = window.getSelection();
  if (!selection) return;
  const range = document.createRange();
  range.selectNodeContents(element);
  selection.removeAllRanges();
  selection.addRange(range);
}

function copyToClipboard(text: string, onCopied: () => void, fallbackElement: HTMLElement): void {
  if (!navigator.clipboard?.writeText) {
    selectElementText(fallbackElement);
    return;
  }
  void navigator.clipboard.writeText(text).then(onCopied, () => selectElementText(fallbackElement));
}

/**
 * Find a provider that has a configured API key, other than the current one.
 */
function findAlternativeProvider(
  config: WebSearchProviderConfig,
  currentProvider: WebSearchProvider,
): WebSearchProvider | undefined {
  for (const provider of WEB_SEARCH_PROVIDERS) {
    if (provider === currentProvider) continue;
    if (getApiKeyForProvider(config, provider)) return provider;
  }
  return undefined;
}

/* ── Detection ──────────────────────────────────────────────── */

async function detectSetupMode(
  details: WebSearchDetails,
  settings: WebSearchConfigStore,
): Promise<SetupContext> {
  const [providerConfig, proxyBaseUrl] = await Promise.all([
    loadWebSearchProviderConfig(settings),
    getEnabledProxyBaseUrl(settings),
  ]);

  const provider = (WEB_SEARCH_PROVIDERS as readonly string[]).includes(details.provider)
    ? details.provider as WebSearchProvider
    : providerConfig.provider;

  const hasKey = Boolean(getApiKeyForProvider(providerConfig, provider));
  const needsKey = !hasKey && isApiKeyRequired(provider);
  const isProxyDown = details.proxyDown === true;

  // When the key is missing, also eagerly probe the proxy so we can show both
  // steps at once instead of surprising the user with a second failure.
  let proxyReachable = !isProxyDown;
  if (needsKey && !isProxyDown) {
    const probeUrl = proxyBaseUrl ?? DEFAULT_LOCAL_PROXY_URL;
    proxyReachable = await probeProxyReachability(probeUrl, 1500);
  }

  const needsProxy = !proxyReachable;

  // Check for "wrong provider" — selected provider has no key but another does
  if (needsKey) {
    const alternative = findAlternativeProvider(providerConfig, provider);
    if (alternative && !needsProxy) {
      return {
        mode: { type: "wrong_provider", availableProvider: alternative },
        provider,
        providerConfig,
        proxyBaseUrl,
      };
    }
  }

  let mode: SetupMode;
  if (needsKey && needsProxy) {
    mode = { type: "needs_both" };
  } else if (needsKey) {
    mode = { type: "needs_key" };
  } else if (needsProxy) {
    mode = { type: "needs_proxy" };
  } else {
    mode = { type: "generic_error" };
  }

  return { mode, provider, providerConfig, proxyBaseUrl };
}

/* ── DOM construction ───────────────────────────────────────── */

function createCopyableCommand(command: string): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "pi-search-setup__code";

  const code = document.createElement("code");
  code.textContent = command;

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "pi-search-setup__copy";
  copyBtn.replaceChildren(lucide(Copy));
  copyBtn.title = "Copy command";

  let resetTimeout: ReturnType<typeof setTimeout> | null = null;

  copyBtn.addEventListener("click", () => {
    copyToClipboard(command, () => {
      copyBtn.replaceChildren(lucide(Check));
      copyBtn.title = "Copied";
      if (resetTimeout) clearTimeout(resetTimeout);
      resetTimeout = setTimeout(() => {
        copyBtn.replaceChildren(lucide(Copy));
        copyBtn.title = "Copy command";
        resetTimeout = null;
      }, 1400);
    }, code);
  });

  row.append(code, copyBtn);
  return row;
}

function createProxyStep(stepNumber: number | null): HTMLDivElement {
  const step = document.createElement("div");
  step.className = "pi-search-setup__step";

  const label = document.createElement("p");
  label.className = "pi-search-setup__step-label";
  label.textContent = stepNumber !== null
    ? `Step ${stepNumber} · Start the helper (keep it running):`
    : "Start the helper (keep it running):";

  const hint = document.createElement("p");
  hint.className = "pi-search-setup__hint";
  hint.textContent = "Open Terminal · paste · press Enter · leave the window open";

  step.append(label, createCopyableCommand(PROXY_COMMAND), hint);
  return step;
}

function createKeyStep(
  provider: WebSearchProvider,
  stepNumber: number | null,
  settings: WebSearchConfigStore,
  proxyBaseUrl: string | undefined,
  onSaved: () => void,
): HTMLDivElement {
  const info = WEB_SEARCH_PROVIDER_INFO[provider];

  const step = document.createElement("div");
  step.className = "pi-search-setup__step";

  const label = document.createElement("p");
  label.className = "pi-search-setup__step-label";
  label.textContent = stepNumber !== null
    ? `Step ${stepNumber} · Set up a ${info.title} API key:`
    : `Set up a ${info.title} API key:`;

  const signupLink = document.createElement("a");
  signupLink.className = "pi-search-setup__link";
  signupLink.href = info.signupUrl;
  signupLink.target = "_blank";
  signupLink.rel = "noopener noreferrer";
  signupLink.textContent = `Get a free key at ${info.signupUrl.replace(/^https?:\/\//u, "")} ↗`;

  const inputRow = document.createElement("div");
  inputRow.className = "pi-search-setup__input-row";

  const input = document.createElement("input");
  input.type = "password";
  input.className = "pi-search-setup__input";
  input.placeholder = info.apiKeyLabel;
  input.autocomplete = "off";

  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "pi-search-setup__save";
  saveBtn.textContent = "Save";

  const status = document.createElement("span");
  status.className = "pi-search-setup__status";

  let saving = false;

  saveBtn.addEventListener("click", () => {
    if (saving) return;
    const key = input.value.trim();
    if (!key) {
      showToast("Enter an API key first.");
      return;
    }

    const formatWarning = checkApiKeyFormat(provider, key);
    if (formatWarning) {
      status.textContent = `⚠️ ${formatWarning}`;
      status.className = "pi-search-setup__status is-warn";
    }

    saving = true;
    saveBtn.disabled = true;
    status.textContent = "Saving…";
    status.className = "pi-search-setup__status";

    void (async () => {
      try {
        await saveWebSearchApiKey(settings, provider, key);

        // Validate the key
        status.textContent = "Validating…";
        const result = await validateWebSearchApiKey({ provider, apiKey: key, proxyBaseUrl });

        if (result.ok) {
          status.textContent = `✓ ${result.message}`;
          status.className = "pi-search-setup__status is-ok";
          input.value = "";
          onSaved();
        } else {
          status.textContent = `Key saved. Validation: ${result.message}`;
          status.className = "pi-search-setup__status is-warn";
        }
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        status.textContent = `Error: ${msg}`;
        status.className = "pi-search-setup__status is-error";
      } finally {
        saving = false;
        saveBtn.disabled = false;
      }
    })();
  });

  inputRow.append(input, saveBtn);
  step.append(label, signupLink, inputRow, status);
  return step;
}

/* ── Card assembly ──────────────────────────────────────────── */

function buildCardContent(
  ctx: SetupContext,
  settings: WebSearchConfigStore,
  onDismiss: () => void,
): { title: string; body: DocumentFragment } {
  const body = document.createDocumentFragment();
  const { mode, provider, proxyBaseUrl } = ctx;

  const markDone = (): void => {
    showToast("✓ Web search is ready — ask the assistant to try again.");
    onDismiss();
  };

  switch (mode.type) {
    case "needs_both": {
      body.append(
        createProxyStep(1),
        createKeyStep(provider, 2, settings, proxyBaseUrl, markDone),
      );
      return { title: "Web search needs setup", body };
    }
    case "needs_key": {
      body.append(createKeyStep(provider, null, settings, proxyBaseUrl, markDone));
      return { title: "Web search needs an API key", body };
    }
    case "needs_proxy": {
      body.append(createProxyStep(null));
      return { title: "Web search can't connect", body };
    }
    case "wrong_provider": {
      const altInfo = WEB_SEARCH_PROVIDER_INFO[mode.availableProvider];
      const currentInfo = WEB_SEARCH_PROVIDER_INFO[provider];

      const hint = document.createElement("p");
      hint.className = "pi-search-setup__text";
      hint.textContent = `No ${currentInfo.apiKeyLabel} found. You have a ${altInfo.title} key configured.`;

      const switchNote = document.createElement("p");
      switchNote.className = "pi-search-setup__text";
      switchNote.textContent = `Switch to ${altInfo.title} in /tools, or set up a ${currentInfo.title} key below:`;

      body.append(hint, switchNote);
      body.append(createKeyStep(provider, null, settings, proxyBaseUrl, markDone));
      return { title: `No ${currentInfo.apiKeyLabel} found`, body };
    }
    case "generic_error": {
      const msg = document.createElement("p");
      msg.className = "pi-search-setup__text";
      msg.textContent = "Check your API key and proxy configuration in /tools.";
      body.append(msg);
      return { title: "Web search failed", body };
    }
  }
}

/* ── Public API ─────────────────────────────────────────────── */

/**
 * Mount the inline search setup card into a container element.
 *
 * Called from the tool renderer via a `ref` callback when a `web_search`
 * tool result has `ok: false`.
 */
export function mountSearchSetupCard(
  container: HTMLElement,
  details: WebSearchDetails,
): void {
  // Prevent double-mount
  if (container.dataset.mounted === "true") return;
  container.dataset.mounted = "true";

  // Show a minimal loading state while we detect the setup mode
  const card = document.createElement("div");
  card.className = "pi-search-setup";

  const header = document.createElement("div");
  header.className = "pi-search-setup__header";

  const icon = lucide(Search);
  icon.classList.add("pi-search-setup__icon");

  const titleEl = document.createElement("span");
  titleEl.className = "pi-search-setup__title";
  titleEl.textContent = "Checking search setup…";

  header.append(icon, titleEl);
  card.append(header);
  container.append(card);

  void (async () => {
    try {
      const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
      const settings = storageModule.getAppStorage().settings;

      const ctx = await detectSetupMode(details, settings);

      const dismiss = (): void => {
        card.classList.add("is-dismissed");
        setTimeout(() => card.remove(), 200);
      };

      const { title, body } = buildCardContent(ctx, settings, dismiss);

      // Update header
      titleEl.textContent = title;

      // Replace icon for error states
      if (ctx.mode.type !== "generic_error") {
        const warningIcon = lucide(AlertTriangle);
        warningIcon.classList.add("pi-search-setup__icon");
        icon.replaceWith(warningIcon);
      }

      // Append body
      const bodyEl = document.createElement("div");
      bodyEl.className = "pi-search-setup__body";
      bodyEl.append(body);

      // Dismiss button
      const footer = document.createElement("div");
      footer.className = "pi-search-setup__footer";

      const dismissBtn = document.createElement("button");
      dismissBtn.type = "button";
      dismissBtn.className = "pi-search-setup__dismiss";
      dismissBtn.textContent = "Dismiss";
      dismissBtn.addEventListener("click", dismiss);

      footer.append(dismissBtn);

      card.append(bodyEl, footer);
    } catch {
      // Detection failed — remove the loading card silently
      card.remove();
    }
  })();
}

/**
 * Returns true when the details indicate a web search failure that should
 * show the inline setup card.
 */
export function shouldShowSearchSetupCard(details: unknown): details is WebSearchDetails {
  if (typeof details !== "object" || details === null) return false;
  const d = details as Record<string, unknown>;
  return d.kind === "web_search" && d.ok === false;
}
