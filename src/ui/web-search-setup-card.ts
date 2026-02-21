/**
 * Inline setup card for web search failures.
 *
 * Rendered inside the chat stream after a failed `web_search` tool card.
 * Detects the failure mode and shows contextual setup steps with inline
 * actions (proxy retry + API key save/validate).
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  DEFAULT_LOCAL_PROXY_URL,
  probeProxyReachability,
} from "../auth/proxy-validation.js";
import {
  detectWebSearchSetupContext,
  type WebSearchSetupContext,
} from "../tools/web-search-setup-detection.js";
import {
  checkApiKeyFormat,
  saveWebSearchApiKey,
  WEB_SEARCH_PROVIDER_INFO,
  type WebSearchConfigStore,
  type WebSearchProvider,
} from "../tools/web-search-config.js";
import { isWebSearchDetails, type WebSearchDetails } from "../tools/tool-details.js";
import { validateWebSearchApiKey } from "../tools/web-search.js";
import { AlertTriangle, Check, Copy, Search, lucide } from "./lucide-icons.js";
import { showToast } from "./toast.js";

const PROXY_COMMAND = "npx pi-for-excel-proxy";

interface ProxyStepOptions {
  stepNumber: number | null;
  proxyBaseUrl: string | undefined;
  onProxyReady?: () => void;
}

function selectElementText(element: HTMLElement): void {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

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

function createCopyableCommand(command: string): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "pi-search-setup__code";

  const code = document.createElement("code");
  code.textContent = command;

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "pi-search-setup__copy";
  copyBtn.title = "Copy command";
  copyBtn.setAttribute("aria-label", "Copy command");
  copyBtn.replaceChildren(lucide(Copy));

  let resetTimeout: ReturnType<typeof setTimeout> | null = null;

  copyBtn.addEventListener("click", () => {
    copyToClipboard(command, () => {
      copyBtn.replaceChildren(lucide(Check));
      copyBtn.title = "Copied";
      copyBtn.setAttribute("aria-label", "Copied");

      if (resetTimeout !== null) {
        clearTimeout(resetTimeout);
      }

      resetTimeout = setTimeout(() => {
        copyBtn.replaceChildren(lucide(Copy));
        copyBtn.title = "Copy command";
        copyBtn.setAttribute("aria-label", "Copy command");
        resetTimeout = null;
      }, 1400);
    }, code);
  });

  row.append(code, copyBtn);
  return row;
}

function createProxyStep(options: ProxyStepOptions): HTMLDivElement {
  const step = document.createElement("div");
  step.className = "pi-search-setup__step";

  const label = document.createElement("p");
  label.className = "pi-search-setup__step-label";
  label.textContent = options.stepNumber !== null
    ? `Step ${options.stepNumber} · Start the helper (keep it running):`
    : "Start the helper (keep it running):";

  const hint = document.createElement("p");
  hint.className = "pi-search-setup__hint";
  hint.textContent = "Open Terminal · paste · press Enter · wait for \"Proxy listening\"";

  const actions = document.createElement("div");
  actions.className = "pi-search-setup__actions";

  const retryBtn = document.createElement("button");
  retryBtn.type = "button";
  retryBtn.className = "pi-search-setup__retry";
  retryBtn.textContent = "Retry";

  const status = document.createElement("span");
  status.className = "pi-search-setup__status";

  let checking = false;

  retryBtn.addEventListener("click", () => {
    if (checking) {
      return;
    }

    checking = true;
    retryBtn.disabled = true;
    retryBtn.textContent = "Checking…";
    status.textContent = "Checking helper…";
    status.className = "pi-search-setup__status";

    const probeUrl = options.proxyBaseUrl ?? DEFAULT_LOCAL_PROXY_URL;

    void probeProxyReachability(probeUrl, 1500).then(
      (reachable) => {
        if (reachable) {
          status.textContent = "✓ Helper detected";
          status.className = "pi-search-setup__status is-ok";
          options.onProxyReady?.();
          return;
        }

        status.textContent = "Helper not detected yet — keep terminal open and try again.";
        status.className = "pi-search-setup__status is-warn";
      },
      () => {
        status.textContent = "Could not check helper status right now.";
        status.className = "pi-search-setup__status is-error";
      },
    ).finally(() => {
      checking = false;
      retryBtn.disabled = false;
      retryBtn.textContent = "Retry";
    });
  });

  actions.append(retryBtn);
  step.append(label, createCopyableCommand(PROXY_COMMAND), hint, actions, status);
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
    if (saving) {
      return;
    }

    const key = input.value.trim();
    if (key.length === 0) {
      showToast("Enter an API key first.");
      return;
    }

    const formatWarning = checkApiKeyFormat(provider, key);

    saving = true;
    saveBtn.disabled = true;

    if (formatWarning) {
      status.textContent = `⚠️ ${formatWarning} Saving anyway…`;
      status.className = "pi-search-setup__status is-warn";
    } else {
      status.textContent = "Saving…";
      status.className = "pi-search-setup__status";
    }

    void (async () => {
      try {
        await saveWebSearchApiKey(settings, provider, key);

        status.textContent = "Validating…";
        status.className = "pi-search-setup__status";

        const result = await validateWebSearchApiKey({ provider, apiKey: key, proxyBaseUrl });

        if (result.ok) {
          status.textContent = `✓ ${result.message}`;
          status.className = "pi-search-setup__status is-ok";
          input.value = "";
          onSaved();
          return;
        }

        status.textContent = `Key saved. Validation: ${result.message}`;
        status.className = "pi-search-setup__status is-warn";
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : String(error);
        status.textContent = `Error: ${message}`;
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

function buildCardContent(
  context: WebSearchSetupContext,
  settings: WebSearchConfigStore,
  onDismiss: () => void,
): { title: string; body: DocumentFragment } {
  const body = document.createDocumentFragment();
  const { mode, provider, proxyBaseUrl } = context;

  const markDone = (): void => {
    showToast("✓ Web search is ready — ask the assistant to try again.");
    onDismiss();
  };

  switch (mode.type) {
    case "needs_both": {
      body.append(
        createProxyStep({
          stepNumber: 1,
          proxyBaseUrl,
        }),
        createKeyStep(provider, 2, settings, proxyBaseUrl, markDone),
      );
      return { title: "Web search needs setup", body };
    }

    case "needs_key": {
      body.append(createKeyStep(provider, null, settings, proxyBaseUrl, markDone));
      return { title: "Web search needs an API key", body };
    }

    case "needs_proxy": {
      body.append(createProxyStep({
        stepNumber: null,
        proxyBaseUrl,
        onProxyReady: markDone,
      }));
      return { title: "Web search can't connect", body };
    }

    case "wrong_provider": {
      const alternativeInfo = WEB_SEARCH_PROVIDER_INFO[mode.availableProvider];
      const currentInfo = WEB_SEARCH_PROVIDER_INFO[provider];

      const hint = document.createElement("p");
      hint.className = "pi-search-setup__text";
      hint.textContent = `No ${currentInfo.apiKeyLabel} found. You have a ${alternativeInfo.title} key configured.`;

      const switchNote = document.createElement("p");
      switchNote.className = "pi-search-setup__text";
      switchNote.textContent = `Switch to ${alternativeInfo.title} in /tools, or set up a ${currentInfo.title} key below:`;

      body.append(hint, switchNote, createKeyStep(provider, null, settings, proxyBaseUrl, markDone));

      return { title: `No ${currentInfo.apiKeyLabel} found`, body };
    }

    case "generic_error": {
      const message = document.createElement("p");
      message.className = "pi-search-setup__text";
      message.textContent = "Check your API key and proxy configuration in /tools.";
      body.append(message);
      return { title: "Web search failed", body };
    }
  }
}

/**
 * Mount the inline search setup card into a container element.
 *
 * Called from the tool renderer via a `ref` callback when a `web_search`
 * tool result has `ok: false`.
 */
export function mountSearchSetupCard(container: HTMLElement, details: WebSearchDetails): void {
  if (container.dataset.mounted === "true") {
    return;
  }

  container.dataset.mounted = "true";

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
      const settings = getAppStorage().settings;
      const context = await detectWebSearchSetupContext(details, settings, {
        isDev: import.meta.env.DEV,
      });

      const dismiss = (): void => {
        card.classList.add("is-dismissed");
        setTimeout(() => card.remove(), 200);
      };

      const { title, body } = buildCardContent(context, settings, dismiss);

      titleEl.textContent = title;

      if (context.mode.type !== "generic_error") {
        const warningIcon = lucide(AlertTriangle);
        warningIcon.classList.add("pi-search-setup__icon");
        icon.replaceWith(warningIcon);
      }

      const bodyEl = document.createElement("div");
      bodyEl.className = "pi-search-setup__body";
      bodyEl.append(body);

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
      card.remove();
    }
  })();
}

/**
 * Returns true when the details indicate a web search failure that should
 * show the inline setup card.
 */
export function shouldShowSearchSetupCard(details: unknown): details is WebSearchDetails {
  return isWebSearchDetails(details) && details.ok === false;
}
