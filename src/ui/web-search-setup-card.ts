/**
 * Inline setup card for web search failures.
 *
 * Rendered inside the chat stream after a failed `web_search` tool card.
 * Detects the failure mode and shows contextual setup steps with inline
 * actions (proxy retry + API key save/validate).
 */

import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

import {
  DEFAULT_PROXY_URL,
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
import { t } from "../language/index.js";

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
  copyBtn.title = t("bridge-setup.copyCommandTitle");
  copyBtn.setAttribute("aria-label", "Copy command");
  copyBtn.replaceChildren(lucide(Copy));

  let resetTimeout: ReturnType<typeof setTimeout> | null = null;

  copyBtn.addEventListener("click", () => {
    copyToClipboard(command, () => {
      copyBtn.replaceChildren(lucide(Check));
      copyBtn.title = t("bridge-setup.copiedTitle");
      copyBtn.setAttribute("aria-label", "Copied");

      if (resetTimeout !== null) {
        clearTimeout(resetTimeout);
      }

      resetTimeout = setTimeout(() => {
        copyBtn.replaceChildren(lucide(Copy));
        copyBtn.title = t("bridge-setup.copyCommandTitle");
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
    ? t("web-search-setup.stepLabel", { n: String(options.stepNumber) })
    : t("web-search-setup.startHelper");

  const hint = document.createElement("p");
  hint.className = "pi-search-setup__hint";
  hint.textContent = t("web-search-setup.helper-instructions");

  const actions = document.createElement("div");
  actions.className = "pi-search-setup__actions";

  const retryBtn = document.createElement("button");
  retryBtn.type = "button";
  retryBtn.className = "pi-search-setup__retry";
  retryBtn.textContent = t("web-search-setup.retry");

  const status = document.createElement("span");
  status.className = "pi-search-setup__status";
  status.setAttribute("role", "status");
  status.setAttribute("aria-live", "polite");

  let checking = false;

  retryBtn.addEventListener("click", () => {
    if (checking) {
      return;
    }

    checking = true;
    retryBtn.disabled = true;
    retryBtn.textContent = t("web-search-setup.checking");
    status.textContent = t("web-search-setup.checking-helper");
    status.className = "pi-search-setup__status";

    const probeUrl = options.proxyBaseUrl ?? DEFAULT_PROXY_URL;

    void probeProxyReachability(probeUrl, 1500).then(
      (reachable) => {
        if (reachable) {
          status.textContent = t("web-search-setup.helperDetected");
          status.className = "pi-search-setup__status is-ok";
          options.onProxyReady?.();
          return;
        }

        status.textContent = t("web-search-setup.helper-not-detected");
        status.className = "pi-search-setup__status is-warn";
      },
      () => {
        status.textContent = t("web-search-setup.check-failed");
        status.className = "pi-search-setup__status is-error";
      },
    ).finally(() => {
      checking = false;
      retryBtn.disabled = false;
      retryBtn.textContent = t("web-search-setup.retry");
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
  signupLink.textContent = t("web-search-setup.freeKeyLink", { url: info.signupUrl.replace(/^https?:\/\//u, "") });

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
  saveBtn.textContent = t("web-search-setup.save");

  const status = document.createElement("span");
  status.className = "pi-search-setup__status";
  status.setAttribute("role", "status");
  status.setAttribute("aria-live", "polite");

  let saving = false;

  saveBtn.addEventListener("click", () => {
    if (saving) {
      return;
    }

    const key = input.value.trim();
    if (key.length === 0) {
      showToast(t("web-search-setup.toast.enterApiKey"));
      return;
    }

    const formatWarning = checkApiKeyFormat(provider, key);

    saving = true;
    saveBtn.disabled = true;

    if (formatWarning) {
      status.textContent = `⚠️ ${formatWarning} Saving anyway…`;
      status.className = "pi-search-setup__status is-warn";
    } else {
      status.textContent = t("web-search-setup.saving");
      status.className = "pi-search-setup__status";
    }

    void (async () => {
      try {
        await saveWebSearchApiKey(settings, provider, key);

        status.textContent = t("web-search-setup.validating");
        status.className = "pi-search-setup__status";

        const result = await validateWebSearchApiKey({ provider, apiKey: key, proxyBaseUrl });

        if (result.ok) {
          status.textContent = `✓ ${result.message}`;
          status.className = "pi-search-setup__status is-ok";
          input.value = "";
          onSaved();
          return;
        }

        status.textContent = t("web-search-setup.keySavedValidation", { message: result.message });
        status.className = "pi-search-setup__status is-warn";
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : String(error);
        status.textContent = t("web-search-setup.error", { message });
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
    showToast(t("web-search-setup.toast.ready"));
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
      return { title: t("web-search-setup.title.needsSetup"), body };
    }

    case "needs_key": {
      body.append(createKeyStep(provider, null, settings, proxyBaseUrl, markDone));
      return { title: t("web-search-setup.title.needsApiKey"), body };
    }

    case "needs_proxy": {
      body.append(createProxyStep({
        stepNumber: null,
        proxyBaseUrl,
        onProxyReady: markDone,
      }));
      return { title: t("web-search-setup.title.cantConnect"), body };
    }

    case "wrong_provider": {
      const alternativeInfo = WEB_SEARCH_PROVIDER_INFO[mode.availableProvider];
      const currentInfo = WEB_SEARCH_PROVIDER_INFO[provider];

      const hint = document.createElement("p");
      hint.className = "pi-search-setup__text";
      hint.textContent = t("web-search-setup.noCurrentKeyHaveAlternative", { current: currentInfo.apiKeyLabel, alternative: alternativeInfo.title });

      const switchNote = document.createElement("p");
      switchNote.className = "pi-search-setup__text";
      switchNote.textContent = t("web-search-setup.switchOrSetup", { alternative: alternativeInfo.title, current: currentInfo.title });

      body.append(hint, switchNote, createKeyStep(provider, null, settings, proxyBaseUrl, markDone));

      return { title: t("web-search-setup.title.noKeyFound", { label: currentInfo.apiKeyLabel }), body };
    }

    case "generic_error": {
      const message = document.createElement("p");
      message.className = "pi-search-setup__text";
      message.textContent = t("web-search-setup.check-config");
      body.append(message);
      return { title: t("web-search-setup.title.failed"), body };
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
  titleEl.textContent = t("web-search-setup.checking-setup");

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
      dismissBtn.textContent = t("web-search-setup.dismiss");
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
