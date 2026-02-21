/**
 * Proxy warning banner.
 *
 * State-driven inline banner shown above chat messages when the local proxy is
 * unavailable. Expands inline with quick setup guidance.
 */

import { AlertTriangle, Check, Copy, lucide } from "./lucide-icons.js";

const PROXY_COMMAND = "npx pi-for-excel-proxy";
const INSTALL_GUIDE_URL = "https://pi.dev/excel#connect";

export type ProxyBannerState = "detected" | "not-detected" | "unknown";

export interface ProxyBannerHandle {
  root: HTMLElement;
  update: (state: ProxyBannerState) => void;
}

function selectElementText(element: HTMLElement): void {
  const selection = window.getSelection();
  if (!selection) return;

  const range = document.createRange();
  range.selectNodeContents(element);
  selection.removeAllRanges();
  selection.addRange(range);
}

export function createProxyBanner(): ProxyBannerHandle {
  const root = document.createElement("section");
  root.className = "pi-proxy-banner";
  root.hidden = true;

  const topRow = document.createElement("div");
  topRow.className = "pi-proxy-banner__row";

  const text = document.createElement("p");
  text.className = "pi-proxy-banner__text";

  const warningIcon = lucide(AlertTriangle);
  warningIcon.classList.add("pi-proxy-banner__text-icon");
  warningIcon.setAttribute("aria-hidden", "true");

  const textLabel = document.createElement("span");
  textLabel.textContent = "Proxy not running · some features won't work.";

  text.append(warningIcon, textLabel);

  const action = document.createElement("button");
  action.type = "button";
  action.className = "pi-proxy-banner__action";
  action.textContent = "How to fix →";

  topRow.append(text, action);

  const details = document.createElement("div");
  details.className = "pi-proxy-banner__details";
  details.hidden = true;

  const detailsIntro = document.createElement("p");
  detailsIntro.className = "pi-proxy-banner__details-text";
  detailsIntro.textContent = "Run this command in a terminal and keep that window open:";

  const codeRow = document.createElement("div");
  codeRow.className = "pi-proxy-banner__code";

  const code = document.createElement("code");
  code.textContent = PROXY_COMMAND;

  const copyButton = document.createElement("button");
  copyButton.type = "button";
  copyButton.className = "pi-proxy-banner__copy";

  let resetCopyIconTimeout: ReturnType<typeof setTimeout> | null = null;

  const renderCopyIcon = (): void => {
    copyButton.replaceChildren(lucide(Copy));
    copyButton.title = "Copy command";
    copyButton.setAttribute("aria-label", "Copy command");
  };

  const renderCopiedIcon = (): void => {
    copyButton.replaceChildren(lucide(Check));
    copyButton.title = "Copied";
    copyButton.setAttribute("aria-label", "Copied");
  };

  renderCopyIcon();

  copyButton.addEventListener("click", () => {
    if (!navigator.clipboard?.writeText) {
      selectElementText(code);
      return;
    }

    void navigator.clipboard.writeText(PROXY_COMMAND).then(
      () => {
        renderCopiedIcon();
        if (resetCopyIconTimeout) {
          clearTimeout(resetCopyIconTimeout);
        }
        resetCopyIconTimeout = setTimeout(() => {
          renderCopyIcon();
          resetCopyIconTimeout = null;
        }, 1400);
      },
      () => {
        selectElementText(code);
      },
    );
  });

  codeRow.append(code, copyButton);

  const hint = document.createElement("p");
  hint.className = "pi-proxy-banner__hint";
  hint.textContent = "Open Terminal · paste · press Enter · type y and Enter if prompted · leave open";

  const guideLink = document.createElement("a");
  guideLink.className = "pi-proxy-banner__link";
  guideLink.href = INSTALL_GUIDE_URL;
  guideLink.target = "_blank";
  guideLink.rel = "noopener noreferrer";
  guideLink.textContent = "No Node.js? See install guide →";

  details.append(detailsIntro, codeRow, hint, guideLink);

  action.addEventListener("click", () => {
    const shouldOpen = details.hidden;
    details.hidden = !shouldOpen;
    root.classList.toggle("is-open", shouldOpen);
    action.textContent = shouldOpen ? "Hide steps" : "How to fix →";
  });

  root.append(topRow, details);

  const update = (state: ProxyBannerState): void => {
    const shouldShow = state === "not-detected";
    root.hidden = !shouldShow;

    if (!shouldShow) {
      details.hidden = true;
      root.classList.remove("is-open");
      action.textContent = "How to fix →";
    }
  };

  return { root, update };
}
