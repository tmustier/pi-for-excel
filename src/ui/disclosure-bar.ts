/**
 * Disclosure bar — non-blocking banner shown after first provider connect.
 *
 * Informs the user about Pi's external capabilities (web search, extensions,
 * MCP, skills) and lets them acknowledge or customize before using the agent.
 */

import { t } from "../language/index.js";
import { createToggleRow } from "./extensions-hub-components.js";

const ACKNOWLEDGED_KEY = "pi.onboarding.disclosure.acknowledged";

function isAcknowledged(): boolean {
  try {
    return localStorage.getItem(ACKNOWLEDGED_KEY) === "1";
  } catch {
    return false;
  }
}

function setAcknowledged(): void {
  try {
    localStorage.setItem(ACKNOWLEDGED_KEY, "1");
  } catch {
    // ignore — private mode / storage unavailable
  }
}

export interface DisclosureBarOptions {
  /** Number of configured providers (bar only shows when ≥1). */
  providerCount: number;
  /** Callback to open Settings overlay. If provided, t("disclosure-bar.changeInSettings") becomes a link. */
  onOpenSettings?: () => void;
}

/**
 * Create and return the disclosure bar element, or `null` if already dismissed.
 *
 * The caller is responsible for inserting the element into the DOM.
 * The returned element removes itself when the user dismisses it.
 */
export function createDisclosureBar(options: DisclosureBarOptions): HTMLElement | null {
  if (isAcknowledged() || options.providerCount < 1) {
    return null;
  }

  const bar = document.createElement("div");
  bar.className = "pi-disclosure-bar";

  const text = document.createElement("div");
  text.className = "pi-disclosure-bar__text";
  text.textContent = t("disclosure-bar.text");
  bar.appendChild(text);

  // --- Expandable picker (hidden by default) ---
  const picker = document.createElement("div");
  picker.className = "pi-disclosure-picker";
  bar.appendChild(picker);

  const toggleRows: { label: string; sublabel: string }[] = [
    { label: t("disclosure-bar.webSearchLabel"), sublabel: "Search engines and read web pages" },
    { label: t("disclosure-bar.extensionsLabel"), sublabel: "Sidebar tools and custom commands" },
    { label: t("disclosure-bar.externalServicesLabel"), sublabel: "Connect to tool servers you configure" },
    { label: t("disclosure-bar.skillsLabel"), sublabel: "Instruction documents the AI follows" },
  ];

  for (const row of toggleRows) {
    const toggleRow = createToggleRow({
      label: row.label,
      sublabel: row.sublabel,
      checked: true,
    });
    picker.appendChild(toggleRow.root);
  }

  // --- Actions row ---
  const actions = document.createElement("div");
  actions.className = "pi-disclosure-bar__actions";
  bar.appendChild(actions);

  const dismiss = () => {
    setAcknowledged();
    bar.remove();
  };

  const gotItBtn = document.createElement("button");
  gotItBtn.className = "pi-overlay-btn pi-overlay-btn--primary pi-overlay-btn--compact";
  gotItBtn.textContent = t("disclosure-bar.gotIt");
  gotItBtn.addEventListener("click", dismiss);
  actions.appendChild(gotItBtn);

  const customizeBtn = document.createElement("button");
  customizeBtn.className = "pi-disclosure-bar__link";
  customizeBtn.textContent = t("disclosure-bar.customize");
  actions.appendChild(customizeBtn);

  let hint: HTMLElement;
  if (options.onOpenSettings) {
    const link = document.createElement("button");
    link.type = "button";
    link.className = "pi-disclosure-bar__settings-link";
    link.textContent = t("disclosure-bar.changeInSettings");
    link.addEventListener("click", () => {
      dismiss();
      options.onOpenSettings?.();
    });
    hint = link;
  } else {
    const span = document.createElement("span");
    span.className = "pi-disclosure-bar__muted";
    span.textContent = t("disclosure-bar.changeInSettingsMuted");
    hint = span;
  }
  actions.appendChild(hint);

  customizeBtn.addEventListener("click", () => {
    const isVisible = picker.classList.toggle("is-visible");
    if (isVisible) {
      gotItBtn.textContent = t("disclosure-bar.done");
      customizeBtn.style.display = "none";
      hint.style.display = "none";
    } else {
      gotItBtn.textContent = t("disclosure-bar.gotIt");
      customizeBtn.style.display = "";
      hint.style.display = "";
    }
  });

  return bar;
}
