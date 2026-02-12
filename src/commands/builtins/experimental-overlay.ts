/**
 * Experimental features overlay.
 */

import {
  getExperimentalFeatureSnapshots,
  isExperimentalFeatureEnabled,
  setExperimentalFeatureEnabled,
  type ExperimentalFeatureSnapshot,
} from "../../experiments/flags.js";
import { showToast } from "../../ui/toast.js";

const OVERLAY_ID = "pi-experimental-overlay";

function applyStatusVisual(statusEl: HTMLSpanElement, enabled: boolean): void {
  statusEl.textContent = enabled ? "Enabled" : "Disabled";
  statusEl.style.color = enabled ? "var(--pi-green)" : "var(--muted-foreground)";
}

function applyToggleButtonVisual(button: HTMLButtonElement, enabled: boolean): void {
  button.textContent = enabled ? "Disable" : "Enable";
  button.style.background = enabled ? "oklch(0.92 0.02 30 / 0.85)" : "var(--pi-green)";
  button.style.color = enabled ? "var(--foreground)" : "white";
  button.style.border = enabled ? "1px solid oklch(0 0 0 / 0.12)" : "none";
}

function buildFeatureRow(feature: ExperimentalFeatureSnapshot): HTMLElement {
  const row = document.createElement("div");
  row.style.cssText =
    "padding: 10px 12px; border: 1px solid oklch(0 0 0 / 0.08); border-radius: 10px; background: oklch(1 0 0 / 0.35);";

  const header = document.createElement("div");
  header.style.cssText = "display: flex; align-items: center; justify-content: space-between; gap: 10px;";

  const title = document.createElement("div");
  title.style.cssText = "font-size: 13px; font-weight: 600; font-family: var(--font-sans);";
  title.textContent = feature.title;

  const status = document.createElement("span");
  status.style.cssText = "font-size: 11px; font-family: var(--font-mono);";

  header.append(title, status);

  const description = document.createElement("div");
  description.style.cssText = "margin-top: 4px; font-size: 12px; color: var(--muted-foreground); font-family: var(--font-sans);";
  description.textContent = feature.description;

  const meta = document.createElement("div");
  meta.style.cssText =
    "margin-top: 6px; display: flex; align-items: center; justify-content: space-between; gap: 10px;";

  const commandHint = document.createElement("code");
  commandHint.style.cssText =
    "font-size: 10px; color: var(--muted-foreground); font-family: var(--font-mono); background: oklch(0 0 0 / 0.04); padding: 2px 6px; border-radius: 6px;";
  commandHint.textContent = `/experimental toggle ${feature.slug}`;

  const toggleBtn = document.createElement("button");
  toggleBtn.type = "button";
  toggleBtn.style.cssText =
    "padding: 6px 10px; border-radius: 8px; font-family: var(--font-sans); font-size: 12px; font-weight: 600; cursor: pointer;";

  meta.append(commandHint, toggleBtn);

  const warning = document.createElement("div");
  warning.style.cssText = "margin-top: 6px; font-size: 11px; color: oklch(0.58 0.16 40); font-family: var(--font-sans);";
  warning.textContent = feature.warning ?? "";
  warning.style.display = feature.warning ? "block" : "none";

  const readiness = document.createElement("div");
  readiness.style.cssText = "margin-top: 6px; font-size: 11px; color: var(--muted-foreground); font-family: var(--font-sans);";
  readiness.textContent =
    feature.wiring === "wired"
      ? "Ready now"
      : "Flag only for now — this capability is planned but not wired yet.";

  const applyState = (enabled: boolean): void => {
    applyStatusVisual(status, enabled);
    applyToggleButtonVisual(toggleBtn, enabled);
  };

  toggleBtn.addEventListener("click", () => {
    const next = !isExperimentalFeatureEnabled(feature.id);
    setExperimentalFeatureEnabled(feature.id, next);
    applyState(next);

    const suffix = feature.wiring === "flag-only"
      ? " (flag saved; feature not wired yet)"
      : "";
    showToast(`${feature.title}: ${next ? "enabled" : "disabled"}${suffix}`);
  });

  applyState(feature.enabled);
  row.append(header, description, meta, warning, readiness);
  return row;
}

export function showExperimentalDialog(): void {
  const existing = document.getElementById(OVERLAY_ID);
  if (existing) {
    existing.remove();
    return;
  }

  const overlay = document.createElement("div");
  overlay.id = OVERLAY_ID;
  overlay.className = "pi-welcome-overlay";

  const card = document.createElement("div");
  card.className = "pi-welcome-card";
  card.style.cssText = "text-align: left; max-width: 560px;";

  const title = document.createElement("h2");
  title.style.cssText = "font-size: 16px; font-weight: 600; margin: 0 0 6px; font-family: var(--font-sans);";
  title.textContent = "Experimental Features";

  const subtitle = document.createElement("p");
  subtitle.style.cssText = "font-size: 12px; color: var(--muted-foreground); margin: 0 0 12px; font-family: var(--font-sans);";
  subtitle.textContent =
    "These toggles are local to this browser profile. Use carefully — some are security-sensitive.";

  const list = document.createElement("div");
  list.style.cssText = "display: flex; flex-direction: column; gap: 8px;";

  for (const feature of getExperimentalFeatureSnapshots()) {
    list.appendChild(buildFeatureRow(feature));
  }

  const footer = document.createElement("p");
  footer.style.cssText = "font-size: 11px; color: var(--muted-foreground); margin: 12px 0 0; font-family: var(--font-sans);";
  footer.textContent =
    "Tip: use /experimental on <feature>, /experimental off <feature>, /experimental toggle <feature>, /experimental tmux-bridge-url <url>, /experimental tmux-bridge-token <token>, /experimental tmux-status, /experimental python-bridge-url <url>, or /experimental python-bridge-token <token>.";

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.style.cssText =
    "margin-top: 12px; width: 100%; padding: 8px; border-radius: 8px; border: 1px solid oklch(0 0 0 / 0.08); background: oklch(0 0 0 / 0.03); cursor: pointer; font-family: var(--font-sans); font-size: 13px;";
  closeBtn.textContent = "Close";

  closeBtn.addEventListener("click", () => {
    overlay.remove();
  });

  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
    }
  });

  card.append(title, subtitle, list, footer, closeBtn);
  overlay.appendChild(card);
  document.body.appendChild(overlay);
}
