/**
 * Experimental features section builder.
 */

import {
  getExperimentalFeatureSnapshots,
  isExperimentalFeatureEnabled,
  setExperimentalFeatureEnabled,
  type ExperimentalFeatureSnapshot,
} from "../../experiments/flags.js";
import { showToast } from "../../ui/toast.js";

const ADVANCED_SECURITY_FEATURE_IDS = new Set<ExperimentalFeatureSnapshot["id"]>([
  "remote_extension_urls",
  "extension_permission_gates",
  "extension_sandbox_runtime",
]);

interface FeatureSectionSpec {
  title: string;
  hint: string;
  features: ExperimentalFeatureSnapshot[];
}


function isAdvancedSecurityFeature(feature: ExperimentalFeatureSnapshot): boolean {
  return ADVANCED_SECURITY_FEATURE_IDS.has(feature.id);
}

function applyStatusVisual(statusEl: HTMLSpanElement, enabled: boolean): void {
  statusEl.textContent = enabled ? "Enabled" : "Disabled";
  statusEl.classList.toggle("is-enabled", enabled);
}

function applyToggleButtonVisual(button: HTMLButtonElement, enabled: boolean): void {
  button.textContent = enabled ? "Disable" : "Enable";
  button.classList.toggle("is-enabled", enabled);
}

function buildFeatureRow(feature: ExperimentalFeatureSnapshot): HTMLElement {
  const row = document.createElement("div");
  row.className = "pi-experimental-row";

  const header = document.createElement("div");
  header.className = "pi-experimental-row__header";

  const title = document.createElement("div");
  title.className = "pi-experimental-row__title";
  title.textContent = feature.title;

  const status = document.createElement("span");
  status.className = "pi-experimental-row__status";

  header.append(title, status);

  const description = document.createElement("div");
  description.className = "pi-experimental-row__description";
  description.textContent = feature.description;

  const meta = document.createElement("div");
  meta.className = "pi-experimental-row__meta";

  const commandHint = document.createElement("code");
  commandHint.className = "pi-experimental-row__command";
  commandHint.textContent = `/experimental toggle ${feature.slug}`;

  const toggleBtn = document.createElement("button");
  toggleBtn.type = "button";
  toggleBtn.className = "pi-overlay-btn pi-experimental-row__toggle";

  meta.append(commandHint, toggleBtn);

  const warning = document.createElement("div");
  warning.className = "pi-experimental-row__warning";
  warning.textContent = feature.warning ?? "";
  warning.hidden = !feature.warning;

  const readiness = document.createElement("div");
  readiness.className = "pi-experimental-row__readiness";
  readiness.textContent = feature.wiring === "wired"
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

function buildFeatureSection(spec: FeatureSectionSpec): HTMLElement {
  const section = document.createElement("section");
  section.className = "pi-overlay-section pi-experimental-section";

  const title = document.createElement("h3");
  title.className = "pi-overlay-section-title";
  title.textContent = spec.title;

  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = spec.hint;

  const list = document.createElement("div");
  list.className = "pi-experimental-list";

  for (const feature of spec.features) {
    list.appendChild(buildFeatureRow(feature));
  }

  section.append(title, hint, list);
  return section;
}

export function buildExperimentalFeatureContent(): HTMLDivElement {
  const content = document.createElement("div");
  content.className = "pi-experimental-content";

  const snapshots = getExperimentalFeatureSnapshots();
  const experimentalFeatures: ExperimentalFeatureSnapshot[] = [];
  const advancedSecurityFeatures: ExperimentalFeatureSnapshot[] = [];

  for (const feature of snapshots) {
    if (isAdvancedSecurityFeature(feature)) {
      advancedSecurityFeatures.push(feature);
      continue;
    }

    experimentalFeatures.push(feature);
  }

  if (experimentalFeatures.length > 0) {
    content.appendChild(buildFeatureSection({
      title: "Experimental capabilities",
      hint: "In-progress features that may evolve quickly.",
      features: experimentalFeatures,
    }));
  }

  if (advancedSecurityFeatures.length > 0) {
    content.appendChild(buildFeatureSection({
      title: "Advanced / security controls",
      hint: "Power-user toggles for extension trust, permissions, and rollback behavior.",
      features: advancedSecurityFeatures,
    }));
  }

  if (snapshots.length === 0) {
    const empty = document.createElement("p");
    empty.className = "pi-overlay-empty";
    empty.textContent = "No experimental features are currently available.";
    content.appendChild(empty);
  }

  return content;
}

export function buildExperimentalFeatureFooter(): HTMLParagraphElement {
  const footer = document.createElement("p");
  footer.className = "pi-experimental-footer";
  footer.textContent =
    "Tip: use /experimental on <feature>, /experimental off <feature>, /experimental toggle <feature>, "
    + "/experimental tmux-bridge-url <url>, /experimental tmux-bridge-token <token>, /experimental tmux-status, "
    + "/experimental python-bridge-url <url>, or /experimental python-bridge-token <token>.";
  return footer;
}

export function showExperimentalDialog(): void {
  void import("./settings-overlay.js").then(({ showSettingsDialog }) => {
    void showSettingsDialog({ section: "experimental" });
  });
}
