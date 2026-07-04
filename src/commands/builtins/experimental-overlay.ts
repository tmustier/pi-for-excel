/**
 * Experimental features section builder.
 */

import {
  getExperimentalFeatureSnapshots,
  setExperimentalFeatureEnabled,
  type ExperimentalFeatureSnapshot,
} from "../../experiments/flags.js";
import { createToggleRow } from "../../ui/extensions-hub-components.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";

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

function buildFeatureRow(feature: ExperimentalFeatureSnapshot): HTMLElement {
  const row = document.createElement("div");
  row.className = "pi-experimental-row";

  const toggleRow = createToggleRow({
    label: feature.title,
    sublabel: feature.description,
    checked: feature.enabled,
    onChange: (checked) => {
      setExperimentalFeatureEnabled(feature.id, checked);

      const suffix = feature.wiring === "flag-only"
        ? " (flag saved; feature not wired yet)"
        : "";
      showToast(`${feature.title}: ${checked ? t("experimental.enabled") : t("experimental.disabled")}${suffix}`);
    },
  });
  toggleRow.root.classList.add("pi-experimental-row__toggle-row");

  const meta = document.createElement("div");
  meta.className = "pi-experimental-row__meta";

  const commandHint = document.createElement("code");
  commandHint.className = "pi-experimental-row__command";
  commandHint.textContent = `/experimental toggle ${feature.slug}`;

  meta.appendChild(commandHint);

  const warning = document.createElement("div");
  warning.className = "pi-experimental-row__warning";
  warning.textContent = feature.warning ?? "";
  warning.hidden = !feature.warning;

  const readiness = document.createElement("div");
  readiness.className = "pi-experimental-row__readiness";
  readiness.textContent = feature.wiring === "wired"
    ? t("experimental.ready")
    : t("experimental.notWired") + " — this capability is planned but not wired yet.";

  row.append(toggleRow.root, meta, warning, readiness);
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
      title: t("experimental.title"),
      hint: t("experimental.hint"),
      features: experimentalFeatures,
    }));
  }

  // Python bridge configuration is in Extensions → Connections → Bridges.

  if (advancedSecurityFeatures.length > 0) {
    content.appendChild(buildFeatureSection({
      title: t("experimental.advancedSecurity"),
      hint: "Power-user toggles for extension trust, permissions, and rollback behavior.",
      features: advancedSecurityFeatures,
    }));
  }

  if (snapshots.length === 0) {
    const empty = document.createElement("p");
    empty.className = "pi-overlay-empty";
    empty.textContent = t("experimental.noFeatures");
    content.appendChild(empty);
  }

  return content;
}

export function buildExperimentalFeatureFooter(): HTMLParagraphElement {
  const footer = document.createElement("p");
  footer.className = "pi-experimental-footer";
  footer.textContent = t("experimental.tip");
  return footer;
}

export function showExperimentalDialog(): void {
  void import("./settings-overlay.js").then(({ showSettingsDialog }) => {
    void showSettingsDialog({ section: "experimental" });
  });
}
