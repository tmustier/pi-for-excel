import {
  createOverlayBadge,
  createOverlayButton,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { showToast } from "../../ui/toast.js";
import type { ExtensionRuntimeStatus } from "../../extensions/runtime-manager.js";
import type { AddonsDialogActions } from "./addons-overlay-types.js";

const HIGH_RISK_CAPABILITIES = new Set<string>([
  "tools.register",
  "agent.read",
  "agent.events.read",
  "llm.complete",
  "http.fetch",
  "agent.context.write",
  "agent.steer",
  "agent.followup",
  "skills.write",
]);

function confirmExtensionEnable(status: ExtensionRuntimeStatus): boolean {
  if (status.trust === "builtin") {
    return true;
  }

  const highRiskCapabilities = status.grantedCapabilities.filter((capability) => HIGH_RISK_CAPABILITIES.has(capability));
  if (highRiskCapabilities.length === 0) {
    return true;
  }

  const lines = [
    `Enable extension "${status.name}" with higher-risk permissions?`,
    "",
    "Granted higher-risk permissions:",
    ...highRiskCapabilities.map((capability) => `- ${capability}`),
    "",
    `Source: ${status.trustLabel}`,
    "",
    "You can edit permissions later in /extensions.",
  ];

  return window.confirm(lines.join("\n"));
}

export function renderExtensionsSection(args: {
  container: HTMLElement;
  actions: AddonsDialogActions;
  busy: boolean;
  onRefresh: () => void;
}): void {
  args.container.replaceChildren();

  const section = document.createElement("section");
  section.className = "pi-overlay-section pi-addons-section";
  section.dataset.addonsSection = "extensions";
  section.appendChild(createOverlaySectionTitle("Extensions"));

  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = "Code plugins running in the sidebar.";
  section.appendChild(hint);

  const list = document.createElement("div");
  list.className = "pi-overlay-list";

  const statuses = args.actions.listExtensions();
  if (statuses.length === 0) {
    const empty = document.createElement("div");
    empty.className = "pi-overlay-empty";
    empty.textContent = "No extensions installed.";
    list.appendChild(empty);
  } else {
    for (const status of statuses) {
      const card = document.createElement("div");
      card.className = "pi-overlay-surface pi-addons-extension";

      const top = document.createElement("div");
      top.className = "pi-addons-extension__top";

      const title = document.createElement("strong");
      title.className = "pi-addons-extension__name";
      title.textContent = status.name;

      const badges = document.createElement("div");
      badges.className = "pi-overlay-badges";
      badges.appendChild(createOverlayBadge(status.enabled ? "enabled" : "disabled", status.enabled ? "ok" : "muted"));
      badges.appendChild(createOverlayBadge(status.loaded ? "loaded" : "not loaded", status.loaded ? "ok" : "warn"));
      badges.appendChild(createOverlayBadge(status.trustLabel, "muted"));

      top.append(title, badges);

      const details = document.createElement("div");
      details.className = "pi-addons-extension__details";
      details.textContent = `${status.sourceLabel} · ${status.runtimeLabel}`;

      const controls = document.createElement("div");
      controls.className = "pi-overlay-actions pi-overlay-actions--inline";

      const toggleLabel = document.createElement("label");
      toggleLabel.className = "pi-addons-extension__toggle";

      const toggle = document.createElement("input");
      toggle.type = "checkbox";
      toggle.checked = status.enabled;
      toggle.disabled = args.busy;
      toggle.addEventListener("change", () => {
        if (toggle.checked && !status.enabled && !confirmExtensionEnable(status)) {
          toggle.checked = false;
          return;
        }

        void args.actions.setExtensionEnabled(status.id, toggle.checked)
          .then(async () => {
            if (args.actions.onChanged) {
              await args.actions.onChanged();
            }

            showToast(`${status.name}: ${toggle.checked ? "enabled" : "disabled"}`);
            args.onRefresh();
          })
          .catch((error: unknown) => {
            const message = error instanceof Error ? error.message : "Unknown error";
            showToast(`Extensions: ${message}`);
            args.onRefresh();
          });
      });

      toggleLabel.append(toggle, document.createTextNode("Enabled"));
      controls.appendChild(toggleLabel);

      card.append(top, details, controls);

      if (status.lastError) {
        const errorLine = document.createElement("p");
        errorLine.className = "pi-overlay-hint pi-overlay-text-warning";
        errorLine.textContent = `Last error: ${status.lastError}`;
        card.appendChild(errorLine);
      }

      list.appendChild(card);
    }
  }

  const actionsRow = document.createElement("div");
  actionsRow.className = "pi-overlay-actions";

  const openButton = createOverlayButton({ text: "Open full Extensions manager…" });
  openButton.addEventListener("click", () => {
    args.actions.openExtensionsManager();
  });

  actionsRow.appendChild(openButton);
  section.append(list, actionsRow);
  args.container.appendChild(section);
}
