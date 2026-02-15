import { type IntegrationDefinition } from "../../integrations/catalog.js";
import { createOverlayBadge } from "../../ui/overlay-dialog.js";
import { type IntegrationsSnapshot } from "./integrations-overlay-types.js";

function isEnabledInList(integrationIds: readonly string[], integrationId: string): boolean {
  return integrationIds.includes(integrationId);
}

export function createIntegrationCard(args: {
  integration: IntegrationDefinition;
  snapshot: IntegrationsSnapshot;
  onToggleSession: (integrationId: string, next: boolean) => Promise<void>;
  onToggleWorkbook: (integrationId: string, next: boolean) => Promise<void>;
}): HTMLElement {
  const { integration, snapshot } = args;

  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-integrations-card";

  const top = document.createElement("div");
  top.className = "pi-integrations-card__top";

  const textWrap = document.createElement("div");
  textWrap.className = "pi-integrations-card__text-wrap";

  const title = document.createElement("strong");
  title.textContent = integration.title;
  title.className = "pi-integrations-card__title";

  const description = document.createElement("span");
  description.textContent = integration.description;
  description.className = "pi-integrations-card__description";

  textWrap.append(title, description);

  const badges = document.createElement("div");
  badges.className = "pi-overlay-badges";

  if (isEnabledInList(snapshot.activeIntegrationIds, integration.id) && snapshot.externalToolsEnabled) {
    badges.appendChild(createOverlayBadge("active", "ok"));
  } else if (isEnabledInList(snapshot.activeIntegrationIds, integration.id) && !snapshot.externalToolsEnabled) {
    badges.appendChild(createOverlayBadge("configured (blocked)", "warn"));
  } else {
    badges.appendChild(createOverlayBadge("inactive", "muted"));
  }

  top.append(textWrap, badges);

  const warning = document.createElement("div");
  warning.className = "pi-integrations-card__warning pi-overlay-text-warning";
  warning.textContent = integration.warning ?? "";
  warning.hidden = !integration.warning;

  const toggles = document.createElement("div");
  toggles.className = "pi-integrations-card__toggles";

  const sessionLabel = document.createElement("label");
  sessionLabel.className = "pi-integrations-card__toggle-label";

  const sessionToggle = document.createElement("input");
  sessionToggle.type = "checkbox";
  sessionToggle.checked = isEnabledInList(snapshot.sessionIntegrationIds, integration.id);
  sessionToggle.addEventListener("change", () => {
    void args.onToggleSession(integration.id, sessionToggle.checked);
  });
  sessionLabel.append(sessionToggle, document.createTextNode("Enable for this session"));

  const workbookLabel = document.createElement("label");
  workbookLabel.className = "pi-integrations-card__toggle-label";

  const workbookToggle = document.createElement("input");
  workbookToggle.type = "checkbox";
  workbookToggle.checked = isEnabledInList(snapshot.workbookIntegrationIds, integration.id);
  workbookToggle.disabled = snapshot.workbookId === null;
  workbookToggle.addEventListener("change", () => {
    void args.onToggleWorkbook(integration.id, workbookToggle.checked);
  });

  const workbookText = snapshot.workbookId
    ? `Enable for workbook (${snapshot.workbookLabel})`
    : "Workbook scope unavailable";

  workbookLabel.append(workbookToggle, document.createTextNode(workbookText));

  toggles.append(sessionLabel, workbookLabel);

  card.append(top, warning, toggles);
  return card;
}
