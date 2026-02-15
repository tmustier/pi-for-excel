/**
 * Add-ons entrypoint overlay.
 *
 * Keeps a single gear-menu item while preserving direct access to
 * Integrations, Skills, and Extensions managers.
 */

import { INTEGRATIONS_MANAGER_LABEL } from "../../integrations/naming.js";
import {
  closeOverlayById,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { ADDONS_OVERLAY_ID } from "../../ui/overlay-ids.js";

export interface AddonsDialogActions {
  openIntegrationsManager: () => void;
  openSkillsManager: () => void;
  openExtensionsManager: () => void;
}

export function showAddonsDialog(actions: AddonsDialogActions): void {
  if (closeOverlayById(ADDONS_OVERLAY_ID)) {
    return;
  }

  const dialog = createOverlayDialog({
    overlayId: ADDONS_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m",
  });

  const { header } = createOverlayHeader({
    onClose: dialog.close,
    closeLabel: "Close add-ons",
    title: "Add-ons",
    subtitle: "Manage tools & MCP, skills, and extensions.",
  });

  const body = document.createElement("div");
  body.className = "pi-overlay-body";

  const section = document.createElement("section");
  section.className = "pi-overlay-section";

  const card = document.createElement("div");
  card.className = "pi-overlay-surface";

  const actionsRow = document.createElement("div");
  actionsRow.className = "pi-overlay-actions";

  const integrationsButton = createOverlayButton({
    text: INTEGRATIONS_MANAGER_LABEL,
    className: "pi-overlay-btn--primary",
  });
  integrationsButton.addEventListener("click", () => {
    dialog.close();
    actions.openIntegrationsManager();
  });

  const skillsButton = createOverlayButton({
    text: "Skills",
  });
  skillsButton.addEventListener("click", () => {
    dialog.close();
    actions.openSkillsManager();
  });

  const extensionsButton = createOverlayButton({
    text: "Extensions",
  });
  extensionsButton.addEventListener("click", () => {
    dialog.close();
    actions.openExtensionsManager();
  });

  actionsRow.append(integrationsButton, skillsButton, extensionsButton);
  card.appendChild(actionsRow);
  section.appendChild(card);
  body.appendChild(section);

  dialog.card.append(header, body);
  dialog.mount();
}
