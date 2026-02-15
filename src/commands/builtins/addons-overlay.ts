/**
 * Add-ons overlay.
 *
 * Unified entry point for:
 * - Connections
 * - Extensions
 * - Skills
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import { dispatchIntegrationsChanged } from "../../integrations/events.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { ADDONS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import {
  buildConnectionsSnapshot,
  renderConnectionsSection,
} from "./addons-overlay-connections.js";
import { renderExtensionsSection } from "./addons-overlay-extensions.js";
import {
  buildSkillsSnapshot,
  renderSkillsSection,
} from "./addons-overlay-skills.js";
import type {
  AddonsDialogActions,
  AddonsSection,
  ShowAddonsDialogOptions,
} from "./addons-overlay-types.js";

export type { AddonsSection, ShowAddonsDialogOptions, AddonsDialogActions } from "./addons-overlay-types.js";

let addonsDialogOpenInFlight: Promise<void> | null = null;
let pendingSectionFocus: AddonsSection | null = null;

function sectionSelector(section: AddonsSection): string {
  return `[data-addons-section=\"${section}\"]`;
}

function focusAddonsSection(overlay: HTMLElement, section: AddonsSection): void {
  const target = overlay.querySelector<HTMLElement>(sectionSelector(section));
  if (!target) {
    return;
  }

  target.scrollIntoView({ behavior: "smooth", block: "start" });
}

export async function showAddonsDialog(
  actions: AddonsDialogActions,
  options: ShowAddonsDialogOptions = {},
): Promise<void> {
  const existing = document.getElementById(ADDONS_OVERLAY_ID);
  if (existing instanceof HTMLElement) {
    if (options.section) {
      focusAddonsSection(existing, options.section);
      return;
    }

    closeOverlayById(ADDONS_OVERLAY_ID);
    return;
  }

  if (addonsDialogOpenInFlight) {
    if (options.section) {
      pendingSectionFocus = options.section;
    }

    await addonsDialogOpenInFlight;

    const mounted = document.getElementById(ADDONS_OVERLAY_ID);
    if (mounted instanceof HTMLElement && options.section) {
      focusAddonsSection(mounted, options.section);
    }
    return;
  }

  pendingSectionFocus = options.section ?? pendingSectionFocus;

  addonsDialogOpenInFlight = (async () => {
    const settings = getAppStorage().settings;

    const dialog = createOverlayDialog({
      overlayId: ADDONS_OVERLAY_ID,
      cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-addons-dialog",
    });

    const managedActions: AddonsDialogActions = {
      ...actions,
      openIntegrationsManager: () => {
        dialog.close();
        actions.openIntegrationsManager();
      },
      openExtensionsManager: () => {
        dialog.close();
        actions.openExtensionsManager();
      },
      openSkillsManager: () => {
        dialog.close();
        actions.openSkillsManager();
      },
    };

    const { header } = createOverlayHeader({
      onClose: dialog.close,
      closeLabel: "Close add-ons",
      title: "Add-ons",
      subtitle: "Connections, extensions, and skills in one place.",
    });

    const body = document.createElement("div");
    body.className = "pi-overlay-body pi-addons-body";

    const connectionsContainer = document.createElement("div");
    const extensionsContainer = document.createElement("div");
    const skillsContainer = document.createElement("div");
    body.append(connectionsContainer, extensionsContainer, skillsContainer);

    dialog.card.append(header, body);

    let disposed = false;
    dialog.addCleanup(() => {
      disposed = true;
    });

    let connectionsBusy = false;

    const refreshExtensions = (): void => {
      if (disposed) {
        return;
      }

      renderExtensionsSection({
        container: extensionsContainer,
        actions: managedActions,
        busy: connectionsBusy,
        onRefresh: refreshExtensions,
      });
    };

    const refreshSkills = async (): Promise<void> => {
      if (disposed) {
        return;
      }

      try {
        const snapshot = await buildSkillsSnapshot(settings);
        if (disposed) {
          return;
        }

        renderSkillsSection({
          container: skillsContainer,
          actions: managedActions,
          snapshot,
        });
      } catch (error: unknown) {
        skillsContainer.replaceChildren();

        const section = document.createElement("section");
        section.className = "pi-overlay-section pi-addons-section";
        section.dataset.addonsSection = "skills";

        const title = document.createElement("h3");
        title.className = "pi-overlay-section-title";
        title.textContent = "Skills";

        const warning = document.createElement("p");
        warning.className = "pi-overlay-hint pi-overlay-text-warning";
        warning.textContent = `Failed to load skills: ${error instanceof Error ? error.message : "Unknown error"}`;

        section.append(title, warning);
        skillsContainer.appendChild(section);
      }
    };

    const refreshConnections = async (): Promise<void> => {
      if (disposed) {
        return;
      }

      try {
        const snapshot = await buildConnectionsSnapshot(settings, managedActions);
        if (disposed) {
          return;
        }

        const onMutate = async (
          mutation: () => Promise<void>,
          reason: "toggle" | "scope" | "external-toggle" | "config",
          successMessage?: string,
        ): Promise<void> => {
          if (connectionsBusy) {
            return;
          }

          connectionsBusy = true;
          try {
            await mutation();
            dispatchIntegrationsChanged({ reason });
            if (managedActions.onChanged) {
              await managedActions.onChanged();
            }
            if (successMessage) {
              showToast(successMessage);
            }
          } catch (error: unknown) {
            const message = error instanceof Error ? error.message : "Unknown error";
            showToast(`Add-ons: ${message}`);
          } finally {
            connectionsBusy = false;
            await refreshConnections();
            refreshExtensions();
          }
        };

        renderConnectionsSection({
          container: connectionsContainer,
          snapshot,
          settings,
          actions: managedActions,
          busy: connectionsBusy,
          onRefresh: () => {
            void refreshConnections();
          },
          onMutate,
        });
      } catch (error: unknown) {
        connectionsContainer.replaceChildren();

        const section = document.createElement("section");
        section.className = "pi-overlay-section pi-addons-section";
        section.dataset.addonsSection = "connections";

        const title = document.createElement("h3");
        title.className = "pi-overlay-section-title";
        title.textContent = "Connections";

        const warning = document.createElement("p");
        warning.className = "pi-overlay-hint pi-overlay-text-warning";
        warning.textContent = `Failed to load connections: ${error instanceof Error ? error.message : "Unknown error"}`;

        section.append(title, warning);
        connectionsContainer.appendChild(section);
      }
    };

    await Promise.all([
      refreshConnections(),
      Promise.resolve(refreshExtensions()),
      refreshSkills(),
    ]);

    dialog.mount();

    if (pendingSectionFocus) {
      const section = pendingSectionFocus;
      pendingSectionFocus = null;
      requestAnimationFrame(() => {
        const mounted = document.getElementById(ADDONS_OVERLAY_ID);
        if (mounted instanceof HTMLElement) {
          focusAddonsSection(mounted, section);
        }
      });
    }
  })();

  try {
    await addonsDialogOpenInFlight;
  } finally {
    addonsDialogOpenInFlight = null;
  }
}
