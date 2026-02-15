/**
 * Extensions overlay (formerly Add-ons).
 *
 * Unified entry point for:
 * - Connections
 * - Plugins
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
import { renderPluginsSection } from "./addons-overlay-extensions.js";
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

const EXTENSIONS_TABS: ReadonlyArray<{ section: AddonsSection; label: string }> = [
  { section: "connections", label: "Connections" },
  { section: "plugins", label: "Plugins" },
  { section: "skills", label: "Skills" },
];

let addonsDialogOpenInFlight: Promise<void> | null = null;
let pendingSectionFocus: AddonsSection | null = null;

function activateExtensionsSection(overlay: HTMLElement, section: AddonsSection): void {
  const tabButtons = overlay.querySelectorAll<HTMLButtonElement>("[data-extensions-tab]");
  for (const button of tabButtons) {
    const isActive = button.dataset.extensionsTab === section;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  }

  const panels = overlay.querySelectorAll<HTMLElement>("[data-extensions-panel]");
  for (const panel of panels) {
    panel.hidden = panel.dataset.extensionsPanel !== section;
  }
}

function focusAddonsSection(overlay: HTMLElement, section: AddonsSection): void {
  activateExtensionsSection(overlay, section);
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
      closeLabel: "Close extensions",
      title: "Extensions",
      subtitle: "Connections, plugins, and skills in one place.",
    });

    const body = document.createElement("div");
    body.className = "pi-overlay-body pi-addons-body";

    const tabs = document.createElement("div");
    tabs.className = "pi-overlay-tabs";
    tabs.setAttribute("role", "tablist");
    tabs.setAttribute("aria-label", "Extensions sections");

    const sectionsWrap = document.createElement("div");
    sectionsWrap.className = "pi-addons-sections";

    const connectionsContainer = document.createElement("div");
    connectionsContainer.dataset.extensionsPanel = "connections";

    const pluginsContainer = document.createElement("div");
    pluginsContainer.dataset.extensionsPanel = "plugins";

    const skillsContainer = document.createElement("div");
    skillsContainer.dataset.extensionsPanel = "skills";

    sectionsWrap.append(connectionsContainer, pluginsContainer, skillsContainer);

    for (const tab of EXTENSIONS_TABS) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "pi-overlay-tab";
      button.textContent = tab.label;
      button.dataset.extensionsTab = tab.section;
      button.setAttribute("role", "tab");
      button.setAttribute("aria-selected", "false");
      button.addEventListener("click", () => {
        activateExtensionsSection(dialog.overlay, tab.section);
      });
      tabs.appendChild(button);
    }

    body.append(tabs, sectionsWrap);
    dialog.card.append(header, body);

    let disposed = false;
    dialog.addCleanup(() => {
      disposed = true;
    });

    let connectionsBusy = false;

    const refreshPlugins = (): void => {
      if (disposed) {
        return;
      }

      renderPluginsSection({
        container: pluginsContainer,
        actions: managedActions,
        busy: connectionsBusy,
        onRefresh: refreshPlugins,
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
            showToast(`Extensions: ${message}`);
          } finally {
            connectionsBusy = false;
            await refreshConnections();
            refreshPlugins();
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
      Promise.resolve(refreshPlugins()),
      refreshSkills(),
    ]);

    dialog.mount();

    const initialSection = pendingSectionFocus ?? "connections";
    pendingSectionFocus = null;
    requestAnimationFrame(() => {
      const mounted = document.getElementById(ADDONS_OVERLAY_ID);
      if (mounted instanceof HTMLElement) {
        activateExtensionsSection(mounted, initialSection);
      }
    });
  })();

  try {
    await addonsDialogOpenInFlight;
  } finally {
    addonsDialogOpenInFlight = null;
  }
}
