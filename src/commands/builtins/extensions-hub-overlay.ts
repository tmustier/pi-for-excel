/**
 * Extensions hub overlay — unified 3-tab overlay replacing the old
 * addons-overlay, integrations-overlay, extensions-overlay, and skills-overlay.
 *
 * Tabs: Connections | Plugins | Skills
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";
import type { ExtensionRuntimeManager } from "../../extensions/runtime-manager.js";
import { dispatchIntegrationsChanged } from "../../integrations/events.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { ADDONS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import { renderConnectionsTab } from "./extensions-hub-connections.js";
import { renderPluginsTab } from "./extensions-hub-plugins.js";
import { renderSkillsTab } from "./extensions-hub-skills.js";

export type ExtensionsHubTab = "connections" | "plugins" | "skills";

export interface WorkbookContextSnapshot {
  workbookId: string | null;
  workbookLabel: string;
}

export interface ExtensionsHubDependencies {
  getActiveSessionId: () => string | null;
  resolveWorkbookContext: () => Promise<WorkbookContextSnapshot>;
  extensionManager: ExtensionRuntimeManager;
  onChanged?: () => Promise<void> | void;
}

const TABS: ReadonlyArray<{ id: ExtensionsHubTab; label: string }> = [
  { id: "connections", label: "Connections" },
  { id: "plugins", label: "Plugins" },
  { id: "skills", label: "Skills" },
];

let openInFlight: Promise<void> | null = null;
let pendingTab: ExtensionsHubTab | null = null;

function activateTab(overlay: HTMLElement, tabId: ExtensionsHubTab): void {
  for (const btn of overlay.querySelectorAll<HTMLButtonElement>("[data-hub-tab]")) {
    const active = btn.dataset.hubTab === tabId;
    btn.classList.toggle("is-active", active);
    btn.setAttribute("aria-selected", active ? "true" : "false");
  }
  for (const panel of overlay.querySelectorAll<HTMLElement>("[data-hub-panel]")) {
    panel.hidden = panel.dataset.hubPanel !== tabId;
  }
}

export async function showExtensionsHubDialog(
  deps: ExtensionsHubDependencies,
  options?: { tab?: ExtensionsHubTab },
): Promise<void> {
  // Toggle or switch tab on existing overlay
  const existing = document.getElementById(ADDONS_OVERLAY_ID);
  if (existing instanceof HTMLElement) {
    if (options?.tab) {
      activateTab(existing, options.tab);
      return;
    }
    closeOverlayById(ADDONS_OVERLAY_ID);
    return;
  }

  // Debounce concurrent opens
  if (openInFlight) {
    if (options?.tab) pendingTab = options.tab;
    await openInFlight;
    const mounted = document.getElementById(ADDONS_OVERLAY_ID);
    if (mounted instanceof HTMLElement && options?.tab) activateTab(mounted, options.tab);
    return;
  }

  pendingTab = options?.tab ?? pendingTab;

  openInFlight = (async () => {
    const settings = getAppStorage().settings;

    const dialog = createOverlayDialog({
      overlayId: ADDONS_OVERLAY_ID,
      cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l",
    });

    const { header } = createOverlayHeader({
      onClose: dialog.close,
      closeLabel: "Close extensions",
      title: "Extensions",
      subtitle: "Connections, plugins, and skills that extend Pi",
    });

    // ── Tab bar ────────────────────────────────────
    const tabBar = document.createElement("div");
    tabBar.className = "pi-overlay-tabs";
    tabBar.setAttribute("role", "tablist");
    tabBar.setAttribute("aria-label", "Extensions tabs");

    for (const tab of TABS) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "pi-overlay-tab";
      btn.textContent = tab.label;
      btn.dataset.hubTab = tab.id;
      btn.setAttribute("role", "tab");
      btn.setAttribute("aria-selected", "false");
      btn.addEventListener("click", () => activateTab(dialog.overlay, tab.id));
      tabBar.appendChild(btn);
    }

    // ── Panels ─────────────────────────────────────
    const connectionsPanel = document.createElement("div");
    connectionsPanel.dataset.hubPanel = "connections";
    connectionsPanel.setAttribute("role", "tabpanel");
    connectionsPanel.className = "pi-hub-stack pi-hub-stack--lg";

    const pluginsPanel = document.createElement("div");
    pluginsPanel.dataset.hubPanel = "plugins";
    pluginsPanel.setAttribute("role", "tabpanel");
    pluginsPanel.className = "pi-hub-stack pi-hub-stack--lg";

    const skillsPanel = document.createElement("div");
    skillsPanel.dataset.hubPanel = "skills";
    skillsPanel.setAttribute("role", "tabpanel");
    skillsPanel.className = "pi-hub-stack pi-hub-stack--lg";

    const body = document.createElement("div");
    body.className = "pi-overlay-body";
    body.append(tabBar, connectionsPanel, pluginsPanel, skillsPanel);
    dialog.card.append(header, body);

    let disposed = false;
    dialog.addCleanup(() => { disposed = true; });

    // ── Shared mutation helper ─────────────────────
    let busy = false;

    const runMutation = async (
      action: () => Promise<void>,
      reason: "toggle" | "scope" | "external-toggle" | "config",
      successMsg?: string,
    ): Promise<void> => {
      if (busy || disposed) return;
      busy = true;
      try {
        await action();
        dispatchIntegrationsChanged({ reason });
        if (deps.onChanged) await deps.onChanged();
        if (successMsg) showToast(successMsg);
      } catch (err: unknown) {
        showToast(`Extensions: ${err instanceof Error ? err.message : String(err)}`);
      } finally {
        busy = false;
        if (!disposed) void refreshAll();
      }
    };

    const isBusy = (): boolean => busy;

    // ── Render each tab ────────────────────────────
    const refreshConnections = async (): Promise<void> => {
      if (disposed) return;
      await renderConnectionsTab({
        container: connectionsPanel,
        settings,
        deps,
        isBusy,
        runMutation,
      });
    };

    const refreshPlugins = (): void => {
      if (disposed) return;
      renderPluginsTab({
        container: pluginsPanel,
        manager: deps.extensionManager,
        isBusy,
        onChanged: async () => {
          if (deps.onChanged) await deps.onChanged();
        },
      });
    };

    const refreshSkills = async (): Promise<void> => {
      if (disposed) return;
      await renderSkillsTab({
        container: skillsPanel,
        settings,
        isBusy,
        runMutation,
      });
    };

    const refreshAll = async (): Promise<void> => {
      await Promise.all([
        refreshConnections(),
        Promise.resolve(refreshPlugins()),
        refreshSkills(),
      ]);
    };

    // ── Mount & initial render ─────────────────────
    try {
      await refreshAll();
    } catch (err: unknown) {
      showToast(`Extensions: ${err instanceof Error ? err.message : String(err)}`);
    }

    dialog.mount();

    const initialTab = pendingTab ?? "connections";
    pendingTab = null;
    requestAnimationFrame(() => {
      const el = document.getElementById(ADDONS_OVERLAY_ID);
      if (el instanceof HTMLElement) activateTab(el, initialTab);
    });
  })();

  try {
    await openInFlight;
  } finally {
    openInFlight = null;
  }
}
