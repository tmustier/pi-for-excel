/**
 * Extensions pages — Connections, Plugins, and Skills.
 *
 * Thin page adapters over the existing tab renderers
 * (`extensions-hub-connections/plugins/skills`), with per-page busy
 * handling and live-refresh subscriptions scoped to the page lifecycle.
 */

import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

import { dispatchIntegrationsChanged } from "../../../integrations/events.js";
import { t } from "../../../language/index.js";
import type { SettingsPageContext, SettingsShellPage } from "../../../ui/settings-shell.js";
import { showToast } from "../../../ui/toast.js";
import { renderConnectionsTab } from "../extensions-hub-connections.js";
import { renderPluginsTab } from "../extensions-hub-plugins.js";
import { createDeferredConnectionsRefreshController } from "../extensions-hub-refresh.js";
import { renderSkillsTab } from "../extensions-hub-skills.js";
import { getSettingsPagesDependencies, type ExtensionsHubDependencies } from "./dependencies.js";

type MutationReason = "toggle" | "scope" | "external-toggle" | "config";

interface PageMutationHelper {
  isBusy: () => boolean;
  runMutation: (
    action: () => Promise<void>,
    reason: MutationReason,
    successMsg?: string,
  ) => Promise<void>;
}

function createPageMutationHelper(args: {
  deps: ExtensionsHubDependencies;
  isDisposed: () => boolean;
  refresh: () => Promise<void>;
}): PageMutationHelper {
  let busy = false;

  return {
    isBusy: () => busy,
    runMutation: async (action, reason, successMsg) => {
      if (busy || args.isDisposed()) return;
      busy = true;
      try {
        await action();
        dispatchIntegrationsChanged({ reason });
        if (args.deps.onChanged) await args.deps.onChanged();
        if (successMsg) showToast(successMsg);
      } catch (err) {
        showToast(t("extensions-hub.toast.error", { error: err instanceof Error ? err.message : String(err) }));
      } finally {
        busy = false;
        if (!args.isDisposed()) {
          await args.refresh();
        }
      }
    },
  };
}

function renderMissingDependencies(container: HTMLElement): void {
  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = t("extensions-hub.toast.error", { error: "not initialized" });
  container.appendChild(hint);
}

function setupPageLifecycle(ctx: SettingsPageContext): { isDisposed: () => boolean } {
  let disposed = false;
  ctx.addCleanup(() => {
    disposed = true;
  });
  return { isDisposed: () => disposed };
}

export function createConnectionsPage(): SettingsShellPage {
  return {
    id: "connections",
    parentId: "root",
    title: () => t("settings.page.connections"),
    subtitle: () => t("settings.page.connections.sub"),
    render: async (ctx) => {
      const deps = getSettingsPagesDependencies().extensionsHub;
      if (!deps) {
        renderMissingDependencies(ctx.body);
        return;
      }

      const lifecycle = setupPageLifecycle(ctx);
      const settings = getAppStorage().settings;

      const container = document.createElement("div");
      container.className = "pi-hub-stack pi-hub-stack--lg";
      ctx.body.appendChild(container);

      const refresh = async (): Promise<void> => {
        if (lifecycle.isDisposed()) return;
        await renderConnectionsTab({
          container,
          settings,
          deps,
          isBusy: helper.isBusy,
          runMutation: helper.runMutation,
        });
      };

      const helper = createPageMutationHelper({
        deps,
        isDisposed: lifecycle.isDisposed,
        refresh,
      });

      // Live refresh on background state changes, deferred while a
      // connection secret input has focus so edits are not wiped.
      const hasActiveSecretInput = (): boolean => {
        const active = document.activeElement;
        return active instanceof HTMLInputElement && container.contains(active);
      };

      const refreshController = createDeferredConnectionsRefreshController({
        isDisposed: lifecycle.isDisposed,
        hasActiveSecretInput,
        refresh: () => {
          if (!lifecycle.isDisposed()) void refresh();
        },
      });

      const onFocusOut = (): void => {
        refreshController.onConnectionsFocusOut();
      };

      container.addEventListener("focusout", onFocusOut);
      const unsubConnection = deps.connectionManager.subscribe(() => {
        refreshController.requestRefresh();
      });
      const unsubExtension = deps.extensionManager.subscribe(() => {
        refreshController.requestRefresh();
      });

      ctx.addCleanup(() => {
        refreshController.dispose();
        unsubConnection();
        unsubExtension();
        container.removeEventListener("focusout", onFocusOut);
      });

      try {
        await refresh();
      } catch (err) {
        showToast(t("extensions-hub.toast.error", { error: err instanceof Error ? err.message : String(err) }));
      }
    },
  };
}

export function createPluginsPage(): SettingsShellPage {
  return {
    id: "plugins",
    parentId: "root",
    title: () => t("settings.page.plugins"),
    subtitle: () => t("settings.page.plugins.sub"),
    render: (ctx) => {
      const deps = getSettingsPagesDependencies().extensionsHub;
      if (!deps) {
        renderMissingDependencies(ctx.body);
        return;
      }

      const lifecycle = setupPageLifecycle(ctx);

      const container = document.createElement("div");
      container.className = "pi-hub-stack pi-hub-stack--lg";
      ctx.body.appendChild(container);

      let busy = false;

      const refresh = (): void => {
        if (lifecycle.isDisposed()) return;
        renderPluginsTab({
          container,
          manager: deps.extensionManager,
          isBusy: () => busy,
          onChanged: async () => {
            if (deps.onChanged) await deps.onChanged();
          },
        });
      };

      const unsubExtension = deps.extensionManager.subscribe(() => {
        if (!lifecycle.isDisposed()) refresh();
      });
      ctx.addCleanup(() => {
        unsubExtension();
      });

      // renderPluginsTab drives its own mutations through the manager; the
      // busy flag only guards re-entrant clicks during manager updates.
      busy = false;
      refresh();
    },
  };
}

export function createSkillsPage(): SettingsShellPage {
  return {
    id: "skills",
    parentId: "root",
    title: () => t("settings.page.skills"),
    subtitle: () => t("settings.page.skills.sub"),
    render: async (ctx) => {
      const deps = getSettingsPagesDependencies().extensionsHub;
      if (!deps) {
        renderMissingDependencies(ctx.body);
        return;
      }

      const lifecycle = setupPageLifecycle(ctx);
      const settings = getAppStorage().settings;

      const container = document.createElement("div");
      container.className = "pi-hub-stack pi-hub-stack--lg";
      ctx.body.appendChild(container);

      const refresh = async (): Promise<void> => {
        if (lifecycle.isDisposed()) return;
        await renderSkillsTab({
          container,
          settings,
          isBusy: helper.isBusy,
          runMutation: helper.runMutation,
        });
      };

      const helper = createPageMutationHelper({
        deps,
        isDisposed: lifecycle.isDisposed,
        refresh,
      });

      try {
        await refresh();
      } catch (err) {
        showToast(t("extensions-hub.toast.error", { error: err instanceof Error ? err.message : String(err) }));
      }
    },
  };
}
