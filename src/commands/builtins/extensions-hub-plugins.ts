/**
 * Extensions hub — Plugins tab.
 *
 * Installed plugins and install from URL.
 */

import type {
  ExtensionRuntimeManager,
  ExtensionRuntimeStatus,
} from "../../extensions/runtime-manager.js";
import { classifyExtensionSource } from "../extension-source-policy.js";
import { isExperimentalFeatureEnabled } from "../../experiments/flags.js";
import {
  describeExtensionCapability,
  deriveStoredExtensionTrust,
  getDefaultPermissionsForTrust,
  isExtensionCapabilityAllowed,
  listAllExtensionCapabilities,
  listGrantedExtensionCapabilities,
  type ExtensionCapability,
} from "../../extensions/permissions.js";
import { requestConfirmationDialog } from "../../ui/confirm-dialog.js";
import { requestTextInputDialog } from "../../ui/text-input-dialog.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";
import {
  createSectionHeader,
  createItemCard,
  createConfigRow,
  createConfigValue,
  createCallout,
  createToggle,
  createToggleRow,
  createAddForm,
  createAddFormRow,
  createAddFormInput,
  createEmptyInline,
  createActionsRow,
  createButton,
} from "../../ui/extensions-hub-components.js";
import { lucide, AlertTriangle, Puzzle } from "../../ui/lucide-icons.js";

// ── Constants ───────────────────────────────────────

const HIGH_RISK_CAPABILITIES = new Set<ExtensionCapability>([
  "tools.register",
  "agent.read",
  "agent.events.read",
  "llm.complete",
  "http.fetch",
  "agent.context.write",
  "agent.steer",
  "agent.followup",
  "skills.write",
  "connections.readwrite",
  "connections.secrets.read",
]);

// ── Helpers ─────────────────────────────────────────

function getHighRiskGranted(status: ExtensionRuntimeStatus): ExtensionCapability[] {
  return listAllExtensionCapabilities().filter(
    (cap) => HIGH_RISK_CAPABILITIES.has(cap) && isExtensionCapabilityAllowed(status.permissions, cap),
  );
}

async function confirmEnable(status: ExtensionRuntimeStatus): Promise<boolean> {
  if (status.trust === "builtin") return true;
  const risky = getHighRiskGranted(status);
  if (risky.length === 0) return true;

  return requestConfirmationDialog({
    title: t("ext-hub-plugins.confirmEnableTitle", { name: status.name }),
    message: [
      t("ext-hub-plugins.confirm.grantedHighRisk"),
      ...risky.map((c) => `- ${describeExtensionCapability(c)}`),
      "",
      t("ext-hub-plugins.confirm.source", { source: status.trustLabel }),
    ].join("\n"),
    confirmLabel: t("ext-hub-plugins.enable"),
    cancelLabel: t("confirm.cancel"),
    confirmButtonTone: "danger",
    restoreFocusOnClose: false,
  });
}

/**
 * One in-taskpane dialog for URL installs: asks for the name and shows the
 * permissions this source will actually get (derived from its trust level).
 * Resolves to the trimmed name, or null when cancelled.
 */
async function requestInstallName(url: string): Promise<string | null> {
  const trust = deriveStoredExtensionTrust("ext.pending", { kind: "module", specifier: url });
  const capabilities = listGrantedExtensionCapabilities(getDefaultPermissionsForTrust(trust));
  const risky = capabilities.filter((c) => HIGH_RISK_CAPABILITIES.has(c));
  const name = await requestTextInputDialog({
    title: t("ext-hub-plugins.installDialogTitle"),
    message: [
      t("ext-hub-plugins.confirm.source", { source: url }),
      "",
      t("ext-hub-plugins.confirm.defaultPermissions"),
      ...(capabilities.length > 0
        ? capabilities.map((c) => `- ${describeExtensionCapability(c)}`)
        : [`- ${t("ext-hub-plugins.confirm.none")}`]),
      ...(risky.length > 0
        ? ["", t("ext-hub-plugins.confirm.higherRisk"), ...risky.map((c) => `- ${describeExtensionCapability(c)}`)]
        : []),
      "",
      t("ext-hub-plugins.installDialogName"),
    ].join("\n"),
    placeholder: t("ext-hub-plugins.installNamePlaceholder"),
    confirmLabel: t("ext-hub-plugins.installButton"),
    cancelLabel: t("confirm.cancel"),
    restoreFocusOnClose: false,
  });
  const trimmed = name?.trim() ?? "";
  return trimmed.length > 0 ? trimmed : null;
}

// ── Main render ─────────────────────────────────────

export function renderPluginsTab(args: {
  container: HTMLElement;
  manager: ExtensionRuntimeManager;
  /** Runs one user action at a time; returns false when another is in flight. */
  runExclusive: (action: () => Promise<void>) => Promise<boolean>;
  onChanged: () => Promise<void>;
}): void {
  const { container, manager, runExclusive, onChanged } = args;
  container.replaceChildren();

  const statuses = manager.list();

  // ── Installed section ─────────────────────────
  container.appendChild(createSectionHeader({
    label: t("ext-hub-plugins.installed"),
    count: statuses.length,
  }));

  if (statuses.length === 0) {
    container.appendChild(createEmptyInline(lucide(Puzzle), t("ext-hub-plugins.empty")));
  } else {
    const list = document.createElement("div");
    list.className = "pi-hub-stack";

    for (const status of statuses) {
      list.appendChild(renderPluginCard(status, manager, runExclusive, onChanged, () => {
        renderPluginsTab(args);
      }));
    }
    container.appendChild(list);
  }

  // ── Install from URL ───────────────────────────
  container.appendChild(createSectionHeader({ label: t("ext-hub-plugins.install") }));

  const installForm = createAddForm();
  const urlRow = createAddFormRow();
  const urlInput = createAddFormInput(t("ext-hub-plugins.pasteUrl"));
  urlRow.append(
    urlInput,
    createButton(t("ext-hub-plugins.installButton"), {
      primary: true,
      compact: true,
      onClick: () => {
        const url = urlInput.value.trim();
        if (!url) { showToast(t("extensions-hub-plugins.toast.enterUrl")); return; }
        void runExclusive(() => installFromUrl(url, manager, onChanged, () => renderPluginsTab(args)));
      },
    }),
  );
  installForm.appendChild(urlRow);
  container.appendChild(installForm);
}

// ── Plugin card ─────────────────────────────────────

function renderPluginCard(
  status: ExtensionRuntimeStatus,
  manager: ExtensionRuntimeManager,
  runExclusive: (action: () => Promise<void>) => Promise<boolean>,
  onChanged: () => Promise<void>,
  refresh: () => void,
): HTMLElement {
  const enableToggle = createToggle({
    checked: status.enabled,
    stopPropagation: true,
    onChange: (checked) => {
      void runExclusive(async () => {
        if (checked && !status.enabled && !(await confirmEnable(status))) {
          enableToggle.input.checked = false;
          return;
        }
        try {
          await manager.setExtensionEnabled(status.id, checked);
          await onChanged();
          showToast(t("ext-hub-plugins.toast.enabledState", { name: status.name, state: t(checked ? "experimental.state.enabled" : "experimental.state.disabled") }));
          refresh();
        } catch (err) {
          showToast(t("ext-hub-plugins.toast.error", { error: err instanceof Error ? err.message : String(err) }));
          refresh();
        }
      }).then((ran) => {
        if (!ran) enableToggle.input.checked = !checked;
      });
    },
  });

  const card = createItemCard({
    icon: lucide(Puzzle),
    iconColor: "purple",
    name: status.name,
    description: `${status.sourceLabel} · ${status.runtimeLabel}`,
    expandable: true,
    rightContent: enableToggle.root,
  });

  // Commands
  if (status.commandNames.length > 0) {
    const cmds = status.commandNames.map((c: string) => `/${c}`).join(", ");
    card.body.appendChild(createConfigRow(t("ext-hub-plugins.commands"), createConfigValue(cmds)));
  }

  // Permissions grid
  const allCaps = listAllExtensionCapabilities();
  if (allCaps.length > 0) {
    card.body.appendChild(createSectionHeader({ label: t("ext-hub-plugins.permissions") }));

    const grid = document.createElement("div");
    grid.className = "pi-item-card__permissions";

    for (const cap of allCaps) {
      const allowed = isExtensionCapabilityAllowed(status.permissions, cap);
      const row = createToggleRow({
        label: describeExtensionCapability(cap),
        checked: allowed,
        onChange: (checked) => {
          void runExclusive(async () => {
            try {
              await manager.setExtensionCapability(status.id, cap, checked);
              showToast(t("ext-hub-plugins.toast.permissionUpdated", { name: status.name, capability: describeExtensionCapability(cap), state: t(checked ? "ext-hub-plugins.granted" : "ext-hub-plugins.revoked") }));
              refresh();
            } catch (err) {
              showToast(t("ext-hub-plugins.toast.error", { error: err instanceof Error ? err.message : String(err) }));
              refresh();
            }
          }).then((ran) => {
            if (!ran) row.input.checked = !checked;
          });
        },
      });
      // Use sublabel styling for compact grid.
      // Avoid querySelector so this remains compatible with fake test DOMs.
      const labels = row.root.firstElementChild;
      if (labels instanceof HTMLElement) {
        const labelEl = labels.firstElementChild;
        if (labelEl instanceof HTMLElement) {
          labelEl.className = "pi-toggle-row__sublabel";
        }
      }
      grid.appendChild(row.root);
    }
    card.body.appendChild(grid);
  }

  // Uninstall
  card.body.appendChild(createActionsRow(
    createButton(t("ext-hub-plugins.uninstall"), {
      danger: true,
      compact: true,
      onClick: () => {
        void runExclusive(async () => {
          const ok = await requestConfirmationDialog({
            title: t("ext-hub-plugins.confirmUninstallTitle", { name: status.name }),
            message: t("ext-hub-plugins.confirm.uninstallMessage"),
            confirmLabel: t("ext-hub-plugins.uninstall"),
            cancelLabel: t("confirm.cancel"),
            confirmButtonTone: "danger",
            restoreFocusOnClose: false,
          });
          if (!ok) return;
          try {
            await manager.uninstallExtension(status.id);
            await onChanged();
            showToast(t("ext-hub-plugins.toast.uninstalled", { name: status.name }));
            refresh();
          } catch (err) {
            showToast(t("ext-hub-plugins.toast.error", { error: err instanceof Error ? err.message : String(err) }));
          }
        });
      },
    }),
  ));

  // Error callout
  if (status.lastError) {
    card.body.appendChild(createCallout("warn", lucide(AlertTriangle), t("ext-hub-plugins.lastError", { error: status.lastError }), { compact: true }));
  }

  return card.root;
}

// ── Install helpers ─────────────────────────────────

async function installFromUrl(
  url: string,
  manager: ExtensionRuntimeManager,
  onChanged: () => Promise<void>,
  refresh: () => void,
): Promise<void> {
  if (classifyExtensionSource(url) !== "remote-url") {
    showToast(t("ext-hub-plugins.toast.urlOnly"));
    return;
  }
  if (!isExperimentalFeatureEnabled("remote_extension_urls")) {
    showToast(t("ext-hub-plugins.toast.remoteOff"));
    return;
  }

  try {
    const name = await requestInstallName(url);
    if (name === null) return;
    await manager.installFromUrl(name, url);
    await onChanged();
    showToast(t("ext-hub-plugins.toast.installed", { name }));
    refresh();
  } catch (err) {
    showToast(t("ext-hub-plugins.toast.installFailed", { error: err instanceof Error ? err.message : String(err) }));
  }
}


