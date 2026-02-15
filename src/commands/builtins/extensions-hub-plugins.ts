/**
 * Extensions hub — Plugins tab.
 *
 * Installed plugins, install from URL / code, and advanced settings.
 */

import type {
  ExtensionRuntimeManager,
  ExtensionRuntimeStatus,
} from "../../extensions/runtime-manager.js";
import {
  describeExtensionCapability,
  getDefaultPermissionsForTrust,
  isExtensionCapabilityAllowed,
  listAllExtensionCapabilities,
  listGrantedExtensionCapabilities,
  type ExtensionCapability,
} from "../../extensions/permissions.js";
import { requestConfirmationDialog } from "../../ui/confirm-dialog.js";
import { showToast } from "../../ui/toast.js";
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
    title: `Enable plugin "${status.name}"?`,
    message: [
      "Granted higher-risk permissions:",
      ...risky.map((c) => `- ${describeExtensionCapability(c)}`),
      "",
      `Source: ${status.trustLabel}`,
    ].join("\n"),
    confirmLabel: "Enable",
    cancelLabel: "Cancel",
    confirmButtonTone: "danger",
    restoreFocusOnClose: false,
  });
}

async function confirmInstall(name: string, sourceLabel: string, capabilities: readonly ExtensionCapability[]): Promise<boolean> {
  const risky = capabilities.filter((c) => HIGH_RISK_CAPABILITIES.has(c));
  return requestConfirmationDialog({
    title: `Install plugin "${name}"?`,
    message: [
      `Source: ${sourceLabel}`,
      "",
      "Default permissions:",
      ...(capabilities.length > 0
        ? capabilities.map((c) => `- ${describeExtensionCapability(c)}`)
        : ["- (none)"]),
      ...(risky.length > 0
        ? ["", "Higher-risk:", ...risky.map((c) => `- ${describeExtensionCapability(c)}`)]
        : []),
    ].join("\n"),
    confirmLabel: "Install",
    cancelLabel: "Cancel",
    confirmButtonTone: risky.length > 0 ? "danger" : "primary",
    restoreFocusOnClose: false,
  });
}

// ── Main render ─────────────────────────────────────

export function renderPluginsTab(args: {
  container: HTMLElement;
  manager: ExtensionRuntimeManager;
  isBusy: () => boolean;
  onChanged: () => Promise<void>;
}): void {
  const { container, manager, isBusy, onChanged } = args;
  container.replaceChildren();

  const statuses = manager.list();

  // ── Installed section ─────────────────────────
  container.appendChild(createSectionHeader({
    label: "Installed",
    count: statuses.length,
  }));

  if (statuses.length === 0) {
    container.appendChild(createEmptyInline("🧩", "No plugins installed.\nInstall from a URL or paste code."));
  } else {
    const list = document.createElement("div");
    list.className = "pi-hub-stack";

    for (const status of statuses) {
      list.appendChild(renderPluginCard(status, manager, isBusy, onChanged, () => {
        renderPluginsTab(args);
      }));
    }
    container.appendChild(list);
  }

  // ── Install section ───────────────────────────
  container.appendChild(createSectionHeader({ label: "Install" }));

  const installForm = createAddForm();

  // URL install
  const urlRow = createAddFormRow();
  const urlInput = createAddFormInput("Extension URL (https://… or file:///…)");
  urlRow.append(
    urlInput,
    createButton("Install", {
      primary: true,
      compact: true,
      onClick: () => {
        if (isBusy()) return;
        const url = urlInput.value.trim();
        if (!url) { showToast("Enter a URL first."); return; }
        void installFromUrl(url, manager, onChanged, () => renderPluginsTab(args));
        urlInput.value = "";
      },
    }),
  );
  installForm.appendChild(urlRow);

  // Code install toggle
  const codeArea = document.createElement("textarea");
  codeArea.className = "pi-overlay-input pi-hub-textarea";
  codeArea.placeholder = "// Paste extension source code…";
  codeArea.hidden = true;

  const codeActions = document.createElement("div");
  codeActions.className = "pi-hub-actions-end";
  codeActions.hidden = true;

  codeActions.appendChild(createButton("Install from code", {
    primary: true,
    compact: true,
    onClick: () => {
      if (isBusy()) return;
      const code = codeArea.value.trim();
      if (!code) { showToast("Paste code first."); return; }
      void installFromCode(code, manager, onChanged, () => renderPluginsTab(args));
      codeArea.value = "";
    },
  }));

  const codeAreaId = "pi-plugin-install-code";
  codeArea.id = codeAreaId;

  const toggleCodeLink = document.createElement("button");
  toggleCodeLink.type = "button";
  toggleCodeLink.className = "pi-section-header__action pi-hub-code-toggle";
  toggleCodeLink.textContent = "Or paste code directly…";
  toggleCodeLink.setAttribute("aria-controls", codeAreaId);
  toggleCodeLink.setAttribute("aria-expanded", "false");
  toggleCodeLink.addEventListener("click", () => {
    const show = codeArea.hidden;
    codeArea.hidden = !show;
    codeActions.hidden = !show;
    toggleCodeLink.textContent = show ? "Hide code editor" : "Or paste code directly…";
    toggleCodeLink.setAttribute("aria-expanded", show ? "true" : "false");
    if (show) {
      codeArea.focus();
    }
  });

  const codeToggleRow = document.createElement("div");
  codeToggleRow.className = "pi-hub-code-toggle-row";
  codeToggleRow.appendChild(toggleCodeLink);

  installForm.append(codeToggleRow, codeArea, codeActions);
  container.appendChild(installForm);

  // ── Advanced section ──────────────────────────
  const details = document.createElement("details");
  

  const summary = document.createElement("summary");
  summary.className = "pi-hub-advanced-summary";
  summary.textContent = "Advanced";

  const advancedBody = document.createElement("div");
  advancedBody.className = "pi-hub-advanced-body";

  advancedBody.appendChild(createToggleRow({
    label: "Allow remote URLs",
    sublabel: "Install extensions from external URLs (requires trust)",
    checked: false, // Read from settings if available
  }).root);

  advancedBody.appendChild(createToggleRow({
    label: "LLM prompt template",
    sublabel: "Custom system prompt preamble",
    checked: false,
  }).root);

  details.append(summary, advancedBody);
  container.appendChild(details);
}

// ── Plugin card ─────────────────────────────────────

function renderPluginCard(
  status: ExtensionRuntimeStatus,
  manager: ExtensionRuntimeManager,
  isBusy: () => boolean,
  onChanged: () => Promise<void>,
  refresh: () => void,
): HTMLElement {
  const enableToggle = createToggle({
    checked: status.enabled,
    stopPropagation: true,
    onChange: (checked) => {
      if (isBusy()) return;
      void (async () => {
        if (checked && !status.enabled && !(await confirmEnable(status))) {
          enableToggle.input.checked = false;
          return;
        }
        try {
          await manager.setExtensionEnabled(status.id, checked);
          await onChanged();
          showToast(`${status.name}: ${checked ? "enabled" : "disabled"}`);
          refresh();
        } catch (err: unknown) {
          showToast(`Plugins: ${err instanceof Error ? err.message : String(err)}`);
          refresh();
        }
      })();
    },
  });

  const card = createItemCard({
    icon: "🧩",
    iconColor: "purple",
    name: status.name,
    description: `${status.sourceLabel} · ${status.runtimeLabel}`,
    expandable: true,
    rightContent: enableToggle.root,
  });

  // Commands
  if (status.commandNames.length > 0) {
    const cmds = status.commandNames.map((c: string) => `/${c}`).join(", ");
    card.body.appendChild(createConfigRow("Commands", createConfigValue(cmds)));
  }

  // Permissions grid
  const allCaps = listAllExtensionCapabilities();
  if (allCaps.length > 0) {
    card.body.appendChild(createSectionHeader({ label: "Permissions" }));

    const grid = document.createElement("div");
    grid.className = "pi-item-card__permissions";

    for (const cap of allCaps) {
      const allowed = isExtensionCapabilityAllowed(status.permissions, cap);
      const row = createToggleRow({
        label: describeExtensionCapability(cap),
        checked: allowed,
        onChange: (checked) => {
          void (async () => {
            try {
              await manager.setExtensionCapability(status.id, cap, checked);
              showToast(`${status.name}: ${describeExtensionCapability(cap)} ${checked ? "granted" : "revoked"}`);
              refresh();
            } catch (err: unknown) {
              showToast(`Plugins: ${err instanceof Error ? err.message : String(err)}`);
              refresh();
            }
          })();
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
    createButton("Uninstall", {
      danger: true,
      compact: true,
      onClick: () => {
        if (isBusy()) return;
        void (async () => {
          const ok = await requestConfirmationDialog({
            title: `Uninstall "${status.name}"?`,
            message: "This removes the plugin and its settings.",
            confirmLabel: "Uninstall",
            cancelLabel: "Cancel",
            confirmButtonTone: "danger",
            restoreFocusOnClose: false,
          });
          if (!ok) return;
          try {
            await manager.uninstallExtension(status.id);
            await onChanged();
            showToast(`Uninstalled: ${status.name}`);
            refresh();
          } catch (err: unknown) {
            showToast(`Plugins: ${err instanceof Error ? err.message : String(err)}`);
          }
        })();
      },
    }),
  ));

  // Error callout
  if (status.lastError) {
    card.body.appendChild(createCallout("warn", "⚠", `Error: ${status.lastError}`, { compact: true }));
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
  try {
    const name = window.prompt("Extension name:", "") ?? "";
    if (!name.trim()) return;
    const perms = getDefaultPermissionsForTrust("remote-url");
    const caps = listGrantedExtensionCapabilities(perms);
    if (!(await confirmInstall(name, url, caps))) return;
    await manager.installFromUrl(name, url);
    await onChanged();
    showToast(`Installed: ${name}`);
    refresh();
  } catch (err: unknown) {
    showToast(`Install failed: ${err instanceof Error ? err.message : String(err)}`);
  }
}

async function installFromCode(
  code: string,
  manager: ExtensionRuntimeManager,
  onChanged: () => Promise<void>,
  refresh: () => void,
): Promise<void> {
  try {
    const name = window.prompt("Extension name:", "my-extension") ?? "";
    if (!name.trim()) return;
    const perms = getDefaultPermissionsForTrust("inline-code");
    const caps = listGrantedExtensionCapabilities(perms);
    if (!(await confirmInstall(name, "inline code", caps))) return;
    await manager.installFromCode(name, code);
    await onChanged();
    showToast(`Installed: ${name}`);
    refresh();
  } catch (err: unknown) {
    showToast(`Install failed: ${err instanceof Error ? err.message : String(err)}`);
  }
}
