/**
 * Extensions manager overlay.
 */

import type { ExtensionRuntimeManager, ExtensionRuntimeStatus } from "../../extensions/runtime-manager.js";
import {
  describeExtensionCapability,
  getDefaultPermissionsForTrust,
  isExtensionCapabilityAllowed,
  listAllExtensionCapabilities,
  listGrantedExtensionCapabilities,
  type ExtensionCapability,
} from "../../extensions/permissions.js";
import { validateOfficeProxyUrl } from "../../auth/proxy-validation.js";
import { dispatchExperimentalToolConfigChanged } from "../../experiments/events.js";
import { isExperimentalFeatureEnabled, setExperimentalFeatureEnabled } from "../../experiments/flags.js";
import { PYTHON_BRIDGE_URL_SETTING_KEY } from "../../tools/experimental-tool-gates.js";
import {
  closeOverlayById,
  createOverlayBadge,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
  createOverlayInput,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { EXTENSIONS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";

const EXTENSION_PROMPT_TEMPLATE = [
  "Write a single-file JavaScript ES module extension for Pi for Excel.",
  "Requirements:",
  "- Export activate(api)",
  "- No external imports",
  "- Use only the ExcelExtensionAPI surface (commands/tools, llm/http/storage/clipboard, agent steering/context, skills, download, overlay/widget/toast/onAgentEvent)",
  "- Keep it self-contained in one file",
  "- Include concise comments",
].join("\n");

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

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  return String(error);
}

function formatCapabilityList(capabilities: readonly string[]): string {
  return capabilities.length > 0 ? capabilities.join(", ") : "(none)";
}

function getCapabilityRiskLabel(capability: ExtensionCapability): string | null {
  if (!HIGH_RISK_CAPABILITIES.has(capability)) {
    return null;
  }

  return "higher risk";
}

function getHighRiskGrantedCapabilities(status: ExtensionRuntimeStatus): ExtensionCapability[] {
  const capabilities: ExtensionCapability[] = [];

  for (const capability of listAllExtensionCapabilities()) {
    if (!HIGH_RISK_CAPABILITIES.has(capability)) {
      continue;
    }

    if (isExtensionCapabilityAllowed(status.permissions, capability)) {
      capabilities.push(capability);
    }
  }

  return capabilities;
}

function confirmExtensionEnable(status: ExtensionRuntimeStatus): boolean {
  if (status.trust === "builtin") {
    return true;
  }

  const highRiskCapabilities = getHighRiskGrantedCapabilities(status);
  if (highRiskCapabilities.length === 0) {
    return true;
  }

  const lines = [
    `Enable extension "${status.name}" with higher-risk permissions?`,
    "",
    "Granted higher-risk permissions:",
    ...highRiskCapabilities.map((capability) => `- ${describeExtensionCapability(capability)}`),
    "",
    `Source: ${status.trustLabel}`,
    "",
    "You can edit permissions later in /extensions.",
  ];

  return window.confirm(lines.join("\n"));
}

function confirmExtensionInstall(args: {
  name: string;
  sourceLabel: string;
  capabilities: readonly ExtensionCapability[];
}): boolean {
  const highRiskCapabilities = args.capabilities.filter((capability) => HIGH_RISK_CAPABILITIES.has(capability));

  const lines = [
    `Install extension "${args.name}" from ${args.sourceLabel}?`,
    "",
    "Default granted permissions:",
    ...(args.capabilities.length > 0
      ? args.capabilities.map((capability) => `- ${describeExtensionCapability(capability)}`)
      : ["- (none)"]),
  ];

  if (highRiskCapabilities.length > 0) {
    lines.push(
      "",
      "Higher-risk default permissions:",
      ...highRiskCapabilities.map((capability) => `- ${describeExtensionCapability(capability)}`),
    );
  } else {
    lines.push("", "No higher-risk permissions are granted by default.");
  }

  lines.push("", "You can review/edit permissions later in /extensions.");

  return window.confirm(lines.join("\n"));
}

function createReadOnlyCodeBlock(text: string): HTMLTextAreaElement {
  const area = document.createElement("textarea");
  area.readOnly = true;
  area.value = text;
  area.className = "pi-overlay-code";
  return area;
}

async function getSettingsStore() {
  const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
  return storageModule.getAppStorage().settings;
}

async function readSettingValue(settingKey: string): Promise<string | undefined> {
  try {
    const settings = await getSettingsStore();
    const value = await settings.get<string>(settingKey);
    if (typeof value !== "string") return undefined;

    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  } catch {
    return undefined;
  }
}

async function writeSettingValue(settingKey: string, value: string): Promise<void> {
  const settings = await getSettingsStore();
  await settings.set(settingKey, value);
}

async function deleteSettingValue(settingKey: string): Promise<void> {
  const settings = await getSettingsStore();
  await settings.delete(settingKey);
}

export function showExtensionsDialog(manager: ExtensionRuntimeManager): void {
  if (closeOverlayById(EXTENSIONS_OVERLAY_ID)) {
    return;
  }

  const dialog = createOverlayDialog({
    overlayId: EXTENSIONS_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-ext-card",
  });

  const closeOverlay = dialog.close;

  const { header } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close extensions manager",
    title: "Extensions Manager",
    subtitle: "Extensions can read/write workbook data. Only enable code you trust.",
  });

  const body = document.createElement("div");
  body.className = "pi-overlay-body";

  const installedSection = document.createElement("section");
  installedSection.className = "pi-overlay-section";
  installedSection.appendChild(createOverlaySectionTitle("Installed"));

  const installedList = document.createElement("div");
  installedList.className = "pi-overlay-list";
  installedSection.appendChild(installedList);

  const sandboxSection = document.createElement("section");
  sandboxSection.className = "pi-overlay-section";
  sandboxSection.appendChild(createOverlaySectionTitle("Sandbox runtime (default for untrusted sources)"));

  const sandboxCard = document.createElement("div");
  sandboxCard.className = "pi-overlay-surface pi-ext-local-bridge-card";

  const sandboxStatusRow = document.createElement("div");
  sandboxStatusRow.className = "pi-ext-local-bridge-status-row";

  const sandboxStatusText = document.createElement("div");
  sandboxStatusText.className = "pi-ext-local-bridge-status-text";

  const sandboxStatusBadgeSlot = document.createElement("div");
  sandboxStatusBadgeSlot.className = "pi-ext-local-bridge-status-badge";

  sandboxStatusRow.append(sandboxStatusText, sandboxStatusBadgeSlot);

  const sandboxActions = document.createElement("div");
  sandboxActions.className = "pi-overlay-actions";

  const sandboxEnableButton = createOverlayButton({ text: "Enable host-runtime fallback" });
  const sandboxDisableButton = createOverlayButton({ text: "Disable host-runtime fallback" });

  sandboxActions.append(sandboxEnableButton, sandboxDisableButton);

  const sandboxHint = document.createElement("p");
  sandboxHint.textContent =
    "Default: inline-code and remote-URL extensions run in sandbox iframes. Enable host-runtime fallback only as a rollback kill switch. Built-in/local extensions always stay on host runtime.";
  sandboxHint.className = "pi-overlay-hint";

  sandboxCard.append(sandboxStatusRow, sandboxActions, sandboxHint);
  sandboxSection.appendChild(sandboxCard);

  const localBridgeSection = document.createElement("section");
  localBridgeSection.className = "pi-overlay-section";
  localBridgeSection.appendChild(createOverlaySectionTitle("Local Python / LibreOffice bridge (experimental)"));

  const localBridgeCard = document.createElement("div");
  localBridgeCard.className = "pi-overlay-surface pi-ext-local-bridge-card";

  const localBridgeStatusRow = document.createElement("div");
  localBridgeStatusRow.className = "pi-ext-local-bridge-status-row";

  const localBridgeStatusText = document.createElement("div");
  localBridgeStatusText.className = "pi-ext-local-bridge-status-text";

  const localBridgeStatusBadgeSlot = document.createElement("div");
  localBridgeStatusBadgeSlot.className = "pi-ext-local-bridge-status-badge";

  localBridgeStatusRow.append(localBridgeStatusText, localBridgeStatusBadgeSlot);

  const localBridgeUrlRow = document.createElement("div");
  localBridgeUrlRow.className = "pi-ext-local-bridge-url-row";

  const localBridgeUrlInput = createOverlayInput({ placeholder: "https://localhost:3340" });
  const localBridgeEnableButton = createOverlayButton({ text: "Enable + save URL" });
  const localBridgeSaveUrlButton = createOverlayButton({ text: "Save URL" });
  const localBridgeDisableButton = createOverlayButton({ text: "Disable" });

  localBridgeUrlRow.append(
    localBridgeUrlInput,
    localBridgeEnableButton,
    localBridgeSaveUrlButton,
    localBridgeDisableButton,
  );

  const localBridgeHint = document.createElement("p");
  localBridgeHint.textContent =
    "Enable the Python bridge and save the URL in one step (equivalent to two /experimental commands).";
  localBridgeHint.className = "pi-overlay-hint";

  localBridgeCard.append(localBridgeStatusRow, localBridgeUrlRow, localBridgeHint);
  localBridgeSection.appendChild(localBridgeCard);

  const installUrlSection = document.createElement("section");
  installUrlSection.className = "pi-overlay-section";
  installUrlSection.appendChild(createOverlaySectionTitle("Install from URL"));

  const installUrlRow = document.createElement("div");
  installUrlRow.className = "pi-ext-install-url-row";

  const installUrlName = createOverlayInput({ placeholder: "Name" });
  const installUrlInput = createOverlayInput({ placeholder: "https://example.com/pi-extension.js" });
  const installUrlButton = createOverlayButton({ text: "Install" });

  installUrlRow.append(installUrlName, installUrlInput, installUrlButton);

  const installUrlHint = document.createElement("p");
  installUrlHint.textContent = "Remote URLs require /experimental on remote-extension-urls.";
  installUrlHint.className = "pi-overlay-hint";

  installUrlSection.append(installUrlRow, installUrlHint);

  const installCodeSection = document.createElement("section");
  installCodeSection.className = "pi-overlay-section";
  installCodeSection.appendChild(createOverlaySectionTitle("Install from pasted code"));

  const installCodeName = createOverlayInput({ placeholder: "Name" });
  const installCodeText = document.createElement("textarea");
  installCodeText.placeholder = "export function activate(api) { ... }";
  installCodeText.className = "pi-ext-install-code";

  const installCodeActions = document.createElement("div");
  installCodeActions.className = "pi-overlay-actions";
  const installCodeButton = createOverlayButton({ text: "Install code" });
  installCodeActions.appendChild(installCodeButton);

  installCodeSection.append(installCodeName, installCodeText, installCodeActions);

  const templateSection = document.createElement("section");
  templateSection.className = "pi-overlay-section";
  templateSection.appendChild(createOverlaySectionTitle("LLM prompt template"));

  const templateCode = createReadOnlyCodeBlock(EXTENSION_PROMPT_TEMPLATE);
  const templateActions = document.createElement("div");
  templateActions.className = "pi-overlay-actions";
  const copyTemplateButton = createOverlayButton({ text: "Copy template" });
  templateActions.appendChild(copyTemplateButton);

  templateSection.append(templateCode, templateActions);

  body.append(
    installedSection,
    sandboxSection,
    localBridgeSection,
    installUrlSection,
    installCodeSection,
    templateSection,
  );

  dialog.card.append(header, body);

  const setBusy = (busy: boolean) => {
    installUrlButton.disabled = busy;
    installCodeButton.disabled = busy;
    copyTemplateButton.disabled = busy;
    localBridgeUrlInput.disabled = busy;
    localBridgeEnableButton.disabled = busy;
    localBridgeSaveUrlButton.disabled = busy;
    localBridgeDisableButton.disabled = busy;
    sandboxEnableButton.disabled = busy;
    sandboxDisableButton.disabled = busy;
  };

  const renderSandboxState = (): void => {
    const statuses = manager.list();
    const rollbackEnabled = isExperimentalFeatureEnabled("extension_sandbox_runtime");
    const untrusted = statuses.filter((status) => status.trust === "inline-code" || status.trust === "remote-url");
    const sandboxed = untrusted.filter((status) => status.runtimeMode === "sandbox-iframe");

    if (rollbackEnabled) {
      if (untrusted.length > 0) {
        sandboxStatusText.textContent =
          `Rollback enabled — ${untrusted.length} untrusted extension${untrusted.length === 1 ? "" : "s"} currently run in host runtime.`;
      } else {
        sandboxStatusText.textContent = "Rollback enabled — no untrusted extensions installed.";
      }
    } else if (untrusted.length > 0) {
      sandboxStatusText.textContent =
        `Default-on — ${sandboxed.length}/${untrusted.length} untrusted extension${untrusted.length === 1 ? "" : "s"} in sandbox runtime.`;
    } else {
      sandboxStatusText.textContent = "Default-on — no untrusted extensions installed.";
    }

    sandboxStatusBadgeSlot.replaceChildren(
      createOverlayBadge(rollbackEnabled ? "rollback" : "default-on", rollbackEnabled ? "warn" : "ok"),
    );

    sandboxEnableButton.disabled = rollbackEnabled;
    sandboxDisableButton.disabled = !rollbackEnabled;
  };

  const renderLocalBridgeState = async (): Promise<void> => {
    const configuredUrl = await readSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY);

    if (configuredUrl && localBridgeUrlInput.value.trim().length === 0) {
      localBridgeUrlInput.value = configuredUrl;
    }

    localBridgeStatusText.textContent = configuredUrl
      ? `Bridge URL: ${configuredUrl}`
      : "Bridge URL not set";

    localBridgeStatusBadgeSlot.replaceChildren(
      createOverlayBadge(configuredUrl ? "configured" : "not set", configuredUrl ? "ok" : "muted"),
    );
  };

  const runAction = async (action: () => Promise<void> | void): Promise<void> => {
    setBusy(true);
    try {
      await action();
    } catch (error: unknown) {
      showToast(`Extensions: ${getErrorMessage(error)}`);
    } finally {
      setBusy(false);
      renderInstalledList();
      renderSandboxState();
      void renderLocalBridgeState();
    }
  };

  const createInstalledRow = (status: ExtensionRuntimeStatus): HTMLDivElement => {
    const row = document.createElement("div");
    row.className = "pi-overlay-surface pi-ext-installed-row";

    const allCapabilities = listAllExtensionCapabilities();
    const grantedCapabilities = new Set(status.grantedCapabilities);

    const top = document.createElement("div");
    top.className = "pi-ext-installed-row__top";

    const info = document.createElement("div");
    info.className = "pi-ext-installed-row__info";

    const name = document.createElement("strong");
    name.textContent = status.name;
    name.className = "pi-ext-installed-row__name";

    const source = document.createElement("code");
    source.textContent = status.sourceLabel;
    source.className = "pi-ext-installed-row__source";

    info.append(name, source);

    const badges = document.createElement("div");
    badges.className = "pi-overlay-badges";

    if (!status.enabled) {
      badges.appendChild(createOverlayBadge("disabled", "muted"));
    } else if (status.lastError) {
      badges.appendChild(createOverlayBadge("error", "warn"));
    } else if (status.loaded) {
      badges.appendChild(createOverlayBadge("loaded", "ok"));
    } else {
      badges.appendChild(createOverlayBadge("pending", "muted"));
    }

    const trustIsUntrusted = status.trust === "remote-url" || status.trust === "inline-code";
    const trustBadgeColor = trustIsUntrusted ? "warn" : "muted";
    badges.appendChild(createOverlayBadge(status.trustLabel, trustBadgeColor));

    const runtimeBadgeColor = status.runtimeMode === "sandbox-iframe"
      ? "ok"
      : trustIsUntrusted
        ? "warn"
        : "muted";
    badges.appendChild(createOverlayBadge(status.runtimeLabel, runtimeBadgeColor));

    badges.appendChild(
      createOverlayBadge(`${status.effectiveCapabilities.length} permission${status.effectiveCapabilities.length === 1 ? "" : "s"}`, "muted"),
    );

    if (!status.permissionsEnforced) {
      badges.appendChild(createOverlayBadge("gates off", "warn"));
    }

    if (status.toolNames.length > 0) {
      badges.appendChild(createOverlayBadge(`${status.toolNames.length} tool${status.toolNames.length === 1 ? "" : "s"}`, "muted"));
    }

    if (status.commandNames.length > 0) {
      badges.appendChild(createOverlayBadge(`${status.commandNames.length} command${status.commandNames.length === 1 ? "" : "s"}`, "muted"));
    }

    top.append(info, badges);

    const details = document.createElement("div");
    details.className = "pi-ext-installed-row__details";

    if (status.commandNames.length > 0) {
      const commands = document.createElement("div");
      commands.textContent = `Commands: ${status.commandNames.map((name) => `/${name}`).join(", ")}`;
      commands.className = "pi-ext-installed-row__line";
      details.appendChild(commands);
    }

    if (status.toolNames.length > 0) {
      const tools = document.createElement("div");
      tools.textContent = `Tools: ${status.toolNames.join(", ")}`;
      tools.className = "pi-ext-installed-row__line";
      details.appendChild(tools);
    }

    const permissions = document.createElement("div");
    if (status.permissionsEnforced) {
      permissions.textContent = `Permissions: ${formatCapabilityList(status.effectiveCapabilities)}`;
    } else {
      permissions.textContent = "Permissions: all capabilities active (extension-permissions experiment is off)";
    }
    permissions.className = "pi-ext-installed-row__line";
    details.appendChild(permissions);

    const runtime = document.createElement("div");
    runtime.textContent = `Runtime: ${status.runtimeLabel}`;
    runtime.className = "pi-ext-installed-row__line";
    details.appendChild(runtime);

    if ((status.trust === "inline-code" || status.trust === "remote-url") && status.runtimeMode === "host") {
      const runtimeWarning = document.createElement("div");
      runtimeWarning.textContent =
        "Host-runtime fallback is enabled: this untrusted extension runs in host runtime. Disable fallback above to restore sandbox mode.";
      runtimeWarning.className = "pi-ext-installed-row__error";
      details.appendChild(runtimeWarning);
    }

    if (!status.permissionsEnforced) {
      const configuredPermissions = document.createElement("div");
      configuredPermissions.textContent = `Configured (inactive): ${formatCapabilityList(status.grantedCapabilities)}`;
      configuredPermissions.className = "pi-ext-installed-row__line";
      details.appendChild(configuredPermissions);
    }

    const permissionsEditor = document.createElement("div");
    permissionsEditor.className = "pi-ext-installed-row__permissions-editor";

    for (const capability of allCapabilities) {
      const toggleLabel = document.createElement("label");
      toggleLabel.className = "pi-ext-installed-row__perm-toggle";

      const toggle = document.createElement("input");
      toggle.type = "checkbox";
      toggle.checked = grantedCapabilities.has(capability)
        || isExtensionCapabilityAllowed(status.permissions, capability);

      const toggleText = document.createElement("span");
      toggleText.textContent = describeExtensionCapability(capability);

      const riskLabel = getCapabilityRiskLabel(capability);
      if (riskLabel) {
        const risk = document.createElement("span");
        risk.textContent = `(${riskLabel})`;
        risk.className = "pi-ext-installed-row__risk";
        toggleText.append(" ", risk);
      }

      toggleLabel.append(toggle, toggleText);
      permissionsEditor.appendChild(toggleLabel);

      toggle.addEventListener("change", () => {
        const nextAllowed = toggle.checked;
        toggle.disabled = true;

        void runAction(async () => {
          await manager.setExtensionCapability(status.id, capability, nextAllowed);

          const updated = manager.list().find((entry) => entry.id === status.id);
          if (!updated) {
            showToast(`Updated permissions for ${status.name}.`);
            return;
          }

          if (!updated.enabled) {
            showToast(`Updated permissions for ${status.name}.`);
            return;
          }

          if (updated.lastError) {
            showToast(`Updated permissions for ${status.name}; reload failed (see Last error).`);
            return;
          }

          if (updated.loaded) {
            showToast(`Updated permissions for ${status.name}; extension reloaded.`);
            return;
          }

          showToast(`Updated permissions for ${status.name}.`);
        });
      });
    }

    details.appendChild(permissionsEditor);

    if (status.lastError) {
      const errorLine = document.createElement("div");
      errorLine.textContent = `Last error: ${status.lastError}`;
      errorLine.className = "pi-ext-installed-row__error";
      details.appendChild(errorLine);
    }

    const actions = document.createElement("div");
    actions.className = "pi-overlay-actions pi-overlay-actions--wrap";

    const toggleButton = createOverlayButton({ text: status.enabled ? "Disable" : "Enable" });
    toggleButton.addEventListener("click", () => {
      const nextEnabled = !status.enabled;

      if (nextEnabled && !confirmExtensionEnable(status)) {
        return;
      }

      void runAction(async () => {
        await manager.setExtensionEnabled(status.id, nextEnabled);
      });
    });

    const reloadButton = createOverlayButton({ text: "Reload" });
    reloadButton.disabled = !status.enabled;
    reloadButton.addEventListener("click", () => {
      void runAction(async () => {
        await manager.reloadExtension(status.id);
      });
    });

    const uninstallButton = createOverlayButton({ text: "Uninstall" });
    uninstallButton.addEventListener("click", () => {
      const confirmed = window.confirm(`Uninstall extension "${status.name}"?`);
      if (!confirmed) {
        return;
      }

      void runAction(async () => {
        await manager.uninstallExtension(status.id);
      });
    });

    actions.append(toggleButton, reloadButton, uninstallButton);
    row.append(top, details, actions);
    return row;
  };

  const renderInstalledList = () => {
    const statuses = manager.list();
    installedList.replaceChildren();

    if (statuses.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = "No extensions installed.";
      empty.className = "pi-overlay-empty";
      installedList.appendChild(empty);
      return;
    }

    for (const status of statuses) {
      installedList.appendChild(createInstalledRow(status));
    }
  };

  localBridgeEnableButton.addEventListener("click", () => {
    void runAction(async () => {
      const candidateUrl = localBridgeUrlInput.value.trim();
      if (candidateUrl.length === 0) {
        throw new Error("Provide a Python bridge URL (example: https://localhost:3340)");
      }

      const normalizedUrl = validateOfficeProxyUrl(candidateUrl);
      await writeSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY, normalizedUrl);
      dispatchExperimentalToolConfigChanged({ configKey: PYTHON_BRIDGE_URL_SETTING_KEY });

      localBridgeUrlInput.value = normalizedUrl;
      showToast(`Python bridge configured at ${normalizedUrl}`);
    });
  });

  localBridgeSaveUrlButton.addEventListener("click", () => {
    void runAction(async () => {
      const candidateUrl = localBridgeUrlInput.value.trim();
      if (candidateUrl.length === 0) {
        await deleteSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY);
        dispatchExperimentalToolConfigChanged({ configKey: PYTHON_BRIDGE_URL_SETTING_KEY });
        showToast("Python bridge URL cleared.");
        return;
      }

      const normalizedUrl = validateOfficeProxyUrl(candidateUrl);
      await writeSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY, normalizedUrl);
      dispatchExperimentalToolConfigChanged({ configKey: PYTHON_BRIDGE_URL_SETTING_KEY });

      localBridgeUrlInput.value = normalizedUrl;
      showToast(`Python bridge URL saved: ${normalizedUrl}`);
    });
  });

  localBridgeDisableButton.addEventListener("click", () => {
    void runAction(async () => {
      await deleteSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY);
      dispatchExperimentalToolConfigChanged({ configKey: PYTHON_BRIDGE_URL_SETTING_KEY });
      localBridgeUrlInput.value = "";
      showToast("Python bridge URL cleared.");
    });
  });

  sandboxEnableButton.addEventListener("click", () => {
    const confirmed = window.confirm(
      "Enable host-runtime fallback? Untrusted extensions will run in host runtime until fallback is disabled.",
    );
    if (!confirmed) {
      return;
    }

    void runAction(() => {
      setExperimentalFeatureEnabled("extension_sandbox_runtime", true);
      showToast("Extension host-runtime fallback enabled.");
    });
  });

  sandboxDisableButton.addEventListener("click", () => {
    void runAction(() => {
      setExperimentalFeatureEnabled("extension_sandbox_runtime", false);
      showToast("Extension host-runtime fallback disabled.");
    });
  });

  installUrlButton.addEventListener("click", () => {
    const name = installUrlName.value.trim();
    const url = installUrlInput.value.trim();
    if (name.length === 0 || url.length === 0) {
      showToast("Extensions: Provide both name and URL");
      return;
    }

    const defaultPermissions = getDefaultPermissionsForTrust("remote-url");
    const defaultCapabilities = listGrantedExtensionCapabilities(defaultPermissions);
    const confirmed = confirmExtensionInstall({
      name,
      sourceLabel: `remote URL (${url})`,
      capabilities: defaultCapabilities,
    });
    if (!confirmed) {
      return;
    }

    void runAction(async () => {
      await manager.installFromUrl(name, url);
      installUrlInput.value = "";
      showToast(`Installed extension: ${name}`);
    });
  });

  installCodeButton.addEventListener("click", () => {
    const name = installCodeName.value.trim();
    const code = installCodeText.value;
    if (name.length === 0) {
      showToast("Extensions: Provide an extension name");
      return;
    }

    const defaultPermissions = getDefaultPermissionsForTrust("inline-code");
    const defaultCapabilities = listGrantedExtensionCapabilities(defaultPermissions);
    const confirmed = confirmExtensionInstall({
      name,
      sourceLabel: "pasted code",
      capabilities: defaultCapabilities,
    });
    if (!confirmed) {
      return;
    }

    void runAction(async () => {
      await manager.installFromCode(name, code);
      installCodeText.value = "";
      showToast(`Installed extension: ${name}`);
    });
  });

  copyTemplateButton.addEventListener("click", () => {
    void runAction(async () => {
      if (!navigator.clipboard?.writeText) {
        throw new Error("Clipboard API not available");
      }

      await navigator.clipboard.writeText(EXTENSION_PROMPT_TEMPLATE);
      showToast("Template copied");
    });
  });

  const unsubscribe = manager.subscribe(() => {
    renderInstalledList();
    renderSandboxState();
    void renderLocalBridgeState();
  });

  dialog.addCleanup(unsubscribe);

  dialog.mount();
  renderInstalledList();
  renderSandboxState();
  void renderLocalBridgeState();
}
