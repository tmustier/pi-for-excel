/**
 * Extensions manager overlay.
 */

import type { ExtensionRuntimeManager, ExtensionRuntimeStatus } from "../../extensions/runtime-manager.js";
import { validateOfficeProxyUrl } from "../../auth/proxy-validation.js";
import { dispatchExperimentalToolConfigChanged } from "../../experiments/events.js";
import { isExperimentalFeatureEnabled, setExperimentalFeatureEnabled } from "../../experiments/flags.js";
import { PYTHON_BRIDGE_URL_SETTING_KEY } from "../../tools/experimental-tool-gates.js";
import { showToast } from "../../ui/toast.js";

const OVERLAY_ID = "pi-extensions-overlay";
const overlayClosers = new WeakMap<HTMLElement, () => void>();

const EXTENSION_PROMPT_TEMPLATE = [
  "Write a single-file JavaScript ES module extension for Pi for Excel.",
  "Requirements:",
  "- Export activate(api)",
  "- No external imports",
  "- Use only the ExcelExtensionAPI surface (registerCommand, registerTool, overlay/widget/toast/onAgentEvent)",
  "- Keep it self-contained in one file",
  "- Include concise comments",
].join("\n");

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  return String(error);
}

function createSectionTitle(text: string): HTMLHeadingElement {
  const title = document.createElement("h3");
  title.textContent = text;
  title.style.cssText = "font-size: 13px; margin: 0; font-weight: 600;";
  return title;
}

function createButton(text: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.textContent = text;
  button.style.cssText =
    "padding: 6px 10px; border-radius: 8px; border: 1px solid oklch(0 0 0 / 0.08); "
    + "background: oklch(0 0 0 / 0.02); cursor: pointer; font-size: 12px;";
  return button;
}

function createInput(placeholder: string): HTMLInputElement {
  const input = document.createElement("input");
  input.type = "text";
  input.placeholder = placeholder;
  input.style.cssText =
    "width: 100%; padding: 7px 9px; border-radius: 8px; border: 1px solid oklch(0 0 0 / 0.12); "
    + "font-size: 12px; font-family: var(--font-sans); background: white;";
  return input;
}

function createBadge(text: string, color: "ok" | "warn" | "muted"): HTMLSpanElement {
  const badge = document.createElement("span");
  badge.textContent = text;
  const palette =
    color === "ok"
      ? "background: oklch(0.58 0.14 160 / 0.12); color: oklch(0.42 0.1 160); border-color: oklch(0.58 0.14 160 / 0.4);"
      : color === "warn"
        ? "background: oklch(0.67 0.17 35 / 0.12); color: oklch(0.5 0.12 35); border-color: oklch(0.67 0.17 35 / 0.35);"
        : "background: oklch(0 0 0 / 0.03); color: var(--muted-foreground); border-color: oklch(0 0 0 / 0.08);";

  badge.style.cssText =
    `font-size: 10px; padding: 2px 6px; border-radius: 999px; border: 1px solid; ${palette}`;
  return badge;
}

function createReadOnlyCodeBlock(text: string): HTMLTextAreaElement {
  const area = document.createElement("textarea");
  area.readOnly = true;
  area.value = text;
  area.style.cssText =
    "width: 100%; min-height: 84px; resize: vertical; border-radius: 8px; "
    + "border: 1px solid oklch(0 0 0 / 0.1); padding: 8px; font-size: 11px; "
    + "font-family: var(--font-mono); background: oklch(0 0 0 / 0.02);";
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
  const existing = document.getElementById(OVERLAY_ID);
  if (existing) {
    const closeExisting = overlayClosers.get(existing);
    if (closeExisting) {
      closeExisting();
    } else {
      existing.remove();
    }

    return;
  }

  const overlay = document.createElement("div");
  overlay.id = OVERLAY_ID;
  overlay.className = "pi-welcome-overlay";

  const card = document.createElement("div");
  card.className = "pi-welcome-card";
  card.style.cssText =
    "text-align: left; width: min(760px, 92vw); max-height: 88vh; overflow: hidden; "
    + "display: flex; flex-direction: column; gap: 12px;";

  const header = document.createElement("div");
  header.style.cssText = "display: flex; align-items: flex-start; justify-content: space-between; gap: 10px;";

  const titleWrap = document.createElement("div");
  titleWrap.style.cssText = "display: flex; flex-direction: column; gap: 4px;";

  const title = document.createElement("h2");
  title.textContent = "Extensions Manager";
  title.style.cssText = "font-size: 16px; font-weight: 600; margin: 0;";

  const subtitle = document.createElement("p");
  subtitle.textContent = "Extensions can read/write workbook data. Only enable code you trust.";
  subtitle.style.cssText = "margin: 0; font-size: 11px; color: var(--muted-foreground);";

  titleWrap.append(title, subtitle);

  let closeOverlay = () => {
    overlay.remove();
  };

  const closeButton = createButton("Close");
  closeButton.addEventListener("click", () => {
    closeOverlay();
  });

  header.append(titleWrap, closeButton);

  const body = document.createElement("div");
  body.style.cssText = "overflow-y: auto; display: flex; flex-direction: column; gap: 12px; padding-right: 4px;";

  const installedSection = document.createElement("section");
  installedSection.style.cssText = "display: flex; flex-direction: column; gap: 8px;";
  installedSection.appendChild(createSectionTitle("Installed"));

  const installedList = document.createElement("div");
  installedList.style.cssText = "display: flex; flex-direction: column; gap: 8px;";
  installedSection.appendChild(installedList);

  const localBridgeSection = document.createElement("section");
  localBridgeSection.style.cssText = "display: flex; flex-direction: column; gap: 8px;";
  localBridgeSection.appendChild(createSectionTitle("Local Python / LibreOffice bridge (experimental)"));

  const localBridgeCard = document.createElement("div");
  localBridgeCard.style.cssText =
    "display: flex; flex-direction: column; gap: 8px; border: 1px solid oklch(0 0 0 / 0.08); "
    + "background: oklch(0 0 0 / 0.015); border-radius: 10px; padding: 9px;";

  const localBridgeStatusRow = document.createElement("div");
  localBridgeStatusRow.style.cssText = "display: flex; justify-content: space-between; align-items: center; gap: 10px;";

  const localBridgeStatusText = document.createElement("div");
  localBridgeStatusText.style.cssText = "font-size: 12px; color: var(--muted-foreground);";

  const localBridgeStatusBadgeSlot = document.createElement("div");
  localBridgeStatusBadgeSlot.style.cssText = "display: flex; align-items: center;";

  localBridgeStatusRow.append(localBridgeStatusText, localBridgeStatusBadgeSlot);

  const localBridgeUrlRow = document.createElement("div");
  localBridgeUrlRow.style.cssText = "display: grid; grid-template-columns: 1fr auto auto auto; gap: 8px; align-items: center;";

  const localBridgeUrlInput = createInput("https://localhost:3340");
  const localBridgeEnableButton = createButton("Enable + save URL");
  const localBridgeSaveUrlButton = createButton("Save URL");
  const localBridgeDisableButton = createButton("Disable");

  localBridgeUrlRow.append(
    localBridgeUrlInput,
    localBridgeEnableButton,
    localBridgeSaveUrlButton,
    localBridgeDisableButton,
  );

  const localBridgeHint = document.createElement("p");
  localBridgeHint.textContent =
    "One-step setup from this menu: enable python-bridge + save URL (same as two /experimental commands).";
  localBridgeHint.style.cssText = "margin: 0; font-size: 11px; color: var(--muted-foreground);";

  localBridgeCard.append(localBridgeStatusRow, localBridgeUrlRow, localBridgeHint);
  localBridgeSection.appendChild(localBridgeCard);

  const installUrlSection = document.createElement("section");
  installUrlSection.style.cssText = "display: flex; flex-direction: column; gap: 8px;";
  installUrlSection.appendChild(createSectionTitle("Install from URL"));

  const installUrlRow = document.createElement("div");
  installUrlRow.style.cssText = "display: grid; grid-template-columns: 160px 1fr auto; gap: 8px; align-items: center;";

  const installUrlName = createInput("Name");
  const installUrlInput = createInput("https://example.com/pi-extension.js");
  const installUrlButton = createButton("Install");

  installUrlRow.append(installUrlName, installUrlInput, installUrlButton);

  const installUrlHint = document.createElement("p");
  installUrlHint.textContent = "Remote URLs require /experimental on remote-extension-urls.";
  installUrlHint.style.cssText = "margin: 0; font-size: 11px; color: var(--muted-foreground);";

  installUrlSection.append(installUrlRow, installUrlHint);

  const installCodeSection = document.createElement("section");
  installCodeSection.style.cssText = "display: flex; flex-direction: column; gap: 8px;";
  installCodeSection.appendChild(createSectionTitle("Install from pasted code"));

  const installCodeName = createInput("Name");
  const installCodeText = document.createElement("textarea");
  installCodeText.placeholder = "export function activate(api) { ... }";
  installCodeText.style.cssText =
    "width: 100%; min-height: 120px; resize: vertical; border-radius: 8px; "
    + "border: 1px solid oklch(0 0 0 / 0.12); padding: 8px; font-size: 12px; "
    + "font-family: var(--font-mono); background: white;";

  const installCodeActions = document.createElement("div");
  installCodeActions.style.cssText = "display: flex; justify-content: flex-end;";
  const installCodeButton = createButton("Install code");
  installCodeActions.appendChild(installCodeButton);

  installCodeSection.append(installCodeName, installCodeText, installCodeActions);

  const templateSection = document.createElement("section");
  templateSection.style.cssText = "display: flex; flex-direction: column; gap: 8px;";
  templateSection.appendChild(createSectionTitle("LLM prompt template"));

  const templateCode = createReadOnlyCodeBlock(EXTENSION_PROMPT_TEMPLATE);
  const templateActions = document.createElement("div");
  templateActions.style.cssText = "display: flex; justify-content: flex-end;";
  const copyTemplateButton = createButton("Copy template");
  templateActions.appendChild(copyTemplateButton);

  templateSection.append(templateCode, templateActions);

  body.append(installedSection, localBridgeSection, installUrlSection, installCodeSection, templateSection);

  card.append(header, body);
  overlay.appendChild(card);

  const setBusy = (busy: boolean) => {
    installUrlButton.disabled = busy;
    installCodeButton.disabled = busy;
    copyTemplateButton.disabled = busy;
    localBridgeUrlInput.disabled = busy;
    localBridgeEnableButton.disabled = busy;
    localBridgeSaveUrlButton.disabled = busy;
    localBridgeDisableButton.disabled = busy;
  };

  const renderLocalBridgeState = async (): Promise<void> => {
    const enabled = isExperimentalFeatureEnabled("python_bridge");
    const configuredUrl = await readSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY);

    if (configuredUrl && localBridgeUrlInput.value.trim().length === 0) {
      localBridgeUrlInput.value = configuredUrl;
    }

    localBridgeStatusText.textContent = configuredUrl
      ? `Bridge URL: ${configuredUrl}`
      : "Bridge URL not set";

    localBridgeStatusBadgeSlot.replaceChildren(
      createBadge(enabled ? "enabled" : "disabled", enabled ? "ok" : "muted"),
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
      void renderLocalBridgeState();
    }
  };

  const createInstalledRow = (status: ExtensionRuntimeStatus): HTMLDivElement => {
    const row = document.createElement("div");
    row.style.cssText =
      "display: flex; flex-direction: column; gap: 8px; border: 1px solid oklch(0 0 0 / 0.08); "
      + "background: oklch(0 0 0 / 0.015); border-radius: 10px; padding: 9px;";

    const top = document.createElement("div");
    top.style.cssText = "display: flex; justify-content: space-between; gap: 10px; align-items: flex-start;";

    const info = document.createElement("div");
    info.style.cssText = "display: flex; flex-direction: column; gap: 4px; min-width: 0;";

    const name = document.createElement("strong");
    name.textContent = status.name;
    name.style.cssText = "font-size: 13px;";

    const source = document.createElement("code");
    source.textContent = status.sourceLabel;
    source.style.cssText =
      "font-size: 10px; color: var(--muted-foreground); font-family: var(--font-mono); "
      + "white-space: nowrap; overflow: hidden; text-overflow: ellipsis; display: block; max-width: 430px;";

    info.append(name, source);

    const badges = document.createElement("div");
    badges.style.cssText = "display: flex; gap: 6px; flex-wrap: wrap; justify-content: flex-end;";

    if (!status.enabled) {
      badges.appendChild(createBadge("disabled", "muted"));
    } else if (status.lastError) {
      badges.appendChild(createBadge("error", "warn"));
    } else if (status.loaded) {
      badges.appendChild(createBadge("loaded", "ok"));
    } else {
      badges.appendChild(createBadge("pending", "muted"));
    }

    if (status.toolNames.length > 0) {
      badges.appendChild(createBadge(`${status.toolNames.length} tool${status.toolNames.length === 1 ? "" : "s"}`, "muted"));
    }

    if (status.commandNames.length > 0) {
      badges.appendChild(createBadge(`${status.commandNames.length} command${status.commandNames.length === 1 ? "" : "s"}`, "muted"));
    }

    top.append(info, badges);

    const details = document.createElement("div");
    details.style.cssText = "display: flex; flex-direction: column; gap: 2px;";

    if (status.commandNames.length > 0) {
      const commands = document.createElement("div");
      commands.textContent = `Commands: ${status.commandNames.map((name) => `/${name}`).join(", ")}`;
      commands.style.cssText = "font-size: 11px; color: var(--muted-foreground);";
      details.appendChild(commands);
    }

    if (status.toolNames.length > 0) {
      const tools = document.createElement("div");
      tools.textContent = `Tools: ${status.toolNames.join(", ")}`;
      tools.style.cssText = "font-size: 11px; color: var(--muted-foreground);";
      details.appendChild(tools);
    }

    if (status.lastError) {
      const errorLine = document.createElement("div");
      errorLine.textContent = `Last error: ${status.lastError}`;
      errorLine.style.cssText = "font-size: 11px; color: oklch(0.52 0.13 35);";
      details.appendChild(errorLine);
    }

    const actions = document.createElement("div");
    actions.style.cssText = "display: flex; gap: 6px; justify-content: flex-end; flex-wrap: wrap;";

    const toggleButton = createButton(status.enabled ? "Disable" : "Enable");
    toggleButton.addEventListener("click", () => {
      void runAction(async () => {
        await manager.setExtensionEnabled(status.id, !status.enabled);
      });
    });

    const reloadButton = createButton("Reload");
    reloadButton.disabled = !status.enabled;
    reloadButton.addEventListener("click", () => {
      void runAction(async () => {
        await manager.reloadExtension(status.id);
      });
    });

    const uninstallButton = createButton("Uninstall");
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
      empty.style.cssText = "font-size: 12px; color: var(--muted-foreground);";
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
      setExperimentalFeatureEnabled("python_bridge", true);

      localBridgeUrlInput.value = normalizedUrl;
      showToast(`Python bridge enabled at ${normalizedUrl}`);
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
    void runAction(() => {
      setExperimentalFeatureEnabled("python_bridge", false);
      showToast("Python bridge disabled.");
    });
  });

  installUrlButton.addEventListener("click", () => {
    void runAction(async () => {
      const name = installUrlName.value.trim();
      const url = installUrlInput.value.trim();
      if (name.length === 0 || url.length === 0) {
        throw new Error("Provide both name and URL");
      }

      await manager.installFromUrl(name, url);
      installUrlInput.value = "";
      showToast(`Installed extension: ${name}`);
    });
  });

  installCodeButton.addEventListener("click", () => {
    void runAction(async () => {
      const name = installCodeName.value.trim();
      const code = installCodeText.value;
      if (name.length === 0) {
        throw new Error("Provide an extension name");
      }

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
    void renderLocalBridgeState();
  });

  let closed = false;
  closeOverlay = () => {
    if (closed) {
      return;
    }

    closed = true;
    overlayClosers.delete(overlay);
    unsubscribe();
    overlay.remove();
  };

  overlayClosers.set(overlay, () => {
    closeOverlay();
  });

  overlay.addEventListener("click", (event) => {
    if (event.target !== overlay) {
      return;
    }

    closeOverlay();
  });

  document.body.appendChild(overlay);
  renderInstalledList();
  void renderLocalBridgeState();
}
