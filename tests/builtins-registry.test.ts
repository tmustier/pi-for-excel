import assert from "node:assert/strict";
import { test } from "node:test";
import { readFile } from "node:fs/promises";

import {
  BUILTIN_SNAKE_EXTENSION_ID,
  EXTENSIONS_REGISTRY_STORAGE_KEY,
  LEGACY_EXTENSIONS_REGISTRY_STORAGE_KEY,
  loadStoredExtensions,
  saveStoredExtensions,
} from "../src/extensions/store.ts";
import {
  isExtensionCapabilityAllowed,
  setExtensionCapabilityAllowed,
  type StoredExtensionPermissions,
} from "../src/extensions/permissions.ts";
import { registerBuiltins, type BuiltinsContext } from "../src/commands/builtins/index.ts";
import { commandRegistry } from "../src/commands/types.ts";

class MemorySettingsStore {
  private readonly values = new Map<string, DynamicValue>();

  get(key: string): Promise<DynamicValue> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: DynamicValue): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  readRaw(key: string): DynamicValue {
    return this.values.has(key) ? this.values.get(key) ?? null : null;
  }

  writeRaw(key: string, value: DynamicValue): void {
    this.values.set(key, value);
  }
}

void test("registerBuiltins registers and routes workspace commands", async () => {
  const previousCommands = commandRegistry.list();
  const openedTabs: Array<string | undefined> = [];
  let filesOpenCount = 0;
  const context: BuiltinsContext = {
    getActiveAgent: () => null,
    openModelSelector: () => {},
    openInstructionsEditor: () => Promise.resolve(),
    getExecutionMode: () => Promise.resolve("safe"),
    setExecutionMode: () => Promise.resolve(),
    renameActiveSession: () => Promise.resolve(),
    createRuntime: () => Promise.resolve(),
    openResumeDialog: () => Promise.resolve(),
    openRecoveryDialog: () => Promise.resolve(),
    reopenLastClosed: () => Promise.resolve(),
    revertLatestCheckpoint: () => Promise.resolve(),
    createManualFullBackup: () => Promise.resolve({ id: "backup", createdAt: 0, sizeBytes: 0 }),
    listManualFullBackups: () => Promise.resolve([]),
    restoreManualFullBackup: () => Promise.resolve(null),
    clearManualFullBackups: () => Promise.resolve(0),
    openExtensionsHub: (tab) => {
      openedTabs.push(tab);
    },
    openFilesWorkspace: () => {
      filesOpenCount += 1;
    },
  };

  try {
    registerBuiltins(context);

    for (const name of ["settings", "login", "experimental", "extensions", "plugins", "tools", "skills", "files"]) {
      assert.equal(commandRegistry.get(name)?.source, "builtin", `expected /${name} to be registered`);
    }
    assert.equal(commandRegistry.get("addons"), undefined);
    assert.equal(commandRegistry.get("integrations"), undefined);

    await commandRegistry.get("extensions")?.execute("");
    await commandRegistry.get("plugins")?.execute("");
    await commandRegistry.get("tools")?.execute("");
    await commandRegistry.get("skills")?.execute("");
    await commandRegistry.get("files")?.execute("");

    assert.deepEqual(openedTabs, [undefined, "plugins", "connections", "skills"]);
    assert.equal(filesOpenCount, 1);
  } finally {
    for (const command of commandRegistry.list()) {
      commandRegistry.unregister(command.name);
    }
    for (const command of previousCommands) {
      commandRegistry.register(command);
    }
  }
});

void test("taskpane init waits for local services probe and refreshes capabilities", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /let localServicesReady: Promise<void> = Promise\.resolve\(\);/);
  assert.match(initSource, /await localServicesReady;/);
  assert.match(
    initSource,
    /localServicesReady\s*=\s*probeLocalServices\(\)\.then\(\s*\(result\) => \{[\s\S]*localServicesSnapshot\s*=\s*result;[\s\S]*void refreshCapabilitiesForAllRuntimes\(\);[\s\S]*\},/,
  );
});

void test("extensions hub connections tab includes MCP test flow", async () => {
  const source = await readFile(
    new URL("../src/commands/builtins/extensions-hub-connections.ts", import.meta.url),
    "utf8",
  );

  assert.match(source, /label: t\("extensions-hub-connections\.mcpSection"\)/);
  assert.match(source, /extensions-hub-connections\.addServer/);
  assert.match(source, /createConfigRow\(t\("extensions-hub-connections\.availability"\)/);
  assert.match(source, /scopeSummary\.textContent = t\("extensions-hub-connections\.scope-controls"\)/);
  assert.match(source, /probeMcpServer/);
});

void test("taskpane init wires Files workspace opener", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /showFilesWorkspaceDialog/);
  assert.match(
    initSource,
    /sidebar\.onOpenFilesWorkspace\s*=\s*\(\)\s*=>\s*\{\s*void showFilesWorkspaceDialog\(\);\s*\};/,
  );
});

void test("taskpane init wires extensions menu opener", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /const openExtensionsHub = \(tab\?: ExtensionsHubTab\): void =>/);
  assert.match(initSource, /openSettings\(tab \?\? "connections"\)/);
  assert.match(initSource, /extensionManager/);
  assert.match(initSource, /configureSettingsPages/);
  assert.match(initSource, /registerBuiltins\([\s\S]*openExtensionsHub/);
  assert.match(initSource, /sidebar\.onOpenExtensions\s*=\s*\(\)\s*=>\s*\{\s*openExtensionsHub\(\);\s*\};/);
});

void test("taskpane init wires gear settings to unified settings overlay", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /openSettings/);
  assert.match(
    initSource,
    /sidebar\.onOpenSettings\s*=\s*\(\)\s*=>\s*\{\s*void openSettings\(\);\s*\};/,
  );
  assert.match(initSource, /configureSettingsPages\(\{[\s\S]*getExecutionMode/);
  assert.match(initSource, /configureSettingsPages\(\{[\s\S]*setExecutionMode/);
  assert.match(initSource, /configureSettingsPages\(\{[\s\S]*getModelSwitchBehavior/);
  assert.match(initSource, /configureSettingsPages\(\{[\s\S]*setModelSwitchBehavior/);
});

void test("taskpane init mounts proxy banner and reacts to proxy state changes", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /createProxyBanner/);
  assert.match(initSource, /document\.addEventListener\("pi:proxy-state-changed"/);
  assert.match(initSource, /proxyBanner\.update\(getProxyState\(\)\)/);
});

void test("status bar keeps model, thinking, context, and mode without rules\/proxy badges", async () => {
  const statusBarSource = await readFile(new URL("../src/taskpane/status-bar.ts", import.meta.url), "utf8");

  assert.match(statusBarSource, /pi-status-model/);
  assert.match(statusBarSource, /pi-status-thinking/);
  assert.match(statusBarSource, /pi-status-ctx__pct/);
  assert.match(statusBarSource, /pi-status-mode/);
  assert.doesNotMatch(statusBarSource, /pi-status-rules/);
  assert.doesNotMatch(statusBarSource, /pi-status-proxy/);
});

void test("sidebar utilities menu includes extensions label", async () => {
  const sidebarSource = await readFile(new URL("../src/ui/pi-sidebar.ts", import.meta.url), "utf8");

  assert.match(sidebarSource, /aria-label=\$\{t\("sidebar\.utilities\.aria"\)\}/);
  assert.match(sidebarSource, /sidebar\.menu\.extensions/);
  assert.match(sidebarSource, /sidebar\.menu\.files/);
  assert.doesNotMatch(sidebarSource, /Extensions…/);
  assert.doesNotMatch(sidebarSource, /Files…/);
  assert.doesNotMatch(sidebarSource, /Add-ons…/);
});

void test("disclosure bar reuses shared toggle rows", async () => {
  const disclosureSource = await readFile(new URL("../src/ui/disclosure-bar.ts", import.meta.url), "utf8");

  assert.match(disclosureSource, /createToggleRow/);
  assert.doesNotMatch(disclosureSource, /pi-toggle__track/);
});

void test("extensions pages expose connections, plugins, and skills in the settings shell", async () => {
  const pagesSource = await readFile(
    new URL("../src/commands/builtins/settings-pages/extensions-pages.ts", import.meta.url),
    "utf8",
  );
  const connectionsSource = await readFile(
    new URL("../src/commands/builtins/extensions-hub-connections.ts", import.meta.url),
    "utf8",
  );
  const pluginsSource = await readFile(
    new URL("../src/commands/builtins/extensions-hub-plugins.ts", import.meta.url),
    "utf8",
  );
  const skillsSource = await readFile(
    new URL("../src/commands/builtins/extensions-hub-skills.ts", import.meta.url),
    "utf8",
  );

  assert.match(pagesSource, /export function createConnectionsPage/);
  assert.match(pagesSource, /export function createPluginsPage/);
  assert.match(pagesSource, /export function createSkillsPage/);
  assert.match(pagesSource, /createDeferredConnectionsRefreshController/);
  assert.match(connectionsSource, /Web search/);
  assert.match(pluginsSource, /Installed/);
  assert.match(skillsSource, /Bundled skills/);
});

void test("context pill headers expose expanded state and controlled body", async () => {
  const sidebarSource = await readFile(new URL("../src/ui/pi-sidebar.ts", import.meta.url), "utf8");

  assert.match(sidebarSource, /private readonly _contextPillBodyId = "pi-context-pill-body";/);
  assert.match(sidebarSource, /class="pi-context-pill__header"[\s\S]*aria-controls=\$\{this\._contextPillBodyId\}/);
  assert.match(sidebarSource, /class="pi-context-pill__header"[\s\S]*aria-expanded=\$\{expanded \? "true" : "false"\}/);
  assert.match(sidebarSource, /class="pi-context-pill__body" id=\$\{this\._contextPillBodyId\}/);
});

void test("input paperclip opens Files workspace through sidebar callback", async () => {
  const inputSource = await readFile(new URL("../src/ui/pi-input.ts", import.meta.url), "utf8");
  const sidebarSource = await readFile(new URL("../src/ui/pi-sidebar.ts", import.meta.url), "utf8");

  assert.match(inputSource, /pi-open-files/);
  assert.match(sidebarSource, /onOpenFilesWorkspace/);
  assert.match(sidebarSource, /@pi-open-files=\$\{this\._onOpenFilesWorkspace\}/);
});

void test("session builtins include recovery and manual-backup commands", async () => {
  const sessionSource = await readFile(new URL("../src/commands/builtins/session.ts", import.meta.url), "utf8");

  assert.match(sessionSource, /name:\s*"history"/);
  assert.match(sessionSource, /openRecoveryDialog/);
  assert.match(sessionSource, /name:\s*"revert"/);
  assert.match(sessionSource, /name:\s*"backup"/);
  assert.match(sessionSource, /createManualFullBackup/);
  assert.match(sessionSource, /restoreManualFullBackup/);
});

void test("resume overlay surfaces recently closed tabs and taskpane wires reopen callback", async () => {
  const resumeSource = await readFile(new URL("../src/commands/builtins/resume-overlay.ts", import.meta.url), "utf8");
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(resumeSource, /resume\.recentlyClosed/);
  assert.match(resumeSource, /getRecentlyClosedItems\?: \(\) => readonly ResumeRecentlyClosedItem\[]/);
  assert.match(resumeSource, /onReopenRecentlyClosed\?: \(item: ResumeRecentlyClosedItem\) => Promise<boolean>/);
  assert.match(resumeSource, /resume\.recentlyClosedMeta/);

  assert.match(initSource, /getRecentlyClosedItems:\s*\(\)\s*=>\s*recentlyClosed\.snapshot\(\)/);
  assert.match(initSource, /onReopenRecentlyClosed:\s*async \(item\) =>/);
  assert.match(initSource, /const reopenRecentlyClosedById = async \(recentlyClosedId: string\): Promise<boolean> =>/);
  assert.match(initSource, /recentlyClosed\.removeById\(recentlyClosedId\)/);
  assert.match(initSource, /if \(reopenResult === "failed"\) \{\s*recentlyClosed\.push\(item\);\s*\}/);
});

void test("experimental overlay remains a settings section alias", async () => {
  const experimentalSource = await readFile(new URL("../src/commands/builtins/experimental-overlay.ts", import.meta.url), "utf8");

  assert.match(experimentalSource, /openSettings\("experimental"\)/);
  assert.match(experimentalSource, /buildExperimentalFeatureContent/);
  assert.match(experimentalSource, /createToggleRow/);
});

void test("extensions and alias commands deep-link to hub tabs", async () => {
  const addonsSource = await readFile(new URL("../src/commands/builtins/addons.ts", import.meta.url), "utf8");
  const toolsSource = await readFile(new URL("../src/commands/builtins/tools.ts", import.meta.url), "utf8");
  const extensionsSource = await readFile(new URL("../src/commands/builtins/extensions.ts", import.meta.url), "utf8");
  const skillsSource = await readFile(new URL("../src/commands/builtins/skills.ts", import.meta.url), "utf8");

  assert.match(addonsSource, /name:\s*"extensions"/);
  assert.doesNotMatch(addonsSource, /name:\s*"addons"/);
  assert.match(addonsSource, /openExtensionsHub\(\)/);

  assert.match(toolsSource, /openExtensionsHub\("connections"\)/);
  assert.match(extensionsSource, /openExtensionsHub\("plugins"\)/);
  assert.match(skillsSource, /openExtensionsHub\("skills"\)/);
});

void test("settings shell guards navigation and pages adopt shared controls", async () => {
  const shellSource = await readFile(new URL("../src/ui/settings-shell.ts", import.meta.url), "utf8");
  const rootSource = await readFile(
    new URL("../src/commands/builtins/settings-pages/root-page.ts", import.meta.url),
    "utf8",
  );
  const providersSource = await readFile(
    new URL("../src/commands/builtins/settings-pages/providers-page.ts", import.meta.url),
    "utf8",
  );
  const proxySource = await readFile(
    new URL("../src/commands/builtins/settings-pages/proxy-page.ts", import.meta.url),
    "utf8",
  );

  assert.match(shellSource, /beforeLeave/);
  assert.match(shellSource, /registerOverlayCloser/);
  assert.match(shellSource, /buildStackFor/);
  assert.match(rootSource, /settings\.section\.execution\.auto_mode/);
  assert.match(rootSource, /settings\.section\.advanced\.fork_label/);
  assert.match(providersSource, /settings\.warning\.provider_state/);
  assert.match(proxySource, /createToggleRow/);
  assert.match(proxySource, /createConfigRow/);
  assert.match(proxySource, /createCallout/);
});

void test("slash-command busy policy is centralized and shared across entry points", async () => {
  const keyboardActionsSource = await readFile(
    new URL("../src/taskpane/keyboard-shortcuts/editor-actions.ts", import.meta.url),
    "utf8",
  );
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");
  const slashExecutionSource = await readFile(
    new URL("../src/commands/slash-command-execution.ts", import.meta.url),
    "utf8",
  );
  const busyPolicySource = await readFile(new URL("../src/commands/busy-command-policy.ts", import.meta.url), "utf8");

  assert.match(keyboardActionsSource, /executeSlashCommand/);
  assert.match(initSource, /executeSlashCommand/);

  assert.match(slashExecutionSource, /isBusyAllowedCommand/);
  assert.match(slashExecutionSource, /commandRegistry\.get\(options\.name\)/);

  assert.match(busyPolicySource, /"yolo"/);
  assert.match(busyPolicySource, /"rules"/);
  assert.match(busyPolicySource, /"files"/);
  assert.match(busyPolicySource, /TOOLS_COMMAND_NAME/);
  assert.match(busyPolicySource, /command\.source === "extension"/);
  assert.match(busyPolicySource, /command\.busyAllowed \?\? true/);
  assert.doesNotMatch(busyPolicySource, /INTEGRATIONS_COMMAND_NAME/);
  assert.doesNotMatch(busyPolicySource, /"addons"/);
});

void test("escape guard scopes widget claims to streaming abort paths", async () => {
  const escapeGuardSource = await readFile(new URL("../src/utils/escape-guard.ts", import.meta.url), "utf8");
  const keyboardShortcutsSource = await readFile(new URL("../src/taskpane/keyboard-shortcuts.ts", import.meta.url), "utf8");
  const inputSource = await readFile(new URL("../src/ui/pi-input.ts", import.meta.url), "utf8");
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(escapeGuardSource, /export function doesExtensionWidgetClaimEscape/);
  assert.match(escapeGuardSource, /export function doesUiClaimStreamingEscape/);
  assert.match(escapeGuardSource, /#pi-widget-slot:not\(:empty\)/);
  assert.match(escapeGuardSource, /#pi-widget-slot-below:not\(:empty\)/);

  assert.match(keyboardShortcutsSource, /doesUiClaimStreamingEscape/);
  assert.match(inputSource, /doesUiClaimStreamingEscape/);
  assert.match(initSource, /doesOverlayClaimEscape\(document\.activeElement\)/);
});

void test("taskpane init wires recovery overlay opener", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /openSettings\("backups"\)/);
  assert.match(initSource, /const openRecoveryDialog = async \(\): Promise<void> =>/);
  assert.match(initSource, /sidebar\.onOpenRecovery\s*=\s*\(\)\s*=>\s*\{\s*void openRecoveryDialog\(\);\s*\};/);
  assert.match(initSource, /onCreateManualFullBackup:\s*async \(\)\s*=>\s*\{\s*return createManualFullBackup\(\);\s*\}/);
});

void test("backups page includes manual full-backup action", async () => {
  const overlaySource = await readFile(
    new URL("../src/commands/builtins/settings-pages/backups-page.ts", import.meta.url),
    "utf8",
  );

  assert.match(overlaySource, /onCreateManualFullBackup\?: \(\) => Promise<ManualFullBackupSummary>/);
  assert.match(overlaySource, /createButton\(t\("recovery\.downloadBackup"\)/);
  assert.match(overlaySource, /recovery\.toast\.backupDownloaded/);
  assert.match(overlaySource, /retentionInput\.max = String\(MAX_RECOVERY_ENTRIES\)/);
});

void test("permission helper updates one capability without mutating others", () => {
  const permissions: StoredExtensionPermissions = {
    commandsRegister: true,
    toolsRegister: false,
    agentRead: false,
    agentEventsRead: false,
    uiOverlay: true,
    uiWidget: true,
    uiToast: true,
    llmComplete: false,
    httpFetch: false,
    storageReadWrite: true,
    connectionsReadWrite: false,
    connectionsSecretsRead: false,
    clipboardWrite: true,
    agentContextWrite: false,
    agentSteer: false,
    agentFollowUp: false,
    skillsRead: true,
    skillsWrite: false,
    downloadFile: true,
  };

  const updated = setExtensionCapabilityAllowed(permissions, "tools.register", true);

  assert.equal(isExtensionCapabilityAllowed(updated, "tools.register"), true);
  assert.equal(isExtensionCapabilityAllowed(updated, "commands.register"), true);
  assert.equal(isExtensionCapabilityAllowed(updated, "agent.read"), false);

  // original object remains unchanged
  assert.equal(isExtensionCapabilityAllowed(permissions, "tools.register"), false);
});

void test("extension registry seeds default snake extension when storage is empty", async () => {
  const settings = new MemorySettingsStore();

  const entries = await loadStoredExtensions(settings);
  assert.equal(entries.length, 1);
  assert.equal(entries[0].id, BUILTIN_SNAKE_EXTENSION_ID);
  assert.equal(entries[0].trust, "builtin");
  assert.equal(entries[0].permissions.commandsRegister, true);
  assert.equal(entries[0].permissions.toolsRegister, true);
  assert.equal(entries[0].permissions.agentRead, true);

  const raw = settings.readRaw(EXTENSIONS_REGISTRY_STORAGE_KEY);
  assert.ok(raw);
});

void test("extension registry preserves explicit empty saved entries", async () => {
  const settings = new MemorySettingsStore();

  await saveStoredExtensions(settings, []);
  const entries = await loadStoredExtensions(settings);
  assert.deepEqual(entries, []);
});

void test("extension registry migrates legacy v1 entries to v2 permissions", async () => {
  const settings = new MemorySettingsStore();
  const timestamp = "2026-02-12T00:00:00.000Z";

  settings.writeRaw(LEGACY_EXTENSIONS_REGISTRY_STORAGE_KEY, {
    version: 1,
    items: [
      {
        id: "ext.legacy.inline",
        name: "Legacy Inline",
        enabled: true,
        source: {
          kind: "inline",
          code: "export function activate(api) { api.toast('hi'); }",
        },
        createdAt: timestamp,
        updatedAt: timestamp,
      },
    ],
  });

  const entries = await loadStoredExtensions(settings);
  assert.equal(entries.length, 1);
  assert.equal(entries[0].id, "ext.legacy.inline");
  assert.equal(entries[0].trust, "inline-code");
  assert.equal(entries[0].permissions.commandsRegister, true);
  assert.equal(entries[0].permissions.toolsRegister, false);
  assert.equal(entries[0].permissions.agentRead, false);

  const migrated = settings.readRaw(EXTENSIONS_REGISTRY_STORAGE_KEY);
  assert.deepEqual(migrated, { version: 2, items: entries });
});

void test("tool disclosure bundles remain centralized in capabilities metadata", async () => {
  const disclosureSource = await readFile(new URL("../src/context/tool-disclosure.ts", import.meta.url), "utf8");
  assert.match(disclosureSource, /type ToolDisclosureBundleId/);
  assert.doesNotMatch(disclosureSource, /TOOL_DISCLOSURE_BUNDLES\s*=/);

  const capabilitiesSource = await readFile(new URL("../src/tools/capabilities.ts", import.meta.url), "utf8");
  assert.match(capabilitiesSource, /TOOL_DISCLOSURE_BUNDLES/);
  assert.match(capabilitiesSource, /core:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /analysis:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /formatting:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /structure:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /comments:\s*buildCoreDisclosureBundle/);
});
