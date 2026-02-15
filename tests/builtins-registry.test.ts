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

class MemorySettingsStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  readRaw(key: string): unknown {
    return this.values.has(key) ? this.values.get(key) ?? null : null;
  }

  writeRaw(key: string, value: unknown): void {
    this.values.set(key, value);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

void test("builtins registry wires /addons, /experimental, /extensions, /integrations, and /files command registration", async () => {
  const source = await readFile(new URL("../src/commands/builtins/index.ts", import.meta.url), "utf8");

  assert.match(source, /createAddonsCommands/);
  assert.match(source, /\.\.\.createAddonsCommands\(context\)/);

  assert.match(source, /createExperimentalCommands/);
  assert.match(source, /\.\.\.createExperimentalCommands\(\)/);

  assert.match(source, /createIntegrationsCommands/);
  assert.match(source, /\.\.\.createIntegrationsCommands\(context\)/);

  assert.match(source, /createExtensionsCommands/);
  assert.match(source, /\.\.\.createExtensionsCommands\(context\)/);

  assert.match(source, /createFilesCommands/);
  assert.match(source, /\.\.\.createFilesCommands\(context\)/);

  const extensionApiSource = await readFile(new URL("../src/commands/extension-api.ts", import.meta.url), "utf8");
  const extensionModuleImportSource = await readFile(
    new URL("../src/commands/extension-module-import.ts", import.meta.url),
    "utf8",
  );

  assert.match(extensionModuleImportSource, /glob\("\.\.\/extensions\/\*\.\{ts,js\}"\)/);
  assert.match(extensionModuleImportSource, /return import\.meta\.env\.DEV === true/);
  assert.doesNotMatch(extensionModuleImportSource, /typeof \(import\.meta.*\)\.glob !== "function"/);
  assert.match(extensionModuleImportSource, /Local extension module/);

  assert.match(extensionApiSource, /isCapabilityEnabled/);
  assert.match(extensionApiSource, /commands\.register/);
  assert.match(extensionApiSource, /tools\.register/);
  assert.match(extensionApiSource, /agent\.events\.read/);
  assert.match(
    extensionApiSource,
    /get raw\(\)\s*\{[\s\S]*assertCapability\("agent\.read"\);[\s\S]*assertCapability\("agent\.events\.read"\);/,
  );

  const runtimeManagerSource = await readFile(new URL("../src/extensions/runtime-manager.ts", import.meta.url), "utf8");
  assert.match(runtimeManagerSource, /effectiveCapabilities/);
  assert.match(runtimeManagerSource, /permissionsEnforced/);
  assert.match(runtimeManagerSource, /async setExtensionCapability\(/);
  assert.match(runtimeManagerSource, /setExtensionCapabilityAllowed\(/);
  assert.match(runtimeManagerSource, /await this\.reloadExtension\(entry\.id\);/);
  assert.match(runtimeManagerSource, /activateExtensionInSandbox/);
  assert.match(runtimeManagerSource, /extension_sandbox_runtime/);

  const extensionsOverlaySource = await readFile(
    new URL("../src/commands/builtins/extensions-overlay.ts", import.meta.url),
    "utf8",
  );
  assert.match(extensionsOverlaySource, /manager\.setExtensionCapability\(/);
  assert.match(extensionsOverlaySource, /toggle\.type = "checkbox"/);
  assert.match(extensionsOverlaySource, /confirmExtensionInstall\(/);
  assert.match(extensionsOverlaySource, /confirmExtensionEnable\(/);
  assert.match(extensionsOverlaySource, /Sandbox runtime \(default for untrusted sources\)/);
  assert.match(extensionsOverlaySource, /setExperimentalFeatureEnabled\("extension_sandbox_runtime", true\)/);
  assert.match(extensionsOverlaySource, /setExperimentalFeatureEnabled\("extension_sandbox_runtime", false\)/);
  assert.match(extensionsOverlaySource, /Host-runtime fallback is enabled: this untrusted extension runs in host runtime/);
  assert.match(extensionsOverlaySource, /higher-risk permissions/);
  assert.match(extensionsOverlaySource, /Updated permissions for/);
  assert.match(extensionsOverlaySource, /Permissions \(\$\{grantedCapabilities\.size\}\)/);
  assert.match(extensionsOverlaySource, /createOverlayBadge\("all permissions", "muted"\)/);
  assert.match(extensionsOverlaySource, /addExtensionSummary\.textContent = "Add extension"/);
  assert.match(extensionsOverlaySource, /advancedSummary\.textContent = "Advanced"/);
  assert.match(extensionsOverlaySource, /errorSummary\.textContent = "Failed to load"/);
  assert.match(extensionsOverlaySource, /reload failed \(see Failed to load\)/);

  const extensionsDocsSource = await readFile(new URL("../docs/extensions.md", import.meta.url), "utf8");
  assert.match(extensionsDocsSource, /## Permission review\/revoke/);
  assert.match(extensionsDocsSource, /Install from URL\/code asks for confirmation/);
  assert.match(extensionsDocsSource, /extensions\.registry\.v2/);
  assert.match(extensionsDocsSource, /extension-widget-v2/);

  const experimentalFlagsSource = await readFile(new URL("../src/experiments/flags.ts", import.meta.url), "utf8");
  assert.match(experimentalFlagsSource, /extension_permission_gates/);
  assert.match(experimentalFlagsSource, /extension-permissions/);
  assert.match(experimentalFlagsSource, /extension_sandbox_runtime/);
  assert.match(experimentalFlagsSource, /extension-sandbox-rollback/);
  assert.match(experimentalFlagsSource, /extension_widget_v2/);
  assert.match(experimentalFlagsSource, /extension-widget-v2/);
  assert.match(experimentalFlagsSource, /ui_dark_mode/);
  assert.match(experimentalFlagsSource, /dark-mode/);
  assert.doesNotMatch(experimentalFlagsSource, /external_skills_discovery/);
  assert.doesNotMatch(experimentalFlagsSource, /id:\s*"mcp_tools"/);
});

void test("taskpane init keeps getIntegrationToolNames imported when used", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");
  if (!/getIntegrationToolNames\(\)/.test(initSource)) {
    return;
  }

  assert.match(
    initSource,
    /import\s*\{[\s\S]*getIntegrationToolNames[\s\S]*\}\s*from "\.\.\/integrations\/catalog\.js";/,
  );
});

void test("integrations builtins expose /tools without /integrations alias", async () => {
  const source = await readFile(new URL("../src/commands/builtins/integrations.ts", import.meta.url), "utf8");

  assert.match(source, /TOOLS_COMMAND_NAME/);
  assert.doesNotMatch(source, /INTEGRATIONS_COMMAND_NAME/);
});

void test("extensions builtins expose /extensions without /addons alias", async () => {
  const source = await readFile(new URL("../src/commands/builtins/addons.ts", import.meta.url), "utf8");

  assert.match(source, /name:\s*"extensions"/);
  assert.doesNotMatch(source, /name:\s*"addons"/);
  assert.match(source, /openExtensionsHub/);
});

void test("add-ons connections section links to detailed Tools & MCP manager", async () => {
  const source = await readFile(
    new URL("../src/commands/builtins/addons-overlay-connections.ts", import.meta.url),
    "utf8",
  );

  assert.match(source, /Open detailed Tools & MCP manager…/);
  assert.match(source, /openIntegrationsManager/);
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
  assert.match(initSource, /showExtensionsHubDialog\(/);
  assert.match(initSource, /extensionManager/);
  assert.match(initSource, /configureSettingsDialogDependencies/);
  assert.match(initSource, /registerBuiltins\([\s\S]*openExtensionsHub/);
  assert.match(initSource, /sidebar\.onOpenExtensions\s*=\s*\(\)\s*=>\s*\{\s*openExtensionsHub\(\);\s*\};/);
});

void test("taskpane init wires gear settings to unified settings overlay", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /showSettingsDialog/);
  assert.match(
    initSource,
    /sidebar\.onOpenSettings\s*=\s*\(\)\s*=>\s*\{\s*void showSettingsDialog\(\);\s*\};/,
  );
});

void test("sidebar utilities menu includes extensions label", async () => {
  const sidebarSource = await readFile(new URL("../src/ui/pi-sidebar.ts", import.meta.url), "utf8");

  assert.match(sidebarSource, /aria-label="Settings and tools"/);
  assert.match(sidebarSource, /Extensions…/);
  assert.match(sidebarSource, /Files…/);
  assert.doesNotMatch(sidebarSource, /Add-ons…/);
});

void test("extensions hub groups connections, plugins, and skills with tabs", async () => {
  const hubSource = await readFile(new URL("../src/commands/builtins/extensions-hub-overlay.ts", import.meta.url), "utf8");
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

  assert.match(hubSource, /title:\s*"Extensions"/);
  assert.match(hubSource, /Connections, plugins, and skills that extend Pi/);
  assert.match(hubSource, /dataset\.hubTab/);
  assert.match(hubSource, /dataset\.hubPanel/);
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

  assert.match(resumeSource, /Recently closed/);
  assert.match(resumeSource, /getRecentlyClosedItems\?: \(\) => readonly ResumeRecentlyClosedItem\[]/);
  assert.match(resumeSource, /onReopenRecentlyClosed\?: \(item: ResumeRecentlyClosedItem\) => Promise<boolean>/);
  assert.match(resumeSource, /REOPENS in new tab/i);

  assert.match(initSource, /getRecentlyClosedItems:\s*\(\)\s*=>\s*recentlyClosed\.snapshot\(\)/);
  assert.match(initSource, /onReopenRecentlyClosed:\s*async \(item\) =>/);
  assert.match(initSource, /const reopenRecentlyClosedById = async \(recentlyClosedId: string\): Promise<boolean> =>/);
  assert.match(initSource, /recentlyClosed\.removeById\(recentlyClosedId\)/);
  assert.match(initSource, /if \(reopenResult === "failed"\) \{\s*recentlyClosed\.push\(item\);\s*\}/);
});

void test("settings builtins route to unified settings overlay", async () => {
  const settingsSource = await readFile(new URL("../src/commands/builtins/settings.ts", import.meta.url), "utf8");

  assert.match(settingsSource, /name:\s*"settings"/);
  assert.match(settingsSource, /showSettingsDialog/);
  assert.match(settingsSource, /name:\s*"login"/);
  assert.match(settingsSource, /showSettingsDialog\(\{ section: "logins" \}\)/);

  assert.match(settingsSource, /name:\s*"yolo"/);
  assert.match(settingsSource, /Toggle execution mode \(Auto vs Confirm\)/);
  assert.match(settingsSource, /Usage:\s*\/yolo/);
});

void test("provider and experimental overlays are aliases into settings sections", async () => {
  const providerSource = await readFile(new URL("../src/commands/builtins/provider-overlay.ts", import.meta.url), "utf8");
  const experimentalSource = await readFile(new URL("../src/commands/builtins/experimental-overlay.ts", import.meta.url), "utf8");

  assert.match(providerSource, /showSettingsDialog\(\{ section: "providers" \}\)/);
  assert.match(experimentalSource, /showSettingsDialog\(\{ section: "experimental" \}\)/);
  assert.match(experimentalSource, /buildExperimentalFeatureContent/);
});

void test("extensions and alias commands deep-link to hub tabs", async () => {
  const addonsSource = await readFile(new URL("../src/commands/builtins/addons.ts", import.meta.url), "utf8");
  const integrationsSource = await readFile(new URL("../src/commands/builtins/integrations.ts", import.meta.url), "utf8");
  const extensionsSource = await readFile(new URL("../src/commands/builtins/extensions.ts", import.meta.url), "utf8");
  const skillsSource = await readFile(new URL("../src/commands/builtins/skills.ts", import.meta.url), "utf8");

  assert.match(addonsSource, /name:\s*"extensions"/);
  assert.doesNotMatch(addonsSource, /name:\s*"addons"/);
  assert.match(addonsSource, /openExtensionsHub\(\)/);

  assert.match(integrationsSource, /openExtensionsHub\("connections"\)/);
  assert.match(extensionsSource, /openExtensionsHub\("plugins"\)/);
  assert.match(skillsSource, /openExtensionsHub\("skills"\)/);
});

void test("settings overlay serializes open flow and tolerates provider storage lookup failure", async () => {
  const settingsOverlaySource = await readFile(
    new URL("../src/commands/builtins/settings-overlay.ts", import.meta.url),
    "utf8",
  );

  assert.match(settingsOverlaySource, /settingsDialogOpenInFlight/);
  assert.match(settingsOverlaySource, /pendingSectionFocus/);
  assert.match(settingsOverlaySource, /await settingsDialogOpenInFlight/);
  assert.match(settingsOverlaySource, /Saved provider state is temporarily unavailable/);
});

void test("slash-command busy policy is centralized and includes /yolo and /addons", async () => {
  const keyboardActionsSource = await readFile(
    new URL("../src/taskpane/keyboard-shortcuts/editor-actions.ts", import.meta.url),
    "utf8",
  );
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");
  const busyPolicySource = await readFile(new URL("../src/commands/busy-command-policy.ts", import.meta.url), "utf8");

  assert.match(keyboardActionsSource, /isBusyAllowedCommand/);
  assert.match(initSource, /isBusyAllowedCommand/);

  assert.match(busyPolicySource, /"yolo"/);
  assert.match(busyPolicySource, /"rules"/);
  assert.match(busyPolicySource, /"files"/);
  assert.match(busyPolicySource, /TOOLS_COMMAND_NAME/);
  assert.doesNotMatch(busyPolicySource, /INTEGRATIONS_COMMAND_NAME/);
  assert.doesNotMatch(busyPolicySource, /"addons"/);
});

void test("taskpane init wires recovery overlay opener", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /showRecoveryDialog/);
  assert.match(initSource, /const openRecoveryDialog = async \(\): Promise<void> =>/);
  assert.match(initSource, /sidebar\.onOpenRecovery\s*=\s*\(\)\s*=>\s*\{\s*void openRecoveryDialog\(\);\s*\};/);
  assert.match(initSource, /onCreateManualFullBackup:\s*async \(\)\s*=>\s*\{\s*return createManualFullBackup\(\);\s*\}/);
});

void test("recovery overlay includes manual full-backup action", async () => {
  const overlaySource = await readFile(new URL("../src/commands/builtins/recovery-overlay.ts", import.meta.url), "utf8");

  assert.match(overlaySource, /onCreateManualFullBackup\?: \(\) => Promise<ManualFullBackupSummary>/);
  assert.match(overlaySource, /fullBackupButton\.textContent = "Full backup"/);
  assert.match(overlaySource, /Full backup downloaded:/);
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
  assert.ok(isRecord(migrated));
  if (!isRecord(migrated)) {
    return;
  }

  assert.equal(migrated.version, 2);
  assert.ok(Array.isArray(migrated.items));
  assert.equal(migrated.items.length, 1);
});

void test("tool disclosure bundles are centralized in capabilities metadata", async () => {
  const disclosureSource = await readFile(new URL("../src/context/tool-disclosure.ts", import.meta.url), "utf8");
  assert.match(disclosureSource, /chooseToolDisclosureBundle/);
  assert.match(disclosureSource, /filterToolsForDisclosureBundle/);

  const capabilitiesSource = await readFile(new URL("../src/tools/capabilities.ts", import.meta.url), "utf8");
  assert.match(capabilitiesSource, /TOOL_DISCLOSURE_BUNDLES/);
  assert.match(capabilitiesSource, /core:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /analysis:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /formatting:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /structure:\s*buildCoreDisclosureBundle/);
  assert.match(capabilitiesSource, /comments:\s*buildCoreDisclosureBundle/);
});
