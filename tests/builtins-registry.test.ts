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

void test("builtins registry wires /experimental, /extensions, and /skills command registration", async () => {
  const source = await readFile(new URL("../src/commands/builtins/index.ts", import.meta.url), "utf8");

  assert.match(source, /createExperimentalCommands/);
  assert.match(source, /\.\.\.createExperimentalCommands\(\)/);

  assert.match(source, /createSkillsCommands/);
  assert.match(source, /\.\.\.createSkillsCommands\(context\)/);

  assert.match(source, /createExtensionsCommands/);
  assert.match(source, /\.\.\.createExtensionsCommands\(context\)/);

  const extensionApiSource = await readFile(new URL("../src/commands/extension-api.ts", import.meta.url), "utf8");
  assert.match(extensionApiSource, /\.glob\("\.\.\/extensions\/\*\.\{ts,js\}"\)/);
  assert.match(extensionApiSource, /Local extension module/);
  assert.match(extensionApiSource, /isCapabilityEnabled/);
  assert.match(extensionApiSource, /commands\.register/);
  assert.match(extensionApiSource, /tools\.register/);
  assert.match(extensionApiSource, /agent\.events\.read/);
  assert.match(
    extensionApiSource,
    /get agent\(\)\s*\{[\s\S]*assertCapability\("agent\.read"\);[\s\S]*assertCapability\("agent\.events\.read"\);/,
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
  assert.match(extensionsOverlaySource, /Sandbox runtime \(experimental\)/);
  assert.match(extensionsOverlaySource, /setExperimentalFeatureEnabled\("extension_sandbox_runtime", true\)/);
  assert.match(extensionsOverlaySource, /setExperimentalFeatureEnabled\("extension_sandbox_runtime", false\)/);
  assert.match(extensionsOverlaySource, /Sandbox runtime is off: this untrusted extension runs in host runtime/);
  assert.match(extensionsOverlaySource, /higher-risk permissions/);
  assert.match(extensionsOverlaySource, /Updated permissions for/);
  assert.match(extensionsOverlaySource, /reload failed \(see Last error\)/);

  const extensionsDocsSource = await readFile(new URL("../docs/extensions.md", import.meta.url), "utf8");
  assert.match(extensionsDocsSource, /## Permission review\/revoke/);
  assert.match(extensionsDocsSource, /Install from URL\/code asks for confirmation/);
  assert.match(extensionsDocsSource, /extensions\.registry\.v2/);

  const experimentalFlagsSource = await readFile(new URL("../src/experiments/flags.ts", import.meta.url), "utf8");
  assert.match(experimentalFlagsSource, /extension_permission_gates/);
  assert.match(experimentalFlagsSource, /extension-permissions/);
  assert.match(experimentalFlagsSource, /extension_sandbox_runtime/);
  assert.match(experimentalFlagsSource, /extension-sandbox/);
});

void test("taskpane init keeps getSkillToolNames imported when used", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");
  if (!/getSkillToolNames\(\)/.test(initSource)) {
    return;
  }

  assert.match(
    initSource,
    /import\s*\{[\s\S]*getSkillToolNames[\s\S]*\}\s*from "\.\.\/skills\/catalog\.js";/,
  );
});

void test("taskpane init wires Files workspace opener when sidebar callback is present", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");
  if (!/sidebar\.onOpenFiles\s*=/.test(initSource)) {
    return;
  }

  assert.match(initSource, /showFilesWorkspaceDialog/);
  assert.match(
    initSource,
    /sidebar\.onOpenFiles\s*=\s*\(\)\s*=>\s*\{\s*void showFilesWorkspaceDialog\(\);\s*\};/,
  );
});

void test("session builtins include recovery history command", async () => {
  const sessionSource = await readFile(new URL("../src/commands/builtins/session.ts", import.meta.url), "utf8");

  assert.match(sessionSource, /name:\s*"history"/);
  assert.match(sessionSource, /openRecoveryDialog/);
  assert.match(sessionSource, /name:\s*"revert"/);
});

void test("taskpane init wires recovery overlay opener", async () => {
  const initSource = await readFile(new URL("../src/taskpane/init.ts", import.meta.url), "utf8");

  assert.match(initSource, /showRecoveryDialog/);
  assert.match(initSource, /const openRecoveryDialog = async \(\): Promise<void> =>/);
  assert.match(initSource, /sidebar\.onOpenRecovery\s*=\s*\(\)\s*=>\s*\{\s*void openRecoveryDialog\(\);\s*\};/);
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
