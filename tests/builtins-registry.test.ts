import assert from "node:assert/strict";
import { test } from "node:test";

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
import { executeSlashCommand, type SlashCommandExecutionResult } from "../src/commands/slash-command-execution.ts";
import { TOOLS_COMMAND_NAME } from "../src/integrations/naming.ts";
import { commandRegistry, type SlashCommand } from "../src/commands/types.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

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

function restoreCommands(previousCommands: SlashCommand[]): void {
  for (const command of commandRegistry.list()) {
    commandRegistry.unregister(command.name);
  }
  for (const command of previousCommands) {
    commandRegistry.register(command);
  }
}

async function executeRegisteredCommand(name: string, args = ""): Promise<void> {
  const command = commandRegistry.get(name);
  assert.ok(command, `expected /${name} to be registered`);
  await command.execute(args);
}

function getToastText(document: Document): string {
  return document.getElementById("pi-toast")?.children[0]?.children[0]?.textContent ?? "";
}

void test("command-layer contract: workspace commands complete, preserve failures, and run while busy", async () => {
  const previousCommands = commandRegistry.list();
  const openedTabs: Array<string | undefined> = [];
  let filesOpenCount = 0;
  let failFiles = false;
  let beforeExecuteCount = 0;
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
    openExtensionsHub: async (tab) => {
      await Promise.resolve();
      openedTabs.push(tab);
    },
    openFilesWorkspace: async () => {
      await Promise.resolve();
      if (failFiles) {
        throw new Error("files overlay failed");
      }
      filesOpenCount += 1;
    },
  };

  try {
    registerBuiltins(context);

    for (const name of ["extensions", "plugins", "tools", "skills", "files"]) {
      await executeRegisteredCommand(name);
    }

    assert.deepEqual(openedTabs, [undefined, "plugins", "connections", "skills"]);
    assert.equal(filesOpenCount, 1);

    const busyExecution = executeSlashCommand({
      name: "files",
      args: "",
      busy: true,
      beforeExecute: () => {
        beforeExecuteCount += 1;
      },
      onError: () => {},
    });
    assert.equal(busyExecution, "executed");
    await Promise.resolve();
    assert.equal(filesOpenCount, 2);
    assert.equal(beforeExecuteCount, 1);

    failFiles = true;
    await assert.rejects(executeRegisteredCommand("files"), /files overlay failed/);

    for (const removedAlias of ["addons", "integrations"]) {
      assert.equal(
        executeSlashCommand({ name: removedAlias, args: "", busy: false, onError: () => {} }),
        "not-found",
      );
    }
  } finally {
    restoreCommands(previousCommands);
  }
});

void test("slash-command dispatcher owns synchronous and asynchronous command failures", async () => {
  const previousCommands = commandRegistry.list();
  let executionCount = 0;
  let failureCount = 0;

  try {
    commandRegistry.register({
      name: "test-sync-failure",
      description: "Test synchronous failure",
      source: "builtin",
      execute: () => {
        executionCount += 1;
        throw new Error("sensitive sync failure");
      },
    });
    commandRegistry.register({
      name: "test-async-failure",
      description: "Test asynchronous failure",
      source: "builtin",
      execute: async () => {
        executionCount += 1;
        await Promise.resolve();
        throw new Error("sensitive async failure");
      },
    });
    commandRegistry.register({
      name: "test-async-success",
      description: "Test asynchronous success",
      source: "builtin",
      execute: async () => {
        executionCount += 1;
        await Promise.resolve();
      },
    });

    assert.equal(
      executeSlashCommand({
        name: "test-sync-failure",
        args: "",
        busy: false,
        onError: () => {
          failureCount += 1;
        },
      }),
      "executed",
    );
    assert.equal(failureCount, 1);

    assert.equal(
      executeSlashCommand({
        name: "test-async-failure",
        args: "",
        busy: false,
        onError: () => {
          failureCount += 1;
        },
      }),
      "executed",
    );
    await Promise.resolve();
    await Promise.resolve();
    assert.equal(failureCount, 2);

    assert.equal(
      executeSlashCommand({
        name: "test-async-success",
        args: "",
        busy: false,
        onError: () => {
          failureCount += 1;
        },
      }),
      "executed",
    );
    await Promise.resolve();
    await Promise.resolve();
    assert.equal(failureCount, 2);

    assert.equal(
      executeSlashCommand({
        name: "test-sync-failure",
        args: "",
        busy: true,
        onError: () => {
          failureCount += 1;
        },
      }),
      "busy-blocked",
    );
    assert.equal(executionCount, 3);
    assert.equal(failureCount, 2);
  } finally {
    restoreCommands(previousCommands);
  }
});

void test("command-layer contract: recovery commands expose completion, errors, busy policy, and queueing", async () => {
  const previousCommands = commandRegistry.list();
  const fakeDom = installFakeDom();
  let recoveryOpenCount = 0;
  let revertCount = 0;
  let failRevert = false;
  let createBackupCount = 0;
  let failBackup = false;
  const restoredBackupIds: Array<string | undefined> = [];
  const queuedCommands: Array<{ name: string; args: string }> = [];
  const context: BuiltinsContext = {
    getActiveAgent: () => null,
    openModelSelector: () => {},
    openInstructionsEditor: () => Promise.resolve(),
    getExecutionMode: () => Promise.resolve("safe"),
    setExecutionMode: () => Promise.resolve(),
    renameActiveSession: () => Promise.resolve(),
    createRuntime: () => Promise.resolve(),
    openResumeDialog: () => Promise.resolve(),
    openRecoveryDialog: async () => {
      await Promise.resolve();
      recoveryOpenCount += 1;
    },
    reopenLastClosed: () => Promise.resolve(),
    revertLatestCheckpoint: async () => {
      await Promise.resolve();
      if (failRevert) {
        throw new Error("restore failed");
      }
      revertCount += 1;
    },
    createManualFullBackup: async () => {
      await Promise.resolve();
      if (failBackup) {
        throw new Error("backup disk unavailable");
      }
      createBackupCount += 1;
      return { id: "backup-created", createdAt: 1, sizeBytes: 2048 };
    },
    listManualFullBackups: () => Promise.resolve([]),
    restoreManualFullBackup: async (backupId) => {
      await Promise.resolve();
      restoredBackupIds.push(backupId);
      return { id: backupId ?? "latest", createdAt: 1, sizeBytes: 2048 };
    },
    clearManualFullBackups: () => Promise.resolve(0),
    openExtensionsHub: () => {},
    openFilesWorkspace: () => {},
  };

  try {
    registerBuiltins(context);

    assert.equal(recoveryOpenCount, 0);
    await executeRegisteredCommand("history");
    assert.equal(recoveryOpenCount, 1);

    const blockedRevert = executeSlashCommand({
      name: "revert",
      args: "",
      busy: true,
      onError: () => {},
    });
    assert.equal(blockedRevert, "busy-blocked");
    assert.equal(revertCount, 0);

    await executeRegisteredCommand("revert");
    assert.equal(revertCount, 1);

    failRevert = true;
    await assert.rejects(executeRegisteredCommand("revert"), /restore failed/);

    const blockedBackup = executeSlashCommand({
      name: "backup",
      args: "",
      busy: true,
      onError: () => {},
    });
    assert.equal(blockedBackup, "busy-blocked");
    assert.equal(createBackupCount, 0);

    await executeRegisteredCommand("backup", "create");
    assert.equal(createBackupCount, 1);
    assert.match(getToastText(fakeDom.document), /Backup created/i);

    await executeRegisteredCommand("backup", "restore backup-123");
    assert.deepEqual(restoredBackupIds, ["backup-123"]);

    failBackup = true;
    await executeRegisteredCommand("backup", "create");
    assert.match(getToastText(fakeDom.document), /backup disk unavailable/i);

    const queuedCompact = executeSlashCommand({
      name: "compact",
      args: "now",
      busy: true,
      enqueueCommand: (name, args) => {
        queuedCommands.push({ name, args });
      },
      onError: () => {},
    });
    assert.equal(queuedCompact, "queued");
    assert.deepEqual(queuedCommands, [{ name: "compact", args: "now" }]);

    assert.equal(
      executeSlashCommand({ name: "compact", args: "", busy: false, onError: () => {} }),
      "missing-queue",
    );
  } finally {
    restoreCommands(previousCommands);
    fakeDom.restore();
  }
});

void test("command-layer contract: busy policy allows workspace commands and blocks ordinary commands", () => {
  const previousCommands = commandRegistry.list();
  const executed: string[] = [];
  const register = (command: Pick<SlashCommand, "name" | "source"> & Partial<Pick<SlashCommand, "busyAllowed">>): void => {
    commandRegistry.register({
      name: command.name,
      description: command.name,
      source: command.source,
      ...(command.busyAllowed !== undefined ? { busyAllowed: command.busyAllowed } : {}),
      execute: () => {
        executed.push(command.name);
      },
    });
  };
  const run = (name: string): SlashCommandExecutionResult =>
    executeSlashCommand({ name, args: "", busy: true, enqueueCommand: () => {}, onError: () => {} });

  try {
    for (const command of commandRegistry.list()) commandRegistry.unregister(command.name);
    const allowedWhileBusy = ["new", "rules", "resume", "history", "reopen", "yolo", "extensions", "plugins", "skills", "files", TOOLS_COMMAND_NAME];
    for (const name of allowedWhileBusy) register({ name, source: "builtin" });
    register({ name: "addons", source: "builtin" });
    register({ name: "integrations", source: "builtin" });
    register({ name: "opted-in", source: "builtin", busyAllowed: true });
    register({ name: "ext-default", source: "extension" });
    register({ name: "ext-opted-out", source: "extension", busyAllowed: false });

    for (const name of allowedWhileBusy) assert.equal(run(name), "executed", `/${name} must run while busy`);
    assert.equal(run("opted-in"), "executed");
    assert.equal(run("ext-default"), "executed", "extension commands run while busy unless they opt out");
    assert.equal(run("addons"), "busy-blocked");
    assert.equal(run("integrations"), "busy-blocked");
    assert.equal(run("ext-opted-out"), "busy-blocked");
    assert.equal(run("compact"), "not-found");
    assert.deepEqual(
      executed,
      [...allowedWhileBusy, "opted-in", "ext-default"],
      "blocked commands must not execute",
    );
  } finally {
    restoreCommands(previousCommands);
  }
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
