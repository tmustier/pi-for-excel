import assert from "node:assert/strict";
import { test } from "node:test";

import { renderPluginsTab } from "../src/commands/builtins/extensions-hub-plugins.ts";
import { createExtensionAPI } from "../src/commands/extension-api.ts";
import { ConnectionManager } from "../src/connections/manager.ts";
import {
  describeStoredExtensionTrust,
  getDefaultPermissionsForTrust,
  listAllExtensionCapabilities,
  listGrantedExtensionCapabilities,
  type StoredExtensionTrust,
} from "../src/extensions/permissions.ts";
import { ExtensionRuntimeManager, type ExtensionRuntimeStatus } from "../src/extensions/runtime-manager.ts";
import { describeExtensionSource } from "../src/extensions/runtime-manager-helpers.ts";
import { describeExtensionRuntimeMode, type ExtensionRuntimeMode } from "../src/extensions/runtime-mode.ts";
import type { ExtensionSettingsStore, StoredExtensionSource } from "../src/extensions/store.ts";
import { EXTENSION_OVERLAY_ID } from "../src/ui/overlay-ids.ts";
import { installFakeDom } from "./fake-dom.test.ts";

class MemorySettingsStore implements ExtensionSettingsStore {
  get(_key: string): Promise<unknown> {
    return Promise.resolve(null);
  }

  set(_key: string, _value: unknown): Promise<void> {
    return Promise.resolve();
  }
}

class StaticExtensionRuntimeManager extends ExtensionRuntimeManager {
  private readonly statuses: ExtensionRuntimeStatus[];

  constructor(statuses: readonly ExtensionRuntimeStatus[]) {
    const settings = new MemorySettingsStore();

    super({
      settings,
      connectionManager: new ConnectionManager({ settings }),
      getActiveAgent: () => null,
      refreshRuntimeTools: async () => {},
      reservedToolNames: new Set<string>(),
    });

    this.statuses = statuses.map(cloneStatus);
  }

  override list(): ExtensionRuntimeStatus[] {
    return this.statuses.map(cloneStatus);
  }

  override subscribe(_listener: () => void): () => void {
    return () => {};
  }
}

function cloneSource(source: StoredExtensionSource): StoredExtensionSource {
  if (source.kind === "module") {
    return {
      kind: "module",
      specifier: source.specifier,
    };
  }

  return {
    kind: "inline",
    code: source.code,
  };
}

function cloneStatus(status: ExtensionRuntimeStatus): ExtensionRuntimeStatus {
  return {
    ...status,
    source: cloneSource(status.source),
    permissions: { ...status.permissions },
    grantedCapabilities: [...status.grantedCapabilities],
    effectiveCapabilities: [...status.effectiveCapabilities],
    commandNames: [...status.commandNames],
    toolNames: [...status.toolNames],
  };
}

function createRuntimeStatus(input: {
  id: string;
  name: string;
  source: StoredExtensionSource;
  trust: StoredExtensionTrust;
  runtimeMode: ExtensionRuntimeMode;
  enabled?: boolean;
  loaded?: boolean;
  permissionsEnforced?: boolean;
  commandNames?: string[];
  toolNames?: string[];
  lastError?: string | null;
}): ExtensionRuntimeStatus {
  const permissions = getDefaultPermissionsForTrust(input.trust);
  const grantedCapabilities = listGrantedExtensionCapabilities(permissions);
  const effectiveCapabilities = input.permissionsEnforced === true
    ? grantedCapabilities
    : listAllExtensionCapabilities();

  return {
    id: input.id,
    name: input.name,
    enabled: input.enabled ?? true,
    loaded: input.loaded ?? true,
    source: cloneSource(input.source),
    sourceLabel: describeExtensionSource(input.source),
    trust: input.trust,
    trustLabel: describeStoredExtensionTrust(input.trust),
    runtimeMode: input.runtimeMode,
    runtimeLabel: describeExtensionRuntimeMode(input.runtimeMode),
    permissions,
    grantedCapabilities,
    effectiveCapabilities,
    permissionsEnforced: input.permissionsEnforced ?? false,
    commandNames: input.commandNames ?? [],
    toolNames: input.toolNames ?? [],
    lastError: input.lastError ?? null,
  };
}

function collectElements(root: HTMLElement, predicate: (element: HTMLElement) => boolean): HTMLElement[] {
  const matches: HTMLElement[] = [];

  const visit = (element: HTMLElement): void => {
    if (predicate(element)) {
      matches.push(element);
    }

    for (const child of Array.from(element.children)) {
      if (!(child instanceof HTMLElement)) {
        continue;
      }

      visit(child);
    }
  };

  visit(root);
  return matches;
}

function collectTextContent(root: HTMLElement): string[] {
  return collectElements(root, () => true)
    .map((element) => (element.textContent ?? "").trim())
    .filter((text) => text.length > 0);
}

void test("extension overlay show/dismiss mounts and tears down shared overlay", () => {
  const { document, restore } = installFakeDom();

  try {
    const api = createExtensionAPI({
      getAgent: () => {
        throw new Error("agent should not be requested in overlay test");
      },
    });

    const first = document.createElement("div");
    first.id = "first-content";

    api.overlay.show(first);

    const mounted = document.getElementById(EXTENSION_OVERLAY_ID);
    assert.notEqual(mounted, null);
    if (!mounted) {
      return;
    }

    assert.equal(mounted.contains(first), true);

    const second = document.createElement("div");
    second.id = "second-content";

    api.overlay.show(second);

    const remounted = document.getElementById(EXTENSION_OVERLAY_ID);
    assert.notEqual(remounted, null);
    if (!remounted) {
      return;
    }

    assert.equal(remounted.contains(second), true);
    assert.equal(remounted.contains(first), false);

    api.overlay.dismiss();
    assert.equal(document.getElementById(EXTENSION_OVERLAY_ID), null);
  } finally {
    restore();
  }
});

void test("extensions hub plugins tab renders installed/extensions sections", () => {
  const { document, restore } = installFakeDom();

  try {
    const builtinStatus = createRuntimeStatus({
      id: "builtin.snake",
      name: "Snake",
      source: {
        kind: "module",
        specifier: "../extensions/snake.js",
      },
      trust: "builtin",
      runtimeMode: "host",
      commandNames: ["snake"],
      toolNames: ["snake_tool"],
    });

    const inlineErrorStatus = createRuntimeStatus({
      id: "ext.inline.error",
      name: "Broken Inline",
      source: {
        kind: "inline",
        code: "export function activate(api) { api.toast('broken'); }",
      },
      trust: "inline-code",
      runtimeMode: "sandbox-iframe",
      loaded: false,
      lastError: "Local extension module \"../extensions/snake.js\" was not bundled.",
    });

    const manager = new StaticExtensionRuntimeManager([builtinStatus, inlineErrorStatus]);

    const container = document.createElement("div");
    renderPluginsTab({
      container,
      manager,
      isBusy: () => false,
      onChanged: async () => {},
    });

    const texts = collectTextContent(container);
    assert.equal(texts.includes("Installed"), true);
    assert.equal(texts.includes("Install"), true);
    assert.equal(texts.includes("Advanced"), false, "Advanced section should be removed");
    assert.equal(texts.includes("Snake"), true);
    assert.equal(texts.includes("Broken Inline"), true);
  } finally {
    restore();
  }
});
