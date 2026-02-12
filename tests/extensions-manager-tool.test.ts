import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createExtensionsManagerTool,
  type ExtensionsManagerToolRuntime,
  type ExtensionsManagerToolStatus,
} from "../src/tools/extensions-manager.ts";

function makeStatus(input: {
  id: string;
  name: string;
  enabled?: boolean;
  loaded?: boolean;
  sourceLabel?: string;
  trustLabel?: string;
  effectiveCapabilities?: string[];
  lastError?: string | null;
}): ExtensionsManagerToolStatus {
  return {
    id: input.id,
    name: input.name,
    enabled: input.enabled ?? true,
    loaded: input.loaded ?? true,
    sourceLabel: input.sourceLabel ?? "inline code",
    trustLabel: input.trustLabel ?? "inline code",
    effectiveCapabilities: input.effectiveCapabilities ?? ["commands.register"],
    lastError: input.lastError ?? null,
  };
}

class FakeExtensionsManager implements ExtensionsManagerToolRuntime {
  private nextId = 1;
  private statuses: ExtensionsManagerToolStatus[] = [];

  constructor(initialStatuses: ExtensionsManagerToolStatus[] = []) {
    this.statuses = [...initialStatuses];
  }

  list(): ExtensionsManagerToolStatus[] {
    return this.statuses.map((status) => ({
      ...status,
      effectiveCapabilities: [...status.effectiveCapabilities],
    }));
  }

  installFromCode(name: string, _code: string): Promise<string> {
    const id = `ext.generated.${this.nextId}`;
    this.nextId += 1;

    this.statuses.push(makeStatus({ id, name }));
    return Promise.resolve(id);
  }

  setExtensionEnabled(entryId: string, enabled: boolean): Promise<void> {
    const status = this.statuses.find((entry) => entry.id === entryId);
    if (!status) {
      return Promise.reject(new Error("Extension not found"));
    }

    status.enabled = enabled;
    status.loaded = enabled;
    return Promise.resolve();
  }

  reloadExtension(entryId: string): Promise<void> {
    const status = this.statuses.find((entry) => entry.id === entryId);
    if (!status) {
      return Promise.reject(new Error("Extension not found"));
    }

    status.loaded = status.enabled;
    return Promise.resolve();
  }

  uninstallExtension(entryId: string): Promise<void> {
    const index = this.statuses.findIndex((entry) => entry.id === entryId);
    if (index < 0) {
      return Promise.reject(new Error("Extension not found"));
    }

    this.statuses.splice(index, 1);
    return Promise.resolve();
  }
}

void test("extensions_manager lists installed extensions", async () => {
  const manager = new FakeExtensionsManager([
    makeStatus({
      id: "ext.alpha",
      name: "Alpha",
      effectiveCapabilities: ["commands.register", "ui.widget"],
    }),
  ]);

  const tool = createExtensionsManagerTool({
    getManager: () => manager,
  });

  const result = await tool.execute("call-list", { action: "list" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Installed extensions:/);
  assert.match(text, /Alpha \(ext\.alpha\)/);
  assert.match(text, /commands\.register, ui\.widget/);
});

void test("extensions_manager install_code replaces extensions with same name by default", async () => {
  const manager = new FakeExtensionsManager([
    makeStatus({ id: "ext.old", name: "Quick KPI" }),
  ]);

  const tool = createExtensionsManagerTool({
    getManager: () => manager,
  });

  const result = await tool.execute("call-install", {
    action: "install_code",
    name: "Quick KPI",
    code: "export function activate(api) { api.toast('hi'); }",
  });

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  const statuses = manager.list();

  assert.equal(statuses.length, 1);
  assert.equal(statuses[0].name, "Quick KPI");
  assert.notEqual(statuses[0].id, "ext.old");
  assert.match(text, /Replaced 1 existing extension/);
});

void test("extensions_manager install_code can fail on duplicate names when replace_existing=false", async () => {
  const manager = new FakeExtensionsManager([
    makeStatus({ id: "ext.old", name: "Quick KPI" }),
  ]);

  const tool = createExtensionsManagerTool({
    getManager: () => manager,
  });

  await assert.rejects(
    () => tool.execute("call-install-duplicate", {
      action: "install_code",
      name: "Quick KPI",
      code: "export function activate(api) {}",
      replace_existing: false,
    }),
    /already exists/i,
  );
});
