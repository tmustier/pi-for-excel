import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer({ token: "extension-connections-browser-token" });
});

after(async () => {
  await env.close();
});

function readInstalledExtensionId(raw: DynamicValue): string | null {
  if (typeof raw !== "object" || raw === null || Array.isArray(raw)) return null;
  const envelope = raw as DynamicObject;
  if (typeof envelope.result !== "object" || envelope.result === null || Array.isArray(envelope.result)) return null;
  const result = envelope.result as DynamicObject;
  return typeof result.extensionId === "string" ? result.extensionId : null;
}

async function openConnectionsWithExtension(code?: string, grantConnections = false): Promise<{
  plugins: ReturnType<Page["locator"]>;
  finish(): Promise<void>;
}> {
  let commandSent = false;
  let capabilitySent = false;
  let installedExtensionId: string | null = null;
  let resolveInstalled: (() => void) | undefined;
  const installed = new Promise<void>((resolve) => {
    resolveInstalled = resolve;
  });

  const opened = await openTaskpane(env, {
    clientId: "extension-connections-test",
    bridge: async (url, route) => {
      if (url.pathname === "/client/poll" && !commandSent) {
        commandSent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify(code
            ? {
              id: "install-connection-extension",
              type: "extensionInstallCode",
              payload: { name: "Connection browser extension", code },
            }
            : {
              id: "uninstall-builtin-extension",
              type: "extensionUninstall",
              payload: { extensionId: "builtin.snake" },
            }),
        });
      } else if (grantConnections && installedExtensionId && url.pathname === "/client/poll" && !capabilitySent) {
        capabilitySent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify({
            id: "grant-connection-capability",
            type: "extensionSetCapability",
            payload: {
              extensionId: installedExtensionId,
              capability: "connections.readwrite",
              allowed: true,
            },
          }),
        });
      } else if (url.pathname === "/client/result") {
        const raw: DynamicValue = JSON.parse(route.request().postData() ?? "{}");
        installedExtensionId = installedExtensionId ?? readInstalledExtensionId(raw);
        if (!grantConnections || capabilitySent) resolveInstalled?.();
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
      } else {
        return false;
      }
      return true;
    },
  });
  const { page } = opened;
  const input = page.locator("pi-input textarea");

  await Promise.race([
    installed,
    new Promise<void>((_resolve, reject) => {
      setTimeout(() => reject(new Error("extension setup did not complete")), 10_000);
    }),
  ]);

  await input.fill("/plugins");
  await input.press("Enter");
  const plugins = page.locator("#pi-settings-overlay");
  await plugins.getByRole("heading", { name: "Plugins" }).waitFor({ state: "visible", timeout: 5_000 });
  await plugins.getByRole("button", { name: "Back" }).click();
  await plugins.getByRole("button", { name: /^Connections/u }).click();
  await plugins.getByRole("heading", { name: "Connections" }).waitFor({ state: "visible" });
  return { plugins, finish: opened.finish };
}

void test("Connections hides extension connections when no extensions are installed", async () => {
  const { plugins, finish } = await openConnectionsWithExtension();
  try {
    await plugins.locator(".pi-section-header__label", { hasText: "MCP servers" }).waitFor({ state: "visible" });
    assert.equal(await plugins.locator(".pi-section-header__label", { hasText: "Extension connections" }).count(), 0);
  } finally {
    await finish();
  }
});

void test("Connections explains when installed extensions register no connections", async () => {
  const { plugins, finish } = await openConnectionsWithExtension("export function activate() {}");
  try {
    await plugins.getByText("Extension connections").waitFor({ state: "visible", timeout: 5_000 });
    await plugins.getByText("Installed extensions haven't registered any connections.").waitFor({ state: "visible", timeout: 5_000 });
  } finally {
    await finish();
  }
});

void test("Connections renders connections registered by an installed extension", async () => {
  const { plugins, finish } = await openConnectionsWithExtension(`
    export function activate(api) {
      api.connections.register({
        id: "apollo",
        title: "Apollo",
        capability: "company enrichment via Apollo API",
        authKind: "api_key",
        secretFields: [{ id: "apiKey", label: "API key", required: true }]
      });
    }
  `, true);
  try {
    const card = plugins.locator(".pi-item-card", { hasText: "Apollo" });
    await card.waitFor({ state: "visible", timeout: 5_000 });
    assert.match((await card.textContent()) ?? "", /company enrichment via Apollo API/u);
  } finally {
    await finish();
  }
});
