// Component-logic contracts only; this suite does not claim native browser behavior.
import assert from "node:assert/strict";
import { test } from "node:test";

import { closeOverlayById } from "../src/ui/overlay-dialog.ts";
import { createSettingsShell, type SettingsPageContext, type SettingsShellPage } from "../src/ui/settings-shell.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

async function flushAsync(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

function createMarker(id: string, text: string): HTMLElement {
  const marker = document.createElement("div");
  marker.id = id;
  marker.textContent = text;
  return marker;
}

void test("settings shell ignores stale async renders after navigation", async () => {
  const { document: fakeDocument, restore } = installFakeDom();
  let resolveSlow: (() => void) | null = null;
  let staleCleanupCalls = 0;

  const slowReady = new Promise<void>((resolve) => {
    resolveSlow = resolve;
  });

  const pages = new Map<string, SettingsShellPage>([
    [
      "root",
      {
        id: "root",
        title: () => "Root",
        render: (ctx) => {
          ctx.body.append(createMarker("root-marker", "Root page"));
        },
      },
    ],
    [
      "slow",
      {
        id: "slow",
        parentId: "root",
        title: () => "Slow",
        render: async (ctx) => {
          await slowReady;

          ctx.addCleanup(() => {
            staleCleanupCalls += 1;
          });
          ctx.setBeforeLeave(() => Promise.resolve(false));

          const footer = createMarker("slow-footer", "Slow footer");
          ctx.setFooter(footer);
          ctx.body.append(createMarker("slow-marker", "Slow page"));
        },
      },
    ],
  ]);

  const shell = createSettingsShell({
    overlayId: "settings-test",
    rootId: "root",
    getPage: (id) => pages.get(id),
    backLabel: () => "Back",
    closeLabel: () => "Close",
  });

  try {
    await shell.open("slow");
    assert.ok(fakeDocument.getElementById("settings-test"));

    await shell.open("root");
    assert.ok(fakeDocument.getElementById("root-marker"));

    if (!resolveSlow) {
      throw new Error("Expected slow page resolver to be initialized");
    }
    resolveSlow();
    await flushAsync();

    assert.ok(fakeDocument.getElementById("root-marker"));
    assert.equal(fakeDocument.getElementById("slow-marker"), null);
    assert.equal(fakeDocument.getElementById("slow-footer"), null);
    assert.equal(staleCleanupCalls, 1);

    assert.equal(closeOverlayById("settings-test"), true);
    await flushAsync();
    assert.equal(fakeDocument.getElementById("settings-test"), null);
  } finally {
    restore();
  }
});

void test("Escape closes a deep-linked page but walks back through pages opened inside the shell", async () => {
  const { document: fakeDocument, restore } = installFakeDom();
  const rendered: { root?: SettingsPageContext } = {};
  const pages = new Map<string, SettingsShellPage>([
    ["root", { id: "root", title: () => "Root", render: (ctx) => { rendered.root = ctx; ctx.body.append(createMarker("root-marker", "Root")); } }],
    ["backups", { id: "backups", parentId: "root", title: () => "Backups", render: (ctx) => { ctx.body.append(createMarker("backups-marker", "Backups")); } }],
  ]);
  const shell = createSettingsShell({
    overlayId: "settings-escape-test",
    rootId: "root",
    getPage: (id) => pages.get(id),
    backLabel: () => "Back",
    closeLabel: () => "Close",
  });
  const pressEscape = async (): Promise<void> => {
    const event = new Event("keydown", { cancelable: true });
    Object.defineProperty(event, "key", { value: "Escape" });
    fakeDocument.dispatchEvent(event);
    await flushAsync();
  };

  try {
    // /history opens Backups directly: Escape closes instead of revealing Settings.
    await shell.open("backups");
    await pressEscape();
    assert.equal(shell.isOpen(), false);

    // Opened at Settings and navigated to Backups: Escape goes back one level.
    await shell.open("root");
    rendered.root?.navigate("backups");
    await flushAsync();
    assert.ok(fakeDocument.getElementById("backups-marker"));
    await pressEscape();
    assert.equal(shell.isOpen(), true);
    assert.ok(fakeDocument.getElementById("root-marker"));
    await pressEscape();
    assert.equal(shell.isOpen(), false);
  } finally {
    closeOverlayById("settings-escape-test");
    restore();
  }
});
