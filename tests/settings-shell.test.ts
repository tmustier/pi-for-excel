import assert from "node:assert/strict";
import { test } from "node:test";

import { closeOverlayById } from "../src/ui/overlay-dialog.ts";
import { createSettingsShell, type SettingsShellPage } from "../src/ui/settings-shell.ts";
import { installFakeDom } from "./fake-dom.test.ts";

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

void test("settings shell registered closer honors before-leave guard", async () => {
  const { document: fakeDocument, restore } = installFakeDom();
  let allowClose = false;

  const page: SettingsShellPage = {
    id: "root",
    title: () => "Root",
    render: (ctx) => {
      ctx.body.append(createMarker("guarded-marker", "Guarded page"));
      ctx.setBeforeLeave(() => Promise.resolve(allowClose));
    },
  };

  const shell = createSettingsShell({
    overlayId: "settings-guard-test",
    rootId: "root",
    getPage: (id) => (id === "root" ? page : undefined),
    backLabel: () => "Back",
    closeLabel: () => "Close",
  });

  try {
    await shell.open();
    assert.ok(fakeDocument.getElementById("settings-guard-test"));

    assert.equal(closeOverlayById("settings-guard-test"), true);
    await flushAsync();
    assert.ok(fakeDocument.getElementById("settings-guard-test"));
    assert.ok(fakeDocument.getElementById("guarded-marker"));

    allowClose = true;
    assert.equal(closeOverlayById("settings-guard-test"), true);
    await flushAsync();
    assert.equal(fakeDocument.getElementById("settings-guard-test"), null);
  } finally {
    restore();
  }
});
