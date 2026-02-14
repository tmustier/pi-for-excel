import assert from "node:assert/strict";
import { test } from "node:test";

import { createExtensionAPI } from "../src/commands/extension-api.ts";
import {
  clearAllExtensionWidgets,
} from "../src/extensions/internal/widget-surface.ts";
import type { ExtensionCapability } from "../src/extensions/permissions.ts";
import { installFakeDom } from "./fake-dom.test.ts";

function createCapabilityGate(allowed: ReadonlySet<ExtensionCapability>) {
  return (capability: ExtensionCapability): boolean => allowed.has(capability);
}

void test("widget v2 methods throw clear guidance when experiment is disabled", () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    widgetApiV2Enabled: false,
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.read",
      "agent.events.read",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
  });

  assert.throws(
    () => {
      api.widget.remove("summary");
    },
    /Widget API v2 is disabled\./,
  );

  assert.throws(
    () => {
      api.widget.clear();
    },
    /extension-widget-v2/,
  );
});

void test("widget v2 methods still enforce ui.widget capability when enabled", () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    widgetApiV2Enabled: true,
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.read",
      "agent.events.read",
      "ui.overlay",
      "ui.toast",
    ])),
    formatCapabilityError: (capability) => `DENIED:${capability}`,
  });

  assert.throws(
    () => {
      api.widget.remove("summary");
    },
    /DENIED:ui\.widget/,
  );
});

void test("widget v2 upsert/remove/clear render and clear slots", () => {
  const { document, restore } = installFakeDom();
  clearAllExtensionWidgets();

  try {
    const rawContainer = document.createElement("div");
    const rawInputArea = document.createElement("div");

    if (!(rawContainer instanceof HTMLElement)) {
      throw new Error("Expected container element to be HTMLElement");
    }

    if (!(rawInputArea instanceof HTMLElement)) {
      throw new Error("Expected input area element to be HTMLElement");
    }

    const container = rawContainer;
    const inputArea = rawInputArea;

    inputArea.className = "pi-input-area";
    container.appendChild(inputArea);
    document.body.appendChild(container);

    Reflect.set(container, "insertBefore", (node: Node): Node => {
      container.appendChild(node);
      return node;
    });

    Reflect.set(document, "querySelector", (selector: string): HTMLElement | null => {
      return selector === ".pi-input-area" ? inputArea : null;
    });

    const api = createExtensionAPI({
      getAgent: () => {
        throw new Error("getAgent should not be called");
      },
      extensionOwnerId: "ext.widgets",
      widgetApiV2Enabled: true,
      isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
        "ui.widget",
      ])),
    });

    const rawSummary = document.createElement("div");
    if (!(rawSummary instanceof HTMLElement)) {
      throw new Error("Expected summary widget element to be HTMLElement");
    }

    const summary = rawSummary;
    summary.textContent = "Summary widget";

    api.widget.upsert({
      id: "summary",
      el: summary,
      title: "Summary",
      minHeightPx: 100,
      maxHeightPx: 200,
    });

    const aboveSlot = document.getElementById("pi-widget-slot");
    assert.ok(aboveSlot);
    assert.equal(aboveSlot.style.display, "flex");
    assert.equal(aboveSlot.children.length, 1);

    const firstCard = aboveSlot.children[0];
    assert.ok(firstCard instanceof HTMLElement);
    assert.equal(firstCard.dataset.widgetId, "summary");

    const firstBody = firstCard.children[1];
    assert.ok(firstBody instanceof HTMLElement);
    assert.equal(firstBody.style.minHeight, "100px");
    assert.equal(firstBody.style.maxHeight, "200px");

    const rawBelow = document.createElement("div");
    if (!(rawBelow instanceof HTMLElement)) {
      throw new Error("Expected below widget element to be HTMLElement");
    }

    const below = rawBelow;
    below.textContent = "Below widget";

    api.widget.upsert({
      id: "below",
      el: below,
      placement: "below-input",
    });

    const belowSlot = document.getElementById("pi-widget-slot-below");
    assert.ok(belowSlot);
    assert.equal(belowSlot.children.length, 1);

    api.widget.remove("summary");
    assert.equal(aboveSlot.children.length, 0);

    api.widget.clear();
    assert.equal(aboveSlot.children.length, 0);
    assert.equal(belowSlot.children.length, 0);
  } finally {
    clearAllExtensionWidgets();
    restore();
  }
});
