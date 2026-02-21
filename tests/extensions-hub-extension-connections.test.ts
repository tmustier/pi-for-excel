import assert from "node:assert/strict";
import { test } from "node:test";

import { renderExtensionConnectionsSection } from "../src/commands/builtins/extensions-hub-extension-connections.ts";
import { ConnectionManager } from "../src/connections/manager.ts";
import type { ConnectionDefinition } from "../src/connections/types.ts";
import { installFakeDom } from "./fake-dom.test.ts";

function createMemorySettings(): {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
} {
  const data = new Map<string, unknown>();
  return {
    get: (key) => Promise.resolve(data.get(key)),
    set: (key, value) => {
      data.set(key, value);
      return Promise.resolve();
    },
  };
}

function collectTextContent(root: HTMLElement): string[] {
  const texts: string[] = [];

  const visit = (node: HTMLElement): void => {
    const text = node.textContent.trim();
    if (text.length > 0) {
      texts.push(text);
    }

    for (const child of Array.from(node.children)) {
      if (child instanceof HTMLElement) {
        visit(child);
      }
    }
  };

  visit(root);
  return texts;
}

const APOLLO_DEFINITION: ConnectionDefinition = {
  id: "ext.apollo.apollo",
  title: "Apollo",
  capability: "company enrichment via Apollo API",
  authKind: "api_key",
  secretFields: [{ id: "apiKey", label: "API key", required: true }],
};

void test("extension connections section is hidden when no extensions are installed", async () => {
  const { document, restore } = installFakeDom();

  try {
    const container = document.createElement("div");
    assert.ok(container instanceof HTMLElement);

    const connectionManager = new ConnectionManager({ settings: createMemorySettings() });
    connectionManager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

    await renderExtensionConnectionsSection({
      container,
      connectionManager,
      extensionManager: {
        list: () => [],
      },
    });

    assert.equal(container.children.length, 0);
  } finally {
    restore();
  }
});

void test("extension connections section shows empty state when extensions are installed but none registered connections", async () => {
  const { document, restore } = installFakeDom();

  try {
    const container = document.createElement("div");
    assert.ok(container instanceof HTMLElement);

    const connectionManager = new ConnectionManager({ settings: createMemorySettings() });

    await renderExtensionConnectionsSection({
      container,
      connectionManager,
      extensionManager: {
        list: () => [{ id: "ext.one" }],
      },
    });

    const text = collectTextContent(container).join("\n");
    assert.match(text, /Extension connections/);
    assert.match(text, /Installed extensions haven't registered any connections\./);
  } finally {
    restore();
  }
});

void test("extension connections section renders cards for registered connections", async () => {
  const { document, restore } = installFakeDom();

  try {
    const container = document.createElement("div");
    assert.ok(container instanceof HTMLElement);

    const connectionManager = new ConnectionManager({ settings: createMemorySettings() });
    connectionManager.registerDefinition("ext.apollo", APOLLO_DEFINITION);

    await renderExtensionConnectionsSection({
      container,
      connectionManager,
      extensionManager: {
        list: () => [{ id: "ext.apollo" }],
      },
    });

    const text = collectTextContent(container).join("\n");
    assert.match(text, /Extension connections/);
    assert.match(text, /Apollo/);
    assert.match(text, /company enrichment via Apollo API/);
    assert.match(text, /Not configured/);
    assert.match(text, /API key/);
    assert.match(text, /Save/);
  } finally {
    restore();
  }
});
