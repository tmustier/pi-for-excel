import assert from "node:assert/strict";
import { test } from "node:test";

import { initLanguage } from "../src/language/index.ts";
import { createCopyableCommand } from "../src/ui/command-copy.ts";
import { installFakeDom } from "./fake-dom.test.ts";

function restoreProperty(key: string, descriptor: PropertyDescriptor | undefined): void {
  if (descriptor) {
    Object.defineProperty(globalThis, key, descriptor);
    return;
  }

  Reflect.deleteProperty(globalThis, key);
}

void test("createCopyableCommand copies, translates its accessible state, and resets it", async () => {
  const { restore } = installFakeDom();
  const navigatorDescriptor = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  const windowDescriptor = Object.getOwnPropertyDescriptor(globalThis, "window");
  let copiedText = "";

  try {
    Object.defineProperty(globalThis, "navigator", {
      configurable: true,
      value: {
        clipboard: {
          writeText: (text: string) => {
            copiedText = text;
            return Promise.resolve();
          },
        },
      },
    });
    Object.defineProperty(globalThis, "window", {
      configurable: true,
      value: { getSelection: () => null },
    });
    initLanguage("zh-CN");

    const row = createCopyableCommand("npx pi-for-excel-proxy");
    const button = row.children[1];
    assert.ok(button instanceof HTMLElement);
    assert.equal(button.getAttribute("aria-label"), "复制命令");

    button.dispatchEvent(new Event("click"));
    await Promise.resolve();

    assert.equal(copiedText, "npx pi-for-excel-proxy");
    assert.equal(button.getAttribute("aria-label"), "已复制");

    await new Promise((resolve) => setTimeout(resolve, 1450));
    assert.equal(button.getAttribute("aria-label"), "复制命令");
  } finally {
    initLanguage("en");
    restoreProperty("navigator", navigatorDescriptor);
    restoreProperty("window", windowDescriptor);
    restore();
  }
});

void test("createCopyableCommand selects the command when clipboard access is unavailable", () => {
  const { document: fakeDocument, restore } = installFakeDom();
  const navigatorDescriptor = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  const windowDescriptor = Object.getOwnPropertyDescriptor(globalThis, "window");
  let selectedElement: HTMLElement | null = null;
  let selectionCleared = false;
  let rangeAdded = false;

  try {
    const range = {
      selectNodeContents: (element: HTMLElement) => {
        selectedElement = element;
      },
    };
    Reflect.set(fakeDocument, "createRange", () => range);
    Object.defineProperty(globalThis, "navigator", {
      configurable: true,
      value: {},
    });
    Object.defineProperty(globalThis, "window", {
      configurable: true,
      value: {
        getSelection: () => ({
          removeAllRanges: () => {
            selectionCleared = true;
          },
          addRange: (addedRange: object) => {
            rangeAdded = addedRange === range;
          },
        }),
      },
    });

    const row = createCopyableCommand("npx pi-for-excel-python-bridge");
    const code = row.children[0];
    const button = row.children[1];
    assert.ok(code instanceof HTMLElement);
    assert.ok(button instanceof HTMLElement);

    button.dispatchEvent(new Event("click"));

    assert.equal(selectedElement, code);
    assert.equal(selectionCleared, true);
    assert.equal(rangeAdded, true);
  } finally {
    restoreProperty("navigator", navigatorDescriptor);
    restoreProperty("window", windowDescriptor);
    restore();
  }
});
