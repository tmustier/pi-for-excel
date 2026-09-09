// Component-logic contracts only; native selection behavior lives in the Chromium suite.
import assert from "node:assert/strict";
import { test } from "node:test";

import { initLanguage } from "../src/language/index.ts";
import { createCopyableCommand } from "../src/ui/command-copy.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

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
