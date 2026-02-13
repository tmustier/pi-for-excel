import assert from "node:assert/strict";
import { test } from "node:test";

import { PI_REQUEST_INPUT_FOCUS_EVENT } from "../src/ui/input-focus.ts";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayDialogManager,
} from "../src/ui/overlay-dialog.ts";
import { installFakeDom } from "./fake-dom.test.ts";

void test("closeOverlayById returns false when overlay does not exist", () => {
  const { restore } = installFakeDom();

  try {
    assert.equal(closeOverlayById("missing-overlay"), false);
  } finally {
    restore();
  }
});

void test("overlay dialog exposes dialog semantics", () => {
  const { restore } = installFakeDom();

  try {
    const dialog = createOverlayDialog({
      overlayId: "overlay-dialog-a11y",
      cardClassName: "overlay-card",
    });

    assert.equal(dialog.overlay.getAttribute("role"), "dialog");
    assert.equal(dialog.overlay.getAttribute("aria-modal"), "true");
  } finally {
    restore();
  }
});

void test("closeOverlayById closes mounted overlay and restores input focus", () => {
  const { document, restore } = installFakeDom();

  try {
    let focusRequests = 0;
    document.addEventListener(PI_REQUEST_INPUT_FOCUS_EVENT, () => {
      focusRequests += 1;
    });

    const dialog = createOverlayDialog({
      overlayId: "overlay-close-by-id",
      cardClassName: "overlay-card",
    });

    dialog.mount();
    assert.equal(document.getElementById("overlay-close-by-id") !== null, true);

    assert.equal(closeOverlayById("overlay-close-by-id"), true);
    assert.equal(document.getElementById("overlay-close-by-id"), null);
    assert.equal(focusRequests, 1);
  } finally {
    restore();
  }
});

void test("overlay closes on backdrop click and runs cleanup", () => {
  const { document, restore } = installFakeDom();

  try {
    let cleanedUp = 0;

    const dialog = createOverlayDialog({
      overlayId: "overlay-backdrop-close",
      cardClassName: "overlay-card",
    });

    dialog.addCleanup(() => {
      cleanedUp += 1;
    });

    dialog.mount();
    dialog.overlay.dispatchEvent(new Event("click"));

    assert.equal(document.getElementById("overlay-backdrop-close"), null);
    assert.equal(cleanedUp, 1);
  } finally {
    restore();
  }
});

void test("overlay closes on Escape key", () => {
  const { document, restore } = installFakeDom();

  try {
    const dialog = createOverlayDialog({
      overlayId: "overlay-escape-close",
      cardClassName: "overlay-card",
    });
    dialog.mount();

    const event = new Event("keydown", { cancelable: true });
    Object.defineProperty(event, "key", {
      configurable: true,
      value: "Escape",
    });

    document.dispatchEvent(event);

    assert.equal(document.getElementById("overlay-escape-close"), null);
    assert.equal(event.defaultPrevented, true);
  } finally {
    restore();
  }
});

void test("overlay dialog manager reuses mounted dialog and resets after dismiss", () => {
  const { document, restore } = installFakeDom();

  try {
    const manager = createOverlayDialogManager({
      overlayId: "overlay-manager",
      cardClassName: "overlay-card",
    });

    const first = manager.ensure();
    first.mount();

    const second = manager.ensure();
    assert.equal(second, first);

    manager.dismiss();
    assert.equal(document.getElementById("overlay-manager"), null);
    assert.equal(manager.getCurrent(), null);

    const third = manager.ensure();
    assert.notEqual(third, first);
  } finally {
    restore();
  }
});
