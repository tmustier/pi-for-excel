import assert from "node:assert/strict";
import { test } from "node:test";

import { PI_REQUEST_INPUT_FOCUS_EVENT } from "../src/ui/input-focus.ts";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayDialogManager,
} from "../src/ui/overlay-dialog.ts";
import {
  CONFIRM_DIALOG_OVERLAY_ID,
  TEXT_INPUT_DIALOG_OVERLAY_ID,
} from "../src/ui/overlay-ids.ts";
import { requestConfirmationDialog } from "../src/ui/confirm-dialog.ts";
import { requestTextInputDialog } from "../src/ui/text-input-dialog.ts";
import { installFakeDom } from "./fake-dom.test.ts";

function findButtonByText(root: HTMLElement, text: string): HTMLElement | null {
  const buttons = root.querySelectorAll("button");

  for (const candidate of buttons) {
    if (!(candidate instanceof HTMLElement) || candidate.tagName !== "BUTTON") {
      continue;
    }

    if (candidate.textContent === text) {
      return candidate;
    }
  }

  return null;
}

function findFirstInput(root: HTMLElement): HTMLInputElement | null {
  const queue: Element[] = Array.from(root.children);

  while (queue.length > 0) {
    const next = queue.shift();
    if (!next) {
      continue;
    }

    if (next instanceof HTMLInputElement) {
      return next;
    }

    queue.push(...Array.from(next.children));
  }

  return null;
}

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

void test("confirmation dialog resolves true when confirm button is clicked", async () => {
  const { document, restore } = installFakeDom();

  try {
    const pendingApproval = requestConfirmationDialog({
      title: "Allow workbook mutation in Confirm mode?",
      message: "Tool: write_cells",
      confirmLabel: "Allow once",
      restoreFocusOnClose: false,
    });

    const overlay = document.getElementById(CONFIRM_DIALOG_OVERLAY_ID);
    assert.ok(overlay);

    const approveButton = findButtonByText(overlay, "Allow once");

    assert.ok(approveButton);
    if (!approveButton) {
      throw new Error("Approval button not found");
    }

    approveButton.dispatchEvent(new Event("click"));

    const approved = await pendingApproval;
    assert.equal(approved, true);
    assert.equal(document.getElementById(CONFIRM_DIALOG_OVERLAY_ID), null);
  } finally {
    restore();
  }
});

void test("text input dialog resolves entered value on confirm", async () => {
  const { document, restore } = installFakeDom();

  try {
    const pendingResult = requestTextInputDialog({
      title: "Rename file",
      initialValue: "notes.md",
      confirmLabel: "Rename",
      cancelLabel: "Cancel",
      restoreFocusOnClose: false,
    });

    const overlay = document.getElementById(TEXT_INPUT_DIALOG_OVERLAY_ID);
    assert.ok(overlay);

    const input = findFirstInput(overlay);
    assert.ok(input);
    if (!input) {
      throw new Error("Text input not found");
    }

    input.value = "notes-renamed.md";

    const confirmButton = findButtonByText(overlay, "Rename");
    assert.ok(confirmButton);
    if (!confirmButton) {
      throw new Error("Confirm button not found");
    }

    confirmButton.dispatchEvent(new Event("click"));

    const value = await pendingResult;
    assert.equal(value, "notes-renamed.md");
    assert.equal(document.getElementById(TEXT_INPUT_DIALOG_OVERLAY_ID), null);
  } finally {
    restore();
  }
});

void test("text input dialog resolves null on cancel", async () => {
  const { document, restore } = installFakeDom();

  try {
    const pendingResult = requestTextInputDialog({
      title: "Rename file",
      confirmLabel: "Rename",
      cancelLabel: "Cancel",
      restoreFocusOnClose: false,
    });

    const overlay = document.getElementById(TEXT_INPUT_DIALOG_OVERLAY_ID);
    assert.ok(overlay);

    const cancelButton = findButtonByText(overlay, "Cancel");
    assert.ok(cancelButton);
    if (!cancelButton) {
      throw new Error("Cancel button not found");
    }

    cancelButton.dispatchEvent(new Event("click"));

    const value = await pendingResult;
    assert.equal(value, null);
    assert.equal(document.getElementById(TEXT_INPUT_DIALOG_OVERLAY_ID), null);
  } finally {
    restore();
  }
});
