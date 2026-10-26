// Component-logic contracts only; this suite does not claim native browser behavior.
import assert from "node:assert/strict";
import { test } from "node:test";

import { requestApproval } from "../src/ui/approval-dialog.ts";
import { TOOL_APPROVAL_OVERLAY_ID } from "../src/ui/overlay-ids.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

async function flushAsync(): Promise<void> {
  for (let index = 0; index < 4; index += 1) await Promise.resolve();
}

function clickButton(label: string): void {
  const overlay = document.getElementById(TOOL_APPROVAL_OVERLAY_ID);
  assert.ok(overlay, "an approval is on screen");
  const button = Array.from(overlay.querySelectorAll("button")).find((candidate) => candidate.textContent === label);
  assert.ok(button, `button ${label}`);
  button.dispatchEvent(new Event("click"));
}

function onScreenTitle(): string | undefined {
  const overlay = document.getElementById(TOOL_APPROVAL_OVERLAY_ID);
  if (!overlay) return undefined;
  const visit = (node: Element): string | undefined => {
    if (node.tagName === "H2") return node.textContent ?? undefined;
    for (const child of Array.from(node.children)) {
      const found = visit(child);
      if (found) return found;
    }
    return undefined;
  };
  return visit(overlay);
}

void test("concurrent approvals queue instead of declining each other", async () => {
  const { restore } = installFakeDom();
  try {
    const composer = document.createElement("div");
    document.body.appendChild(composer);

    const outcomes: string[] = [];
    const first = requestApproval({ title: "Write D1", confirmLabel: "Allow" }).then((ok) => outcomes.push(`D1:${ok}`));
    const second = requestApproval({ title: "Write D2", confirmLabel: "Allow" }).then((ok) => outcomes.push(`D2:${ok}`));
    await flushAsync();

    // The second request neither replaced nor declined the first.
    assert.deepEqual(outcomes, []);
    assert.equal(onScreenTitle(), "Write D1");
    // Nothing behind the question can be used until it is answered.
    assert.equal((composer as HTMLElement).inert, true);

    clickButton("Decline");
    await flushAsync();
    assert.equal(onScreenTitle(), "Write D2");

    clickButton("Allow");
    await Promise.all([first, second]);
    assert.deepEqual(outcomes, ["D1:false", "D2:true"]);
    assert.equal(document.getElementById(TOOL_APPROVAL_OVERLAY_ID), null);
    assert.equal((composer as HTMLElement).inert, false);
  } finally {
    restore();
  }
});

void test("stopping the turn withdraws its approval without touching the others", async () => {
  const { restore } = installFakeDom();
  try {
    const stopped = new AbortController();
    const outcomes: string[] = [];
    const shown = requestApproval({ title: "Shown", confirmLabel: "Allow" }).then((ok) => outcomes.push(`shown:${ok}`));
    const queued = requestApproval({ title: "Queued", confirmLabel: "Allow", signal: stopped.signal })
      .then((ok) => outcomes.push(`queued:${ok}`));
    await flushAsync();

    stopped.abort();
    await queued;
    assert.deepEqual(outcomes, ["queued:false"]);
    assert.equal(onScreenTitle(), "Shown");

    clickButton("Allow");
    await shown;
    assert.deepEqual(outcomes, ["queued:false", "shown:true"]);
    await flushAsync();
    assert.equal(document.getElementById(TOOL_APPROVAL_OVERLAY_ID), null);
  } finally {
    restore();
  }
});
