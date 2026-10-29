import assert from "node:assert/strict";
import { test } from "node:test";

import { getInlineApprovals, registerInlineApprovalHost, requestApproval } from "../src/ui/approval-dialog.ts";
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
    const announcer = document.createElement("div");
    announcer.setAttribute("aria-live", "polite");
    document.body.appendChild(announcer);

    const outcomes: string[] = [];
    const first = requestApproval({ title: "Write D1", confirmLabel: "Allow" }).then((ok) => outcomes.push(`D1:${ok}`));
    const second = requestApproval({ title: "Write D2", confirmLabel: "Allow" }).then((ok) => outcomes.push(`D2:${ok}`));
    await flushAsync();

    assert.deepEqual(outcomes, []);
    assert.equal(onScreenTitle(), "Write D1");
    assert.equal((composer as HTMLElement).inert, true);
    assert.notEqual((announcer as HTMLElement).inert, true);

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

void test("Stop Pi declines the queue and stops the turns", async () => {
  const { restore } = installFakeDom();
  try {
    const stopped: string[] = [];
    const stopA = () => stopped.push("A");
    const stopB = () => stopped.push("B");
    const outcomes: string[] = [];
    const first = requestApproval({ title: "A1", confirmLabel: "Allow", onStop: stopA }).then((ok) => outcomes.push(`A1:${ok}`));
    const second = requestApproval({ title: "A2", confirmLabel: "Allow", onStop: stopA }).then((ok) => outcomes.push(`A2:${ok}`));
    const third = requestApproval({ title: "B1", confirmLabel: "Allow", onStop: stopB }).then((ok) => outcomes.push(`B1:${ok}`));
    await flushAsync();

    clickButton("Stop Pi");
    await Promise.all([first, second, third]);
    assert.deepEqual(outcomes.sort(), ["A1:false", "A2:false", "B1:false"]);
    assert.deepEqual(stopped.sort(), ["A", "B"]);
    await flushAsync();
    assert.equal(document.getElementById(TOOL_APPROVAL_OVERLAY_ID), null);
  } finally {
    restore();
  }
});

void test("turn consent covers only eligible queued and future changes in that run", async () => {
  const { restore } = installFakeDom();
  const turn = new AbortController();
  const otherTurn = new AbortController();
  try {
    const request = { title: "Write", confirmLabel: "Allow", signal: turn.signal, allowForTurn: true };
    const first = requestApproval(request);
    const queued = requestApproval(request);
    const destructive = requestApproval({ ...request, title: "Delete chart", allowForTurn: false });
    clickButton("Allow for the rest of this turn");
    assert.equal(await first, true);
    assert.equal(await queued, true);
    await flushAsync();
    assert.equal(onScreenTitle(), "Delete chart");
    clickButton("Decline");
    assert.equal(await destructive, false);
    assert.equal(await requestApproval(request), true);
    const next = requestApproval({ ...request, title: "New run", signal: otherTurn.signal });
    assert.equal(onScreenTitle(), "New run");
    otherTurn.abort();
    assert.equal(await next, false);
    turn.abort();
    assert.equal(await requestApproval(request), false);
  } finally {
    turn.abort();
    otherTurn.abort();
    await flushAsync();
    restore();
  }
});

void test("secondary approval waits for preparation and cannot approve after Stop", async () => {
  const { restore } = installFakeDom();
  const turn = new AbortController();
  try {
    let finish = (): void => {};
    const preparation = new Promise<void>((resolve) => { finish = resolve; });
    const pending = requestApproval({
      title: "Delete", confirmLabel: "Allow", signal: turn.signal,
      secondaryApproval: { label: "Backup first", onApprove: () => preparation },
    });
    clickButton("Backup first");
    await flushAsync();
    assert.equal(onScreenTitle(), "Delete");
    turn.abort();
    assert.equal(await pending, false);
    finish();
    await flushAsync();
    assert.equal(document.getElementById(TOOL_APPROVAL_OVERLAY_ID), null);
  } finally {
    turn.abort();
    await flushAsync();
    restore();
  }
});

void test("a question about a call the transcript shows is asked in its row, and returns to a dialog when the transcript goes away", async () => {
  const { restore } = installFakeDom();
  try {
    const shown = new Set(["call-1"]);
    let changes = 0;
    const host = registerInlineApprovalHost({ claims: (id) => shown.has(id), onChange: () => { changes += 1; } });

    const outcomes: string[] = [];
    const inlineAsk = requestApproval({ title: "Write D1", toolCallId: "call-1", confirmLabel: "Write" }).then((ok) => outcomes.push(`D1:${ok}`));
    const dialogAsk = requestApproval({ title: "Write D2", toolCallId: "call-2", confirmLabel: "Write" }).then((ok) => outcomes.push(`D2:${ok}`));
    await flushAsync();

    assert.deepEqual(getInlineApprovals().map((approval) => approval.toolCallId), ["call-1"]);
    assert.equal(changes, 1);
    assert.equal(onScreenTitle(), "Write D2");

    getInlineApprovals()[0]?.answer(true);
    await flushAsync();
    assert.deepEqual(outcomes, ["D1:true"]);
    assert.deepEqual(getInlineApprovals(), []);

    clickButton("Decline");
    await Promise.all([inlineAsk, dialogAsk]);
    assert.deepEqual(outcomes, ["D1:true", "D2:false"]);

    const orphan = requestApproval({ title: "Write D3", toolCallId: "call-1", confirmLabel: "Write" }).then((ok) => outcomes.push(`D3:${ok}`));
    await flushAsync();
    assert.equal(getInlineApprovals().length, 1);
    host.unregister();
    await flushAsync();
    assert.deepEqual(getInlineApprovals(), []);
    assert.equal(onScreenTitle(), "Write D3");
    clickButton("Write");
    await orphan;
    assert.deepEqual(outcomes, ["D1:true", "D2:false", "D3:true"]);
  } finally {
    restore();
  }
});

void test("an inline question is withdrawn when its turn stops, and turn consent covers later inline questions", async () => {
  const { restore } = installFakeDom();
  try {
    const host = registerInlineApprovalHost({ claims: () => true, onChange: () => {} });
    const controller = new AbortController();
    const outcomes: string[] = [];
    const first = requestApproval({ title: "Write D1", toolCallId: "call-1", confirmLabel: "Write", allowForTurn: true, signal: controller.signal }).then((ok) => outcomes.push(`D1:${ok}`));
    await flushAsync();
    getInlineApprovals()[0]?.allowForTurn();
    await first;
    assert.deepEqual(outcomes, ["D1:true"]);

    const second = await requestApproval({ title: "Write D2", toolCallId: "call-2", confirmLabel: "Write", allowForTurn: true, signal: controller.signal });
    assert.equal(second, true);

    const third = requestApproval({ title: "Delete sheet", toolCallId: "call-3", confirmLabel: "Delete", signal: controller.signal }).then((ok) => outcomes.push(`D3:${ok}`));
    await flushAsync();
    assert.equal(getInlineApprovals().length, 1);
    controller.abort();
    await third;
    assert.deepEqual(outcomes, ["D1:true", "D3:false"]);
    assert.deepEqual(getInlineApprovals(), []);
    host.unregister();
  } finally {
    restore();
  }
});

void test("changing a host's claims re-asks its unanswered question, removing every abort listener once", async (context) => {
  const { restore } = installFakeDom();
  const controller = new AbortController();
  const added = context.mock.method(controller.signal, "addEventListener");
  const removed = context.mock.method(controller.signal, "removeEventListener");
  let visible = true;
  const host = registerInlineApprovalHost({ claims: () => visible, onChange() {} });
  try {
    const outcomes: boolean[] = [];
    const pending = requestApproval({ title: "Write", toolCallId: "call", confirmLabel: "Write", signal: controller.signal })
      .then((approved) => outcomes.push(approved));
    const stale = getInlineApprovals()[0];
    assert.ok(stale);
    visible = false;
    host.refresh();
    assert.deepEqual(getInlineApprovals(), []);
    assert.equal(onScreenTitle(), "Write");
    assert.deepEqual(outcomes, [], "changing tabs must not answer the question");
    stale.answer(true);
    assert.deepEqual(outcomes, [], "the detached row no longer owns the answer");
    controller.abort();
    await pending;
    assert.deepEqual(outcomes, [false]);
    assert.equal(document.getElementById(TOOL_APPROVAL_OVERLAY_ID), null);
    assert.equal(added.mock.calls.length, 2);
    assert.equal(removed.mock.calls.length, 2);
    assert.deepEqual(removed.mock.calls.map((call) => call.arguments[1]), added.mock.calls.map((call) => call.arguments[1]));
  } finally {
    host.unregister();
    controller.abort();
    await flushAsync();
    restore();
  }
});

void test("another host can keep an inline question visible until the last host disconnects", async () => {
  const { restore } = installFakeDom();
  const first = registerInlineApprovalHost({ claims: () => true, onChange() {} });
  const second = registerInlineApprovalHost({ claims: () => true, onChange() {} });
  try {
    const pending = requestApproval({ title: "Write", toolCallId: "shared", confirmLabel: "Write" });
    first.unregister();
    assert.equal(getInlineApprovals().length, 1);
    assert.equal(onScreenTitle(), undefined);
    second.unregister();
    assert.deepEqual(getInlineApprovals(), []);
    assert.equal(onScreenTitle(), "Write");
    clickButton("Write");
    assert.equal(await pending, true);
  } finally {
    first.unregister();
    second.unregister();
    restore();
  }
});

void test("Stop Pi stops turns with inline questions as well as the dialog queue", async () => {
  const { restore } = installFakeDom();
  const host = registerInlineApprovalHost({ claims: (id) => id === "inline", onChange() {} });
  try {
    let stopped = 0;
    const inline = requestApproval({ title: "Inline", toolCallId: "inline", confirmLabel: "Write", onStop: () => { stopped++; } });
    const dialog = requestApproval({ title: "Dialog", confirmLabel: "Write", onStop() {} });
    clickButton("Stop Pi");
    assert.deepEqual(await Promise.all([inline, dialog]), [false, false]);
    assert.equal(stopped, 1);
    assert.deepEqual(getInlineApprovals(), []);
  } finally {
    host.unregister();
    restore();
  }
});

void test("turn consent answers already waiting inline changes without granting destructive ones", async () => {
  const { restore } = installFakeDom();
  const host = registerInlineApprovalHost({ claims: () => true, onChange() {} });
  const controller = new AbortController();
  try {
    const request = { title: "Write", confirmLabel: "Write", allowForTurn: true, signal: controller.signal };
    const first = requestApproval({ ...request, toolCallId: "a" });
    const second = requestApproval({ ...request, toolCallId: "b" });
    const destructive = requestApproval({ ...request, toolCallId: "c", allowForTurn: false });
    const approval = getInlineApprovals()[0];
    assert.ok(approval);
    approval.allowForTurn();
    assert.deepEqual(await Promise.all([first, second]), [true, true]);
    assert.deepEqual(getInlineApprovals().map((item) => item.toolCallId), ["c"]);
    controller.abort();
    assert.equal(await destructive, false);
  } finally {
    controller.abort();
    host.unregister();
    restore();
  }
});

void test("inline preparation runs once and cannot approve after the host has released it", async () => {
  const { restore } = installFakeDom();
  const host = registerInlineApprovalHost({ claims: () => true, onChange() {} });
  try {
    let finish = (): void => {};
    let preparations = 0;
    const prepared = new Promise<void>((resolve) => { finish = resolve; });
    const pending = requestApproval({ title: "Delete", toolCallId: "call", confirmLabel: "Delete", secondaryApproval: {
      label: "Backup first", onApprove: () => { preparations++; return prepared; },
    } });
    const approval = getInlineApprovals()[0];
    assert.ok(approval);
    const first = approval.prepareAndApprove();
    await approval.prepareAndApprove();
    assert.equal(preparations, 1);
    host.unregister();
    finish();
    await first;
    assert.equal(onScreenTitle(), "Delete");
    clickButton("Decline");
    assert.equal(await pending, false);
  } finally {
    host.unregister();
    restore();
  }
});
