import assert from "node:assert/strict";
import { test } from "node:test";


import {
  getAdjacentTabDirectionFromShortcut,
  isCloseActiveTabShortcut,
  isCreateTabShortcut,
  isFocusInputShortcut,
  isReopenLastClosedShortcut,
  isRestoreQueuedMessagesShortcut,
  shouldAbortFromEscape,
  shouldBlurEditorFromEscape,
} from "../src/taskpane/keyboard-shortcuts.ts";
import { RecentlyClosedStack } from "../src/taskpane/recently-closed.ts";

void test("recently closed stack reopens newest first and enforces max size", () => {
  const stack = new RecentlyClosedStack(2);

  stack.push({
    id: "closed-1",
    sessionId: "session-1",
    title: "One",
    closedAt: "2026-02-11T10:00:00.000Z",
    workbookId: "wb-a",
  });
  stack.push({
    id: "closed-2",
    sessionId: "session-2",
    title: "Two",
    closedAt: "2026-02-11T10:01:00.000Z",
    workbookId: "wb-a",
  });
  stack.push({
    id: "closed-3",
    sessionId: "session-3",
    title: "Three",
    closedAt: "2026-02-11T10:02:00.000Z",
    workbookId: "wb-b",
  });

  assert.equal(stack.size, 2);
  assert.deepEqual(
    stack.snapshot().map((item) => item.sessionId),
    ["session-3", "session-2"],
  );

  assert.equal(stack.popMostRecent()?.sessionId, "session-3");
  assert.equal(stack.popMostRecent()?.sessionId, "session-2");
  assert.equal(stack.popMostRecent(), null);
});

void test("recently closed stack keeps one entry per session, the most recent close", () => {
  const stack = new RecentlyClosedStack(5);

  stack.push({
    id: "closed-a",
    sessionId: "session-dup",
    title: "First",
    closedAt: "2026-02-11T10:00:00.000Z",
    workbookId: "wb-a",
  });
  stack.push({
    id: "closed-b",
    sessionId: "session-dup",
    title: "Second",
    closedAt: "2026-02-11T10:05:00.000Z",
    workbookId: "wb-a",
  });

  assert.deepEqual(stack.snapshot().map((item) => item.id), ["closed-b"]);
  assert.equal(stack.removeById("closed-a"), null);
});

void test("Ctrl+Alt+N creates tabs without taking browser shortcuts", () => {
  assert.equal(
    isCreateTabShortcut({
      key: "n",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isCreateTabShortcut({
      key: "N",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isCreateTabShortcut({
      key: "t",
      metaKey: true,
      ctrlKey: false,
      shiftKey: true,
      altKey: false,
    }),
    false,
  );

  assert.equal(
    isCreateTabShortcut({
      key: "t",
      metaKey: true,
      ctrlKey: false,
      shiftKey: false,
      altKey: true,
    }),
    false,
  );
});

void test("Ctrl+Alt+W closes tabs without taking browser shortcuts", () => {
  assert.equal(
    isCloseActiveTabShortcut({
      key: "w",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isCloseActiveTabShortcut({
      key: "W",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isCloseActiveTabShortcut({
      key: "w",
      metaKey: true,
      ctrlKey: false,
      shiftKey: true,
      altKey: false,
    }),
    false,
  );

  assert.equal(
    isCloseActiveTabShortcut({
      key: "w",
      metaKey: true,
      ctrlKey: false,
      shiftKey: false,
      altKey: true,
    }),
    false,
  );
});

void test("Ctrl+Alt+Shift+T reopens tabs without taking browser shortcuts", () => {
  assert.equal(
    isReopenLastClosedShortcut({
      key: "t",
      metaKey: false,
      ctrlKey: true,
      shiftKey: true,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isReopenLastClosedShortcut({
      key: "T",
      metaKey: false,
      ctrlKey: true,
      shiftKey: true,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isReopenLastClosedShortcut({
      key: "t",
      metaKey: true,
      ctrlKey: false,
      shiftKey: true,
      altKey: true,
    }),
    false,
  );
});

void test("F2 focuses chat input only without modifiers", () => {
  assert.equal(
    isFocusInputShortcut({
      key: "F2",
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isFocusInputShortcut({
      key: "F2",
      metaKey: false,
      ctrlKey: false,
      shiftKey: true,
      altKey: false,
    }),
    false,
  );

  assert.equal(
    isFocusInputShortcut({
      key: "f2",
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    false,
  );
});

void test("Arrow and Ctrl+Alt tab-switch shortcuts resolve expected direction", () => {
  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "ArrowLeft",
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    -1,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "ArrowRight",
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    1,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "Right",
      code: "ArrowRight",
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    1,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "Unidentified",
      keyCode: 37,
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    -1,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "ArrowLeft",
      repeat: false,
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: true,
    }),
    -1,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "ArrowRight",
      repeat: false,
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: true,
    }),
    1,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "PageUp",
      repeat: false,
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: false,
    }),
    null,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "PageDown",
      repeat: false,
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: false,
    }),
    null,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "ArrowLeft",
      repeat: true,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    null,
  );
});

void test("Alt+Up detection matches host fallbacks and rejects modified chords", () => {
  assert.equal(
    isRestoreQueuedMessagesShortcut({
      key: "ArrowUp",
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isRestoreQueuedMessagesShortcut({
      key: "PageUp",
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isRestoreQueuedMessagesShortcut({
      key: "Unidentified",
      code: "PageUp",
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: true,
    }),
    true,
  );

  assert.equal(
    isRestoreQueuedMessagesShortcut({
      key: "ArrowUp",
      repeat: false,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    false,
  );

  assert.equal(
    isRestoreQueuedMessagesShortcut({
      key: "ArrowUp",
      repeat: false,
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: true,
    }),
    false,
  );

  assert.equal(
    isRestoreQueuedMessagesShortcut({
      key: "ArrowUp",
      repeat: true,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: true,
    }),
    false,
  );
});

void test("Escape exits editor focus only when not streaming", () => {
  assert.equal(
    shouldBlurEditorFromEscape({
      key: "Escape",
      isInEditor: true,
      isStreaming: false,
    }),
    true,
  );

  assert.equal(
    shouldBlurEditorFromEscape({
      key: "Esc",
      isInEditor: true,
      isStreaming: false,
    }),
    true,
  );

  assert.equal(
    shouldBlurEditorFromEscape({
      key: "Escape",
      isInEditor: true,
      isStreaming: true,
    }),
    false,
  );

  assert.equal(
    shouldBlurEditorFromEscape({
      key: "Enter",
      isInEditor: true,
      isStreaming: false,
    }),
    false,
  );
});

void test("Escape abort is suppressed when overlay UI claims Escape", () => {
  assert.equal(
    shouldAbortFromEscape({
      isStreaming: true,
      hasAgent: true,
      escapeClaimedByOverlay: true,
    }),
    false,
  );

  assert.equal(
    shouldAbortFromEscape({
      isStreaming: true,
      hasAgent: true,
      escapeClaimedByOverlay: false,
    }),
    true,
  );

  assert.equal(
    shouldAbortFromEscape({
      isStreaming: false,
      hasAgent: true,
      escapeClaimedByOverlay: false,
    }),
    false,
  );
});
