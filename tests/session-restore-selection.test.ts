import assert from "node:assert/strict";
import { test } from "node:test";

import {
  getCrossWorkbookResumeConfirmMessage,
  getResumeTargetLabel,
} from "../src/commands/builtins/resume-target.ts";
import {
  getAdjacentTabDirectionFromShortcut,
  isCloseActiveTabShortcut,
  isCreateTabShortcut,
  isFocusInputShortcut,
  isReopenLastClosedShortcut,
  isUndoCloseTabShortcut,
  shouldAbortFromEscape,
  shouldBlurEditorFromEscape,
  shouldHandleUndoCloseTabShortcut,
} from "../src/taskpane/keyboard-shortcuts.ts";
import { RecentlyClosedStack } from "../src/taskpane/recently-closed.ts";
import {
  getRestoreCandidateSessionIds,
  shouldPersistSession,
} from "../src/taskpane/sessions.ts";

void test("known workbook restores only workbook-linked latest session", () => {
  const candidates = getRestoreCandidateSessionIds({
    workbookId: "url_sha256:workbook-a",
    workbookLatestSessionId: "session-a",
    globalLatestSessionId: "session-global",
  });

  assert.deepEqual(candidates, ["session-a"]);
});

void test("known workbook with no workbook-linked latest does not fall back to global", () => {
  const candidates = getRestoreCandidateSessionIds({
    workbookId: "url_sha256:workbook-b",
    workbookLatestSessionId: null,
    globalLatestSessionId: "session-global",
  });

  assert.deepEqual(candidates, []);
});

void test("unknown workbook falls back to global latest", () => {
  const candidates = getRestoreCandidateSessionIds({
    workbookId: null,
    workbookLatestSessionId: "session-workbook",
    globalLatestSessionId: "session-global",
  });

  assert.deepEqual(candidates, ["session-global"]);
});

void test("candidate selection normalizes empty session IDs", () => {
  const candidates = getRestoreCandidateSessionIds({
    workbookId: null,
    workbookLatestSessionId: "",
    globalLatestSessionId: "   ",
  });

  assert.deepEqual(candidates, []);
});

void test("session persistence guard allows forced saves before first assistant response", () => {
  assert.equal(
    shouldPersistSession({ firstAssistantSeen: false }),
    false,
  );
  assert.equal(
    shouldPersistSession({ firstAssistantSeen: false, force: true }),
    true,
  );
  assert.equal(
    shouldPersistSession({ firstAssistantSeen: true }),
    true,
  );
});

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

void test("recently closed stack removes duplicate-session entries by unique id", () => {
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

  const removed = stack.removeById("closed-a");
  assert.equal(removed?.id, "closed-a");
  assert.deepEqual(stack.snapshot().map((item) => item.id), ["closed-b"]);
});

void test("resume target labels and workbook confirmation copy follow selected target", () => {
  assert.equal(getResumeTargetLabel("new_tab"), "Open in new tab");
  assert.equal(getResumeTargetLabel("replace_current"), "Replace current tab");

  assert.match(
    getCrossWorkbookResumeConfirmMessage("new_tab"),
    /new tab/i,
  );
  assert.match(
    getCrossWorkbookResumeConfirmMessage("replace_current"),
    /replace the current chat/i,
  );
});

void test("Cmd/Ctrl+T detection ignores Shift/Alt-modified chords", () => {
  assert.equal(
    isCreateTabShortcut({
      key: "t",
      metaKey: true,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isCreateTabShortcut({
      key: "T",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: false,
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

void test("Cmd/Ctrl+W detection ignores Shift/Alt-modified chords", () => {
  assert.equal(
    isCloseActiveTabShortcut({
      key: "w",
      metaKey: true,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isCloseActiveTabShortcut({
      key: "W",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: false,
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

void test("Cmd/Ctrl+Z detection ignores Shift/Alt-modified chords", () => {
  assert.equal(
    isUndoCloseTabShortcut({
      key: "z",
      metaKey: true,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isUndoCloseTabShortcut({
      key: "Z",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isUndoCloseTabShortcut({
      key: "Unidentified",
      code: "KeyZ",
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isUndoCloseTabShortcut({
      key: "Unidentified",
      keyCode: 90,
      metaKey: false,
      ctrlKey: true,
      shiftKey: false,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isUndoCloseTabShortcut({
      key: "z",
      metaKey: true,
      ctrlKey: false,
      shiftKey: true,
      altKey: false,
    }),
    false,
  );

  assert.equal(
    isUndoCloseTabShortcut({
      key: "z",
      metaKey: true,
      ctrlKey: false,
      shiftKey: false,
      altKey: true,
    }),
    false,
  );
});

void test("Undo-close handling supports empty input + non-input contexts while preserving text undo", () => {
  assert.equal(
    shouldHandleUndoCloseTabShortcut({
      canUndoCloseTab: false,
      isTextEntry: false,
      textEntryHasContent: false,
      actionToastVisible: true,
    }),
    false,
  );

  assert.equal(
    shouldHandleUndoCloseTabShortcut({
      canUndoCloseTab: true,
      isTextEntry: false,
      textEntryHasContent: false,
      actionToastVisible: false,
    }),
    true,
  );

  assert.equal(
    shouldHandleUndoCloseTabShortcut({
      canUndoCloseTab: true,
      isTextEntry: true,
      textEntryHasContent: false,
      actionToastVisible: false,
    }),
    true,
  );

  assert.equal(
    shouldHandleUndoCloseTabShortcut({
      canUndoCloseTab: true,
      isTextEntry: true,
      textEntryHasContent: true,
      actionToastVisible: false,
    }),
    false,
  );

  assert.equal(
    shouldHandleUndoCloseTabShortcut({
      canUndoCloseTab: true,
      isTextEntry: true,
      textEntryHasContent: true,
      actionToastVisible: true,
    }),
    true,
  );
});

void test("Cmd/Ctrl+Shift+T detection ignores Alt-modified chords", () => {
  assert.equal(
    isReopenLastClosedShortcut({
      key: "t",
      metaKey: true,
      ctrlKey: false,
      shiftKey: true,
      altKey: false,
    }),
    true,
  );

  assert.equal(
    isReopenLastClosedShortcut({
      key: "T",
      metaKey: false,
      ctrlKey: true,
      shiftKey: true,
      altKey: false,
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

void test("Arrow and fallback tab-switch shortcuts resolve expected direction", () => {
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
      key: "[",
      repeat: false,
      metaKey: true,
      ctrlKey: false,
      shiftKey: true,
      altKey: false,
    }),
    -1,
  );

  assert.equal(
    getAdjacentTabDirectionFromShortcut({
      key: "]",
      repeat: false,
      metaKey: true,
      ctrlKey: false,
      shiftKey: true,
      altKey: false,
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
    -1,
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
    1,
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
