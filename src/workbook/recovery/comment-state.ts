/** Comment-thread capture/apply for workbook recovery snapshots. */

import { excelRun, getRange } from "../../excel/helpers.js";
import { runMutation } from "../../tools/mutation/run-mutation.js";
import { firstCellAddress } from "./address.js";
import { cloneRecoveryCommentThreadState } from "./clone.js";
import type { RecoveryCommentThreadState } from "./types.js";

function emptyCommentThreadState(): RecoveryCommentThreadState {
  return {
    exists: false,
    content: "",
    resolved: false,
    replies: [],
  };
}

interface LoadedCommentThread {
  state: RecoveryCommentThreadState;
  comment: Excel.Comment | null;
}

async function loadCommentThreadInRange(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
): Promise<LoadedCommentThread> {
  range.load("address");

  const commentCollection = sheet.comments;
  commentCollection.load("items");
  await context.sync();

  if (commentCollection.items.length === 0) {
    return {
      state: emptyCommentThreadState(),
      comment: null,
    };
  }

  const entries = commentCollection.items.map((comment) => {
    comment.load("content,resolved");
    comment.replies.load("items");
    const location = comment.getLocation();
    location.load("address");
    return { comment, location };
  });

  await context.sync();

  const targetCell = firstCellAddress(range.address).toUpperCase();
  let match: { comment: Excel.Comment } | null = null;

  for (const entry of entries) {
    if (firstCellAddress(entry.location.address).toUpperCase() === targetCell) {
      match = { comment: entry.comment };
      break;
    }
  }

  if (!match) {
    return {
      state: emptyCommentThreadState(),
      comment: null,
    };
  }

  for (const reply of match.comment.replies.items) {
    reply.load("content");
  }

  if (match.comment.replies.items.length > 0) {
    await context.sync();
  }

  return {
    state: {
      exists: true,
      content: match.comment.content,
      resolved: match.comment.resolved,
      replies: match.comment.replies.items.map((reply) => reply.content),
    },
    comment: match.comment,
  };
}

export async function captureCommentThreadState(address: string): Promise<RecoveryCommentThreadState> {
  return excelRun<RecoveryCommentThreadState>(async (context) => {
    const { sheet, range } = getRange(context, address);
    const loaded = await loadCommentThreadInRange(context, sheet, range);
    return cloneRecoveryCommentThreadState(loaded.state);
  });
}

export async function applyCommentThreadState(
  address: string,
  targetState: RecoveryCommentThreadState,
): Promise<RecoveryCommentThreadState> {
  return runMutation<RecoveryCommentThreadState>(async (context) => {
    const { sheet, range } = getRange(context, address);
    const loaded = await loadCommentThreadInRange(context, sheet, range);

    if (!targetState.exists) {
      if (loaded.comment) {
        loaded.comment.delete();
        await context.sync();
      }

      return cloneRecoveryCommentThreadState(loaded.state);
    }

    if (loaded.comment) {
      loaded.comment.delete();
      await context.sync();
    }

    const restoredComment = sheet.comments.add(range, targetState.content);

    for (const reply of targetState.replies) {
      restoredComment.replies.add(reply);
    }

    restoredComment.resolved = targetState.resolved;
    await context.sync();

    return cloneRecoveryCommentThreadState(loaded.state);
  });
}
