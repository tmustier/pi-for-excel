import { copyJson, type Context } from "@earendil-works/chord";
import { InboxDoc, type Conversation, type InboxItem, type SubmissionId, type UserInput } from "@earendil-works/pi-durable";

export type QueuedInput = Exclude<InboxItem, { mode: "write" }>;

/** Change an admitted input in place, preserving its ID, mode, and queue position.
 * The inbox is the sole owner of queued content; SubmissionRecord stores only
 * identity and lifecycle. Checking and changing the item in one Harness commit
 * serializes with placement. A lost race never becomes a second submission.
 */
export function editQueuedInput(
  conversation: Conversation,
  id: SubmissionId,
  content: UserInput,
  context: Context,
): Promise<boolean> {
  return conversation.commit(async (tx) => {
    const inbox = await tx.doc(InboxDoc, conversation.id);
    const item = inbox.items.find((candidate) => candidate.id === id);
    if (!item || item.mode === "write") return false;
    item.content = content;
    return true;
  }, context);
}

/** Withdraw the inputs actually still queued, atomically with their receipts.
 * Writes (including a pending compaction/reset) keep their original positions.
 * Return detached content only after persistence succeeds, for Alt+Up editing.
 */
export function takeQueuedInputs(conversation: Conversation, context: Context): Promise<QueuedInput[]> {
  return conversation.commit(async (tx) => {
    const inbox = await tx.doc(InboxDoc, conversation.id);
    const taken: QueuedInput[] = [];
    for (let index = inbox.items.length - 1; index >= 0; index--) {
      const item = inbox.items[index];
      if (!item || item.mode === "write") continue;
      // The upstream typed InboxDoc supplies this shape. copyJson detaches its
      // revocable draft without changing the JSON representation.
      taken.unshift(copyJson(item) as QueuedInput);
      tx.settleSubmission(item.id, { status: "unanswered", reason: "aborted" });
      inbox.items.splice(index, 1);
    }
    return taken;
  }, context);
}
