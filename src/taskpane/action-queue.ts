/**
 * Per-tab ordered action queue and turn lifecycle.
 *
 * Needed because some actions (notably `/compact`) run outside the Agent loop
 * (they call `agent.streamFunction(...)` directly) and therefore don't set
 * `agent.state.isStreaming`. Without a queue, user input can be lost when
 * compaction rewrites the message list.
 *
 * The queue owns, for its own agent (never "whichever tab is active"):
 * - queued prompts and commands, run one at a time when the tab is idle;
 *   a queued `/compact` compacts this tab's agent
 * - automatic context protection (#566): pre-prompt and mid-turn
 *   auto-compaction, and one compact-and-retry after a context overflow
 * - bounded retries of transient provider failures (rate limits, 5xx,
 *   dropped connections), shown as the tab's activity line
 * - Stop: one behavior whatever the trigger. The run and any compaction or
 *   retry wait are cancelled, and every message the user queued that has not
 *   been delivered stays above the composer as its own "Not sent" row, never
 *   sent automatically and never merged with other messages. Queued commands
 *   are cancelled (and reported), since they are work Pi would have done,
 *   not text the user wrote.
 * - when a turn ends, steers and follow-ups it never delivered become "Not
 *   sent" rows the same way instead of leaking into the next turn.
 * - each queue row's Edit (puts that one row's text back in the composer)
 *   and Remove; Alt+Up edits the last row.
 */

import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";
import { isRetryableAssistantError, retryDelayMs, type AssistantMessage } from "@earendil-works/pi-ai";
import { t } from "../language/index.js";

import { commandRegistry } from "../commands/types.js";
import {
  maybeAutoCompactBeforeContinuation,
  maybeAutoCompactBeforePrompt,
} from "../compaction/auto-compaction.js";
import {
  findTrailingContextOverflowError,
  recoverFromContextOverflow,
} from "../compaction/overflow-recovery.js";
import type { QueueDisplay, QueueRow } from "./queue-display.js";
import type { TabState } from "./tab-state.js";

export type QueuedAction =
  | { type: "prompt"; text: string }
  | { type: "command"; name: string; args: string };

/** What a stop or a take-back removed from the queue. */
export interface TakenBack {
  /** Steer, follow-up and queued prompt texts, oldest first. */
  texts: string[];
  /** Labels of queued commands that were cancelled (e.g. "/compact"). */
  cancelledCommands: string[];
}

export interface TurnOutcome {
  /** True when the run ended in a context overflow after compact-and-retry. */
  overflowAfterRecovery: boolean;
}

export interface ActionQueue {
  enqueuePrompt: (text: string) => void;
  enqueueCommand: (name: string, args: string) => void;
  /** Alt+Up: puts the last row holding the user's text back in the composer; false when there is none. */
  editLast: () => boolean;
  /** Stops the tab: cancels the run, compaction or retry wait; queued messages become "Not sent" rows. */
  stop: () => TakenBack;
  isBusy: () => boolean;
  shutdown: () => void;
}

type ActionQueueDisplay = Pick<
  QueueDisplay,
  "setActionQueue" | "drainQueuedMessages" | "removeQueuedMessage" | "setRowHandlers" | "lastEditableRow"
>;

type ActivityHost = Pick<TabState, "setActivity" | "setBusy" | "returnToDraft" | "held" | "setHeld" | "hold">;

/** Transient provider failures are retried up to this many times (2s, 4s, 8s). */
export const PROVIDER_RETRY_LIMIT = 3;
const PROVIDER_RETRY_BASE_DELAY_MS = 2_000;

class StoppedError extends Error {
  constructor() {
    super("Stopped");
  }
}

function waitUnlessStopped(ms: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal.aborted) {
      reject(new StoppedError());
      return;
    }
    const onAbort = () => {
      clearTimeout(timer);
      reject(new StoppedError());
    };
    const timer = setTimeout(() => {
      signal.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal.addEventListener("abort", onAbort, { once: true });
  });
}

function lastAssistant(messages: readonly AgentMessage[]): AssistantMessage | null {
  const last = messages[messages.length - 1];
  return last && last.role === "assistant" ? last : null;
}

export function createActionQueue(opts: {
  agent: Agent;
  tab: ActivityHost;
  queueDisplay: ActionQueueDisplay;
  autoCompactEnabled: boolean;
  /** Runs compaction for this queue's agent (not the active tab's agent). */
  runCompact: (args: string, signal: AbortSignal) => Promise<void>;
  /** Called after each queued prompt's turn has fully settled (retries and recovery included). */
  onTurnSettled?: (outcome: TurnOutcome) => void;
  /** Called when a turn ended with steers/follow-ups it never delivered (now "Not sent" rows). */
  onLeftUnsent?: (count: number) => void;
}): ActionQueue {
  const { agent, tab, queueDisplay, autoCompactEnabled } = opts;

  const actions: QueuedAction[] = [];
  let running = false;
  let closed = false;
  /** Aborted by stop(); renewed for each queue item so a stop only cancels what is running. */
  let stopController = new AbortController();

  const syncDisplay = () => {
    queueDisplay.setActionQueue(
      actions.map((a) => {
        if (a.type === "prompt") return { type: "prompt", label: t("action-queue.queued"), text: a.text };
        return { type: "command", label: `/${a.name}`, text: a.args ? a.args : "" };
      }),
    );
  };

  const isBusy = () => running || agent.state.isStreaming;

  const runCompactWithIndicator = async (args: string, signal: AbortSignal) => {
    tab.setActivity({ label: t("action-queue.compacting"), hint: t("action-queue.compactingHint") });
    try {
      await opts.runCompact(args, signal);
    } finally {
      tab.setActivity(null);
    }
  };

  // Mid-turn auto-compaction: checked between tool-loop continuations so a
  // single tool-heavy turn can't overflow a small context window. The run's
  // own signal cancels it on Stop.
  agent.prepareNextTurn = async (signal) => {
    if (closed) return undefined;
    return maybeAutoCompactBeforeContinuation({
      agent,
      enabled: autoCompactEnabled,
      runCompact: () => runCompactWithIndicator("", signal ?? stopController.signal),
    });
  };

  /** Removes undelivered steers/follow-ups from the strip and from the agent. */
  const drainMessages = (): string[] => {
    const texts = queueDisplay.drainQueuedMessages().map((item) => item.text);
    agent.clearAllQueues();
    return texts;
  };

  const takeBack = (): TakenBack => {
    const texts = drainMessages();
    const cancelledCommands: string[] = [];
    for (const action of actions) {
      if (action.type === "prompt") texts.push(action.text);
      else cancelledCommands.push(`/${action.name}`);
    }
    actions.length = 0;
    syncDisplay();
    return { texts, cancelledCommands };
  };

  const removeRow = (row: QueueRow): string | null => {
    if (row.kind === "held") {
      const text = tab.held[row.index];
      if (text === undefined) return null;
      tab.setHeld(tab.held.filter((_, i) => i !== row.index));
      return text;
    }
    if (row.kind === "message") {
      queueDisplay.removeQueuedMessage(row.item);
      return row.item.text;
    }
    const action = actions[row.index];
    if (!action) return null;
    actions.splice(row.index, 1);
    syncDisplay();
    return action.type === "prompt" ? action.text : null;
  };

  /** Edit puts that one row's text back in the composer, above any draft. */
  const editRow = (row: QueueRow): boolean => {
    if (row.kind === "action" && actions[row.index]?.type !== "prompt") return false;
    const text = removeRow(row);
    if (text === null) return false;
    tab.returnToDraft([text]);
    return true;
  };

  queueDisplay.setRowHandlers({
    edit: (row) => { editRow(row); },
    remove: (row) => { removeRow(row); },
  });

  const editLast = (): boolean => {
    const row = queueDisplay.lastEditableRow();
    return row !== null && editRow(row);
  };

  const stop = (): TakenBack => {
    const taken = takeBack();
    stopController.abort();
    agent.abort();
    tab.hold(taken.texts);
    return taken;
  };

  const shutdown = () => {
    closed = true;
    stopController.abort();
    delete agent.prepareNextTurn;
    actions.length = 0;
    syncDisplay();
  };

  /**
   * Retries a run that failed with a transient provider error: drops the
   * failed response, waits with a visible countdown, and continues the turn.
   * If the turn still ends in an error (Stop during the wait, a retry that
   * fails differently, or no retries left), the transcript keeps the first
   * failed response: it holds any answer that streamed before the failure,
   * which a retry that fails before streaming would otherwise replace. It
   * carries the latest error, since that is why the turn ended.
   */
  async function retryTransientFailure(signal: AbortSignal): Promise<void> {
    let firstFailure: AssistantMessage | null = null;
    for (let attempt = 1; attempt <= PROVIDER_RETRY_LIMIT; attempt++) {
      const failure = lastAssistant(agent.state.messages);
      if (!failure || failure.stopReason !== "error" || signal.aborted) break;
      if (!isRetryableAssistantError(failure) || findTrailingContextOverflowError(agent.state)) break;
      firstFailure ??= failure;

      agent.state.messages = agent.state.messages.slice(0, -1);
      const delayMs = retryDelayMs({ baseDelayMs: PROVIDER_RETRY_BASE_DELAY_MS }, attempt);
      const hint = t("action-queue.retryingHint");
      try {
        for (let remaining = delayMs; remaining > 0; remaining -= 1_000) {
          tab.setActivity({
            label: t("action-queue.retrying", {
              seconds: String(Math.ceil(remaining / 1_000)),
              attempt: String(attempt),
              max: String(PROVIDER_RETRY_LIMIT),
            }),
            hint,
          });
          await waitUnlessStopped(Math.min(1_000, remaining), signal);
        }
      } catch (error) {
        if (!(error instanceof StoppedError)) throw error;
        agent.state.messages = [...agent.state.messages, failure];
        break;
      } finally {
        tab.setActivity(null);
      }

      if (closed) return;
      await agent.continue();
    }

    const last = lastAssistant(agent.state.messages);
    if (firstFailure && last && last !== firstFailure && last.stopReason === "error" && !findTrailingContextOverflowError(agent.state)) {
      const kept: AssistantMessage = last.errorMessage === undefined
        ? firstFailure
        : { ...firstFailure, errorMessage: last.errorMessage };
      agent.state.messages = [...agent.state.messages.slice(0, -1), kept];
    }
  }

  async function runCommand(name: string, args: string, signal: AbortSignal): Promise<void> {
    // A queued /compact compacts this tab, even if another tab is active by
    // the time it runs.
    if (name === "compact") {
      await runCompactWithIndicator(args, signal);
      return;
    }

    const cmd = commandRegistry.get(name);
    if (!cmd) throw new Error(`Unknown command: /${name}`);
    await cmd.execute(args);
  }

  /**
   * A turn that ends (stopped, failed or finished) never keeps steers and
   * follow-ups it did not deliver; they become "Not sent" rows rather than
   * going into the next turn. Only a message sent while the turn was already
   * ending can still be here.
   */
  function returnLeftovers(): void {
    const leftovers = drainMessages();
    if (leftovers.length === 0) return;
    tab.hold(leftovers);
    opts.onLeftUnsent?.(leftovers.length);
  }

  async function runPrompt(text: string, signal: AbortSignal): Promise<void> {
    try {
      await runTurn(text, signal);
    } finally {
      if (!closed) returnLeftovers();
    }
  }

  async function runTurn(text: string, signal: AbortSignal): Promise<void> {
    await agent.prompt(text);
    if (closed || signal.aborted) return;

    await retryTransientFailure(signal);
    if (closed || signal.aborted) return;

    let overflowAfterRecovery = false;
    if (autoCompactEnabled) {
      // One compact-and-retry attempt when the run ended in a
      // context-overflow error. Never loops: a second overflow stays in
      // the transcript with an actionable error banner.
      const retried = await recoverFromContextOverflow({
        agent,
        runCompact: () => runCompactWithIndicator("", signal),
      });
      if (retried && !signal.aborted) {
        await retryTransientFailure(signal);
        overflowAfterRecovery = findTrailingContextOverflowError(agent.state) !== null;
      }
    }

    returnLeftovers();
    opts.onTurnSettled?.({ overflowAfterRecovery });
  }

  async function process(): Promise<void> {
    if (running || closed) return;
    running = true;
    tab.setBusy(true);

    try {
      while (!closed && actions.length > 0) {
        // Never start queued actions while the agent is still streaming.
        await agent.waitForIdle();
        if (closed) break;

        const next = actions[0];
        if (!next) break;
        if (stopController.signal.aborted) stopController = new AbortController();
        const signal = stopController.signal;

        if (next.type === "command") {
          actions.shift();
          syncDisplay();
          try {
            await runCommand(next.name, next.args, signal);
          } catch (error) {
            console.warn(`[pi] Queued /${next.name} failed:`, error);
          }
          continue;
        }

        // The prompt stays in the queue strip (and can be taken back or
        // stopped) while pre-prompt compaction runs.
        await maybeAutoCompactBeforePrompt({
          agent,
          nextUserText: next.text,
          enabled: autoCompactEnabled,
          runCompact: () => runCompactWithIndicator("", signal),
        });
        if (closed) break;
        if (actions[0] !== next) continue;

        actions.shift();
        syncDisplay();
        await runPrompt(next.text, signal);
      }
    } finally {
      running = false;
      syncDisplay();
      tab.setBusy(false);
    }
  }

  const enqueuePrompt = (text: string) => {
    if (closed) return;

    const trimmed = text.trim();
    if (!trimmed) return;

    actions.push({ type: "prompt", text: trimmed });
    syncDisplay();
    void process();
  };

  const enqueueCommand = (name: string, args: string) => {
    if (closed) return;

    const cmdName = name.trim();
    if (!cmdName) return;

    actions.push({ type: "command", name: cmdName, args: args.trim() });
    syncDisplay();
    void process();
  };

  return { enqueuePrompt, enqueueCommand, editLast, stop, isBusy, shutdown };
}
