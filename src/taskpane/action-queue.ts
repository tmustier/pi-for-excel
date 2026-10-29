import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";
import { retryDelayMs, type AssistantMessage } from "@earendil-works/pi-ai";
import { t } from "../language/index.js";
import { providerRetryDelay } from "../auth/provider-failure.js";
import "../messages/interrupted-response.js";

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

export interface TakenBack {
  texts: string[];
  cancelledCommands: string[];
}

export interface TurnOutcome {
  overflowAfterRecovery: boolean;
}

export interface ActionQueue {
  enqueuePrompt: (text: string) => void;
  enqueueCommand: (name: string, args: string) => void;
  retry: () => boolean;
  editLast: () => boolean;
  stop: () => TakenBack;
  isBusy: () => boolean;
  shutdown: () => void;
}

type ActionQueueDisplay = Pick<
  QueueDisplay,
  "setActionQueue" | "drainQueuedMessages" | "removeQueuedMessage" | "setRowHandlers" | "lastEditableRow"
>;

type ActivityHost = Pick<TabState, "setActivity" | "setBusy" | "returnToDraft" | "held" | "setHeld" | "hold">;

export const PROVIDER_RETRY_LIMIT = 6;
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
  runCompact: (args: string, signal: AbortSignal) => Promise<void>;
  onTurnSettled?: (outcome: TurnOutcome) => void;
  onLeftUnsent?: (count: number) => void;
}): ActionQueue {
  const { agent, tab, queueDisplay, autoCompactEnabled } = opts;

  const actions: QueuedAction[] = [];
  let running = false;
  let retryRequested = false;
  let closed = false;
  let stopController = new AbortController();

  const syncDisplay = () => {
    queueDisplay.setActionQueue(
      actions.map((a) => {
        if (a.type === "prompt") return { type: "prompt", label: t("action-queue.queued"), text: a.text };
        return { type: "command", label: `/${a.name}`, text: a.args };
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

  agent.prepareNextTurn = async (signal) => {
    if (closed) return undefined;
    return maybeAutoCompactBeforeContinuation({
      agent,
      enabled: autoCompactEnabled,
      runCompact: () => runCompactWithIndicator("", signal ?? stopController.signal),
    });
  };

  const drainMessages = (): string[] => {
    const texts = queueDisplay.drainQueuedMessages().map((item) => item.text);
    agent.clearAllQueues();
    return texts;
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

  const stop = (): TakenBack => {
    if (isBusy()) tab.setActivity({ label: t("working.stopping"), hint: null });
    const texts = drainMessages();
    const cancelledCommands: string[] = [];
    for (const action of actions) {
      if (action.type === "prompt") texts.push(action.text);
      else cancelledCommands.push(`/${action.name}`);
    }
    actions.length = 0;
    syncDisplay();
    stopController.abort();
    agent.abort();
    tab.hold(texts);
    return { texts, cancelledCommands };
  };

  const shutdown = () => {
    closed = true;
    stopController.abort();
    delete agent.prepareNextTurn;
    actions.length = 0;
    syncDisplay();
  };

  function prepareRetry(): void {
    const failure = lastAssistant(agent.state.messages);
    if (!failure || failure.stopReason !== "error") return;
    agent.state.messages = agent.state.messages.slice(0, -1);
    if (failure.content.some((part) => part.type === "text" && part.text.trim())) {
      agent.state.messages.push({ role: "interruptedResponse", message: failure, timestamp: failure.timestamp });
    }
  }

  async function retryTransientFailure(signal: AbortSignal): Promise<void> {
    for (let attempt = 1; attempt <= PROVIDER_RETRY_LIMIT; attempt++) {
      const failure = lastAssistant(agent.state.messages);
      if (!failure || failure.stopReason !== "error" || signal.aborted) break;
      if (findTrailingContextOverflowError(agent.state)) break;
      const delayMs = providerRetryDelay(failure, retryDelayMs({ baseDelayMs: PROVIDER_RETRY_BASE_DELAY_MS, maxAgentDelayMs: 16_000 }, attempt));
      if (delayMs === null) break;
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
        break;
      } finally {
        tab.setActivity(null);
      }

      if (closed) return;
      prepareRetry();
      await agent.continue();
    }
  }

  function returnLeftovers(): void {
    const leftovers = drainMessages();
    if (leftovers.length === 0) return;
    tab.hold(leftovers);
    opts.onLeftUnsent?.(leftovers.length);
  }

  async function process(): Promise<void> {
    if (running || closed) return;
    running = true;
    tab.setBusy(true);

    try {
      while (!closed && (actions.length > 0 || retryRequested)) {
        await agent.waitForIdle();
        if (closed) break;

        const next = actions[0];
        const retry = retryRequested;
        retryRequested = false;
        if (!next && !retry) break;
        if (stopController.signal.aborted) stopController = new AbortController();
        const signal = stopController.signal;

        if (next?.type === "command") {
          actions.shift();
          syncDisplay();
          try {
            if (next.name === "compact") {
              await runCompactWithIndicator(next.args, signal);
            } else {
              const cmd = commandRegistry.get(next.name);
              if (!cmd) throw new Error(`Unknown command: /${next.name}`);
              await cmd.execute(next.args);
            }
          } catch (error) {
            console.warn(`[pi] Queued /${next.name} failed:`, error);
          }
          continue;
        }

        if (next?.type === "prompt") await maybeAutoCompactBeforePrompt({
          agent,
          nextUserText: next.text,
          enabled: autoCompactEnabled,
          runCompact: () => runCompactWithIndicator("", signal),
        });
        if (closed) break;
        if (!retry && actions[0] !== next) continue;

        actions.shift();
        syncDisplay();
        try {
          if (retry) {
            prepareRetry();
            await agent.continue();
          } else if (next?.type === "prompt") {
            await agent.prompt(next.text);
          }
          if (closed || signal.aborted) continue;

          await retryTransientFailure(signal);
          if (closed || signal.aborted) continue;

          const retried = autoCompactEnabled && await recoverFromContextOverflow({
            agent,
            runCompact: () => runCompactWithIndicator("", signal),
          });
          if (retried && !signal.aborted) await retryTransientFailure(signal);

          returnLeftovers();
          opts.onTurnSettled?.({
            overflowAfterRecovery: retried && findTrailingContextOverflowError(agent.state) !== null,
          });
        } finally {
          if (!closed) returnLeftovers();
        }
      }
    } finally {
      running = false;
      syncDisplay();
      tab.setActivity(null);
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

  return {
    enqueuePrompt,
    enqueueCommand,
    retry: () => {
      const failure = lastAssistant(agent.state.messages);
      if (closed || isBusy() || failure?.stopReason !== "error") return false;
      retryRequested = true;
      void process();
      return true;
    },
    editLast: () => {
      const row = queueDisplay.lastEditableRow();
      return row !== null && editRow(row);
    },
    stop,
    isBusy,
    shutdown,
  };
}
