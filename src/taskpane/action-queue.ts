/**
 * UI-level ordered action queue.
 *
 * Needed because some actions (notably `/compact`) run outside the Agent loop
 * (they call `agent.streamFn(...)` directly) and therefore don't set
 * `agent.state.isStreaming`. Without a queue, user input can be lost when
 * compaction rewrites the message list.
 *
 * The queue also owns automatic context protection for its agent (#566):
 * - pre-prompt auto-compaction (existing behavior)
 * - mid-turn auto-compaction between tool-loop continuations (via
 *   `agent.prepareNextTurn`)
 * - one compact-and-retry recovery when a run ends in a context-overflow error
 */

import type { Agent } from "@earendil-works/pi-agent-core";

import { commandRegistry } from "../commands/types.js";
import {
  maybeAutoCompactBeforeContinuation,
  maybeAutoCompactBeforePrompt,
} from "../compaction/auto-compaction.js";
import { recoverFromContextOverflow } from "../compaction/overflow-recovery.js";

export type QueuedAction =
  | { type: "prompt"; text: string }
  | { type: "command"; name: string; args: string };

export interface ActionQueue {
  enqueuePrompt: (text: string) => void;
  enqueueCommand: (name: string, args: string) => void;
  drainQueuedActions: () => QueuedAction[];
  isBusy: () => boolean;
  shutdown: () => void;
}

interface ActionQueueDisplay {
  setActionQueue: (items: Array<{ type: "prompt" | "command"; label: string; text: string }>) => void;
}

interface BusyIndicatorHost {
  setBusyIndicator: (label: string | null, hint?: string | null) => void;
}

export function createActionQueue(opts: {
  agent: Agent;
  sidebar: BusyIndicatorHost;
  queueDisplay: ActionQueueDisplay;
  autoCompactEnabled: boolean;
  /** Runs compaction for this queue's agent (not the active tab's agent). */
  runCompact: () => Promise<void>;
}): ActionQueue {
  const { agent, sidebar, queueDisplay, autoCompactEnabled } = opts;

  const actions: QueuedAction[] = [];
  let running = false;
  let closed = false;

  const syncDisplay = () => {
    queueDisplay.setActionQueue(
      actions.map((a) => {
        if (a.type === "prompt") return { type: "prompt", label: "Queued", text: a.text };
        return { type: "command", label: `/${a.name}`, text: a.args ? a.args : "" };
      }),
    );
  };

  const isBusy = () => running || agent.state.isStreaming;

  const runCompactWithIndicator = async () => {
    sidebar.setBusyIndicator(
      "Compacting context…",
      "Send messages and Pi will see them after compaction",
    );
    try {
      await opts.runCompact();
    } finally {
      sidebar.setBusyIndicator(null);
    }
  };

  // Mid-turn auto-compaction: checked between tool-loop continuations so a
  // single tool-heavy turn can't overflow a small context window.
  agent.prepareNextTurn = async () => {
    if (closed) return undefined;
    return maybeAutoCompactBeforeContinuation({
      agent,
      enabled: autoCompactEnabled,
      runCompact: runCompactWithIndicator,
    });
  };

  const shutdown = () => {
    closed = true;
    agent.prepareNextTurn = undefined;
    actions.length = 0;
    syncDisplay();
  };

  async function runCommand(name: string, args: string): Promise<void> {
    const cmd = commandRegistry.get(name);
    if (!cmd) throw new Error(`Unknown command: /${name}`);

    // Special-case: show an explicit non-streaming indicator for compaction.
    if (name === "compact") {
      sidebar.setBusyIndicator(
        "Compacting context…",
        "Send messages and Pi will see them after compaction",
      );
      try {
        await cmd.execute(args);
      } finally {
        sidebar.setBusyIndicator(null);
      }
      return;
    }

    await cmd.execute(args);
  }

  async function process(): Promise<void> {
    if (running || closed) return;
    running = true;

    try {
      // Drain sequentially.
      while (!closed && actions.length > 0) {
        // Never start queued actions while the agent is still streaming.
        await agent.waitForIdle();

        if (closed) break;

        const next = actions.shift();
        if (!next) break;
        syncDisplay();

        if (closed) break;

        if (next.type === "command") {
          await runCommand(next.name, next.args);
          continue;
        }

        // next.type === "prompt"
        await maybeAutoCompactBeforePrompt({
          agent,
          nextUserText: next.text,
          enabled: autoCompactEnabled,
          runCompact: runCompactWithIndicator,
        });

        if (closed) break;
        await agent.prompt(next.text);

        if (closed) break;
        if (autoCompactEnabled) {
          // One compact-and-retry attempt when the run ended in a
          // context-overflow error. Never loops: a second overflow stays in
          // the transcript with an actionable error banner.
          await recoverFromContextOverflow({
            agent,
            runCompact: runCompactWithIndicator,
          });
        }
      }
    } finally {
      running = false;
      syncDisplay();
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

  const drainQueuedActions = (): QueuedAction[] => {
    if (actions.length === 0) {
      return [];
    }

    const drained = [...actions];
    actions.length = 0;
    syncDisplay();
    return drained;
  };

  return { enqueuePrompt, enqueueCommand, drainQueuedActions, isBusy, shutdown };
}
