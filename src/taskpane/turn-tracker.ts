import type { Agent } from "@earendil-works/pi-agent-core";

export interface RuntimeTurn {
  readonly id: string;
  readonly toolCallIds: Set<string>;
  readonly officeJsCallIds: Set<string>;
  missingCredentialsFor: string | null;
  expiredLoginFor: string | null;
}

export interface TurnTracker {
  current(): RuntimeTurn;
  running(): RuntimeTurn | null;
  turnIdOf(toolCallId: string): string | undefined;
  dispose(): void;
}

function startTurn(): RuntimeTurn {
  return {
    id: crypto.randomUUID(),
    toolCallIds: new Set(),
    officeJsCallIds: new Set(),
    missingCredentialsFor: null,
    expiredLoginFor: null,
  };
}

export function trackTurns(agent: Agent): TurnTracker {
  let turn = startTurn();
  let isRunning = false;
  const toolTurns = new Map<string, string>();

  const unsubscribe = agent.subscribe((event) => {
    if (event.type === "agent_start") {
      turn = startTurn();
      isRunning = true;
    } else if (event.type === "tool_execution_start") {
      toolTurns.set(event.toolCallId, turn.id);
    } else if (event.type === "agent_end") {
      isRunning = false;
    }
  });

  return {
    current: () => turn,
    running: () => (isRunning ? turn : null),
    turnIdOf: (toolCallId) => toolTurns.get(toolCallId),
    dispose: unsubscribe,
  };
}
