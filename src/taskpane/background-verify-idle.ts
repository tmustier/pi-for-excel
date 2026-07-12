/**
 * Pure wait-until-idle decision logic for the background verification bridge.
 *
 * The old waiter clamped to 120s and used an in-call `sawProgress` flag that
 * raced the run start, producing false immediate-idle. This module derives
 * "started" from durable runtime signals (busy now, or the message count grew
 * past the pre-run baseline) so the decision is correct across independent
 * poll calls and never reports idle before the run actually starts.
 */

export type IdleReason = "awaiting-start" | "start-timeout" | "running" | "idle";

export interface IdleDecision {
  started: boolean;
  idle: boolean;
  /** True when the poll loop should stop (terminal state reached). */
  done: boolean;
  reason: IdleReason;
}

export interface IdleDecisionInput {
  baselineMessageCount: number;
  currentMessageCount: number;
  isBusy: boolean;
  /** Sticky flag: has a run start been observed in any prior tick of this wait? */
  sawStart: boolean;
  /** How long this wait call has been observing (ms). */
  observedElapsedMs: number;
  /** Grace window to wait for the run to start before giving up. */
  startupGraceMs: number;
}

/**
 * Decide the current idle state from durable signals. A run is considered
 * started once the runtime is busy or the message count has grown past the
 * pre-run baseline; idle is only ever reported after a start is observed.
 */
export function decideRuntimeIdle(input: IdleDecisionInput): IdleDecision {
  const startedNow = input.isBusy || input.currentMessageCount > input.baselineMessageCount;
  const started = input.sawStart || startedNow;

  if (!started) {
    if (input.observedElapsedMs >= input.startupGraceMs) {
      return { started: false, idle: false, done: true, reason: "start-timeout" };
    }
    return { started: false, idle: false, done: false, reason: "awaiting-start" };
  }

  if (input.isBusy) {
    return { started: true, idle: false, done: false, reason: "running" };
  }

  return { started: true, idle: true, done: true, reason: "idle" };
}
