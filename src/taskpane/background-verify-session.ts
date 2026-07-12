/**
 * Pure fresh-session transition guard for the background verification bridge.
 *
 * `newSession` must fail closed: returning `runtimeChanged/sessionChanged`
 * booleans alone can falsely claim success. A real fresh session must produce
 * an active runtime that actually changed (new runtime or new session) and
 * starts empty. This logic is DOM-free so it can be unit-tested deterministically;
 * the bridge feeds it the before/after runtime + session ids and message count.
 */

export interface FreshSessionTransitionInput {
  beforeRuntimeId: string | null;
  afterRuntimeId: string | null;
  beforeSessionId: string | null;
  afterSessionId: string | null;
  afterMessageCount: number;
}

export interface FreshSessionTransition {
  runtimeChanged: boolean;
  sessionChanged: boolean;
  afterMessageCount: number;
}

/**
 * Validate that a `newSession` actually produced a fresh, empty session.
 * Throws (fail closed) when there is no active runtime, when neither the
 * runtime nor the session changed, or when the new runtime already has
 * messages.
 */
export function validateFreshSessionTransition(input: FreshSessionTransitionInput): FreshSessionTransition {
  if (!input.afterRuntimeId) {
    throw new Error("newSession produced no active runtime");
  }

  const runtimeChanged = input.beforeRuntimeId !== input.afterRuntimeId;
  const sessionChanged = input.beforeSessionId !== input.afterSessionId;
  if (!runtimeChanged && !sessionChanged) {
    throw new Error("newSession did not change the active runtime or session");
  }

  if (input.afterMessageCount !== 0) {
    throw new Error(`newSession produced a non-empty runtime (${input.afterMessageCount} messages)`);
  }

  return { runtimeChanged, sessionChanged, afterMessageCount: input.afterMessageCount };
}
