import {
  createAssistantMessageEventStream,
  type AssistantMessage,
  type AssistantMessageEventStream,
} from "@earendil-works/pi-ai";

const SUSPEND_EVENT = "pi-durable-suspend";

/** Seal Durable before asking WebKit to cancel this page's network requests. */
export function prepareDurableReload(target: EventTarget = window): void {
  target.dispatchEvent(new Event(SUSPEND_EVENT));
}

export interface ModelRecoveryLifecycle {
  readonly suspending: boolean;
  suspend(): void;
  dispose(): void;
}

/**
 * Harness.close seals commit admission synchronously and joins running tasks
 * without committing their outcomes. Do not abort the conversation: that would
 * persist an intentional cancellation rather than leave its work resumable.
 */
export function createModelRecoveryLifecycle(
  close: () => Promise<void>,
  target: EventTarget = window,
  report: (error: Error) => void = (error) => console.warn(error.message),
): ModelRecoveryLifecycle {
  let suspending = false;
  const suspend = (): void => {
    if (suspending) return;
    suspending = true;
    void close().catch(() => report(new Error("Could not finish closing Durable before the pane unloaded.")));
  };
  const events = [SUSPEND_EVENT, "beforeunload", "pagehide"];
  for (const event of events) target.addEventListener(event, suspend);
  return {
    get suspending() { return suspending; },
    suspend,
    dispose: () => {
      for (const event of events) target.removeEventListener(event, suspend);
    },
  };
}

/**
 * WebKit can deliver a fetch error before its pagehide event. Yield terminal
 * errors to the browser event queue so an already pending unload can seal the
 * harness first. Ordinary network/authentication errors still reach Durable;
 * this is not an error retry or a heuristic based on the provider's wording.
 */
export function protectModelStream(
  source: AssistantMessageEventStream,
  lifecycle: ModelRecoveryLifecycle,
  failure: () => AssistantMessage,
): AssistantMessageEventStream {
  const output = createAssistantMessageEventStream();
  const cancelled = (message: AssistantMessage): AssistantMessage => ({
    ...message,
    stopReason: "aborted",
    errorMessage: "Pane unloaded; the unfinished request will resume when the pane opens.",
  });
  void (async () => {
    for await (const event of source) {
      if (event.type === "error") {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
        if (lifecycle.suspending) {
          const error = cancelled(event.error);
          output.push({ type: "error", reason: "aborted", error });
          return;
        }
      }
      if (lifecycle.suspending) continue;
      output.push(event);
    }
    const result = await source.result();
    output.end(lifecycle.suspending ? cancelled(result) : result);
  })().catch(() => {
    // Never expose an arbitrary provider exception, which may contain secrets.
    const error = lifecycle.suspending ? cancelled(failure()) : failure();
    output.push({ type: "error", reason: lifecycle.suspending ? "aborted" : "error", error });
  });
  return output;
}
