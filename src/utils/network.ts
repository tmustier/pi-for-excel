export function isAbortError(error: unknown): boolean {
  if (error instanceof DOMException) {
    return error.name === "AbortError";
  }

  if (error instanceof Error) {
    return error.name === "AbortError";
  }

  return false;
}

export function getHttpErrorReason(status: number, responseBody: string): string {
  const trimmed = responseBody.trim();
  if (trimmed.length > 0) {
    return trimmed;
  }

  return `HTTP ${status}`;
}

export async function runWithTimeoutAbort<TResult>(args: {
  signal: AbortSignal | undefined;
  timeoutMs: number;
  timeoutErrorMessage: string;
  run: (signal: AbortSignal) => Promise<TResult>;
}): Promise<TResult> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => {
    controller.abort();
  }, args.timeoutMs);

  const abortFromCaller = () => {
    controller.abort();
  };

  const callerSignal = args.signal;
  if (callerSignal) {
    if (callerSignal.aborted) {
      controller.abort();
    } else {
      callerSignal.addEventListener("abort", abortFromCaller, { once: true });
    }
  }

  let rejectAbort: ((reason?: unknown) => void) | null = null;
  const abortPromise = new Promise<never>((_resolve, reject) => {
    rejectAbort = reject;
  });

  const onControllerAbort = () => {
    if (rejectAbort) {
      rejectAbort(new DOMException("aborted", "AbortError"));
    }
  };

  if (controller.signal.aborted) {
    onControllerAbort();
  } else {
    controller.signal.addEventListener("abort", onControllerAbort, { once: true });
  }

  try {
    const runPromise = args.run(controller.signal);
    return await Promise.race([runPromise, abortPromise]);
  } catch (error: unknown) {
    if (isAbortError(error)) {
      if (callerSignal?.aborted) {
        throw new Error("Aborted");
      }

      throw new Error(args.timeoutErrorMessage);
    }

    throw error;
  } finally {
    clearTimeout(timeoutId);
    controller.signal.removeEventListener("abort", onControllerAbort);
    if (callerSignal) {
      callerSignal.removeEventListener("abort", abortFromCaller);
    }
  }
}
