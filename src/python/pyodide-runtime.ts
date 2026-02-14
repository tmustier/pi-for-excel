/**
 * Pyodide runtime adapter — main-thread API for in-browser Python execution.
 *
 * Manages a Web Worker that loads Pyodide (Python compiled to WebAssembly).
 * Provides the same request/response interface as the native Python bridge,
 * so callers don't need to know which backend is being used.
 *
 * Key design decisions:
 * - Lazy initialization: Pyodide (~15MB WASM) is only loaded on first use
 * - Web Worker isolation: Python runs off the main thread
 * - CDN loading: Pyodide served from jsDelivr, cached by the browser
 * - Timeout support: Worker calls are bounded by configurable timeouts
 */

import type {
  PythonBridgeRequest,
  PythonBridgeResponse,
} from "../tools/python-run.js";
import type {
  PyodideWorkerRequest,
  PyodideWorkerResponse,
} from "./pyodide-worker.js";

const DEFAULT_TIMEOUT_MS = 30_000;

let worker: Worker | null = null;
let requestCounter = 0;
const pendingRequests = new Map<string, {
  resolve: (response: PyodideWorkerResponse) => void;
  reject: (error: Error) => void;
  timeoutId: ReturnType<typeof setTimeout>;
}>();

function getOrCreateWorker(): Worker {
  if (worker) return worker;

  // The worker file is bundled by Vite as a separate chunk.
  // Using `new URL(..., import.meta.url)` tells Vite to emit the worker file.
  worker = new Worker(
    new URL("./pyodide-worker.ts", import.meta.url),
    { type: "module" },
  );

  worker.addEventListener("message", (event: MessageEvent<PyodideWorkerResponse>) => {
    const response = event.data;
    const pending = pendingRequests.get(response.id);
    if (!pending) return;

    clearTimeout(pending.timeoutId);
    pendingRequests.delete(response.id);
    pending.resolve(response);
  });

  worker.addEventListener("error", (event) => {
    // Reject all pending requests on worker crash
    const error = new Error(`Pyodide worker error: ${event.message}`);
    for (const [id, pending] of pendingRequests) {
      clearTimeout(pending.timeoutId);
      pending.reject(error);
      pendingRequests.delete(id);
    }

    // Reset so next call creates a fresh worker
    worker = null;
  });

  return worker;
}

function sendToWorker(
  request: PyodideWorkerRequest,
  timeoutMs: number,
): Promise<PyodideWorkerResponse> {
  return new Promise<PyodideWorkerResponse>((resolve, reject) => {
    const timeoutId = setTimeout(() => {
      pendingRequests.delete(request.id);

      // Terminate the hung worker so it doesn't block future calls
      if (worker) {
        worker.terminate();
        worker = null;
      }

      reject(new Error(`Pyodide execution timed out after ${timeoutMs}ms. Worker was terminated.`));
    }, timeoutMs);

    pendingRequests.set(request.id, { resolve, reject, timeoutId });

    const w = getOrCreateWorker();
    w.postMessage(request);
  });
}

/**
 * Execute Python code via Pyodide (in-browser WebAssembly runtime).
 *
 * Accepts the same request shape as the native Python bridge and returns
 * a compatible response, so callers can swap between backends transparently.
 */
export async function callPyodideRuntime(
  request: PythonBridgeRequest,
  signal?: AbortSignal,
): Promise<PythonBridgeResponse> {
  if (signal?.aborted) {
    return {
      ok: false,
      action: "run_python",
      error: "Aborted",
    };
  }

  const id = `pyodide-${++requestCounter}`;
  const timeoutMs = request.timeout_ms ?? DEFAULT_TIMEOUT_MS;

  const workerRequest: PyodideWorkerRequest = {
    id,
    type: "run",
    code: request.code,
    inputJson: request.input_json,
  };

  // Set up abort handling
  let abortHandler: (() => void) | null = null;
  if (signal) {
    abortHandler = () => {
      const pending = pendingRequests.get(id);
      if (pending) {
        clearTimeout(pending.timeoutId);
        pendingRequests.delete(id);
        pending.reject(new Error("Aborted"));
      }
    };
    signal.addEventListener("abort", abortHandler, { once: true });
  }

  try {
    const workerResponse = await sendToWorker(workerRequest, timeoutMs);

    return {
      ok: workerResponse.ok,
      action: "run_python",
      exit_code: workerResponse.ok ? 0 : 1,
      stdout: workerResponse.stdout,
      stderr: workerResponse.stderr,
      result_json: workerResponse.resultJson,
      error: workerResponse.error,
      metadata: {
        backend: "pyodide",
        loadTimeMs: workerResponse.loadTimeMs,
        runTimeMs: workerResponse.runTimeMs,
      },
    };
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      ok: false,
      action: "run_python",
      error: message,
    };
  } finally {
    if (signal && abortHandler) {
      signal.removeEventListener("abort", abortHandler);
    }
  }
}

/** Check if Pyodide is available in this environment (Web Workers + WASM). */
export function isPyodideAvailable(): boolean {
  if (typeof Worker === "undefined") return false;
  if (typeof WebAssembly === "undefined") return false;
  return true;
}

/** Check if Pyodide has already been loaded (warm). */
export function isPyodideLoaded(): boolean {
  return worker !== null;
}

/** Terminate the Pyodide worker and release resources. */
export function terminatePyodide(): void {
  if (!worker) return;

  for (const [id, pending] of pendingRequests) {
    clearTimeout(pending.timeoutId);
    pending.reject(new Error("Pyodide runtime terminated"));
    pendingRequests.delete(id);
  }

  worker.terminate();
  worker = null;
}
