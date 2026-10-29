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
    const error = new Error(`Pyodide worker error: ${event.message}`);
    for (const [id, pending] of pendingRequests) {
      clearTimeout(pending.timeoutId);
      pending.reject(error);
      pendingRequests.delete(id);
    }

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
  };
  if (request.input_json !== undefined) {
    workerRequest.inputJson = request.input_json;
  }

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

    const response: PythonBridgeResponse = {
      ok: workerResponse.ok,
      action: "run_python",
      exit_code: workerResponse.ok ? 0 : 1,
      metadata: {
        backend: "pyodide",
        loadTimeMs: workerResponse.loadTimeMs,
        runTimeMs: workerResponse.runTimeMs,
      },
    };
    if (workerResponse.stdout !== undefined) response.stdout = workerResponse.stdout;
    if (workerResponse.stderr !== undefined) response.stderr = workerResponse.stderr;
    if (workerResponse.resultJson !== undefined) response.result_json = workerResponse.resultJson;
    if (workerResponse.error !== undefined) response.error = workerResponse.error;
    return response;
  } catch (error) {
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

export function isPyodideAvailable(): boolean {
  if (typeof Worker === "undefined") return false;
  if (typeof WebAssembly === "undefined") return false;
  return true;
}

export function isPyodideLoaded(): boolean {
  return worker !== null;
}

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
