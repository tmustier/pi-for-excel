import type {
  PythonBridgeRequest,
  PythonBridgeResponse,
} from "../tools/python-run.js";
import type {
  PyodideWorkerMessage,
  PyodideWorkerRequest,
  PyodideWorkerResponse,
} from "./pyodide-worker.js";
import { ToolStoppedError } from "../tools/tool-stopped.js";

const DEFAULT_TIMEOUT_MS = 30_000;

interface PendingRequest {
  request: PyodideWorkerRequest;
  timeoutMs: number;
  resolve: (response: PyodideWorkerResponse) => void;
  reject: (error: Error) => void;
}

type WorkerPhase =
  | { kind: "idle" }
  | { kind: "running"; pending: PendingRequest; timeoutId: ReturnType<typeof setTimeout> }
  | { kind: "terminated" };

interface WorkerInstance {
  worker: Worker;
  loaded: boolean;
  phase: WorkerPhase;
  queue: PendingRequest[];
}

type RuntimeState = { kind: "stopped" } | { kind: "active"; instance: WorkerInstance };
let runtime: RuntimeState = { kind: "stopped" };
let requestCounter = 0;

function terminateInstance(instance: WorkerInstance, error: Error): void {
  const phase = instance.phase;
  if (phase.kind === "terminated") return;
  instance.phase = { kind: "terminated" };
  if (runtime.kind === "active" && runtime.instance === instance) runtime = { kind: "stopped" };
  instance.worker.terminate();
  if (phase.kind === "running") {
    clearTimeout(phase.timeoutId);
    phase.pending.reject(error);
  }
  for (const pending of instance.queue.splice(0)) pending.reject(error);
}

function dispatchNext(instance: WorkerInstance): void {
  if (instance.phase.kind !== "idle") return;
  const pending = instance.queue.shift();
  if (!pending) return;
  const timeoutId = setTimeout(() => {
    terminateInstance(instance, new Error(`Pyodide execution timed out after ${pending.timeoutMs}ms. Worker was terminated.`));
  }, pending.timeoutMs);
  instance.phase = { kind: "running", pending, timeoutId };
  try {
    instance.worker.postMessage(pending.request);
  } catch (error) {
    terminateInstance(instance, error instanceof Error ? error : new Error(String(error)));
  }
}

function getOrCreateWorker(): WorkerInstance {
  if (runtime.kind === "active") return runtime.instance;
  const worker = new Worker(new URL("./pyodide-worker.ts", import.meta.url), { type: "module" });
  const instance: WorkerInstance = { worker, loaded: false, phase: { kind: "idle" }, queue: [] };
  runtime = { kind: "active", instance };
  worker.addEventListener("message", (event: MessageEvent<PyodideWorkerMessage>) => {
    if (instance.phase.kind === "terminated") return;
    if ("type" in event.data) {
      instance.loaded = true;
      return;
    }
    const phase = instance.phase;
    if (phase.kind !== "running" || phase.pending.request.id !== event.data.id) return;
    clearTimeout(phase.timeoutId);
    instance.phase = { kind: "idle" };
    phase.pending.resolve(event.data);
    dispatchNext(instance);
  });
  worker.addEventListener("error", (event) => {
    terminateInstance(instance, new Error(`Pyodide worker error: ${event.message}`));
  });
  return instance;
}

function sendToWorker(
  request: PyodideWorkerRequest,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<PyodideWorkerResponse> {
  return new Promise<PyodideWorkerResponse>((resolve, reject) => {
    const instance = getOrCreateWorker();
    const abort = (): void => {
      cleanup();
      reject(new ToolStoppedError());
      if (instance.phase.kind === "running" && instance.phase.pending.request.id === request.id) {
        terminateInstance(instance, new Error("Aborted"));
      } else {
        const index = instance.queue.findIndex((pending) => pending.request.id === request.id);
        if (index >= 0) instance.queue.splice(index, 1)[0]?.reject(new Error("Aborted"));
      }
    };
    const cleanup = (): void => signal?.removeEventListener("abort", abort);
    instance.queue.push({
      request, timeoutMs,
      resolve: (response) => { cleanup(); resolve(response); },
      reject: (error) => { cleanup(); reject(error); },
    });
    signal?.addEventListener("abort", abort, { once: true });
    if (signal?.aborted) abort();
    dispatchNext(instance);
  });
}

export async function callPyodideRuntime(
  request: PythonBridgeRequest,
  signal?: AbortSignal,
): Promise<PythonBridgeResponse> {
  if (signal?.aborted) throw new ToolStoppedError();

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

  try {
    const workerResponse = await sendToWorker(workerRequest, timeoutMs, signal);

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
    if (error instanceof ToolStoppedError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    return {
      ok: false,
      action: "run_python",
      error: message,
    };
  }
}

export function isPyodideAvailable(): boolean {
  if (typeof Worker === "undefined") return false;
  if (typeof WebAssembly === "undefined") return false;
  return true;
}

export function isPyodideLoaded(): boolean {
  return runtime.kind === "active" && runtime.instance.loaded;
}

export function terminatePyodide(): void {
  if (runtime.kind === "active") terminateInstance(runtime.instance, new Error("Pyodide runtime terminated"));
}
