/** Execution mode persistence + helpers (YOLO vs Safe). */

export const EXECUTION_MODE_SETTING_KEY = "execution.mode.v1";
export const PI_EXECUTION_MODE_CHANGED_EVENT = "pi:execution-mode-changed";

export type ExecutionMode = "yolo" | "safe";

export interface ExecutionModeStore {
  get: (key: string) => Promise<unknown>;
  set: (key: string, value: unknown) => Promise<void>;
}

const EXECUTION_MODE_VALUES = new Set<ExecutionMode>(["yolo", "safe"]);

export function isExecutionMode(value: unknown): value is ExecutionMode {
  return typeof value === "string" && EXECUTION_MODE_VALUES.has(value as ExecutionMode);
}

export function normalizeExecutionMode(value: unknown): ExecutionMode {
  return isExecutionMode(value) ? value : "yolo";
}

export async function getStoredExecutionMode(store: ExecutionModeStore): Promise<ExecutionMode> {
  const value = await store.get(EXECUTION_MODE_SETTING_KEY);
  return normalizeExecutionMode(value);
}

export async function setStoredExecutionMode(
  store: ExecutionModeStore,
  mode: ExecutionMode,
): Promise<ExecutionMode> {
  await store.set(EXECUTION_MODE_SETTING_KEY, mode);
  return mode;
}

export function toggleExecutionMode(mode: ExecutionMode): ExecutionMode {
  return mode === "yolo" ? "safe" : "yolo";
}

export function formatExecutionModeLabel(mode: ExecutionMode): string {
  return mode === "yolo" ? "YOLO" : "Safe";
}

export function dispatchExecutionModeChanged(mode: ExecutionMode): void {
  if (typeof document === "undefined") {
    return;
  }

  document.dispatchEvent(new CustomEvent(PI_EXECUTION_MODE_CHANGED_EVENT, {
    detail: {
      mode,
    },
  }));

  document.dispatchEvent(new CustomEvent("pi:status-update"));
}
