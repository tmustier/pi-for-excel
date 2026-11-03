import { createConcurrencyLimit } from "../utils/concurrency.js";
import {
  dispatchExecutionModeChanged,
  getStoredExecutionMode,
  setStoredExecutionMode,
  type ExecutionMode,
  type ExecutionModeStore,
} from "./mode.js";

export interface ExecutionModeController {
  getMode: () => ExecutionMode;
  readMode: () => Promise<ExecutionMode>;
  setMode: (mode: ExecutionMode) => Promise<void>;
}

export interface CreateExecutionModeControllerOptions {
  settings: ExecutionModeStore;
  channel?: Pick<BroadcastChannel, "postMessage" | "addEventListener"> | null;
}

const EXECUTION_MODE_CHANNEL = "pi-execution-mode";

export async function createExecutionModeController(
  options: CreateExecutionModeControllerOptions,
): Promise<ExecutionModeController> {
  let mode = await getStoredExecutionMode(options.settings);
  const channel = options.channel === undefined && typeof BroadcastChannel !== "undefined"
    ? new BroadcastChannel(EXECUTION_MODE_CHANNEL)
    : options.channel;

  const serialize = createConcurrencyLimit(1);
  let generation = 0;

  const readMode = (): Promise<ExecutionMode> => {
    const currentGeneration = ++generation;
    return serialize(async () => {
      const stored = await getStoredExecutionMode(options.settings);
      if (currentGeneration === generation && stored !== mode) {
        mode = stored;
        dispatchExecutionModeChanged(mode);
      }
      return mode;
    });
  };

  channel?.addEventListener("message", () => {
    void readMode().catch(() => {});
  });

  return {
    getMode: () => mode,
    readMode,
    setMode: (nextMode) => {
      generation += 1;
      return serialize(async () => {
        mode = await setStoredExecutionMode(options.settings, nextMode);
        dispatchExecutionModeChanged(mode);
        channel?.postMessage(mode);
      });
    },
  };
}
