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

function defaultChannel(): Pick<BroadcastChannel, "postMessage" | "addEventListener"> | null {
  return typeof BroadcastChannel === "undefined" ? null : new BroadcastChannel(EXECUTION_MODE_CHANNEL);
}

export async function createExecutionModeController(
  options: CreateExecutionModeControllerOptions,
): Promise<ExecutionModeController> {
  let mode = await getStoredExecutionMode(options.settings);
  const channel = options.channel === undefined ? defaultChannel() : options.channel;

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

  const applyMode = (nextMode: ExecutionMode): Promise<void> => {
    const currentGeneration = ++generation;
    return serialize(async () => {
      const stored = await setStoredExecutionMode(options.settings, nextMode);
      if (currentGeneration === generation) {
        mode = stored;
        dispatchExecutionModeChanged(mode);
      }
      channel?.postMessage(stored);
    });
  };

  return {
    getMode: () => mode,
    readMode,
    setMode: async (nextMode: ExecutionMode) => {
      await applyMode(nextMode);
    },
  };
}
