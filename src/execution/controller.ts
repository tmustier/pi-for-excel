/** Runtime controller for persisted execution mode state. */

import {
  dispatchExecutionModeChanged,
  formatExecutionModeLabel,
  getStoredExecutionMode,
  setStoredExecutionMode,
  toggleExecutionMode,
  type ExecutionMode,
  type ExecutionModeStore,
} from "./mode.js";

export interface ExecutionModeController {
  getMode: () => ExecutionMode;
  setMode: (mode: ExecutionMode) => Promise<void>;
  toggleFromUi: () => Promise<void>;
}

export interface CreateExecutionModeControllerOptions {
  settings: ExecutionModeStore;
  showToast?: (message: string) => void;
}

export async function createExecutionModeController(
  options: CreateExecutionModeControllerOptions,
): Promise<ExecutionModeController> {
  let mode = await getStoredExecutionMode(options.settings);

  const applyMode = async (nextMode: ExecutionMode): Promise<void> => {
    if (mode === nextMode) {
      return;
    }

    mode = await setStoredExecutionMode(options.settings, nextMode);
    dispatchExecutionModeChanged(mode);
  };

  return {
    getMode: () => mode,
    setMode: async (nextMode: ExecutionMode) => {
      await applyMode(nextMode);
    },
    toggleFromUi: async () => {
      const nextMode = toggleExecutionMode(mode);
      await applyMode(nextMode);
      options.showToast?.(`${formatExecutionModeLabel(nextMode)} mode.`);
    },
  };
}
