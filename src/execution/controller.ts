/** Runtime controller for persisted execution mode state. */

import {
  dispatchExecutionModeChanged,
  getStoredExecutionMode,
  setStoredExecutionMode,
  type ExecutionMode,
  type ExecutionModeStore,
} from "./mode.js";

export interface ExecutionModeController {
  /** Last known mode (for display). */
  getMode: () => ExecutionMode;
  /**
   * The mode as stored now. The mode is one setting for every taskpane on
   * this computer, so approval decisions use this, never a value cached when
   * another taskpane may have changed it since.
   */
  readMode: () => Promise<ExecutionMode>;
  setMode: (mode: ExecutionMode) => Promise<void>;
}

export interface CreateExecutionModeControllerOptions {
  settings: ExecutionModeStore;
}

export async function createExecutionModeController(
  options: CreateExecutionModeControllerOptions,
): Promise<ExecutionModeController> {
  let mode = await getStoredExecutionMode(options.settings);
  const channel = options.channel === undefined ? defaultChannel() : options.channel;

  const readMode = async (): Promise<ExecutionMode> => {
    const stored = await getStoredExecutionMode(options.settings);
    if (stored !== mode) {
      mode = stored;
      dispatchExecutionModeChanged(mode);
    }
    return mode;
  };

  // Another taskpane changed the mode: show it here too.
  channel?.addEventListener("message", () => {
    void readMode().catch(() => {});
  });

  const applyMode = async (nextMode: ExecutionMode): Promise<void> => {
    if (mode === nextMode) {
      return;
    }

    mode = await setStoredExecutionMode(options.settings, nextMode);
    dispatchExecutionModeChanged(mode);
    channel?.postMessage(mode);
  };

  return {
    getMode: () => mode,
    readMode,
    setMode: async (nextMode: ExecutionMode) => {
      await applyMode(nextMode);
    },
  };
}
