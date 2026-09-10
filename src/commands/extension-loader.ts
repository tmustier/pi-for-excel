function isCommandsExtensionLoaderPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}


export type ExtensionCleanup = () => void | Promise<void>;
export type ExtensionActivateResult = void | ExtensionCleanup | readonly ExtensionCleanup[];
export type ExtensionActivator<TApi> = (
  api: TApi,
) => ExtensionActivateResult | Promise<ExtensionActivateResult>;
export type ExtensionDeactivator = () => void | Promise<void>;

export interface LoadedExtensionHandle {
  deactivate: () => Promise<void>;
}

function isExtensionActivator<TApi>(value: unknown): value is ExtensionActivator<TApi> {
  return typeof value === "function";
}

function isExtensionDeactivator(value: unknown): value is ExtensionDeactivator {
  return typeof value === "function";
}

function isExtensionCleanup(value: unknown): value is ExtensionCleanup {
  return typeof value === "function";
}

export function getExtensionActivator<TApi>(mod: unknown): ExtensionActivator<TApi> | null {
  if (!isCommandsExtensionLoaderPayloadShape(mod)) return null;

  const activate = mod.activate;
  if (isExtensionActivator<TApi>(activate)) {
    return activate;
  }

  const fallback = mod.default;
  if (isExtensionActivator<TApi>(fallback)) {
    return fallback;
  }

  return null;
}

export function getExtensionDeactivator(mod: unknown): ExtensionDeactivator | null {
  if (!isCommandsExtensionLoaderPayloadShape(mod)) return null;

  const deactivate = mod.deactivate;
  return isExtensionDeactivator(deactivate) ? deactivate : null;
}

function toErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  return String(error);
}

export function collectActivationCleanups(result: unknown): ExtensionCleanup[] {
  if (typeof result === "undefined") {
    return [];
  }

  if (isExtensionCleanup(result)) {
    return [result];
  }

  if (!Array.isArray(result)) {
    throw new Error("activate(api) must return void, a cleanup function, or an array of cleanup functions");
  }

  const cleanups: ExtensionCleanup[] = [];
  for (const value of result) {
    if (!isExtensionCleanup(value)) {
      throw new Error("activate(api) returned an invalid cleanup entry; expected a function");
    }

    cleanups.push(value);
  }

  return cleanups;
}

export function createLoadedExtensionHandle(
  cleanups: readonly ExtensionCleanup[],
  moduleDeactivate: ExtensionDeactivator | null,
): LoadedExtensionHandle {
  let deactivated = false;

  return {
    deactivate: async () => {
      if (deactivated) {
        return;
      }
      deactivated = true;

      const failures: string[] = [];

      for (let i = cleanups.length - 1; i >= 0; i -= 1) {
        const cleanup = cleanups[i];
        if (!cleanup) {
          continue;
        }

        try {
          await cleanup();
        } catch (error) {
          failures.push(toErrorMessage(error));
        }
      }

      if (moduleDeactivate) {
        try {
          await moduleDeactivate();
        } catch (error) {
          failures.push(toErrorMessage(error));
        }
      }

      if (failures.length > 0) {
        throw new Error(`Extension cleanup failed:\n- ${failures.join("\n- ")}`);
      }
    },
  };
}
