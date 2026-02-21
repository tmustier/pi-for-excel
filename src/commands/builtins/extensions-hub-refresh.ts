export interface DeferredConnectionsRefreshController {
  requestRefresh: () => void;
  onConnectionsFocusOut: () => void;
  dispose: () => void;
}

export function createDeferredConnectionsRefreshController(args: {
  isDisposed: () => boolean;
  hasActiveSecretInput: () => boolean;
  refresh: () => void;
}): DeferredConnectionsRefreshController {
  let pendingRefresh = false;
  let focusOutTimer: ReturnType<typeof setTimeout> | null = null;

  const clearFocusOutTimer = (): void => {
    if (focusOutTimer === null) return;
    clearTimeout(focusOutTimer);
    focusOutTimer = null;
  };

  const flushIfReady = (): void => {
    if (!pendingRefresh || args.isDisposed()) return;
    if (args.hasActiveSecretInput()) return;

    pendingRefresh = false;
    args.refresh();
  };

  return {
    requestRefresh: () => {
      if (args.isDisposed()) return;

      if (args.hasActiveSecretInput()) {
        pendingRefresh = true;
        return;
      }

      pendingRefresh = false;
      args.refresh();
    },
    onConnectionsFocusOut: () => {
      clearFocusOutTimer();
      focusOutTimer = setTimeout(() => {
        focusOutTimer = null;
        flushIfReady();
      }, 0);
    },
    dispose: () => {
      pendingRefresh = false;
      clearFocusOutTimer();
    },
  };
}
