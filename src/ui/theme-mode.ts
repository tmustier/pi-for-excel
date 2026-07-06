function isUiThemeModePayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Dark/light mode synchronization.
 *
 * Dark mode is gated behind /experimental dark-mode.
 * When disabled, UI remains in light mode.
 */

import {
  PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT,
  type ExperimentalFeatureChangedDetail,
} from "../experiments/events.js";
import {
  isExperimentalFeatureEnabled,
  type ExperimentalFeatureId,
} from "../experiments/flags.js";
import {
  createSpreadsheetHost,
  detectSpreadsheetHost,
  getCurrentSpreadsheetHost,
  type SpreadsheetHost,
} from "../host/index.js";

const DARK_MODE_EXPERIMENT_ID: ExperimentalFeatureId = "ui_dark_mode";

function isDarkModeExperimentEnabled(): boolean {
  return isExperimentalFeatureEnabled(DARK_MODE_EXPERIMENT_ID);
}

function resolveHostThemeDark(): boolean | null {
  const currentHost = getCurrentSpreadsheetHost();
  const currentTheme = currentHost.resolveThemeDark();
  if (currentTheme !== null) {
    return currentTheme;
  }

  // Preserve the old Office retry behavior for local browser loads where
  // Office.js appears after this module is installed but before/on ready.
  if (currentHost.kind !== "office" && detectSpreadsheetHost() === "office") {
    return createSpreadsheetHost("office").resolveThemeDark();
  }

  return null;
}

function resolvePreferredDark(mediaMatches: boolean): boolean {
  if (!isDarkModeExperimentEnabled()) {
    return false;
  }

  const hostDark = resolveHostThemeDark();
  if (hostDark !== null) {
    return hostDark;
  }

  return mediaMatches;
}

function isExperimentalFeatureChangedEvent(
  event: Event,
): event is CustomEvent<ExperimentalFeatureChangedDetail> {
  if (!(event instanceof CustomEvent)) {
    return false;
  }

  const detail: DynamicValue = event.detail;
  if (!isUiThemeModePayloadShape(detail)) {
    return false;
  }

  return typeof detail.featureId === "string"
    && typeof detail.enabled === "boolean";
}

function getOfficeReadyHostForTheme(): SpreadsheetHost | null {
  const currentHost = getCurrentSpreadsheetHost();
  if (currentHost.kind === "office") {
    return currentHost;
  }

  if (detectSpreadsheetHost() === "office") {
    return createSpreadsheetHost("office");
  }

  return null;
}

export function installThemeModeSync(): () => void {
  if (typeof document === "undefined" || typeof window === "undefined") {
    return () => {};
  }

  const root = document.documentElement;
  const media = window.matchMedia("(prefers-color-scheme: dark)");

  const apply = () => {
    root.classList.toggle("dark", resolvePreferredDark(media.matches));
  };

  apply();

  let disposed = false;
  let officeReadyHooked = false;
  let officeReadyDispose: (() => void) | null = null;
  let officeRetryTimer: number | null = null;
  let officeRetryStopTimer: number | null = null;

  const registerOfficeReadyHook = (): void => {
    if (officeReadyHooked) {
      return;
    }

    const officeHost = getOfficeReadyHostForTheme();
    if (!officeHost) {
      return;
    }

    officeReadyHooked = true;
    officeReadyDispose = officeHost.onReady(() => {
      if (disposed) return;
      apply();
    });
  };

  registerOfficeReadyHook();

  if (!officeReadyHooked) {
    officeRetryTimer = window.setInterval(() => {
      registerOfficeReadyHook();
      if (officeReadyHooked && officeRetryTimer !== null) {
        clearInterval(officeRetryTimer);
        officeRetryTimer = null;
      }
    }, 500);

    officeRetryStopTimer = window.setTimeout(() => {
      if (officeRetryTimer !== null) {
        clearInterval(officeRetryTimer);
        officeRetryTimer = null;
      }
      officeRetryStopTimer = null;
    }, 15_000);
  }

  const onMediaChange = () => {
    apply();
  };

  const onExperimentalFeatureChange = (event: Event) => {
    if (!isExperimentalFeatureChangedEvent(event)) {
      return;
    }

    if (event.detail.featureId !== DARK_MODE_EXPERIMENT_ID) {
      return;
    }

    apply();
  };

  document.addEventListener(
    PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT,
    onExperimentalFeatureChange,
  );

  const cleanup = (): void => {
    disposed = true;
    document.removeEventListener(
      PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT,
      onExperimentalFeatureChange,
    );
    officeReadyDispose?.();
    if (officeRetryTimer !== null) {
      clearInterval(officeRetryTimer);
    }
    if (officeRetryStopTimer !== null) {
      clearTimeout(officeRetryStopTimer);
    }
  };

  if (typeof media.addEventListener === "function") {
    media.addEventListener("change", onMediaChange);

    return () => {
      cleanup();
      media.removeEventListener("change", onMediaChange);
    };
  }

  media.addListener(onMediaChange);
  return () => {
    cleanup();
    media.removeListener(onMediaChange);
  };
}
