/**
 * Taskpane bootstrap.
 *
 * Runs immediately when the add-in loads:
 * - renders loading UI
 * - installs global fetch + UI monkey patches
 * - waits for Office.onReady (with fallback) and then initializes the app
 */

import { render } from "lit";

import { installFetchInterceptor } from "../auth/cors-proxy.js";
import { installModelSelectorPatch } from "../compat/model-selector-patch.js";
import { installProcessEnvShim } from "../compat/process-env-shim.js";
import {
  resolveSpreadsheetHostForBoot,
  setCurrentSpreadsheetHost,
} from "../host/index.js";
import { t } from "../language/index.js";
import { renderLoading, renderError } from "../ui/loading.js";
import { getErrorMessage } from "../utils/errors.js";

import { initTaskpane } from "./init.js";

function getRequiredElement<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) {
    throw new Error(`[pi] Missing required element #${id}`);
  }
  return el as T;
}

function showFatalError(errorRoot: HTMLElement, message: string): void {
  render(renderError(message), errorRoot);
}

export function bootstrapTaskpane(): void {
  const appEl = getRequiredElement<HTMLElement>("app");
  const loadingRoot = getRequiredElement<HTMLElement>("loading-root");
  const errorRoot = getRequiredElement<HTMLElement>("error-root");

  // Initial loading UI
  render(renderLoading(), loadingRoot);

  // Global patches
  installProcessEnvShim();
  installFetchInterceptor();
  installModelSelectorPatch();

  // Host bootstrap (Office/WPS/browser fallback for local dev)
  let initialized = false;

  const runInit = () => {
    if (initialized) return;

    initialized = true;

    let initComplete = false;

    const markInitComplete = () => {
      if (initComplete) return false;
      initComplete = true;
      return true;
    };

    const slowInitTimer = setTimeout(() => {
      if (initComplete) return;
      console.warn(t("bootstrap.initTimeoutWarning"));
    }, 12_000);

    const hardTimeoutTimer = setTimeout(() => {
      if (!markInitComplete()) return;
      loadingRoot.innerHTML = "";
      showFatalError(
        errorRoot,
        t("bootstrap.fatalTimeout"),
      );
      console.error("[pi] Init error: Taskpane initialization timed out after 60000ms");
    }, 60_000);

    void initTaskpane({ appEl, errorRoot })
      .then(() => {
        if (!markInitComplete()) return;
        clearTimeout(slowInitTimer);
        clearTimeout(hardTimeoutTimer);
      })
      .catch((error: DynamicValue) => {
        if (!markInitComplete()) {
          console.error("[pi] Init error after timeout:", error);
          return;
        }

        clearTimeout(slowInitTimer);
        clearTimeout(hardTimeoutTimer);
        loadingRoot.innerHTML = "";
        showFatalError(errorRoot, t("bootstrap.fatalError", { msg: getErrorMessage(error) }));
        console.error("[pi] Init error:", error);
      });
  };

  void resolveSpreadsheetHostForBoot({ officeReadyTimeoutMs: 3000 })
    .then(({ host, readyInfo }) => {
      setCurrentSpreadsheetHost(host);

      if (readyInfo.reason === "office-ready") {
        const nativeHost = readyInfo.nativeHost ?? "unknown";
        const nativePlatform = readyInfo.nativePlatform ?? "unknown";
        console.log(
          `[pi] Office.js ready: host=${nativeHost}, platform=${nativePlatform}`,
        );
      } else if (readyInfo.reason === "wps-jsapi") {
        console.log("[pi] WPS JSAPI detected — initializing WPS host");
      } else if (readyInfo.reason === "office-timeout") {
        console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
      } else {
        console.warn(t("bootstrap.officeUnavailable"));
      }

      runInit();
    })
    .catch((error: DynamicValue) => {
      console.warn("[pi] Host detection failed — initializing without Excel:", error);
      runInit();
    });
}
