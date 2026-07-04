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
import { renderLoading, renderError } from "../ui/loading.js";
import { getErrorMessage } from "../utils/errors.js";
import { t } from "../language/index.js";

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

  // Office bootstrap (with fallback for local dev)
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
      .catch((error: unknown) => {
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

  if (typeof Office === "undefined") {
    console.warn(t("bootstrap.officeUnavailable"));
    runInit();
    return;
  }

  void Office.onReady((info) => {
    console.log(`[pi] Office.js ready: host=${info.host}, platform=${info.platform}`);
    runInit();
  });

  setTimeout(() => {
    if (initialized) return;

    console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
    runInit();
  }, 3000);
}
