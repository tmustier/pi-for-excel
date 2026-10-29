import { render } from "lit";

import { installFetchInterceptor } from "../auth/cors-proxy.js";
import { installProcessEnvShim } from "../compat/process-env-shim.js";
import {
  resolveSpreadsheetHostForBoot,
  setCurrentSpreadsheetHost,
} from "../host/index.js";
import { t } from "../language/index.js";
import {
  STORAGE_UPGRADE_BLOCKED_EVENT,
  STORAGE_VERSION_CHANGED_EVENT,
} from "../storage/local/indexeddb-storage-backend.js";
import { renderLoading, renderError } from "../ui/loading.js";
import { getErrorMessage } from "../utils/errors.js";

import { initTaskpane } from "./init.js";

function getRequiredElement(id: string): HTMLElement {
  const el = document.getElementById(id);
  if (!el) {
    throw new Error(`[pi] Missing required element #${id}`);
  }
  return el;
}

function showFatalError(errorRoot: HTMLElement, message: string): void {
  render(renderError(message), errorRoot);
}

export function bootstrapTaskpane(): void {
  const appEl = getRequiredElement("app");
  const loadingRoot = getRequiredElement("loading-root");
  const errorRoot = getRequiredElement("error-root");

  render(renderLoading(), loadingRoot);

  installProcessEnvShim();
  installFetchInterceptor();

  window.addEventListener(STORAGE_UPGRADE_BLOCKED_EVENT, () => {
    render(renderLoading(t("bootstrap.storageBlocked")), loadingRoot);
  });
  window.addEventListener(STORAGE_VERSION_CHANGED_EVENT, () => {
    showFatalError(errorRoot, t("bootstrap.storageVersionChanged"));
  });

  let initialized = false;

  const runInit = () => {
    if (initialized) return;

    initialized = true;

    const slowInitTimer = setTimeout(() => {
      render(renderLoading(t("bootstrap.slowStart")), loadingRoot);
      console.warn("[pi] Taskpane initialization taking longer than expected (>12s)");
    }, 12_000);

    const verySlowInitTimer = setTimeout(() => {
      render(renderLoading(t("bootstrap.stillStarting")), loadingRoot);
      console.error("[pi] Taskpane initialization still running after 60s");
    }, 60_000);

    void initTaskpane({
      appEl,
      errorRoot,
      autoResizeImages: import.meta.env.VITE_PI_AUTO_RESIZE_IMAGES !== "false",
    })
      .catch((error: unknown) => {
        loadingRoot.replaceChildren();
        showFatalError(errorRoot, t("bootstrap.fatalError", { msg: getErrorMessage(error) }));
        console.error("[pi] Init error:", error);
      })
      .finally(() => {
        clearTimeout(slowInitTimer);
        clearTimeout(verySlowInitTimer);
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
    .catch((error: unknown) => {
      console.warn("[pi] Host detection failed — initializing without Excel:", error);
      runInit();
    });
}
