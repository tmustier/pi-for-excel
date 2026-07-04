/** Boot-time host resolution. */

import { BrowserHost } from "./browser-host.js";
import { detectSpreadsheetHost } from "./detection.js";
import { OfficeHost } from "./office-host.js";
import type { SpreadsheetHost, SpreadsheetHostReadyInfo } from "./types.js";
import { WpsHost } from "./wps-host.js";

export interface SpreadsheetHostBootResult {
  host: SpreadsheetHost;
  readyInfo: SpreadsheetHostReadyInfo;
}

export interface ResolveSpreadsheetHostForBootOptions {
  officeReadyTimeoutMs?: number;
}

function browserReadyInfo(reason: "office-unavailable"): SpreadsheetHostReadyInfo {
  return {
    kind: "browser",
    nativeHost: null,
    nativePlatform: null,
    reason,
  };
}

function officeTimeoutReadyInfo(): SpreadsheetHostReadyInfo {
  return {
    kind: "office",
    nativeHost: null,
    nativePlatform: null,
    reason: "office-timeout",
  };
}

export async function resolveSpreadsheetHostForBoot(
  options: ResolveSpreadsheetHostForBootOptions = {},
): Promise<SpreadsheetHostBootResult> {
  const detected = detectSpreadsheetHost();

  if (detected === "wps") {
    const host = new WpsHost();
    return { host, readyInfo: await host.whenReady() };
  }

  if (detected === "browser") {
    const host = new BrowserHost();
    return { host, readyInfo: browserReadyInfo("office-unavailable") };
  }

  const timeoutMs = options.officeReadyTimeoutMs ?? 3000;
  const officeHost = new OfficeHost();

  return new Promise<SpreadsheetHostBootResult>((resolve) => {
    let settled = false;
    let timeout: ReturnType<typeof setTimeout> | null = null;

    const finish = (result: SpreadsheetHostBootResult): void => {
      if (settled) return;
      settled = true;
      if (timeout !== null) {
        clearTimeout(timeout);
      }
      resolve(result);
    };

    timeout = setTimeout(() => {
      // Keep the Office host: before the host seam existed, a slow
      // `Office.onReady` only delayed init while workbook identity and theme
      // were still read lazily from Office globals at call time. Pinning the
      // browser host here would permanently drop workbook identity for slow
      // Office startups, which would be a behavior change.
      finish({
        host: officeHost,
        readyInfo: officeTimeoutReadyInfo(),
      });
    }, timeoutMs);

    void officeHost.whenReady()
      .then((readyInfo) => {
        finish({ host: officeHost, readyInfo });
      })
      .catch(() => {
        // Match pre-host-seam behavior: a failing `Office.onReady` never
        // resolved early; init happened at the timeout with Office globals
        // still read lazily afterwards.
      });
  });
}
