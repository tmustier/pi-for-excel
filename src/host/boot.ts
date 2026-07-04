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

function browserReadyInfo(reason: "office-unavailable" | "office-timeout"): SpreadsheetHostReadyInfo {
  return {
    kind: "browser",
    nativeHost: null,
    nativePlatform: null,
    reason,
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
      finish({
        host: new BrowserHost(),
        readyInfo: browserReadyInfo("office-timeout"),
      });
    }, timeoutMs);

    void officeHost.whenReady()
      .then((readyInfo) => {
        finish({ host: officeHost, readyInfo });
      })
      .catch(() => {
        finish({
          host: new BrowserHost(),
          readyInfo: browserReadyInfo("office-unavailable"),
        });
      });
  });
}
