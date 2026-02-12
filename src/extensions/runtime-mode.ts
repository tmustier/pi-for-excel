/**
 * Extension runtime mode resolution.
 */

import type { StoredExtensionTrust } from "./permissions.js";

export type ExtensionRuntimeMode = "host" | "sandbox-iframe";

export function isSandboxCandidateTrust(trust: StoredExtensionTrust): boolean {
  return trust === "inline-code" || trust === "remote-url";
}

export function resolveExtensionRuntimeMode(
  trust: StoredExtensionTrust,
  sandboxRuntimeEnabled: boolean,
): ExtensionRuntimeMode {
  if (!sandboxRuntimeEnabled) {
    return "host";
  }

  return isSandboxCandidateTrust(trust) ? "sandbox-iframe" : "host";
}

export function describeExtensionRuntimeMode(mode: ExtensionRuntimeMode): string {
  return mode === "sandbox-iframe" ? "sandbox iframe" : "host runtime";
}
