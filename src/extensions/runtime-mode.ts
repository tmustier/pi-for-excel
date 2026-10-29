import type { StoredExtensionTrust } from "./permissions.js";

export type ExtensionRuntimeMode = "host" | "sandbox-iframe";

export function isSandboxCandidateTrust(trust: StoredExtensionTrust): boolean {
  return trust === "inline-code" || trust === "remote-url";
}

export function resolveExtensionRuntimeMode(
  trust: StoredExtensionTrust,
  sandboxHostFallbackEnabled: boolean,
): ExtensionRuntimeMode {
  if (!isSandboxCandidateTrust(trust)) {
    return "host";
  }

  return sandboxHostFallbackEnabled ? "host" : "sandbox-iframe";
}

export function describeExtensionRuntimeMode(mode: ExtensionRuntimeMode): string {
  return mode === "sandbox-iframe" ? "sandbox iframe" : "host runtime";
}
