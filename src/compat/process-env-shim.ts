/**
 * Browser runtime shim for libraries that read `process.env` unguarded.
 *
 * Some provider code paths in upstream deps (e.g. Google Antigravity headers)
 * still access `process.env` directly. Office WebViews do not expose `process`.
 */

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

type ProcessShimTarget = { process?: unknown };

export function installProcessEnvShim(target: ProcessShimTarget = globalThis): void {
  const processValue = target.process;

  if (processValue === undefined) {
    target.process = { env: {} };
    return;
  }

  if (!isRecord(processValue)) {
    return;
  }

  const envValue = processValue.env;
  if (!isRecord(envValue)) {
    processValue.env = {};
  }
}
