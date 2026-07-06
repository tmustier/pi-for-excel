/**
 * Browser runtime shim for libraries that read `process.env` unguarded.
 *
 * Some provider code paths in upstream deps (e.g. Google Antigravity headers)
 * still access `process.env` directly. Office WebViews do not expose `process`.
 */

function isCompatProcessEnvShimPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

type ProcessShimTarget = { process?: DynamicValue };

export function installProcessEnvShim(target: ProcessShimTarget = globalThis): void {
  const processValue = target.process;

  if (processValue === undefined) {
    target.process = { env: {} };
    return;
  }

  if (!isCompatProcessEnvShimPayloadShape(processValue)) {
    return;
  }

  const envValue = processValue.env;
  if (!isCompatProcessEnvShimPayloadShape(envValue)) {
    processValue.env = {};
  }
}
