export const PI_INTEGRATIONS_CHANGED_EVENT = "pi:integrations-changed";

export interface IntegrationsChangedDetail {
  reason: "toggle" | "scope" | "external-toggle" | "config";
}

export function dispatchIntegrationsChanged(detail: IntegrationsChangedDetail): void {
  if (typeof document === "undefined") return;

  document.dispatchEvent(
    new CustomEvent<IntegrationsChangedDetail>(PI_INTEGRATIONS_CHANGED_EVENT, { detail }),
  );
}
