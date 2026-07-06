function isExtensionsSandboxProtocolPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}


export const SANDBOX_CHANNEL = "pi.extension.sandbox.rpc.v1";
export const SANDBOX_REQUEST_TIMEOUT_MS = 15_000;
export const SANDBOX_BOOTSTRAP_KIND = "bootstrap";

export type SandboxDirection = "sandbox_to_host" | "host_to_sandbox";

export type SandboxEnvelopeKind = "request" | "response" | "event";

interface SandboxEnvelopeBase {
  channel: string;
  instanceId: string;
  direction: SandboxDirection;
  kind: SandboxEnvelopeKind;
}

export interface SandboxBootstrapEnvelope {
  channel: string;
  instanceId: string;
  direction: "host_to_sandbox";
  kind: typeof SANDBOX_BOOTSTRAP_KIND;
}

export interface SandboxRequestEnvelope extends SandboxEnvelopeBase {
  kind: "request";
  requestId: string;
  method: string;
  params?: DynamicValue;
}

export interface SandboxResponseEnvelope extends SandboxEnvelopeBase {
  kind: "response";
  requestId: string;
  ok: boolean;
  result?: DynamicValue;
  error?: string;
}

export interface SandboxEventEnvelope extends SandboxEnvelopeBase {
  kind: "event";
  event: string;
  data?: DynamicValue;
}

export type SandboxEnvelope = SandboxRequestEnvelope | SandboxResponseEnvelope | SandboxEventEnvelope;

function hasValidSandboxEnvelopeBase(value: DynamicValue): value is DynamicObject & {
  channel: string;
  instanceId: string;
  direction: SandboxDirection;
  kind: string;
} {
  if (!isExtensionsSandboxProtocolPayloadShape(value)) {
    return false;
  }

  const channel = value.channel;
  const instanceId = value.instanceId;
  const direction = value.direction;
  const kind = value.kind;

  if (channel !== SANDBOX_CHANNEL) {
    return false;
  }

  if (typeof instanceId !== "string") {
    return false;
  }

  if (direction !== "sandbox_to_host" && direction !== "host_to_sandbox") {
    return false;
  }

  return typeof kind === "string";
}

export function isSandboxBootstrapEnvelope(value: DynamicValue): value is SandboxBootstrapEnvelope {
  if (!hasValidSandboxEnvelopeBase(value)) {
    return false;
  }

  return value.direction === "host_to_sandbox" && value.kind === SANDBOX_BOOTSTRAP_KIND;
}

export function isSandboxEnvelope(value: DynamicValue): value is SandboxEnvelope {
  if (!hasValidSandboxEnvelopeBase(value)) {
    return false;
  }

  const kind = value.kind;
  if (kind !== "request" && kind !== "response" && kind !== "event") {
    return false;
  }

  if (kind === "request") {
    return typeof value.requestId === "string" && typeof value.method === "string";
  }

  if (kind === "response") {
    return typeof value.requestId === "string" && typeof value.ok === "boolean";
  }

  return typeof value.event === "string";
}

export function serializeForSandboxInlineScript(value: DynamicValue): string {
  return JSON.stringify(value).replace(/</g, "\\u003c");
}
