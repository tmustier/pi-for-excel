export interface PiAuthRequestPolicyInput {
  remoteAddress?: string;
  hostHeader?: string | string[];
  allowNonLocalHost?: boolean;
}

export function isLoopbackAddress(addr: string | undefined): boolean {
  if (!addr) return false;
  if (addr === "::1" || addr === "0:0:0:0:0:0:0:1") return true;
  if (addr.startsWith("127.")) return true;
  if (addr.startsWith("::ffff:127.")) return true;
  return false;
}

export function hostnameFromHostHeader(hostHeader: string | string[] | undefined): string | null {
  const raw = Array.isArray(hostHeader) ? hostHeader[0] : hostHeader;
  const trimmed = raw?.trim().toLowerCase();
  if (!trimmed) return null;

  if (trimmed.startsWith("[")) {
    const end = trimmed.indexOf("]");
    return end > 1 ? trimmed.slice(1, end) : null;
  }

  const colon = trimmed.indexOf(":");
  return colon >= 0 ? trimmed.slice(0, colon) : trimmed;
}

export function isLocalPiAuthHost(hostHeader: string | string[] | undefined): boolean {
  const hostname = hostnameFromHostHeader(hostHeader);
  return hostname === "localhost" || hostname === "127.0.0.1" || hostname === "::1";
}

export function isPiAuthRequestAllowed(input: PiAuthRequestPolicyInput): boolean {
  if (!isLoopbackAddress(input.remoteAddress)) return false;
  if (input.allowNonLocalHost === true) return true;
  return isLocalPiAuthHost(input.hostHeader);
}
