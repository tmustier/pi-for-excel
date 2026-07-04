/**
 * Client address policy for the CORS proxy server.
 *
 * By default the proxy only accepts loopback clients. For org-hosted central
 * deployments (see docs/central-proxy.md), operators can opt in to additional
 * IPv4 client ranges via the ALLOWED_CLIENT_CIDRS env var.
 *
 * SECURITY: this module is deliberately fail-closed:
 * - invalid CIDR entries are reported, never silently ignored
 * - 0.0.0.0/0 ("allow everyone") is rejected — use network-level controls
 * - IPv6 ranges are not supported; IPv4-mapped addresses (::ffff:a.b.c.d)
 *   are normalized to their IPv4 form before matching
 */

function ipv4ToInt(ip) {
  const parts = ip.split(".");
  if (parts.length !== 4) return null;
  let value = 0;
  for (const part of parts) {
    if (!/^\d{1,3}$/.test(part)) return null;
    const octet = Number(part);
    if (octet > 255) return null;
    value = value * 256 + octet;
  }
  return value;
}

export function isLoopbackAddress(addr) {
  if (typeof addr !== "string" || addr.length === 0) return false;
  const lower = addr.toLowerCase();
  if (lower === "::1" || lower === "0:0:0:0:0:0:0:1") return true;

  // Strictly parse IPv4 (optionally IPv4-mapped IPv6) and check 127/8.
  // Prefix-string matching would accept junk like "127.evil" — harmless for
  // Node-provided socket addresses, but this helper must stay safe if ever
  // reused with less-trusted address sources.
  const cleaned = lower.startsWith("::ffff:") ? lower.slice(7) : lower;
  const value = ipv4ToInt(cleaned);
  if (value === null) return false;
  return Math.floor(value / 2 ** 24) === 127;
}

/**
 * Parse an ALLOWED_CLIENT_CIDRS value ("10.96.0.0/13, 192.168.1.5") into
 * matchable ranges. Bare IPv4 addresses are treated as /32.
 *
 * Returns { cidrs, invalid } — callers should treat any invalid entry as a
 * fatal configuration error (fail closed).
 */
export function parseClientCidrAllowlist(raw) {
  const cidrs = [];
  const invalid = [];

  const entries = String(raw ?? "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  for (const entry of entries) {
    const [ipPart, bitsPart, ...rest] = entry.split("/");
    if (rest.length > 0) {
      invalid.push(entry);
      continue;
    }

    const base = ipv4ToInt(ipPart);
    if (base === null) {
      invalid.push(entry);
      continue;
    }

    let bits = 32;
    if (bitsPart !== undefined) {
      if (!/^\d{1,2}$/.test(bitsPart)) {
        invalid.push(entry);
        continue;
      }
      bits = Number(bitsPart);
    }

    // Reject "allow everyone" and out-of-range prefixes. A /0 would turn the
    // proxy into an open relay; operators wanting that must use network-level
    // controls they consciously own.
    if (bits < 1 || bits > 32) {
      invalid.push(entry);
      continue;
    }

    cidrs.push({ entry, base, bits });
  }

  return { cidrs, invalid };
}

function normalizeClientAddress(addr) {
  if (typeof addr !== "string") return null;
  const lower = addr.toLowerCase();
  const cleaned = lower.startsWith("::ffff:") ? lower.slice(7) : lower;
  return ipv4ToInt(cleaned) === null ? null : cleaned;
}

/**
 * Returns true when the client address is loopback or matches one of the
 * parsed CIDR ranges. Non-IPv4 (and unparseable) addresses only pass the
 * loopback check.
 */
export function isAllowedClientAddress(addr, cidrs = []) {
  if (isLoopbackAddress(addr)) return true;
  if (cidrs.length === 0) return false;

  const normalized = normalizeClientAddress(addr);
  if (normalized === null) return false;

  const value = ipv4ToInt(normalized);
  if (value === null) return false;

  for (const { base, bits } of cidrs) {
    const shift = 32 - bits;
    // Compare network prefixes without bitwise ops to avoid 32-bit sign issues.
    if (Math.floor(value / 2 ** shift) === Math.floor(base / 2 ** shift)) {
      return true;
    }
  }

  return false;
}
