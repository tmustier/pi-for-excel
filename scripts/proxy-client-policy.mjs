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

  const cleaned = lower.startsWith("::ffff:") ? lower.slice(7) : lower;
  const value = ipv4ToInt(cleaned);
  if (value === null) return false;
  return Math.floor(value / 2 ** 24) === 127;
}

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

export function isAllowedClientAddress(addr, cidrs = []) {
  if (isLoopbackAddress(addr)) return true;
  if (cidrs.length === 0) return false;

  const normalized = normalizeClientAddress(addr);
  if (normalized === null) return false;

  const value = ipv4ToInt(normalized);
  if (value === null) return false;

  for (const { base, bits } of cidrs) {
    const shift = 32 - bits;
    if (Math.floor(value / 2 ** shift) === Math.floor(base / 2 ** shift)) {
      return true;
    }
  }

  return false;
}
