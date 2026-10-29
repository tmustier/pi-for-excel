import net from "node:net";

const LOOPBACK_IPV4 = new net.BlockList();
LOOPBACK_IPV4.addSubnet("127.0.0.0", 8, "ipv4");

const LOOPBACK_IPV6 = new net.BlockList();
LOOPBACK_IPV6.addAddress("::1", "ipv6");

const PRIVATE_LOCAL_IPV4 = new net.BlockList();
PRIVATE_LOCAL_IPV4.addSubnet("10.0.0.0", 8, "ipv4");
PRIVATE_LOCAL_IPV4.addSubnet("172.16.0.0", 12, "ipv4");
PRIVATE_LOCAL_IPV4.addSubnet("192.168.0.0", 16, "ipv4");
PRIVATE_LOCAL_IPV4.addSubnet("169.254.0.0", 16, "ipv4");
PRIVATE_LOCAL_IPV4.addSubnet("127.0.0.0", 8, "ipv4");

const PRIVATE_LOCAL_IPV6 = new net.BlockList();
PRIVATE_LOCAL_IPV6.addAddress("::1", "ipv6");
PRIVATE_LOCAL_IPV6.addSubnet("fc00::", 7, "ipv6");
PRIVATE_LOCAL_IPV6.addSubnet("fe80::", 10, "ipv6");

const IPV4_MAPPED_IPV6_RE = /^::ffff:(\d{1,3}(?:\.\d{1,3}){3})$/i;

export function normalizeHost(hostname) {
  if (typeof hostname !== "string") return "";

  let host = hostname.trim().toLowerCase();
  if (!host) return "";

  if (host.startsWith("[") && host.endsWith("]")) {
    host = host.slice(1, -1);
  }

  const zoneIndex = host.indexOf("%");
  if (zoneIndex >= 0) {
    host = host.slice(0, zoneIndex);
  }

  return host;
}

function mappedIPv4FromIPv6(hostname) {
  const host = normalizeHost(hostname);
  const match = IPV4_MAPPED_IPV6_RE.exec(host);
  if (!match) return null;

  const candidate = match[1];
  return net.isIP(candidate) === 4 ? candidate : null;
}

function checkBlockList(list, ip) {
  const family = net.isIP(ip);
  if (family === 4) return list.check(ip, "ipv4");
  if (family === 6) return list.check(ip, "ipv6");
  return false;
}

export function parseAllowedTargetHosts(raw) {
  if (typeof raw !== "string" || raw.trim().length === 0) {
    return new Set();
  }

  const out = new Set();
  for (const entry of raw.split(",")) {
    const trimmed = entry.trim();
    if (!trimmed) continue;

    let host = "";

    if (trimmed.includes("://")) {
      try {
        host = normalizeHost(new URL(trimmed).hostname);
      } catch {
        host = "";
      }
    } else {
      host = normalizeHost(trimmed);
    }

    if (host) out.add(host);
  }

  return out;
}

export function isIpLiteral(hostname) {
  const host = normalizeHost(hostname);
  if (!host) return false;
  return net.isIP(host) !== 0;
}

export function isLoopbackHostname(hostname) {
  const host = normalizeHost(hostname);
  if (!host) return false;

  if (host === "localhost") return true;

  if (checkBlockList(LOOPBACK_IPV4, host)) return true;
  if (checkBlockList(LOOPBACK_IPV6, host)) return true;

  const mapped = mappedIPv4FromIPv6(host);
  return mapped !== null && checkBlockList(LOOPBACK_IPV4, mapped);
}

export function isPrivateOrLocalIp(ip) {
  const host = normalizeHost(ip);
  if (!host) return false;

  if (checkBlockList(PRIVATE_LOCAL_IPV4, host)) return true;
  if (checkBlockList(PRIVATE_LOCAL_IPV6, host)) return true;

  const mapped = mappedIPv4FromIPv6(host);
  return mapped !== null && checkBlockList(PRIVATE_LOCAL_IPV4, mapped);
}

export function isAllowedTargetHost(hostname, allowedHosts) {
  const host = normalizeHost(hostname);
  if (!host) return false;

  if (!(allowedHosts instanceof Set) || allowedHosts.size === 0) {
    return true;
  }

  return allowedHosts.has(host);
}

export function getBlockedTargetReasonForHostname(hostname, opts = {}) {
  const {
    allowLoopbackTargets = false,
    allowPrivateTargets = false,
    allowedHosts = new Set(),
    requireAllowlistForOverriddenTargets = false,
  } = opts;

  const host = normalizeHost(hostname);
  if (!host) return "blocked_target_invalid_host";

  const loopback = isLoopbackHostname(host);
  if (loopback && !allowLoopbackTargets) {
    return "blocked_target_loopback";
  }

  if (loopback && allowLoopbackTargets) {
    if (requireAllowlistForOverriddenTargets && !isAllowedTargetHost(host, allowedHosts)) {
      return "blocked_target_not_allowlisted";
    }
    return null;
  }

  const privateOrLocalLiteral = isIpLiteral(host) && isPrivateOrLocalIp(host);
  if (!allowPrivateTargets && privateOrLocalLiteral) {
    return "blocked_target_private_ip";
  }

  if (allowPrivateTargets && privateOrLocalLiteral) {
    if (requireAllowlistForOverriddenTargets && !isAllowedTargetHost(host, allowedHosts)) {
      return "blocked_target_not_allowlisted";
    }
    return null;
  }

  if (!isAllowedTargetHost(host, allowedHosts)) {
    return "blocked_target_not_allowlisted";
  }

  return null;
}

export function getBlockedTargetReasonForResolvedIps(resolvedIps, opts = {}) {
  const {
    allowLoopbackTargets = false,
    allowPrivateTargets = false,
  } = opts;

  if (!Array.isArray(resolvedIps) || resolvedIps.length === 0) {
    return null;
  }

  for (const ip of resolvedIps) {
    const normalized = normalizeHost(ip);
    if (!normalized) continue;

    const loopback = isLoopbackHostname(normalized);
    if (loopback && !allowLoopbackTargets) {
      return "blocked_target_loopback";
    }

    if (loopback && allowLoopbackTargets) {
      continue;
    }

    if (!allowPrivateTargets && isPrivateOrLocalIp(normalized)) {
      return "blocked_target_private_ip";
    }
  }

  return null;
}

export function evaluateTargetHostPolicy(opts = {}) {
  const {
    hostname,
    resolvedIps = [],
    allowLoopbackTargets = false,
    allowPrivateTargets = false,
    allowedHosts = new Set(),
    requireAllowlistForOverriddenTargets = false,
  } = opts;

  const hostReason = getBlockedTargetReasonForHostname(hostname, {
    allowLoopbackTargets,
    allowPrivateTargets,
    allowedHosts,
    requireAllowlistForOverriddenTargets,
  });

  if (hostReason) {
    return { allowed: false, reason: hostReason };
  }

  const dnsReason = getBlockedTargetReasonForResolvedIps(resolvedIps, {
    allowLoopbackTargets,
    allowPrivateTargets,
  });

  if (dnsReason) {
    return { allowed: false, reason: dnsReason };
  }

  return { allowed: true };
}

export function isBlockedTargetByHostname(hostname) {
  return getBlockedTargetReasonForHostname(hostname) !== null;
}
