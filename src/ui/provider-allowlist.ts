/**
 * Build-time provider allowlist for org/central deployments.
 *
 * Orgs that only approve specific LLM platforms can set
 * VITE_PI_ALLOWED_PROVIDERS (comma-separated provider ids) so the connect UI
 * only shows approved providers.
 *
 * NOTE: this is a UI-level filter, not a security boundary. Enforcement
 * belongs at the network/proxy layer (ALLOWED_TARGET_HOSTS on the CORS proxy
 * plus your org's egress controls) — see docs/central-proxy.md.
 */

/**
 * Parse a raw allowlist value into a set of provider ids.
 * Returns null when no restriction is configured.
 */
export function resolveAllowedProviderIds(raw: unknown): Set<string> | null {
  if (typeof raw !== "string") return null;

  const ids = raw
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

  if (ids.length === 0) return null;
  return new Set(ids);
}

/**
 * Filter provider definitions by an allowlist.
 *
 * Fail-open on a fully mismatched allowlist (e.g. typo'd ids): an empty
 * connect UI would brick logins entirely, and this filter is cosmetic —
 * actual enforcement is the proxy target allowlist. We log loudly so a
 * misconfigured org build is caught in smoke testing.
 */
export function filterProvidersByAllowlist<T extends { id: string }>(
  providers: readonly T[],
  allowed: Set<string> | null,
): T[] {
  if (allowed === null) return [...providers];

  const filtered = providers.filter((p) => allowed.has(p.id.toLowerCase()));

  const known = new Set(providers.map((p) => p.id.toLowerCase()));
  const unknown = [...allowed].filter((id) => !known.has(id));
  if (unknown.length > 0) {
    console.warn(`[pi-for-excel] VITE_PI_ALLOWED_PROVIDERS contains unknown provider ids: ${unknown.join(", ")}`);
  }

  if (filtered.length === 0) {
    console.error(
      "[pi-for-excel] VITE_PI_ALLOWED_PROVIDERS matched no providers; showing all providers. Check the configured ids.",
    );
    return [...providers];
  }

  return filtered;
}
