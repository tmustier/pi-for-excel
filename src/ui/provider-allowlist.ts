export function resolveAllowedProviderIds(raw: unknown): Set<string> | null {
  if (typeof raw !== "string") return null;

  const ids = raw
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

  if (ids.length === 0) return null;
  return new Set(ids);
}

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
