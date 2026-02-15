/**
 * Model ordering + version/recency scoring.
 *
 * Pure helpers (no DOM/Office dependencies) so we can unit test them.
 */

export type ModelRef = { provider: string; id: string };

const PROVIDER_ORDER: Record<string, number> = {
  anthropic: 1,
  "openai-codex": 2,
  openai: 3,
  google: 4,
  "google-gemini-cli": 4,
  "google-antigravity": 4,
  "github-copilot": 5,
};

export function providerPriority(provider: string): number {
  return PROVIDER_ORDER[provider] ?? 999;
}

export function familyPriority(provider: string, id: string): number {
  if (provider === "anthropic") {
    if (id.startsWith("claude-opus-")) return 0;
    if (id.startsWith("claude-sonnet-")) return 1;
    if (id.startsWith("claude-haiku-")) return 2;
    return 9;
  }

  if (provider === "openai-codex" || provider === "openai") {
    if (id.includes("codex")) return 0;
    if (id.startsWith("gpt-")) return 1;
    if (id.startsWith("o")) return 2;
    return 9;
  }

  if (provider === "google" || provider === "google-gemini-cli" || provider === "google-antigravity") {
    // Prefer Pro-ish variants first, then Flash-ish, then any Gemini.
    if (/^gemini-.*-pro/i.test(id)) return 0;
    if (/^gemini-.*-flash/i.test(id)) return 1;
    if (id.includes("gemini")) return 2;
    return 9;
  }

  return 9;
}

export function parseMajorMinor(id: string): number {
  // Extract a comparable major/minor number from common model ID formats.
  // Important: don't misinterpret 8-digit date suffixes (e.g. 20250514) as "minor".
  // Examples:
  // - claude-opus-4-5           -> 45
  // - claude-opus-4-6           -> 46
  // - claude-opus-4-20250514    -> 40 (major only; date handled separately)
  // - gpt-5.3-codex             -> 53
  // - gemini-2.5-pro            -> 25
  // - gemini-3-pro-preview      -> 30

  const pack = (major: number, minor: number | null): number => {
    if (minor === null) return major * 10;
    // minor < 10 => major*10 + minor (4.6 -> 46)
    if (minor < 10) return major * 10 + minor;
    // allow 2-digit minors (e.g. 5.12 -> 512)
    return major * 100 + minor;
  };

  // Claude-style: -4-6 (but NOT -4-20250514)
  const hyphenVer = id.match(/-(\d+)-(\d{1,2})(?:-|$)/);
  if (hyphenVer) {
    return pack(parseInt(hyphenVer[1], 10), parseInt(hyphenVer[2], 10));
  }

  // OpenAI/Gemini-style: 5.3 / 2.5
  const dotVer = id.match(/(\d+)\.(\d{1,2})/);
  if (dotVer) {
    return pack(parseInt(dotVer[1], 10), parseInt(dotVer[2], 10));
  }

  // Fallback: first major number after hyphen
  const majorMatch = id.match(/-(\d+)(?:-|$)/);
  if (majorMatch) {
    return pack(parseInt(majorMatch[1], 10), null);
  }

  return 0;
}

export function modelRecencyScore(id: string): number {
  // Prefer higher major/minor first, then higher date suffix.
  const majorMinor = parseMajorMinor(id);

  let date = 0;
  const dateMatch = id.match(/(\d{8})/);
  if (dateMatch) date = parseInt(dateMatch[1], 10);

  // date is at most 8 digits → multiplier must exceed that range
  return majorMinor * 100_000_000 + date;
}

export function compareModels(a: ModelRef, b: ModelRef): number {
  const aProv = providerPriority(a.provider);
  const bProv = providerPriority(b.provider);
  if (aProv !== bProv) return aProv - bProv;

  const aFam = familyPriority(a.provider, a.id);
  const bFam = familyPriority(b.provider, b.id);
  if (aFam !== bFam) return aFam - bFam;

  const aRec = modelRecencyScore(a.id);
  const bRec = modelRecencyScore(b.id);
  if (aRec !== bRec) return bRec - aRec;

  return a.id.localeCompare(b.id);
}
