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

const OPENAI_CODEX_RE = /^gpt-5(?:\.(\d+))?-codex(?:-|$)/;
const OPENAI_PLAIN_GPT_RE = /^gpt-5(?:\.(\d+))?$/;
const OPENAI_GPT_RE = /^gpt-5(?:[.-]|$)/;

const CLAUDE_HYPHEN_VERSION_RE = /^claude(?:-[a-z]+)*-(\d+)-(\d{1,2})(?:-|$)/i;
const CLAUDE_MAJOR_RE = /^claude(?:-[a-z]+)*-(\d+)(?:-|$)/i;
const GPT_DOT_VERSION_RE = /^gpt-(\d+)\.(\d{1,2})(?:-|$)/i;
const GPT_MAJOR_RE = /^gpt-(\d+)(?:[a-z]+)?(?:-|$)/i;
const GEMINI_DOT_VERSION_RE = /^gemini(?:-[a-z]+)*-(\d+)\.(\d{1,2})(?:-|$)/i;
const GEMINI_MAJOR_RE = /^gemini(?:-[a-z]+)*-(\d+)(?:-|$)/i;

export function isOpenAiCodexModelId(id: string): boolean {
  return OPENAI_CODEX_RE.test(id);
}

export function isOpenAiGeneralGptModelId(id: string): boolean {
  return OPENAI_GPT_RE.test(id) && !isOpenAiCodexModelId(id);
}

export function openAiFamilyPriority(id: string): number {
  // Prefer the latest general GPT-5 model first, then other GPT-5 variants,
  // then Codex-specialized variants, then older o-series fallbacks.
  if (OPENAI_PLAIN_GPT_RE.test(id)) return 0;
  if (isOpenAiGeneralGptModelId(id)) return 1;
  if (isOpenAiCodexModelId(id)) return 2;
  if (id.startsWith("gpt-")) return 3;
  if (id.startsWith("o")) return 4;
  return 9;
}

export function familyPriority(provider: string, id: string): number {
  if (provider === "anthropic") {
    if (id.startsWith("claude-opus-")) return 0;
    if (id.startsWith("claude-sonnet-")) return 1;
    if (id.startsWith("claude-haiku-")) return 2;
    return 9;
  }

  if (provider === "openai-codex" || provider === "openai") {
    return openAiFamilyPriority(id);
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
  // Important: only parse the leading version segment, not later date-like suffixes.
  // Examples:
  // - claude-opus-4-6                 -> 46
  // - claude-opus-4-20250514          -> 40 (major only; date handled separately)
  // - gpt-5.4                         -> 54
  // - gpt-4o-2024-11-20               -> 40 (not 202411)
  // - gemini-2.5-pro-preview-06-05    -> 25 (not 65)
  // - gemini-3-pro-preview            -> 30

  const pack = (major: number, minor: number | null): number => {
    if (minor === null) return major * 10;
    // minor < 10 => major*10 + minor (4.6 -> 46)
    if (minor < 10) return major * 10 + minor;
    // allow 2-digit minors (e.g. 5.12 -> 512)
    return major * 100 + minor;
  };

  const claudeHyphenVer = id.match(CLAUDE_HYPHEN_VERSION_RE);
  if (claudeHyphenVer) {
    return pack(parseInt(claudeHyphenVer[1], 10), parseInt(claudeHyphenVer[2], 10));
  }

  const gptDotVer = id.match(GPT_DOT_VERSION_RE);
  if (gptDotVer) {
    return pack(parseInt(gptDotVer[1], 10), parseInt(gptDotVer[2], 10));
  }

  const geminiDotVer = id.match(GEMINI_DOT_VERSION_RE);
  if (geminiDotVer) {
    return pack(parseInt(geminiDotVer[1], 10), parseInt(geminiDotVer[2], 10));
  }

  const claudeMajor = id.match(CLAUDE_MAJOR_RE);
  if (claudeMajor) {
    return pack(parseInt(claudeMajor[1], 10), null);
  }

  const gptMajor = id.match(GPT_MAJOR_RE);
  if (gptMajor) {
    return pack(parseInt(gptMajor[1], 10), null);
  }

  const geminiMajor = id.match(GEMINI_MAJOR_RE);
  if (geminiMajor) {
    return pack(parseInt(geminiMajor[1], 10), null);
  }

  return 0;
}

function parseDateSuffixScore(id: string): number {
  const compactDateMatch = id.match(/(\d{8})$/);
  if (compactDateMatch) {
    return parseInt(compactDateMatch[1], 10);
  }

  const yearMonthDayMatch = id.match(/(\d{4})-(\d{2})-(\d{2})$/);
  if (yearMonthDayMatch) {
    return parseInt(
      `${yearMonthDayMatch[1]}${yearMonthDayMatch[2]}${yearMonthDayMatch[3]}`,
      10,
    );
  }

  const monthYearMatch = id.match(/(\d{2})-(\d{4})$/);
  if (monthYearMatch) {
    return parseInt(`${monthYearMatch[2]}${monthYearMatch[1]}00`, 10);
  }

  const monthDayMatch = id.match(/(\d{2})-(\d{2})$/);
  if (monthDayMatch) {
    return parseInt(`${monthDayMatch[1]}${monthDayMatch[2]}`, 10);
  }

  return 0;
}

export function modelRecencyScore(id: string): number {
  // Prefer higher major/minor first, then higher date suffix.
  const majorMinor = parseMajorMinor(id);
  const date = parseDateSuffixScore(id);

  // date is at most 8 digits → multiplier must exceed that range
  return majorMinor * 100_000_000 + date;
}

export function compareOpenAiModelIds(aId: string, bId: string): number {
  const recency = modelRecencyScore(bId) - modelRecencyScore(aId);
  if (recency !== 0) return recency;

  const family = openAiFamilyPriority(aId) - openAiFamilyPriority(bId);
  if (family !== 0) return family;

  return aId.localeCompare(bId);
}

export function shouldPreferOpenAiGeneralModel(generalId: string, codexId: string): boolean {
  return parseMajorMinor(generalId) >= parseMajorMinor(codexId);
}

export function compareModels(a: ModelRef, b: ModelRef): number {
  const aProv = providerPriority(a.provider);
  const bProv = providerPriority(b.provider);
  if (aProv !== bProv) return aProv - bProv;

  if (a.provider === b.provider && (a.provider === "openai-codex" || a.provider === "openai")) {
    // OpenAI is recency-first so newer Codex variants still outrank older GPT variants.
    return compareOpenAiModelIds(a.id, b.id);
  }

  const aFam = familyPriority(a.provider, a.id);
  const bFam = familyPriority(b.provider, b.id);
  if (aFam !== bFam) return aFam - bFam;

  const aRec = modelRecencyScore(a.id);
  const bRec = modelRecencyScore(b.id);
  if (aRec !== bRec) return bRec - aRec;

  return a.id.localeCompare(b.id);
}
