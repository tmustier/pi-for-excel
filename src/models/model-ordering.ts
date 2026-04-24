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

const MODEL_NAME_BOUNDARY = String.raw`(?:^|[\\/~.:])`;
const MODEL_VERSION_BOUNDARY = String.raw`(?=$|[-/:])`;

const CLAUDE_VERSION_RE = new RegExp(
  `${MODEL_NAME_BOUNDARY}claude(?:-[a-z]+)*-(\\d+)[.-](\\d{1,2})${MODEL_VERSION_BOUNDARY}`,
  "i",
);
const CLAUDE_MAJOR_RE = new RegExp(
  `${MODEL_NAME_BOUNDARY}claude(?:-[a-z]+)*-(\\d+)${MODEL_VERSION_BOUNDARY}`,
  "i",
);
const GPT_DOT_VERSION_RE = new RegExp(`${MODEL_NAME_BOUNDARY}gpt-(\\d+)\\.(\\d{1,2})${MODEL_VERSION_BOUNDARY}`, "i");
const GPT_MAJOR_RE = new RegExp(`${MODEL_NAME_BOUNDARY}gpt-(\\d+)(?:[a-z]+)?${MODEL_VERSION_BOUNDARY}`, "i");
const GEMINI_DOT_VERSION_RE = new RegExp(
  `${MODEL_NAME_BOUNDARY}gemini(?:-[a-z]+)*-(\\d+)\\.(\\d{1,2})${MODEL_VERSION_BOUNDARY}`,
  "i",
);
const GEMINI_MAJOR_RE = new RegExp(
  `${MODEL_NAME_BOUNDARY}gemini(?:-[a-z]+)*-(\\d+)${MODEL_VERSION_BOUNDARY}`,
  "i",
);
const LETTER_PREFIXED_DOT_VERSION_RE = /(?:^|[\/_.-])[a-z]+(\d{1,2})\.(\d{1,2})(?=$|[-/:._])/i;
const GENERIC_VERSION_RE = /(?:^|[\w~/-][.-]|[-_/])v?(\d{1,2})(?:[.-](\d{1,2})([a-z]+)?)?(?=$|[-/:._])/gi;

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
  // - claude-opus-4-6                         -> 46
  // - claude-opus-4.7                         -> 47
  // - anthropic.claude-opus-4-1-20250805-v1:0 -> 41 (date handled separately)
  // - claude-opus-4-20250514                  -> 40 (major only; date handled separately)
  // - gpt-5.5                                 -> 55
  // - gpt-4o-2024-11-20                       -> 40 (not 202411)
  // - gemini-2.5-pro-preview-06-05            -> 25 (not 65)
  // - google/gemini-3.1-pro-preview           -> 31
  // - gemini-3-pro-preview                    -> 30
  // - MiniMax-M2.7                            -> 27 (letter-prefixed fallback)
  // - Qwen/Qwen3.5                            -> 35 (letter-prefixed fallback)
  // - gemma-4-31b-it                          -> 40 (generic fallback; ignores size suffix)
  // - zai.glm-5                               -> 50 (generic fallback)

  const pack = (major: number, minor: number | null): number => {
    if (minor === null) return major * 10;
    // minor < 10 => major*10 + minor (4.6 -> 46)
    if (minor < 10) return major * 10 + minor;
    // allow 2-digit minors (e.g. 5.12 -> 512)
    return major * 100 + minor;
  };

  const claudeVer = id.match(CLAUDE_VERSION_RE);
  if (claudeVer) {
    return pack(parseInt(claudeVer[1], 10), parseInt(claudeVer[2], 10));
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

  const letterPrefixedDotVersion = id.match(LETTER_PREFIXED_DOT_VERSION_RE);
  if (letterPrefixedDotVersion) {
    return pack(parseInt(letterPrefixedDotVersion[1], 10), parseInt(letterPrefixedDotVersion[2], 10));
  }

  for (const genericVersion of id.matchAll(GENERIC_VERSION_RE)) {
    const matchIndex = genericVersion.index ?? 0;
    const digitOffset = genericVersion[0].search(/\d/);
    const digitIndex = matchIndex + Math.max(digitOffset, 0);
    const surrounding = id.slice(Math.max(0, digitIndex - 5), digitIndex + 8);
    if (/\d{4}-\d{2}-\d{2}/.test(surrounding)) continue;

    const major = parseInt(genericVersion[1], 10);
    const minor = genericVersion[2] && !genericVersion[3] ? parseInt(genericVersion[2], 10) : null;
    return pack(major, minor);
  }

  return 0;
}

function parseDateSuffixScore(id: string): number {
  let score = 0;

  for (const match of id.matchAll(/(?:^|[-/:])(\d{8})(?=$|[-/:])/g)) {
    score = Math.max(score, parseInt(match[1], 10));
  }

  for (const match of id.matchAll(/(?:^|[-/:])(\d{4})-(\d{2})-(\d{2})(?=$|[-/:])/g)) {
    score = Math.max(score, parseInt(`${match[1]}${match[2]}${match[3]}`, 10));
  }

  for (const match of id.matchAll(/(?:^|[-/:])(\d{2})-(\d{4})(?=$|[-/:])/g)) {
    score = Math.max(score, parseInt(`${match[2]}${match[1]}00`, 10));
  }

  for (const match of id.matchAll(/(?:^|[-/:])(\d{2})-(\d{2})(?=$|[-/:])/g)) {
    score = Math.max(score, parseInt(`${match[1]}${match[2]}`, 10));
  }

  return score;
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
