import type { SkillsSourceKind } from "../tools/tool-details.js";

export interface SkillReadCacheEntry {
  skillName: string;
  sourceKind: SkillsSourceKind;
  location: string;
  markdown: string;
  cachedAt: number;
  readCount: number;
}

export interface SkillReadCacheSetInput {
  skillName: string;
  sourceKind: SkillsSourceKind;
  location: string;
  markdown: string;
}

export interface SkillReadCache {
  get: (sessionId: string, skillName: string) => SkillReadCacheEntry | null;
  set: (sessionId: string, skill: SkillReadCacheSetInput) => SkillReadCacheEntry;
  clearSession: (sessionId: string) => void;
  clearAll: () => void;
}

function normalizeSkillName(name: string): string {
  return name.trim().toLowerCase();
}

export function createSkillReadCache(): SkillReadCache {
  const bySession = new Map<string, Map<string, SkillReadCacheEntry>>();

  const ensureSessionCache = (sessionId: string): Map<string, SkillReadCacheEntry> => {
    const existing = bySession.get(sessionId);
    if (existing) {
      return existing;
    }

    const created = new Map<string, SkillReadCacheEntry>();
    bySession.set(sessionId, created);
    return created;
  };

  return {
    get(sessionId: string, skillName: string): SkillReadCacheEntry | null {
      const sessionCache = bySession.get(sessionId);
      if (!sessionCache) {
        return null;
      }

      const cached = sessionCache.get(normalizeSkillName(skillName));
      return cached ?? null;
    },
    set(sessionId: string, skill: SkillReadCacheSetInput): SkillReadCacheEntry {
      const sessionCache = ensureSessionCache(sessionId);
      const normalizedName = normalizeSkillName(skill.skillName);
      const previous = sessionCache.get(normalizedName);

      const next: SkillReadCacheEntry = {
        skillName: skill.skillName,
        sourceKind: skill.sourceKind,
        location: skill.location,
        markdown: skill.markdown,
        cachedAt: Date.now(),
        readCount: previous ? previous.readCount + 1 : 1,
      };

      sessionCache.set(normalizedName, next);
      return next;
    },
    clearSession(sessionId: string): void {
      bySession.delete(sessionId);
    },
    clearAll(): void {
      bySession.clear();
    },
  };
}
