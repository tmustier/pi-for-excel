/**
 * Persistent agent rules storage.
 *
 * We store two scopes:
 * - user-level rules (global to this install)
 * - workbook-level rules (scoped by workbook identity)
 *
 * Storage keys are intentionally unchanged from the original "instructions"
 * naming to preserve backward compatibility with persisted data.
 */

export const USER_RULES_SOFT_LIMIT = 2_000;
export const WORKBOOK_RULES_SOFT_LIMIT = 4_000;

const USER_RULES_KEY = "user.instructions";
const WORKBOOK_RULES_PREFIX = "workbook.instructions.v1.";

export type RuleLevel = "user" | "workbook";
export type RuleAction = "append" | "replace";

export interface RulesStore {
  get: (key: string) => Promise<unknown>;
  set: (key: string, value: unknown) => Promise<void>;
}

function normalizeStoredText(value: unknown): string | null {
  if (typeof value !== "string") return null;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizeDraftText(value: string): string | null {
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function workbookRulesKey(workbookId: string): string {
  return `${WORKBOOK_RULES_PREFIX}${workbookId}`;
}

export async function getUserRules(store: RulesStore): Promise<string | null> {
  const value = await store.get(USER_RULES_KEY);
  return normalizeStoredText(value);
}

export async function setUserRules(
  store: RulesStore,
  nextValue: string | null,
): Promise<string | null> {
  const normalized = nextValue === null ? null : normalizeDraftText(nextValue);
  await store.set(USER_RULES_KEY, normalized ?? "");
  return normalized;
}

export async function getWorkbookRules(
  store: RulesStore,
  workbookId: string | null,
): Promise<string | null> {
  if (!workbookId) return null;

  const value = await store.get(workbookRulesKey(workbookId));
  return normalizeStoredText(value);
}

export async function setWorkbookRules(
  store: RulesStore,
  workbookId: string,
  nextValue: string | null,
): Promise<string | null> {
  const normalized = nextValue === null ? null : normalizeDraftText(nextValue);
  await store.set(workbookRulesKey(workbookId), normalized ?? "");
  return normalized;
}

export function applyRuleAction(args: {
  currentValue: string | null;
  action: RuleAction;
  content: string;
}): string | null {
  const current = normalizeStoredText(args.currentValue);

  if (args.action === "replace") {
    return normalizeDraftText(args.content);
  }

  const addition = normalizeDraftText(args.content);
  if (!addition) {
    throw new Error("content is required for append");
  }

  if (!current) return addition;
  return `${current}\n${addition}`;
}

export function hasAnyRules(values: {
  userRules: string | null;
  workbookRules: string | null;
}): boolean {
  return Boolean(values.userRules || values.workbookRules);
}
