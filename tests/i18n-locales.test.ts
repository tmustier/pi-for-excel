// Static contract: locale key parity and source-level t() usage are dependency restrictions;
// runtime behavior cannot exhaustively prove that every shipped string remains localized.
import assert from "node:assert/strict";
import { test } from "node:test";
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const localesDir = join(root, "src", "language", "locales");

function parseLocaleJson(raw: unknown, label: string): Record<string, string> {
  if (typeof raw !== "object" || raw === null || Array.isArray(raw)) {
    throw new Error(`${label} locale must be an object`);
  }

  const rawObject = raw as Record<string, unknown>;
  const parsed: Record<string, string> = {};
  for (const [key, value] of Object.entries(rawObject)) {
    if (typeof value !== "string") {
      throw new Error(`${label} locale value for ${key} must be a string`);
    }
    parsed[key] = value;
  }
  return parsed;
}

const en = parseLocaleJson(JSON.parse(readFileSync(join(localesDir, "en.json"), "utf8")) as unknown, "en");
const zh = parseLocaleJson(JSON.parse(readFileSync(join(localesDir, "zh-CN.json"), "utf8")) as unknown, "zh-CN");

function requireMatchGroup(match: RegExpMatchArray, index: number): string {
  const value = match[index];
  if (value === undefined) {
    throw new Error(`Expected regex capture group ${index}.`);
  }
  return value;
}

function placeholders(value: string): Set<string> {
  const found = new Set<string>();
  for (const m of value.matchAll(/\{([a-zA-Z0-9_]+)\}/g)) found.add(requireMatchGroup(m, 1));
  return found;
}

function collectSourceFiles(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) {
      if (entry === "locales" || entry === "node_modules") continue;
      collectSourceFiles(full, out);
    } else if (/\.(ts|html)$/.test(entry)) {
      out.push(full);
    }
  }
  return out;
}

const sourceFiles = collectSourceFiles(join(root, "src"));
const localizedUiSourceFiles = ["ui", "taskpane", "commands", "compaction", "files"]
  .flatMap((dir) => collectSourceFiles(join(root, "src", dir)));
const corpus = sourceFiles.map((f) => readFileSync(f, "utf8")).join("\n");

void test("en and zh-CN locales have identical key sets", () => {
  const missingInZh = Object.keys(en).filter((k) => !(k in zh)).sort();
  const extraInZh = Object.keys(zh).filter((k) => !(k in en)).sort();
  assert.deepEqual(missingInZh, [], `keys missing in zh-CN.json: ${missingInZh.join(", ")}`);
  assert.deepEqual(extraInZh, [], `keys in zh-CN.json but not en.json: ${extraInZh.join(", ")}`);
});

void test("zh-CN placeholders are a subset of en placeholders per key", () => {
  // zh may drop English plural-helper vars (e.g. {cue}), but must never
  // reference a placeholder the caller does not provide.
  const violations: string[] = [];
  for (const [key, enValue] of Object.entries(en)) {
    const zhValue = zh[key];
    if (typeof zhValue !== "string") continue;
    const enVars = placeholders(enValue);
    for (const v of placeholders(zhValue)) {
      if (!enVars.has(v)) violations.push(`${key}: {${v}}`);
    }
  }
  assert.deepEqual(violations, [], `zh-CN placeholders missing from en: ${violations.join(", ")}`);
});

void test("zh-CN drops English-only plural helper placeholders", () => {
  // Chinese does not inflect nouns for singular/plural, so keeping these
  // English helper placeholders leaks strings like "file s" into zh-CN UI.
  const banned = new Set(["plural", "cue"]);
  const violations: string[] = [];
  for (const [key, value] of Object.entries(zh)) {
    for (const v of placeholders(value)) {
      if (banned.has(v)) violations.push(`${key}: {${v}}`);
    }
  }
  assert.deepEqual(violations, [], `zh-CN should not include English plural helpers: ${violations.join(", ")}`);
});

void test("zh-CN keeps command syntax placeholders copyable", () => {
  const violations = Object.entries(zh)
    .filter(([key]) => key.startsWith("experimental."))
    .filter(([, value]) => /<[^>]*[\u4e00-\u9fff][^>]*>/.test(value))
    .map(([key, value]) => `${key}: ${value}`);
  assert.deepEqual(violations, [], `localized command placeholders in zh-CN: ${violations.join("\n")}`);
});

void test("en locale has no empty values", () => {
  const empty = Object.entries(en)
    .filter(([, v]) => typeof v !== "string" || v.length === 0)
    .map(([k]) => k);
  assert.deepEqual(empty, [], `empty en values: ${empty.join(", ")}`);
});

void test("POLICY: locale keys, UI sinks, and initialization stay localization-safe", () => {
  const labelKeys = new Set(
    [...corpus.matchAll(/\bl\("([^"]+)"\)/g)].map(
      (match) => `humanize.label.${requireMatchGroup(match, 1).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "")}`,
    ),
  );
  const valueKeys = new Set(
    [...corpus.matchAll(/\bv\("([^"]+)"/g)].map((match) => `humanize.value.${requireMatchGroup(match, 1)}`),
  );
  const unusedKeys = Object.keys(en).filter((key) => {
    if (labelKeys.has(key) || valueKeys.has(key)) return false;
    if (key.startsWith("humanize.unit.") || key.startsWith("perm.trust.")) return false;
    return !corpus.includes(`"${key}"`) && !corpus.includes(`'${key}'`) && !corpus.includes("`" + key + "`");
  });

  const staticSink = /(?:textContent|innerHTML|placeholder|title|subtitle|message|\w*[Ll]abel|showToast|createButton|createConfigRow|\.text)\s*(?:=|:|\()\s*["`][A-Z]/;
  const toastTemplate = /\b(?:showToast|resolved\.showToast)\(\s*`/;
  const allowed = /aria-|data-|className|\.css|https?:\/\/|icon\(|throw new Error|const message = error instanceof Error|externalLoadError|activationLoadError/;
  const hardcodedUiStrings: string[] = [];
  for (const file of localizedUiSourceFiles) {
    const relative = file.slice(root.length + 1);
    for (const [index, line] of readFileSync(file, "utf8").split("\n").entries()) {
      if ((staticSink.test(line) || toastTemplate.test(line)) && !/\bt\(/.test(line) && !allowed.test(line)) {
        hardcodedUiStrings.push(`${relative}:${index + 1}: ${line.trim()}`);
      }
    }
  }

  const moduleScopeCalls: string[] = [];
  for (const file of sourceFiles) {
    let depth = 0;
    for (const [index, line] of readFileSync(file, "utf8").split("\n").entries()) {
      if (depth === 0 && /\bt\(\s*"/.test(line) && !/^\s*(\*|\/\/)/.test(line)) {
        moduleScopeCalls.push(`${file.slice(root.length + 1)}:${index + 1}`);
      }
      for (const character of line) {
        if (character === "{" || character === "(") depth++;
        else if (character === "}" || character === ")") depth = Math.max(0, depth - 1);
      }
    }
  }

  assert.deepEqual(
    { unusedKeys, hardcodedUiStrings, moduleScopeCalls },
    { unusedKeys: [], hardcodedUiStrings: [], moduleScopeCalls: [] },
  );
});
