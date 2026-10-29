import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = fileURLToPath(new URL("../", import.meta.url));
const oxlint = fileURLToPath(new URL("../node_modules/oxlint/bin/oxlint", import.meta.url));
const plugin = fileURLToPath(new URL("../tools/oxlint/pi-for-excel/index.ts", import.meta.url));

function lint(rule, source) {
  const directory = mkdtempSync(join(tmpdir(), "pi-repo-lint-"));
  try {
    const config = join(directory, ".oxlintrc.json");
    writeFileSync(config, JSON.stringify({
      categories: { correctness: "off" },
      jsPlugins: [{ name: "pi-for-excel", specifier: plugin }],
      rules: { [`pi-for-excel/${rule}`]: "error" },
    }));
    const file = join(directory, "example.ts");
    writeFileSync(file, source);
    const result = spawnSync(process.execPath, [oxlint, "--config", config, "--format", "json", file], {
      cwd: root, encoding: "utf8", timeout: 30_000,
    });
    assert.equal(result.error, undefined);
    assert.equal(result.signal, null, result.stderr);
    return JSON.parse(result.stdout).diagnostics.map((diagnostic) => diagnostic.message);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}

test("return contract rule rejects inferred exports regardless of export syntax", () => {
  for (const source of [
    "export function run() { return 1; }",
    "export const run = () => 1;",
    "const run = () => 1; export { run as start };",
    "const run = () => 1; export default run;",
    "function run() { return 1; } export default run;",
    "export default () => 1;",
    "export default function () { return 1; }",
  ]) {
    const messages = lint("exported-function-return-types", source);
    assert.equal(messages.length, 1, source);
    assert.match(messages[0], /Declare a return type/);
  }
});

test("return contract rule permits private inference and explicit public contracts", () => {
  assert.deepEqual(lint("exported-function-return-types", `
    const privateHelper = () => 1;
    export function run(): number { return privateHelper(); }
    const start: () => number = () => 1;
    export default start;
    export declare function declared(): void;
  `), []);
});

test("boundary cast rule rejects decoded JSON cast straight to a domain type", () => {
  assert.deepEqual(lint("no-typed-json-boundary-cast", `
    interface Settings { readonly theme: string }
    declare const text: string;
    declare const response: Response;
    export async function load(): Promise<void> {
      const a = JSON.parse(text) as Settings;
      const b = (await response.json()) as Settings;
      const c = (await response?.json()) as Settings;
      void [a, b, c];
    }
  `), [
    "JSON.parse cast to Settings: decoded JSON/fetch payloads must first land as unknown, then be parsed/refined by a concrete boundary parser.",
    "Response.json cast to Settings: decoded JSON/fetch payloads must first land as unknown, then be parsed/refined by a concrete boundary parser.",
    "Response.json cast to Settings: decoded JSON/fetch payloads must first land as unknown, then be parsed/refined by a concrete boundary parser.",
  ]);
});

test("boundary cast rule permits unknown and JsonValue landings", () => {
  assert.deepEqual(lint("no-typed-json-boundary-cast", `
    type JsonValue = string | number | boolean | null;
    declare const text: string;
    declare const response: Response;
    export async function load(): Promise<void> {
      const a = JSON.parse(text) as unknown;
      const b = (await response.json()) as JsonValue;
      void [a, b];
    }
  `), []);
});

test("auth storage rule rejects every localStorage reference", () => {
  assert.equal(lint("no-auth-local-storage", `
    export const a = localStorage.getItem("token");
    export const b = window.localStorage;
    export const { localStorage: c } = window;
    export const d = sessionStorage.getItem("token");
  `).length, 3);
});

test("module-scope translation rule rejects t() that runs during module evaluation", () => {
  assert.equal(lint("no-module-scope-translation", `
    declare function t(key: string): string;
    export const title = t("title");
    export const labels = { save: t("save") };
    export class Panel { static heading = t("heading"); }
    export const lazy = () => t("lazy");
    export function render(): string { return t("render"); }
    export class Card { label = t("label"); }
    export const fallback = (value = t("fallback")): string => value;
  `).length, 3);
});

test("repository lint config scopes each local rule to the code it protects", () => {
  const config = JSON.parse(readFileSync(join(root, ".oxlintrc.json"), "utf8"));
  const scopes = Object.fromEntries(config.overrides.flatMap((override) =>
    Object.keys(override.rules)
      .filter((rule) => rule.startsWith("pi-for-excel/"))
      .map((rule) => [rule, override.files]),
  ));
  assert.deepEqual(scopes, {
    "pi-for-excel/exported-function-return-types": ["src/**/*.ts"],
    "pi-for-excel/no-auth-local-storage": ["src/auth/**/*.ts"],
    "pi-for-excel/no-module-scope-translation": ["src/**/*.ts"],
    "pi-for-excel/no-typed-json-boundary-cast": ["src/**/*.ts", "tests/**/*.ts"],
  });
});
