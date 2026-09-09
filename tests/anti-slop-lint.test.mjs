import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = fileURLToPath(new URL("../", import.meta.url));
const oxlint = fileURLToPath(new URL("../node_modules/oxlint/bin/oxlint", import.meta.url));

function lint(fixtures) {
  const directory = mkdtempSync(join(tmpdir(), "pi-excel-anti-slop-"));
  try {
    const files = Object.entries(fixtures).map(([name, source]) => {
      const path = join(directory, name);
      writeFileSync(path, source);
      return path;
    });
    const result = spawnSync(
      process.execPath,
      [oxlint, "--config", ".oxlintrc.json", "--quiet", "--format", "json", ...files],
      { cwd: root, encoding: "utf8", timeout: 30_000 },
    );
    assert.equal(result.signal, null, result.stderr);
    assert.equal(result.error, undefined);
    return { status: result.status, report: JSON.parse(result.stdout) };
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}

function diagnosticCodes(report) {
  return report.diagnostics.map((diagnostic) => diagnostic.code).sort();
}

test("anti-slop lint rejects each enabled TypeScript policy", () => {
  const result = lint({
    "invalid.ts": `
      declare const items: ReadonlyArray<{ readonly id: string }>;
      export const spreadIndex = items.reduce(
        (accumulator, item) => ({ ...accumulator, [item.id]: item }),
        {},
      );
      export const copiedIndex = items.reduce(
        (accumulator, item) => Object.assign({}, accumulator, { [item.id]: item }),
        {},
      );
      const source = { id: "known" };
      const widened: unknown = source;
      export const asserted = widened as { readonly id: string };
      vi.mock("./workbook-store.ts");
    `,
  });

  assert.equal(result.status, 1);
  assert.deepEqual(diagnosticCodes(result.report), [
    "anti-slop(no-module-mocking)",
    "anti-slop(no-reduce-accumulator-copy)",
    "anti-slop(no-widen-then-assert)",
    "oxc(no-accumulating-spread)",
  ]);
});

test("anti-slop module-mocking policy also covers MJS tests", () => {
  const result = lint({
    "invalid.mjs": `
      jest.unstable_mockModule("./workbook-store.mjs", () => ({ load: () => [] }));
    `,
  });

  assert.equal(result.status, 1);
  assert.deepEqual(diagnosticCodes(result.report), ["anti-slop(no-module-mocking)"]);
});

test("anti-slop lint accepts mutation-based accumulation and injected fakes", () => {
  const result = lint({
    "valid.ts": `
      interface WorkbookStore {
        load(): readonly string[];
      }
      class FakeWorkbookStore implements WorkbookStore {
        load(): readonly string[] {
          return ["Sheet1"];
        }
      }
      export function loadSheetNames(store: WorkbookStore): readonly string[] {
        return store.load();
      }
      export const sheetNames = loadSheetNames(new FakeWorkbookStore());
      declare const entries: ReadonlyArray<readonly [string, string]>;
      export const entryIndex = entries.reduce<Record<string, string>>(
        (accumulator, [key, value]) => {
          accumulator[key] = value;
          return accumulator;
        },
        {},
      );
    `,
    "valid.mjs": `
      export function loadSheetNames(store) {
        return store.load();
      }
      const fakeStore = { load: () => ["Sheet1"] };
      export const sheetNames = loadSheetNames(fakeStore);
    `,
  });

  assert.equal(result.status, 0);
  assert.deepEqual(result.report.diagnostics, []);
});
