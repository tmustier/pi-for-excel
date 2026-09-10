import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const checker = fileURLToPath(new URL("../scripts/check-top-level-return-types.mjs", import.meta.url));

function check(source) {
  const directory = mkdtempSync(join(tmpdir(), "pi-return-contract-"));
  try {
    mkdirSync(join(directory, "src"));
    writeFileSync(join(directory, "src", "example.ts"), source);
    const result = spawnSync(process.execPath, [checker], {
      cwd: directory, encoding: "utf8", timeout: 30_000,
    });
    assert.equal(result.error, undefined);
    assert.equal(result.signal, null);
    return result;
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}

test("return contract check rejects inferred exports regardless of export syntax", () => {
  for (const source of [
    "export function run() { return 1; }",
    "const run = () => 1; export { run as start };",
    "const run = () => 1; export default run;",
    "function run() { return 1; } export default run;",
    "export default () => 1;",
    "export default function () { return 1; }",
  ]) {
    const result = check(source);
    assert.equal(result.status, 1, source);
    assert.match(result.stderr, /Declare return types/);
  }
});

test("return contract check permits private inference and explicit public contracts", () => {
  const result = check(`
    const privateHelper = () => 1;
    export function run(): number { return privateHelper(); }
    const start: () => number = () => 1;
    export default start;
  `);
  assert.equal(result.status, 0, result.stderr);
});
