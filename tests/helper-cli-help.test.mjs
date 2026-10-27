// `--help` must print usage without starting a server or touching the user's
// certificate directory (the HTTPS path can install mkcert and create certs).
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

for (const pkg of ["proxy", "python-bridge", "tmux-bridge"]) {
  const bin = `pi-for-excel-${pkg}`;
  for (const args of [["--help"], ["-h"], ["--install-missing", "--http", "--help"]]) {
    test(`${bin} ${args.join(" ")} prints usage and exits without side effects`, (t) => {
      const home = mkdtempSync(path.join(tmpdir(), "pi-cli-help-"));
      t.after(() => rmSync(home, { recursive: true, force: true }));
      const result = spawnSync(process.execPath, [path.join(root, "pkg", pkg, "cli.mjs"), ...args], {
        encoding: "utf8",
        env: { ...process.env, HOME: home, USERPROFILE: home },
        timeout: 5000,
      });
      assert.equal(result.status, 0, result.error?.message ?? result.stderr);
      assert.match(result.stdout, new RegExp(`^Usage: npx ${bin} `));
      assert.deepEqual(readdirSync(home), []);
    });
  }
}
