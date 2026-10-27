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

for (const [pkg, bin] of [
  ["proxy", "pi-for-excel-proxy"],
  ["python-bridge", "pi-for-excel-python-bridge"],
  ["tmux-bridge", "pi-for-excel-tmux-bridge"],
]) {
  for (const flag of ["--help", "-h"]) {
    test(`${bin} ${flag} prints usage and exits without side effects`, () => {
      const home = mkdtempSync(path.join(tmpdir(), "pi-cli-help-"));
      try {
        const result = spawnSync(process.execPath, [path.join(root, "pkg", pkg, "cli.mjs"), flag], {
          encoding: "utf8",
          env: { ...process.env, HOME: home, USERPROFILE: home },
          timeout: 5000,
        });
        assert.equal(result.signal, null, "process must exit on its own");
        assert.equal(result.status, 0, result.stderr);
        assert.match(result.stdout, new RegExp(`^Usage: npx ${bin} `));
        assert.deepEqual(readdirSync(home), []);
      } finally {
        rmSync(home, { recursive: true, force: true });
      }
    });
  }
}
