import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

const reader = new URL("../.agents/skills/excel-background-verification/scripts/read-taskpane-setting.py", import.meta.url).pathname;
const fixture = `
import json, os, pathlib, sqlite3, struct, sys
base = pathlib.Path(sys.argv[1]) / 'Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/Default'
for index, spec in enumerate(json.loads(sys.argv[2])):
    path = base / spec.get('origin', 'origin') / 'profile/IndexedDB' / str(index) / 'IndexedDB.sqlite3'
    path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(path) as db:
        db.execute('create table IDBDatabaseInfo (key text, value blob)')
        name = spec['name']
        db.execute('insert into IDBDatabaseInfo values (?,?)', ('DatabaseName', name if spec.get('textName') else name.encode('utf-16-le')))
        db.execute('create table ObjectStoreInfo (id integer, name text)')
        db.execute('insert into ObjectStoreInfo values (7,?)', (spec.get('store', 'settings'),))
        # Non-app databases deliberately lack Records: selection must read metadata only.
        if name == 'pi-for-excel' and spec.get('store', 'settings') == 'settings':
            db.execute('create table Records (objectStoreID integer, key blob, value blob)')
            text = 'synthetic.acceptance'
            key = b'\\x00\\x60' + struct.pack('<I', len(text)) + text.encode('utf-16-le')
            value = struct.pack('<I', 15) + b'\\x05' + struct.pack('<i', spec.get('value', 42))
            db.execute('insert into Records values (7,?,?)', (key, value))
    os.utime(path, (spec['mtime'], spec['mtime']))
`;
const app = { name: "pi-for-excel", mtime: 1000 };
const durable = { name: "pi-for-excel-durable:scratch", store: "journal", mtime: 2000 };

function setup(t, databases = [app, durable]) {
  const home = mkdtempSync(join(tmpdir(), "settings-reader-test-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const prepared = spawnSync("python3", ["-c", fixture, home, JSON.stringify(databases)], { encoding: "utf8" });
  assert.equal(prepared.status, 0, prepared.stderr);
  const run = (...args) => spawnSync("python3", [reader, ...args], {
    encoding: "utf8", env: { ...process.env, HOME: home, PYTHONDONTWRITEBYTECODE: "1" },
  });
  run.profile = join(home, "Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/Default");
  return run;
}

void test("settings readback ignores newer Durable database and decodes the older exact app setting", (t) => {
  const run = setup(t);
  const result = run("synthetic.acceptance");
  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(result.stdout), 42);
  const listed = run("--list");
  assert.equal(listed.status, 0, listed.stderr);
  assert.deepEqual(JSON.parse(listed.stdout), ["synthetic.acceptance"]);
});

void test("settings selection requires exact app identity and the settings store", (t) => {
  const run = setup(t, [
    app,
    { name: "pi-for-excel-backup", mtime: 2000 },
    { name: "pi-for-excel", store: "journal", mtime: 3000 },
    { name: "pi-for-excel-durable:scratch", mtime: 4000 },
  ]);
  const result = run("synthetic.acceptance");
  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(result.stdout), 42);
});

void test("multiple qualified origins retain newest app-file selection, not newest arbitrary database", (t) => {
  const run = setup(t, [
    app,
    { ...app, origin: "second-origin", textName: true, mtime: 2000, value: 73 },
    { ...durable, origin: "third-origin", mtime: 3000 },
  ]);
  const result = run("synthetic.acceptance");
  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(result.stdout), 73);
});

void test("missing app identity or settings store fails instead of returning an empty listing", (t) => {
  for (const databases of [[], [durable], [{ ...app, store: "journal" }]]) {
    const result = setup(t, databases)("--list");
    assert.notEqual(result.status, 0);
    assert.equal(result.stdout, "");
    assert.match(result.stderr, /pi-for-excel.*settings|storage was not found/);
  }
});

for (const nested of [false, true]) {
  void test(`denied ${nested ? "nested origin" : "container"} access is not reported as absent settings`, { skip: process.getuid?.() === 0 }, (t) => {
    const run = setup(t);
    const deniedPath = nested ? join(run.profile, "origin") : run.profile;
    chmodSync(deniedPath, 0);
    try {
      const result = run("--list");
      assert.notEqual(result.status, 0);
      assert.equal(result.stdout, "");
      assert.match(result.stderr, /permission was denied/);
      assert.doesNotMatch(result.stderr, /no `pi-for-excel`/);
    } finally {
      chmodSync(deniedPath, 0o700);
    }
  });
}
