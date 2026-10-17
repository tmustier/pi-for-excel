import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

const reader = new URL("../scripts/read-excel-durable.py", import.meta.url).pathname;
const fixture = `
import pathlib, sqlite3, struct, sys
base = pathlib.Path(sys.argv[1]) / 'Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/Default/origin/profile/IndexedDB'
for name in ['pi-for-excel', 'pi-for-excel-durable:scratch']:
    path = base / name.replace(':', '_') / 'IndexedDB.sqlite3'
    path.parent.mkdir(parents=True)
    with sqlite3.connect(path) as db:
        db.execute('create table IDBDatabaseInfo (key text, value text)')
        db.execute('insert into IDBDatabaseInfo values (?,?)', ('DatabaseName', name.encode('utf-16-le')))
        db.execute('create table ObjectStoreInfo (id integer, name text)')
        db.execute('insert into ObjectStoreInfo values (1,?)', ('settings' if name == 'pi-for-excel' else 'journal',))
        db.execute('create table Records (objectStoreID integer, key text, value blob)')
        text = b'not-a-real-secret' if name == 'pi-for-excel' else b'persisted scratch turn'
        value = struct.pack('<I', 15) + b'\\x10' + struct.pack('<I', 0x80000000 | len(text)) + text
        # Numeric keys occupy SQLite TEXT but are not valid UTF-8 in WebKit.
        db.execute('insert into Records values (1,cast(? as text),?)', (b'\\x00\\x10\\xff', value))
`;

function setup(t) {
  const home = mkdtempSync(join(tmpdir(), "durable-reader-test-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const prepared = spawnSync("python3", ["-c", fixture, home], { encoding: "utf8" });
  assert.equal(prepared.status, 0, prepared.stderr);
  return (...args) => spawnSync("python3", [reader, ...args], {
    encoding: "utf8", env: { ...process.env, HOME: home, PYTHONDONTWRITEBYTECODE: "1" },
  });
}

void test("independent Durable listing excludes legacy settings and returns only store counts", (t) => {
  const run = setup(t);
  const result = run();
  assert.equal(result.status, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), [{
    database: "pi-for-excel-durable:scratch", stores: { journal: { records: 1 } },
  }]);
  assert.doesNotMatch(result.stdout, /not-a-real-secret|persisted scratch turn/);
});

void test("independent Durable inspection decodes WebKit records with binary numeric keys", (t) => {
  const run = setup(t);
  const result = run("--database", "pi-for-excel-durable:scratch");
  assert.equal(result.status, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), [{
    database: "pi-for-excel-durable:scratch",
    stores: { journal: [{ key: "0010ff", value: "persisted scratch turn" }] },
  }]);
});

void test("independent Durable reader rejects legacy settings and absent databases", (t) => {
  const run = setup(t);
  assert.notEqual(run("--database", "pi-for-excel").status, 0);
  const absent = run("--database", "pi-for-excel-durable:absent");
  assert.notEqual(absent.status, 0);
  assert.match(absent.stderr, /was not found/);
});
