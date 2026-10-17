import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

async function openStoragePage(): ReturnType<typeof openBrowserPage> {
  return openBrowserPage(env, {
    path: "durable-storage-test",
    prepareContext: async (context) => {
      await context.route("**/durable-storage-test", (route) => route.fulfill({
        contentType: "text/html", body: "<!doctype html><title>Durable storage contract</title>",
      }));
    },
  });
}

void test("passes Durable 1.3.0's public storage conformance cases in Chromium", async () => {
  const opened = await openStoragePage();
  try {
    const names = await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      const { createStorageConformance } = await import('/node_modules/@earendil-works/pi-durable/dist/testing/storage-conformance.js');
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const { Value } = await import('/node_modules/typebox/build/value/index.mjs');
      const check = (condition, message = 'Storage assertion failed') => { if (!condition) throw new Error(message); };
      const partial = (actual, expected) => {
        if (expected === null || typeof expected !== 'object') return Value.Equal(actual, expected);
        return actual !== null && typeof actual === 'object' && Object.entries(expected).every(([key, value]) => partial(actual[key], value));
      };
      let index = 0;
      const cases = createStorageConformance({
        assertions: {
          ok: check,
          strictEqual: (actual, expected) => check(actual === expected),
          deepEqual: (actual, expected) => check(Value.Equal(actual, expected)),
          partialDeepEqual: (actual, expected) => check(partial(actual, expected)),
          greaterThan: (actual, expected) => check(actual > expected),
          rejects: async (operation, includes) => {
            try { await operation; } catch (error) { check(error.message.includes(includes), error.message); return; }
            throw new Error('Expected storage rejection: ' + includes);
          },
        },
        withStorage: async (use) => {
          const storage = await openIndexedDbStorage('durable-conformance-' + index++);
          try { await use(storage); } finally { await storage.close(ctx); }
        },
      });
      for (const item of cases) {
        try { await item.run(); } catch (error) { throw new Error(item.name + ': ' + error.message); }
      }
      return cases.map(item => item.name);
    })()`);
    assert.ok(Array.isArray(names) && names.length === 24, "all upstream 1.3.0 conformance cases ran");
  } finally { await opened.finish(); }
});

void test("commits survive reopen with detached values, fork queries, documents, sequence and unused IDs", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      let storage = await openIndexedDbStorage('durable-reopen');
      const [root, entry, fork, document, unused] = await Promise.all(
        Array.from({ length: 5 }, () => storage.mintId())
      );
      const value = { nested: { text: 'before' } };
      const writes = [
        { type: 'conversation', value: { id: root } },
        { type: 'entry', value: { id: entry, conversationId: root, kind: 'test', data: value, head: entry } },
        { type: 'conversation', value: { id: fork, parent: { conversationId: root, at: entry } } },
        { type: 'document.create', record: {
          id: document, kind: 'test-document', scope: { kind: 'conversation', conversationId: root },
          history: 'rewindable', fork: 'asOf'
        }, content: { kind: 'base', version: 1, value } }
      ];
      const pending = storage.commit(writes, ctx);
      value.nested.text = 'mutated';
      const seq = await pending;
      const read = await storage.document(document, 'current', ctx);
      read.value.nested.text = 'read mutation';
      const changeSeq = await storage.commit([{ type: 'document.change', id: document,
        content: { kind: 'delta', version: 1, ops: [['s', ['nested', 'text'], 'after']] }
      }], ctx);
      await storage.close(ctx);
      let closedRejected = false;
      try { await storage.conversation(root, ctx); } catch { closedRejected = true; }
      storage = await openIndexedDbStorage('durable-reopen');
      const next = await storage.mintId();
      const inherited = await storage.entry(fork, entry, ctx);
      const entries = await storage.scanEntries({ conversationId: fork, order: 'ascending' }, 1, undefined, ctx);
      const marker = await storage.findLatestHeadMarker(fork, undefined, ctx);
      const old = await storage.document(document, seq, ctx);
      const current = await storage.document(document, 'current', ctx);
      const found = await storage.findDocument({ kind: 'test-document', scope: { kind: 'conversation', conversationId: root } }, 'current', ctx);
      const docs = await storage.scanDocuments({ scope: { kind: 'conversation', conversationId: root }, at: 'current' }, 10, undefined, ctx);
      const nextSeq = await storage.commit([], ctx);
      await storage.close(ctx);
      return {
        idsUnique: new Set([root, entry, fork, document, unused]).size === 5,
        unusedPreserved: next > unused,
        seqs: [seq, changeSeq, nextSeq], closedRejected,
        entryText: inherited.entry.data.nested.text, entrySeq: inherited.commitSeq,
        inherited: entries.items.map(item => item.id), expectedEntry: entry,
        marker: marker.head === entry,
        old: old.value.nested.text, current: current.value.nested.text,
        found: found.id === document, docs: docs.items.length
      };
    })()`);
    assert.deepEqual(result, {
      idsUnique: true, unusedPreserved: true, seqs: [1, 2, 3], closedRejected: true,
      entryText: "before", entrySeq: 1, inherited: [3], expectedEntry: 3,
      marker: true, old: "before", current: "after", found: true, docs: 1,
    });
  } finally { await opened.finish(); }
});

void test("a genuinely aborted IndexedDB transaction publishes neither writes nor sequence", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      let storage = await openIndexedDbStorage('durable-abort');
      const id = await storage.mintId();
      const original = IDBDatabase.prototype.transaction;
      let aborted = false;
      IDBDatabase.prototype.transaction = function (...args) {
        const tx = original.apply(this, args);
        if (this.name === 'durable-abort' && args[1] === 'readwrite') {
          queueMicrotask(() => tx.abort());
        }
        return tx;
      };
      try {
        await storage.commit([{ type: 'conversation', value: { id } }], ctx);
      } catch { aborted = true; }
      finally { IDBDatabase.prototype.transaction = original; }
      const absentInMemory = await storage.conversation(id, ctx) === undefined;
      await storage.close(ctx);
      storage = await openIndexedDbStorage('durable-abort');
      const absentAfterReopen = await storage.conversation(id, ctx) === undefined;
      const seq = await storage.commit([{ type: 'conversation', value: { id } }], ctx);
      await storage.close(ctx);
      return { aborted, absentInMemory, absentAfterReopen, seq };
    })()`);
    assert.deepEqual(result, { aborted: true, absentInMemory: true, absentAfterReopen: true, seq: 1 });
  } finally { await opened.finish(); }
});

void test("two panes fail closed, then ownership transfers after close and reload", async () => {
  const opened = await openStoragePage();
  try {
    const second = await opened.context.newPage();
    await second.goto(`${env.baseUrl}durable-storage-test`);
    const acquire = `(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      try { window.storage = await openIndexedDbStorage('durable-owner'); return true; }
      catch { return false; }
    })()`;
    const close = `(async () => {
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      await window.storage.close(ctx);
    })()`;
    assert.equal(await opened.page.evaluate(acquire), true);
    const conversationId = await opened.page.evaluate(`(async () => {
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const id = await window.storage.mintId();
      await window.storage.commit([{ type: 'conversation', value: { id } }], ctx);
      return id;
    })()`);
    assert.equal(await second.evaluate(acquire), false);
    await opened.page.evaluate(close);
    assert.equal(await second.evaluate(acquire), true);
    assert.deepEqual(await second.evaluate(`(async () => {
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      return (await window.storage.scanConversations({}, 10, undefined, ctx)).items;
    })()`), [{ id: conversationId }]);
    assert.equal(await opened.page.evaluate(acquire), false);
    await second.reload();
    assert.equal(await opened.page.evaluate(acquire), true);
    await opened.page.evaluate(close);
    assert.equal(await second.evaluate(acquire), true);
    await second.close();
    assert.equal(await opened.page.evaluate(acquire), true);
    await opened.page.evaluate(close);
  } finally { await opened.finish(); }
});

void test("missing Web Locks fails closed without opening a database", async () => {
  const opened = await openStoragePage();
  try {
    assert.equal(await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      Object.defineProperty(navigator, 'locks', { value: undefined });
      try { await openIndexedDbStorage('durable-no-locks'); return false; }
      catch { return !(await indexedDB.databases()).some(db => db.name === 'durable-no-locks'); }
    })()`), true);
  } finally { await opened.finish(); }
});

void test("failed journal reads reject open and release ownership rather than restoring an empty store", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      let storage = await openIndexedDbStorage('durable-read-failure');
      const id = await storage.mintId();
      await storage.commit([{ type: 'conversation', value: { id } }], ctx);
      await storage.close(ctx);
      const original = IDBDatabase.prototype.transaction;
      IDBDatabase.prototype.transaction = function (...args) {
        const tx = original.apply(this, args);
        if (this.name === 'durable-read-failure' && args[1] === 'readonly') queueMicrotask(() => tx.abort());
        return tx;
      };
      let rejected = false;
      try { await openIndexedDbStorage('durable-read-failure'); } catch { rejected = true; }
      finally { IDBDatabase.prototype.transaction = original; }
      storage = await openIndexedDbStorage('durable-read-failure');
      const restored = await storage.conversation(id, ctx);
      await storage.close(ctx);
      return { rejected, restored: restored.id === id };
    })()`);
    assert.deepEqual(result, { rejected: true, restored: true });
  } finally { await opened.finish(); }
});

void test("corrupt persisted journal bytes fail closed", async () => {
  const opened = await openStoragePage();
  try {
    assert.equal(await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const storage = await openIndexedDbStorage('durable-corrupt');
      await storage.commit([], ctx);
      await storage.close(ctx);
      const db = await new Promise((resolve, reject) => {
        const request = indexedDB.open('durable-corrupt');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      await new Promise((resolve, reject) => {
        const tx = db.transaction('journal', 'readwrite');
        tx.objectStore('journal').put({ seq: 1, writes: '[]', digest: 'damaged' });
        tx.oncomplete = resolve;
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      try { await openIndexedDbStorage('durable-corrupt'); return false; } catch { return true; }
    })()`), true);
  } finally { await opened.finish(); }
});
