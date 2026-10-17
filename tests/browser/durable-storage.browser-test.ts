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
      let reopenAfterCommit = false;
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
          const name = 'durable-conformance-' + index++;
          let storage = await openIndexedDbStorage(name);
          const proxy = new Proxy({}, { get: (_, key) => {
            if (key === 'commit' && reopenAfterCommit) return async (...args) => {
              const seq = await storage.commit(...args);
              await storage.checkpoint();
              const expectedDatabaseId = storage.databaseId;
              await storage.close(ctx);
              storage = await openIndexedDbStorage(name, { expectedDatabaseId });
              return seq;
            };
            return typeof storage[key] === 'function' ? storage[key].bind(storage) : storage[key];
          } });
          try { await use(proxy); } finally { await storage.close(ctx); }
        },
      });
      for (const reopen of [false, true]) {
        reopenAfterCommit = reopen;
        for (const item of cases) {
          try { await item.run(); } catch (error) { throw new Error(item.name + ' (reopen=' + reopen + '): ' + error.message); }
        }
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

void test("physical checkpoints and pruning preserve every historical value, delta count, version and fork across restart", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      const { MemoryStorage } = await import('/node_modules/@earendil-works/pi-durable/dist/storage/memory.js');
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const { Value } = await import('/node_modules/typebox/build/value/index.mjs');
      const check = (a, b) => { if (!Value.Equal(a,b)) throw new Error('Mismatch: ' + JSON.stringify([a,b])); };
      const reference = new MemoryStorage();
      let storage = await openIndexedDbStorage('durable-history');
      const generation = storage.databaseId;
      const commit = async writes => check(await storage.commit(writes,ctx), await reference.commit(writes,ctx));
      const originalGetAll = IDBObjectStore.prototype.getAll;
      IDBObjectStore.prototype.getAll = function() { throw new Error('Unbounded getAll'); };
      const doc = { id: 2, kind: 'history', scope: { kind: 'conversation', conversationId: 1 }, history: 'rewindable', fork: 'asOf' };
      await commit([{type:'conversation',value:{id:1}}, {type:'document.create',record:doc,content:{kind:'base',version:1,value:{n:0}}}]);
      for (let i=1;i<=140;i++) {
        await commit([
          {type:'document.change',id:2,content:{kind:'delta',version:1,ops:[['s',['n'],i*3]]}},
          {type:'entry',value:{id:10+i,conversationId:1,kind:'history',data:{i}, ...(i===70 ? {head:80} : {})}}
        ]);
      }
      await commit([{type:'conversation',value:{id:200,parent:{conversationId:1,at:80}}},
        {type:'document.copy',record:{...doc,id:201,scope:{kind:'conversation',conversationId:200}},source:{id:2,at:66}}]);
      await commit([{type:'document.change',id:2,content:{kind:'base',version:2,value:{n:1001}}}]);
      await commit([{type:'document.change',id:2,content:{kind:'delta',version:2,ops:[['s',['n'],1007]]}}]);
      await commit([{type:'document.retire',id:2}, {type:'document.create',record:{...doc,id:202},content:{kind:'base',version:3,value:{n:9001}}}]);
      const checkpoint = await storage.checkpoint();
      await storage.close(ctx);
      let cursorReads=0, historyReads=0, revisionRows=0;
      const originalCursor=IDBObjectStore.prototype.openCursor;
      const originalIndexCursor=IDBIndex.prototype.openCursor;
      const originalGet=IDBObjectStore.prototype.get;
      const observe=(source,request)=>{
        if(source==='records'||source==='revisions')historyReads++;
        if(source==='revisions')request.addEventListener('success',()=>{if(request.result)revisionRows++;});
        return request;
      };
      IDBObjectStore.prototype.get=function(...args) {
        if(this.name==='records'||this.name==='revisions')historyReads++;
        return originalGet.apply(this,args);
      };
      IDBObjectStore.prototype.openCursor=function(...args) {cursorReads++;return observe(this.name,originalCursor.apply(this,args));};
      IDBIndex.prototype.openCursor=function(...args) {return observe(this.objectStore.name,originalIndexCursor.apply(this,args));};
      storage=await openIndexedDbStorage('durable-history',{expectedDatabaseId:generation});
      const openingCursorReads=cursorReads,openingHistoryReads=historyReads;
      check(await storage.document(2,141,ctx),await reference.document(2,141,ctx));
      const documentRevisionRows=revisionRows;
      IDBObjectStore.prototype.openCursor=originalCursor;
      IDBIndex.prototype.openCursor=originalIndexCursor;
      IDBObjectStore.prototype.get=originalGet;
      for(let seq=1;seq<=145;seq++) {
        check(await storage.document(2,seq,ctx), await reference.document(2,seq,ctx));
        check(await storage.findDocument(doc,seq,ctx), await reference.findDocument(doc,seq,ctx));
      }
      check(await storage.document(201,'current',ctx),await reference.document(201,'current',ctx));
      check(await storage.document(202,'current',ctx),await reference.document(202,'current',ctx));
      check(await storage.findLatestHeadMarker(200,undefined,ctx),await reference.findLatestHeadMarker(200,undefined,ctx));
      for (const order of ['ascending','descending']) {
        for (const maxEntryId of [73,74,75]) {
          const boundary={conversationId:1,minEntryId:11,maxEntryId,order};
          check(await storage.scanEntries(boundary,100,undefined,ctx),await reference.scanEntries(boundary,100,undefined,ctx));
        }
        const end={after:order==='ascending'?Number.MAX_SAFE_INTEGER:0,order};
        check(await storage.scanConversations({},10,end,ctx),{items:[]});
        let cursor;
        do {
          const query={conversationId:200,order};
          const actual=await storage.scanEntries(query,13,cursor,ctx);
          check(actual,await reference.scanEntries(query,13,cursor,ctx));
          for(const entry of actual.items) check(await storage.entry(200,entry.id,ctx),await reference.entry(200,entry.id,ctx));
          cursor=actual.next;
        } while(cursor);
      }
      IDBObjectStore.prototype.getAll=originalGetAll;
      await storage.close(ctx);
      return {sequence:checkpoint.sequence,pruned:checkpoint.pruned,openingCursorReads,openingHistoryReads,documentRevisionRows};
    })()`);
    assert.deepEqual(result, { sequence: 145, pruned: 17, openingCursorReads: 1, openingHistoryReads: 0, documentRevisionRows: 13 });
  } finally { await opened.finish(); }
});

void test("version-one journal migration resumes an aborted batch without losing original sequences or unused IDs", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
      const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const name='durable-v1-migration';
      const batches=[
        [{type:'conversation',value:{id:1}},{type:'document.create',record:{id:2,kind:'old',scope:{kind:'conversation',conversationId:1},history:'rewindable',fork:'asOf'},content:{kind:'base',version:1,value:{n:7}}}],
        [{type:'entry',value:{id:3,conversationId:1,kind:'old',data:'preserve'}}],
        [{type:'document.change',id:2,content:{kind:'delta',version:1,ops:[['s',['n'],19]]}}]
      ];
      const rows=[];
      for(const [i,batch] of batches.entries()) {
        const writes=JSON.stringify(batch),seq=i+1;
        const hash=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(seq+':'+writes));
        rows.push({seq,writes,digest:Array.from(new Uint8Array(hash),n=>n.toString(16).padStart(2,'0')).join('')});
      }
      const db=await new Promise((resolve,reject)=>{const r=indexedDB.open(name,1);r.onupgradeneeded=()=>{r.result.createObjectStore('meta');r.result.createObjectStore('journal',{keyPath:'seq'});};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
      await new Promise((resolve,reject)=>{const tx=db.transaction(['meta','journal'],'readwrite');for(const row of rows)tx.objectStore('journal').put(row);tx.objectStore('meta').put(3,'sequence');tx.objectStore('meta').put(900,'highWatermark');tx.oncomplete=resolve;tx.onabort=()=>reject(tx.error);});
      db.close();
      const original=IDBDatabase.prototype.transaction;
      let writes=0,rejected=false;
      IDBDatabase.prototype.transaction=function(...args){const tx=original.apply(this,args);if(this.name===name&&args[1]==='readwrite'&&++writes===2)queueMicrotask(()=>tx.abort());return tx;};
      try{await openIndexedDbStorage(name);}catch{rejected=true;}finally{IDBDatabase.prototype.transaction=original;}
      const storage=await openIndexedDbStorage(name);
      const entry=await storage.entry(3,ctx);
      const old=await storage.document(2,1,ctx),current=await storage.document(2,'current',ctx);
      const next=await storage.mintId();
      const seq=await storage.commit([],ctx);
      await storage.close(ctx);
      return {rejected,entrySeq:entry.commitSeq,data:entry.entry.data,old:old.value.n,current:current.value.n,count:current.deltasSinceBase,next,seq};
    })()`);
    assert.deepEqual(result, { rejected: true, entrySeq: 2, data: "preserve", old: 7, current: 19, count: 1, next: 901, seq: 4 });
  } finally { await opened.finish(); }
});

void test("missing, replaced and corrupt storage fail distinctly without creating an empty replacement", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const {openIndexedDbStorage}=await import('/src/durable/indexeddb-storage.ts');
      const {BACKGROUND_CONTEXT:ctx}=await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const errorCode=async fn=>{try{await fn();return 'success';}catch(e){return e.code;}};
      const missing=await errorCode(()=>openIndexedDbStorage('durable-missing',{expectedDatabaseId:'old-generation'}));
      const absent=!(await indexedDB.databases()).some(db=>db.name==='durable-missing');
      let storage=await openIndexedDbStorage('durable-damage');
      const databaseId=storage.databaseId;
      await storage.commit([{type:'conversation',value:{id:1}}],ctx);
      await storage.checkpoint();await storage.close(ctx);
      const replaced=await errorCode(()=>openIndexedDbStorage('durable-damage',{expectedDatabaseId:'wrong-generation'}));
      const db=await new Promise((resolve,reject)=>{const r=indexedDB.open('durable-damage');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
      await new Promise((resolve,reject)=>{const tx=db.transaction('records','readwrite');tx.objectStore('records').delete(1);tx.oncomplete=resolve;tx.onabort=()=>reject(tx.error);});db.close();
      const corrupt=await errorCode(()=>openIndexedDbStorage('durable-damage',{expectedDatabaseId:databaseId}));
      return {missing,absent,replaced,corrupt};
    })()`);
    assert.deepEqual(result, { missing: "missing", absent: true, replaced: "replaced", corrupt: "corrupt" });
  } finally { await opened.finish(); }
});

void test("retention never infers workbook absence; deletion requires authorization, generation and exclusive ownership", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const {openIndexedDbStorage,deleteDurableWorkbookStorage}=await import('/src/durable/indexeddb-storage.ts');
      const {BACKGROUND_CONTEXT:ctx}=await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const name='pi-for-excel-durable:obsolete-test',keep='pi-for-excel-durable:not-recently-opened';
      const retained=await openIndexedDbStorage(keep);await retained.commit([{type:'conversation',value:{id:1}}],ctx);await retained.close(ctx);
      const storage=await openIndexedDbStorage(name);const databaseId=storage.databaseId;
      const errorCode=async fn=>{try{await fn();return 'success';}catch(e){return e.code;}};
      const options={name,databaseId,authorization:'user-confirmed-obsolete-workbook'};
      const owned=await errorCode(()=>deleteDurableWorkbookStorage(options));await storage.close(ctx);
      const unauthorized=await errorCode(()=>deleteDurableWorkbookStorage({...options,authorization:undefined}));
      const replaced=await errorCode(()=>deleteDurableWorkbookStorage({...options,databaseId:'wrong'}));
      const connection=await new Promise((resolve,reject)=>{const r=indexedDB.open(name);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
      let settled=false;
      const blocked=new Promise(resolve=>{connection.onversionchange=resolve;});
      const deletion=deleteDurableWorkbookStorage(options).then(()=>{settled=true;});
      await blocked;await new Promise(resolve=>setTimeout(resolve,30));
      const remainedPending=!settled;
      const lockedWhileBlocked=await errorCode(()=>openIndexedDbStorage(name));
      connection.close();await deletion;
      const absent=!(await indexedDB.databases()).some(db=>db.name===name);
      const retainedAgain=await openIndexedDbStorage(keep);
      const untouched=(await retainedAgain.conversation(1,ctx)).id===1;await retainedAgain.close(ctx);
      return {owned,unauthorized,replaced,remainedPending,lockedWhileBlocked,absent,untouched};
    })()`);
    assert.deepEqual(result, { owned: "owned", unauthorized: "corrupt", replaced: "replaced", remainedPending: true, lockedWhileBlocked: "owned", absent: true, untouched: true });
  } finally { await opened.finish(); }
});

void test("aborted checkpoint and ID reservation retain journal and allocator state on restart", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const {openIndexedDbStorage}=await import('/src/durable/indexeddb-storage.ts');
      const {BACKGROUND_CONTEXT:ctx}=await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      let storage=await openIndexedDbStorage('durable-maintenance-abort');
      await storage.commit([{type:'conversation',value:{id:5}}],ctx);
      const original=IDBDatabase.prototype.transaction;
      IDBDatabase.prototype.transaction=function(...args){const tx=original.apply(this,args);if(this.name==='durable-maintenance-abort'&&args[1]==='readwrite')queueMicrotask(()=>tx.abort());return tx;};
      let checkpointFailed=false,mintFailed=false;
      try{await storage.checkpoint();}catch{checkpointFailed=true;}
      try{await storage.mintId();}catch{mintFailed=true;}
      IDBDatabase.prototype.transaction=original;
      await storage.close(ctx);storage=await openIndexedDbStorage('durable-maintenance-abort');
      const checkpoint=await storage.checkpoint(),id=await storage.mintId();
      const conversation=await storage.conversation(5,ctx);await storage.close(ctx);
      return {checkpointFailed,mintFailed,checkpoint,id,conversation};
    })()`);
    assert.deepEqual(result, { checkpointFailed: true, mintFailed: true, checkpoint: { sequence: 1, pruned: 1 }, id: 6, conversation: { id: 5 } });
  } finally { await opened.finish(); }
});

void test("checksums reject damaged materialized records and revisions after journal pruning", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const {openIndexedDbStorage}=await import('/src/durable/indexeddb-storage.ts');
      const {BACKGROUND_CONTEXT:ctx}=await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      const results=[];
      for(const store of ['records','revisions']) {
        const name='durable-checksum-'+store;
        let storage=await openIndexedDbStorage(name);
        await storage.commit([{type:'conversation',value:{id:1}},{type:'document.create',record:{id:2,kind:'test',scope:{kind:'session'}},content:{kind:'base',version:1,value:{text:'precious'}}}],ctx);
        await storage.checkpoint();await storage.close(ctx);
        const db=await new Promise((resolve,reject)=>{const r=indexedDB.open(name);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
        await new Promise((resolve,reject)=>{const tx=db.transaction(store,'readwrite');const key=store==='records'?2:[2,1];const r=tx.objectStore(store).get(key);r.onsuccess=()=>tx.objectStore(store).put({...r.result,text:r.result.text.replace('precious','damaged').replace('test','lost')});tx.oncomplete=resolve;tx.onabort=()=>reject(tx.error);});db.close();
        storage=await openIndexedDbStorage(name);
        try{await storage.document(2,'current',ctx);results.push('success');}catch(e){results.push(e.code);}
        await storage.close(ctx);
      }
      return results;
    })()`);
    assert.deepEqual(result, ["corrupt", "corrupt"]);
  } finally { await opened.finish(); }
});

void test("persistence requests report denial and unsupported APIs without claiming protection", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const {openIndexedDbStorage}=await import('/src/durable/indexeddb-storage.ts');
      const {BACKGROUND_CONTEXT:ctx}=await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      let requested=0;
      const statuses=[];
      for (const mode of ['granted','denied','throws','unsupported']) {
        Object.defineProperty(navigator,'storage',{configurable:true,value:mode==='unsupported'?{}:{persisted:async()=>false,persist:async()=>{requested++;if(mode==='throws')throw new Error('denied');return mode==='granted';}}});
        const storage=await openIndexedDbStorage('durable-persistence-'+mode);
        statuses.push(storage.persistence);await storage.close(ctx);
      }
      return {statuses,requested};
    })()`);
    assert.deepEqual(result, { statuses: ["persistent", "best-effort", "best-effort", "unsupported"], requested: 3 });
  } finally { await opened.finish(); }
});

void test("request deduplication follows commit order without resurrecting overwritten owners", async () => {
  const opened = await openStoragePage();
  try {
    const result = await opened.page.evaluate(`(async () => {
      const {openIndexedDbStorage}=await import('/src/durable/indexeddb-storage.ts');
      const {BACKGROUND_CONTEXT:ctx}=await import('/node_modules/@earendil-works/chord/dist/context/index.js');
      let storage=await openIndexedDbStorage('durable-request-owner');
      const write=(id,requestId)=>({type:'submission',value:{id,conversationId:1,type:'input',status:'queued',requestId}});
      await storage.commit([{type:'conversation',value:{id:1}},write(20,'same'),write(10,'same')],ctx);
      const first=(await storage.submissionByRequest(1,'same',ctx)).id;
      await storage.commit([write(10,'moved')],ctx);
      await storage.checkpoint();await storage.close(ctx);storage=await openIndexedDbStorage('durable-request-owner');
      const notResurrected=await storage.submissionByRequest(1,'same',ctx)===undefined;
      const moved=(await storage.submissionByRequest(1,'moved',ctx)).id;
      const original=(await storage.submission(20,ctx)).requestId;
      await storage.close(ctx);return {first,notResurrected,moved,original};
    })()`);
    assert.deepEqual(result, { first: 10, notResurrected: true, moved: 10, original: "same" });
  } finally { await opened.finish(); }
});
