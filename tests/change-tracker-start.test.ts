import assert from "node:assert/strict";
import { test } from "node:test";

import { ChangeTracker } from "../src/context/change-tracker.ts";
import { deferred } from "./fixtures/workbook-deferred.ts";

type ChangeEvent = {
  worksheetId: string;
  address: string;
  changeType: string;
  source: string;
  triggerSource: string;
};

function installWorkbook(sync: () => Promise<void>, options: { abandonFirstRun?: boolean } = {}) {
  const changed = new Set<(event: ChangeEvent) => Promise<void>>();
  const unusedEvent = { add: () => ({ remove: () => {} }) };
  const context = {
    workbook: {
      worksheets: {
        items: [{ id: "sheet", name: "Data" }],
        load: () => {},
        onChanged: {
          add: (handler: (event: ChangeEvent) => Promise<void>) => {
            changed.add(handler);
            return { remove: () => { changed.delete(handler); } };
          },
        },
        onAdded: unusedEvent,
        onDeleted: unusedEvent,
        onNameChanged: unusedEvent,
        onMoved: unusedEvent,
        onVisibilityChanged: unusedEvent,
      },
    },
    sync,
  };
  type Context = typeof context;
  type Callback = (host: Context) => Promise<void>;
  const previous: unknown = Reflect.get(globalThis, "Excel");
  const hadExcel = Reflect.has(globalThis, "Excel");
  let firstRun = true;
  let abandoned: Promise<void> | undefined;
  Reflect.set(globalThis, "Excel", {
    run: (contextOrCallback: Context | Callback, callback?: Callback) => {
      const completion = typeof contextOrCallback === "function"
        ? contextOrCallback(context)
        : callback?.(contextOrCallback);
      if (firstRun && options.abandonFirstRun) {
        firstRun = false;
        abandoned = completion;
        return Promise.reject(new Error("host run ended while its batch was still pending"));
      }
      firstRun = false;
      return completion;
    },
  });
  return {
    finishAbandoned: () => abandoned,
    edit: async () => {
      for (const handler of changed) {
        await handler({ worksheetId: "sheet", address: "A1", changeType: "RangeEdited", source: "Local", triggerSource: "Unknown" });
      }
    },
    restore: () => {
      if (hadExcel) Reflect.set(globalThis, "Excel", previous);
      else Reflect.deleteProperty(globalThis, "Excel");
    },
  };
}

void test("concurrent change-tracker starts await the same listener readiness", async () => {
  const entered = deferred<void>();
  const release = deferred<void>();
  let syncs = 0;
  const workbook = installWorkbook(() => {
    syncs += 1;
    if (syncs === 1) { entered.resolve(); return release.promise; }
    return Promise.resolve();
  });
  try {
    const tracker = new ChangeTracker();
    const first = tracker.start();
    await entered.promise;
    let joinedReady = false;
    const joined = tracker.start().then(() => { joinedReady = true; });
    await Promise.resolve();
    const readyBeforeRegistration = joinedReady;
    release.resolve();
    await Promise.all([first, joined]);
    assert.equal(readyBeforeRegistration, false, "start must not report readiness during registration");
    await workbook.edit();
    assert.equal(tracker.cursor(), 1);
    assert.match((await tracker.describeSince(0, "tab")).text ?? "", /Data: A1/u);
  } finally {
    release.resolve();
    workbook.restore();
  }
});

void test("retrying a partially failed change-tracker registration does not duplicate edit events", async () => {
  let syncs = 0;
  const workbook = installWorkbook(() => {
    syncs += 1;
    return syncs === 2 ? Promise.reject(new Error("host failed after installing listeners")) : Promise.resolve();
  });
  try {
    const tracker = new ChangeTracker();
    await tracker.start();
    await tracker.start();
    await workbook.edit();
    assert.equal(tracker.cursor(), 1, "a retried listener must record each edit only once");
    assert.match((await tracker.describeSince(0, "tab")).text ?? "", /Data: A1/u);
  } finally {
    workbook.restore();
  }
});

void test("a failed registration cannot add late listeners after a successful retry", async () => {
  const release = deferred<void>();
  let syncs = 0;
  const workbook = installWorkbook(() => {
    syncs += 1;
    return syncs === 1 ? release.promise : Promise.resolve();
  }, { abandonFirstRun: true });
  try {
    const tracker = new ChangeTracker();
    await tracker.start();
    await tracker.start();
    release.resolve();
    await workbook.finishAbandoned();
    await workbook.edit();
    assert.equal(tracker.cursor(), 1, "abandoned registration must not install late duplicate listeners");
  } finally {
    release.resolve();
    workbook.restore();
  }
});
