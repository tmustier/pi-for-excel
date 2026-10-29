import assert from "node:assert/strict";
import { test } from "node:test";

import {
  cancelCellEditWaits,
  ExcelCellEditModeError,
  excelRun,
  ExcelTimeoutError,
  ExcelWaitStoppedError,
  onCellEditWait,
} from "../src/excel/helpers.ts";

const CELL_EDIT = { code: "InvalidOperationInCellEditMode", message: "Excel is in cell-editing mode." };

function installHost(isEditing: (run: number, sync: number) => boolean): { runs: () => number; restore: () => void } {
  let runs = 0;
  const previous: unknown = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: async <T>(callback: (ctx: unknown) => Promise<T>): Promise<T> => {
      runs += 1;
      const run = runs;
      let syncs = 0;
      const context = {
        sync: () => {
          syncs += 1;
          return isEditing(run, syncs)
            ? Promise.reject(Object.assign(new Error(CELL_EDIT.message), { code: CELL_EDIT.code }))
            : Promise.resolve();
        },
      };
      return callback(context);
    },
  });
  return { runs: () => runs, restore: () => Reflect.set(globalThis, "Excel", previous) };
}

void test("a batch waits while the user is editing a cell, then runs, with a status while it waits", async () => {
  const host = installHost((run) => run <= 2);
  const states: boolean[] = [];
  const stop = onCellEditWait((waiting) => states.push(waiting));
  try {
    const result = await excelRun(async (context) => {
      await context.sync();
      return "read";
    });
    assert.equal(result, "read");
    assert.equal(host.runs(), 3);
    assert.deepEqual(states, [true, false]);
  } finally {
    stop();
    host.restore();
  }
});

void test("when the user never leaves edit mode, the agent gets a precise error", async () => {
  const host = installHost(() => true);
  try {
    await assert.rejects(
      excelRun(async (context) => { await context.sync(); }, { cellEditWaitMs: 600 }),
      (error: unknown) => {
        assert.ok(error instanceof ExcelCellEditModeError);
        assert.match(error.message, /cell-edit mode/u);
        assert.match(error.message, /Nothing was read or changed/u);
        assert.match(error.message, /Enter, Tab or Esc/u);
        return true;
      },
    );
    assert.ok(host.runs() > 1, "retried before giving up");
  } finally {
    host.restore();
  }
});

void test("a batch is not re-run once an earlier sync reached Excel", async () => {
  const host = installHost((_run, sync) => sync >= 2);
  try {
    await assert.rejects(
      excelRun(async (context) => {
        await context.sync();
        await context.sync();
      }, { cellEditWaitMs: 600 }),
      (error: unknown) => error instanceof ExcelCellEditModeError && error.earlierStepReachedExcel,
    );
    assert.equal(host.runs(), 1);
  } finally {
    host.restore();
  }
});

void test("auto-context reads do not wait for edit mode", async () => {
  const host = installHost(() => true);
  try {
    await assert.rejects(
      excelRun(async (context) => { await context.sync(); }, { waitForCellEdit: false }),
      ExcelCellEditModeError,
    );
    assert.equal(host.runs(), 1);
  } finally {
    host.restore();
  }
});

void test("a batch Excel never answers times out with an outcome-unknown error", async () => {
  const previous: unknown = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", { run: () => new Promise(() => {}) });
  try {
    await assert.rejects(
      excelRun(() => Promise.resolve("never"), { timeoutMs: 50 }),
      (error: unknown) => {
        assert.ok(error instanceof ExcelTimeoutError);
        assert.equal(error.code, "excel_timeout");
        assert.equal(error.outcomeUnknown, true);
        assert.match(error.message, /outcome is unknown/u);
        assert.match(error.message, /Read the affected range back before retrying/u);
        return true;
      },
    );
  } finally {
    Reflect.set(globalThis, "Excel", previous);
  }
});

void test("Stop ends a wait for the user to leave cell-edit mode before anything reaches Excel", async () => {
  const host = installHost(() => true);
  try {
    const run = excelRun(async (context) => { await context.sync(); });
    await new Promise((resolve) => setTimeout(resolve, 50));
    cancelCellEditWaits();
    await assert.rejects(run, ExcelWaitStoppedError);
    const runsAtStop = host.runs();
    await new Promise((resolve) => setTimeout(resolve, 600));
    assert.equal(host.runs(), runsAtStop, "no retry after Stop");
  } finally {
    host.restore();
  }
});
