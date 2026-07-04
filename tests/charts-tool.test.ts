import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import {
  createChartsTool,
  executeChartsAction,
  toExcelChartType,
} from "../src/tools/charts.ts";
import { applyChartState } from "../src/workbook/recovery/chart-state.ts";
import {
  createPersistedWorkbookRecoveryPayload,
  parsePersistedSnapshots,
} from "../src/workbook/recovery/log-codec.ts";
import type { ChartsDetails } from "../src/tools/tool-details.ts";
import { getToolContextImpact, getToolExecutionMode } from "../src/tools/execution-policy.ts";
import { WorkbookRecoveryLog, type WorkbookRecoverySnapshot } from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import type { AppendWorkbookChangeAuditEntryArgs } from "../src/audit/workbook-change-audit.ts";
import type {
  RecoveryChartAbsentState,
  RecoveryChartPresentState,
  RecoveryChartState,
} from "../src/workbook/recovery-states.ts";
import {
  createInMemorySettingsStore,
  findSnapshotById,
  withoutUndefined,
} from "./recovery-log-test-helpers.test.ts";

function firstText(result: AgentToolResult<ChartsDetails>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") {
    throw new Error("Expected first content block to be text.");
  }

  return first.text;
}

function imageBlock(result: AgentToolResult<ChartsDetails>): { data: string; mimeType: string } {
  const block = result.content.find((item) => item.type === "image");
  if (!block || block.type !== "image") {
    throw new Error("Expected image content block.");
  }

  return {
    data: block.data,
    mimeType: block.mimeType,
  };
}

class FakeLoadable {
  load(_propertyNames?: string | string[]): void {}
}

class FakeRange extends FakeLoadable {
  readonly address: string;

  constructor(address: string) {
    super();
    this.address = address;
  }
}

class FakeChartTitle extends FakeLoadable {
  text: string;
  visible: boolean;

  constructor(text = "", visible = false) {
    super();
    this.text = text;
    this.visible = visible;
  }
}

class FakeChartAxis {
  readonly title = new FakeChartTitle();
}

class FakeChartAxes {
  readonly categoryAxis = new FakeChartAxis();
  readonly valueAxis = new FakeChartAxis();
}

class FakeChartLegend extends FakeLoadable {
  position = "Right";
  visible = true;
}

let fakeChartIdCounter = 0;

class FakeChart extends FakeLoadable {
  readonly id = `fake-chart-id-${(fakeChartIdCounter += 1)}`;
  readonly title = new FakeChartTitle();
  readonly legend = new FakeChartLegend();
  readonly axes = new FakeChartAxes();
  top = 10;
  left = 20;
  width = 500;
  height = 300;
  deleted = false;
  sourceAddress = "";
  seriesBy = "";
  positionStart = "";
  positionEnd: string | undefined;
  imageBase64 = "iVBORw0KGgo=";
  imageRequest: { width?: number; height?: number; fittingMode?: string } | null = null;

  name: string;
  chartType: string;
  private readonly collection: FakeChartCollection;

  constructor(name: string, chartType: string, collection: FakeChartCollection) {
    super();
    this.name = name;
    this.chartType = chartType;
    this.collection = collection;
  }

  setPosition(start: string, end?: string): void {
    this.positionStart = start;
    this.positionEnd = end;
  }

  setData(range: FakeRange, seriesBy?: string): void {
    this.sourceAddress = range.address;
    this.seriesBy = seriesBy ?? "";
  }

  delete(): void {
    this.deleted = true;
    this.collection.remove(this);
  }

  getImage(width?: number, height?: number, fittingMode?: string): { value: string } {
    this.imageRequest = { width, height, fittingMode };
    return { value: this.imageBase64 };
  }
}

class FakeChartCollection extends FakeLoadable {
  readonly items: FakeChart[] = [];

  /**
   * Mirrors Office.js load-path semantics enough to reject collection loads
   * that reference item-level fields without the `items/` prefix (e.g. the
   * invalid "items/id,name", where `name` targets the collection itself).
   */
  override load(propertyNames?: string | string[]): void {
    if (typeof propertyNames !== "string") return;

    for (const segment of propertyNames.split(",")) {
      const trimmed = segment.trim();
      if (trimmed === "items" || trimmed === "count" || trimmed.startsWith("items/")) {
        continue;
      }

      throw new Error(`Invalid ChartCollection load path segment: "${trimmed}"`);
    }
  }

  get count(): number {
    return this.items.length;
  }

  add(chartType: string, sourceRange: FakeRange, seriesBy?: string): FakeChart {
    const chart = new FakeChart(`Chart ${this.items.length + 1}`, chartType, this);
    chart.sourceAddress = sourceRange.address;
    chart.seriesBy = seriesBy ?? "";
    this.items.push(chart);
    return chart;
  }

  remove(chart: FakeChart): void {
    const index = this.items.indexOf(chart);
    if (index >= 0) {
      this.items.splice(index, 1);
    }
  }
}

class FakeWorksheet extends FakeLoadable {
  readonly charts = new FakeChartCollection();
  readonly name: string;

  constructor(name: string) {
    super();
    this.name = name;
  }

  getRange(address: string): FakeRange {
    return new FakeRange(address);
  }
}

class FakeWorksheetCollection extends FakeLoadable {
  readonly items: FakeWorksheet[];

  constructor(sheets: FakeWorksheet[]) {
    super();
    this.items = sheets;
  }

  getActiveWorksheet(): FakeWorksheet {
    const first = this.items[0];
    if (!first) {
      throw new Error("No active worksheet.");
    }
    return first;
  }

  getItem(name: string): FakeWorksheet {
    const sheet = this.items.find((candidate) => candidate.name === name);
    if (!sheet) {
      throw new Error(`Worksheet ${name} not found.`);
    }
    return sheet;
  }
}

class FakeWorkbook {
  readonly worksheets: FakeWorksheetCollection;

  constructor(sheets: FakeWorksheet[]) {
    this.worksheets = new FakeWorksheetCollection(sheets);
  }
}

class FakeContext {
  readonly workbook: FakeWorkbook;
  syncCount = 0;

  constructor(sheets: FakeWorksheet[]) {
    this.workbook = new FakeWorkbook(sheets);
  }

  sync(): Promise<void> {
    this.syncCount += 1;
    return Promise.resolve();
  }
}

async function withFakeExcel<T>(context: FakeContext, fn: () => Promise<T>): Promise<T> {
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  const fakeExcel = {
    run: <TResult>(callback: (ctx: FakeContext) => Promise<TResult>): Promise<TResult> => callback(context),
  };

  Reflect.set(globalThis, "Excel", fakeExcel);
  try {
    return await fn();
  } finally {
    if (hadExcel) {
      Reflect.set(globalThis, "Excel", previousExcel);
    } else {
      Reflect.deleteProperty(globalThis, "Excel");
    }
  }
}

function createChartState(name = "Sales"): RecoveryChartPresentState {
  return {
    kind: "chart_present",
    sheetName: "Sheet1",
    name,
    chartType: "ColumnClustered",
    title: { text: "Old title", visible: true },
    legend: { position: "Right", visible: true },
    xAxisTitle: { text: "Month", visible: true },
    yAxisTitle: { text: "Revenue", visible: true },
    position: { top: 10, left: 20, width: 500, height: 300 },
  };
}

void test("maps friendly chart type names to Office.js chart types", () => {
  assert.equal(toExcelChartType("column"), "ColumnClustered");
  assert.equal(toExcelChartType("scatter"), "XYScatter");
  assert.equal(toExcelChartType("scatter_lines"), "XYScatterLines");
  assert.equal(toExcelChartType("scatter_smooth"), "XYScatterSmooth");
  assert.throws(() => toExcelChartType("combo"), /Invalid chart_type/u);
});

void test("validates action-specific required params", async () => {
  let executeCalled = false;
  const tool = createChartsTool({
    executeAction: () => {
      executeCalled = true;
      return Promise.resolve({
        result: {
          content: [{ type: "text", text: "should not run" }],
          details: { kind: "charts", action: "create" },
        },
      });
    },
    appendAuditEntry: () => Promise.resolve(),
  });

  const createResult = await tool.execute("tc-invalid-create", { action: "create" });
  assert.match(firstText(createResult), /source_range is required/u);
  assert.equal(executeCalled, false);

  const imageResult = await tool.execute("tc-invalid-image", { action: "get_image" });
  assert.match(firstText(imageResult), /name is required/u);
});

void test("creates charts from a mocked worksheet range", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const context = new FakeContext([sheet]);

  const result = await withFakeExcel(context, () => executeChartsAction({
    action: "create",
    source_range: "A1:B12",
    chart_type: "column",
    series_by: "columns",
    name: "Sales",
    title: "Sales by month",
    legend_position: "right",
    x_axis_title: "Month",
    y_axis_title: "Revenue",
    position: "D2:J18",
  }));

  const chart = sheet.charts.items[0];
  assert.ok(chart);
  assert.equal(chart.name, "Sales");
  assert.equal(chart.chartType, "ColumnClustered");
  assert.equal(chart.sourceAddress, "A1:B12");
  assert.equal(chart.seriesBy, "Columns");
  assert.equal(chart.title.text, "Sales by month");
  assert.equal(chart.title.visible, true);
  assert.equal(chart.legend.position, "Right");
  assert.equal(chart.axes.categoryAxis.title.text, "Month");
  assert.equal(chart.axes.valueAxis.title.text, "Revenue");
  assert.equal(chart.positionStart, "D2");
  assert.equal(chart.positionEnd, "J18");
  assert.match(firstText(result.result), /Chart 'Sales' created — Sheet1!A1:B12/u);
  assert.deepEqual(withoutUndefined(result.recoveryState), {
    kind: "chart_absent",
    sheetName: "Sheet1",
    name: "Sales",
    chartId: chart.id,
  });
});

void test("lists charts across mocked worksheets", async () => {
  const sheet1 = new FakeWorksheet("Sheet1");
  const sheet2 = new FakeWorksheet("Dash");
  const chart1 = sheet1.charts.add("ColumnClustered", sheet1.getRange("A1:B12"), "Columns");
  chart1.name = "Sales";
  chart1.title.text = "Sales by month";
  chart1.title.visible = true;
  const chart2 = sheet2.charts.add("Line", sheet2.getRange("C1:D12"), "Rows");
  chart2.name = "Margin";
  const context = new FakeContext([sheet1, sheet2]);

  const result = await withFakeExcel(context, () => executeChartsAction({ action: "list" }));

  assert.match(firstText(result.result), /Charts \(2\)/u);
  assert.equal(result.result.details.action, "list");
  assert.equal(result.result.details.count, 2);
  assert.deepEqual(result.result.details.charts?.map((chart) => chart.name), ["Sales", "Margin"]);
  assert.equal(context.syncCount, 3);
});

void test("updates chart properties and source data through mocked Office objects", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const chart = sheet.charts.add("ColumnClustered", sheet.getRange("A1:B12"), "Columns");
  chart.name = "Sales";
  const context = new FakeContext([sheet]);

  const result = await withFakeExcel(context, () => executeChartsAction({
    action: "update",
    name: "Sales",
    new_name: "Sales Chart",
    source_range: "B1:C12",
    series_by: "rows",
    chart_type: "line",
    title: "Updated sales",
    legend_position: "bottom",
  }));

  assert.equal(chart.name, "Sales Chart");
  assert.equal(chart.chartType, "Line");
  assert.equal(chart.sourceAddress, "B1:C12");
  assert.equal(chart.seriesBy, "Rows");
  assert.equal(chart.title.text, "Updated sales");
  assert.equal(chart.legend.position, "Bottom");
  assert.equal(result.sourceRangeChanged, true);
  assert.match(firstText(result.result), /Chart 'Sales Chart' updated — source Sheet1!B1:C12/u);
});

void test("deletes a chart through mocked Office objects", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const chart = sheet.charts.add("ColumnClustered", sheet.getRange("A1:B12"), "Columns");
  chart.name = "Sales";
  const context = new FakeContext([sheet]);

  const result = await withFakeExcel(context, () => executeChartsAction({
    action: "delete",
    name: "Sales",
  }));

  assert.equal(chart.deleted, true);
  assert.equal(sheet.charts.items.length, 0);
  assert.match(firstText(result.result), /Chart 'Sales' deleted/u);
});

void test("get_image returns text plus image content and structured PNG details", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const chart = sheet.charts.add("ColumnClustered", sheet.getRange("A1:B12"), "Columns");
  chart.name = "Sales";
  chart.width = 500;
  chart.height = 300;
  chart.imageBase64 = "base64-png";
  const context = new FakeContext([sheet]);

  const result = await withFakeExcel(context, () => executeChartsAction({
    action: "get_image",
    name: "Sales",
    width: 1_600,
  }));

  const text = firstText(result.result);
  assert.match(text, /Captured chart 'Sales' \(1200×720 px\)/u);
  assert.doesNotMatch(text, /base64-png/u);
  assert.deepEqual(imageBlock(result.result), {
    data: "base64-png",
    mimeType: "image/png",
  });
  assert.deepEqual(result.result.details.image, {
    base64: "base64-png",
    mimeType: "image/png",
    width: 1200,
    height: 720,
  });
  assert.deepEqual(chart.imageRequest, {
    width: 1200,
    height: 720,
    fittingMode: "Fit",
  });
});

void test("classifies charts actions by execution policy and context impact", () => {
  assert.equal(getToolExecutionMode("charts", { action: "list" }), "read");
  assert.equal(getToolContextImpact("charts", { action: "list" }), "none");

  assert.equal(getToolExecutionMode("charts", { action: "get_image" }), "read");
  assert.equal(getToolContextImpact("charts", { action: "get_image" }), "none");

  assert.equal(getToolExecutionMode("charts", { action: "create" }), "mutate");
  assert.equal(getToolContextImpact("charts", { action: "create" }), "structure");

  assert.equal(getToolExecutionMode("charts", { action: "update" }), "mutate");
  assert.equal(getToolContextImpact("charts", { action: "update" }), "content");

  assert.equal(getToolExecutionMode("charts", { action: "delete" }), "mutate");
  assert.equal(getToolContextImpact("charts", { action: "delete" }), "structure");
});

void test("update captures a chart checkpoint and appends audit metadata", async () => {
  const beforeState = createChartState("Sales");
  let appendedState: RecoveryChartState | null = null;
  let auditEntry: AppendWorkbookChangeAuditEntryArgs | null = null;
  let dispatchedSnapshotId = "";

  const tool = createChartsTool({
    captureChartPresent: () => Promise.resolve(beforeState),
    executeAction: () => Promise.resolve({
      result: {
        content: [{ type: "text", text: "Chart 'Sales' updated." }],
        details: {
          kind: "charts",
          action: "update",
          name: "Sales",
          address: "Sheet1!Sales",
        },
      },
      outputAddress: "Sheet1!Sales",
      changedCount: 1,
      auditSummary: "updated chart Sales",
      sourceRangeChanged: true,
    }),
    appendRecoverySnapshot: (args) => {
      appendedState = args.chartState;
      const snapshot: WorkbookRecoverySnapshot = {
        id: "snap-chart-1",
        at: 1700000000000,
        toolName: "charts",
        toolCallId: args.toolCallId,
        address: args.address,
        changedCount: args.changedCount ?? 1,
        cellCount: 1,
        beforeValues: [],
        beforeFormulas: [],
        snapshotKind: "chart_state",
        chartState: args.chartState,
        workbookId: "url_sha256:workbook",
      };
      return Promise.resolve(snapshot);
    },
    appendAuditEntry: (entry) => {
      auditEntry = entry;
      return Promise.resolve();
    },
    dispatchSnapshotCreated: (snapshot) => {
      dispatchedSnapshotId = snapshot.id;
    },
  });

  const result = await tool.execute("tc-update-checkpoint", {
    action: "update",
    name: "Sales",
    source_range: "B1:C12",
    title: "Updated",
  });

  assert.deepEqual(appendedState, beforeState);
  assert.equal(result.details.recovery?.status, "checkpoint_created");
  assert.equal(result.details.recovery?.snapshotId, "snap-chart-1");
  assert.equal(dispatchedSnapshotId, "snap-chart-1");
  assert.equal(auditEntry?.toolName, "charts");
  assert.equal(auditEntry?.blocked, false);
  assert.match(firstText(result), /property backup only/u);
});

void test("delete explicitly signals no backup and does not append a recovery snapshot", async () => {
  let appendRecoveryCalls = 0;

  const tool = createChartsTool({
    executeAction: () => Promise.resolve({
      result: {
        content: [{ type: "text", text: "Chart 'Sales' deleted." }],
        details: {
          kind: "charts",
          action: "delete",
          name: "Sales",
          address: "Sheet1!Sales",
        },
      },
      outputAddress: "Sheet1!Sales",
      changedCount: 1,
      auditSummary: "deleted chart Sales",
    }),
    appendRecoverySnapshot: () => {
      appendRecoveryCalls += 1;
      return Promise.resolve(null);
    },
    appendAuditEntry: () => Promise.resolve(),
  });

  const result = await tool.execute("tc-delete-no-backup", {
    action: "delete",
    name: "Sales",
  });

  assert.equal(appendRecoveryCalls, 0);
  assert.equal(result.details.recovery?.status, "not_available");
  assert.match(firstText(result), /Backup not created/u);
  assert.match(firstText(result), /cannot faithfully recreate deleted charts/u);
});

void test("workbook recovery log persists and restores chart checkpoints", async () => {
  const settingsStore = createInMemorySettingsStore();
  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:charts-workbook",
    workbookName: "Charts.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-chart-${idCounter}`;
  };

  const restoredState = createChartState("Sales");
  const currentState = createChartState("Sales (current)");
  let appliedAddress = "";
  let appliedState: RecoveryChartState | null = null;

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000000100,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyChartSnapshot: (address, state) => {
      appliedAddress = address;
      appliedState = state;
      // Simulate a restore that renamed the chart: the post-restore identity
      // differs from the checkpoint address.
      return Promise.resolve({ state: currentState, address: "Sheet1!Sales (restored)" });
    },
  });

  const appended = await log.appendChart({
    toolName: "charts",
    toolCallId: "tc-recovery-chart",
    address: "Sheet1!Sales",
    changedCount: 1,
    chartState: restoredState,
  });

  assert.ok(appended);
  const restored = await log.restore(appended.id);

  assert.equal(restored.restoredSnapshotId, appended.id);
  assert.equal(restored.inverseSnapshotId, "snap-chart-2");
  assert.equal(appliedAddress, "Sheet1!Sales");
  assert.deepEqual(withoutUndefined(appliedState), withoutUndefined(restoredState));

  const snapshots = await log.listForCurrentWorkbook(10);
  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse.snapshotKind, "chart_state");
  assert.deepEqual(withoutUndefined(inverse.chartState), withoutUndefined(currentState));
  // The inverse must live at the post-restore identity so that restoring the
  // rollback can find the chart after a rename-restore.
  assert.equal(inverse.address, "Sheet1!Sales (restored)");
});

void test("chart_absent restore can skip inverse snapshot when recreation would be impossible", async () => {
  const settingsStore = createInMemorySettingsStore();
  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:charts-create-workbook",
    workbookName: "Charts.xlsx",
    source: "document.url",
  };

  const absentState: RecoveryChartAbsentState = {
    kind: "chart_absent",
    sheetName: "Sheet1",
    name: "New Chart",
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000000200,
    createId: () => "snap-create-chart",
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyChartSnapshot: () => Promise.resolve({ state: null, address: "Sheet1!New Chart" }),
  });

  const appended = await log.appendChart({
    toolName: "charts",
    toolCallId: "tc-create-recovery",
    address: "Sheet1!New Chart",
    changedCount: 1,
    chartState: absentState,
  });

  assert.ok(appended);
  const restored = await log.restore(appended.id);
  assert.equal(restored.inverseSnapshotId, null);
});

void test("chart rename restore round-trips through the inverse address", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const context = new FakeContext([sheet]);
  const chart = sheet.charts.add("ColumnClustered", new FakeRange("A1:B12"));
  chart.name = "Sales Chart";
  chart.title.text = "New title";

  const targetState = createChartState("Sales");

  // Restore the pre-rename checkpoint: chart goes back to "Sales".
  const applied = await withFakeExcel(context, () => applyChartState("Sheet1!Sales Chart", targetState));

  assert.equal(chart.name, "Sales");
  assert.ok(applied.state);
  assert.equal(applied.state.kind, "chart_present");
  assert.equal(applied.state.name, "Sales Chart");
  // Inverse must be addressed at the post-restore identity, not the stale one.
  assert.equal(applied.address, "Sheet1!Sales");

  // Restoring the rollback backup at its recorded address must find the chart
  // and rename it forward again.
  const rolledBack = await withFakeExcel(context, () => applyChartState(applied.address, applied.state as RecoveryChartState));

  assert.equal(chart.name, "Sales Chart");
  assert.equal(rolledBack.address, "Sheet1!Sales Chart");
});

void test("chart_absent restore deletes by stable id, surviving rename and name reuse", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const context = new FakeContext([sheet]);

  const created = sheet.charts.add("ColumnClustered", new FakeRange("A1:B12"));
  created.name = "New Chart";
  const checkpoint: RecoveryChartAbsentState = {
    kind: "chart_absent",
    sheetName: "Sheet1",
    name: created.name,
    chartId: created.id,
  };

  // The created chart is renamed, and an unrelated chart takes its old name.
  created.name = "Quarterly";
  const impostor = sheet.charts.add("Pie", new FakeRange("D1:E4"));
  impostor.name = "New Chart";

  const applied = await withFakeExcel(context, () => applyChartState("Sheet1!New Chart", checkpoint));

  assert.equal(applied.state, null);
  assert.equal(created.deleted, true, "the originally created chart should be deleted");
  assert.equal(impostor.deleted, false, "the unrelated chart reusing the name must survive");
  assert.deepEqual(sheet.charts.items.map((item) => item.name), ["New Chart"]);
});

void test("chart_absent restore fails safely when a stored id cannot be verified", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const context = new FakeContext([sheet]);

  const survivor = sheet.charts.add("ColumnClustered", new FakeRange("A1:B12"));
  survivor.name = "New Chart";

  const checkpoint: RecoveryChartAbsentState = {
    kind: "chart_absent",
    sheetName: "Sheet1",
    name: "New Chart",
    chartId: survivor.id,
  };

  // Simulate a host that cannot read chart ids during restore.
  const brokenLoad = (path?: string): void => {
    if (typeof path === "string" && path.includes("items/id")) {
      throw new Error("Chart.id is not supported on this host.");
    }
  };
  Object.defineProperty(sheet.charts, "load", { value: brokenLoad });

  await assert.rejects(
    () => withFakeExcel(context, () => applyChartState("Sheet1!New Chart", checkpoint)),
    /Could not verify the created chart's identity/u,
  );

  assert.equal(survivor.deleted, false, "never degrade to name-based deletion when the id cannot be verified");
  assert.equal(sheet.charts.items.length, 1);
});

void test("chart_absent restore with a missing id deletes nothing", async () => {
  const sheet = new FakeWorksheet("Sheet1");
  const context = new FakeContext([sheet]);

  const survivor = sheet.charts.add("ColumnClustered", new FakeRange("A1:B12"));
  survivor.name = "New Chart";

  const checkpoint: RecoveryChartAbsentState = {
    kind: "chart_absent",
    sheetName: "Sheet1",
    name: "New Chart",
    chartId: "fake-chart-id-does-not-exist",
  };

  const applied = await withFakeExcel(context, () => applyChartState("Sheet1!New Chart", checkpoint));

  assert.equal(applied.state, null);
  assert.equal(survivor.deleted, false, "a different chart holding the name must not be deleted");
  assert.equal(sheet.charts.items.length, 1);
});

void test("persisted chart_state snapshots omit range grids so older codecs drop them", () => {
  const chartSnapshot: WorkbookRecoverySnapshot = {
    id: "snap-chart-persist",
    at: 1700000000300,
    toolName: "restore_snapshot",
    toolCallId: "restore:snap-1",
    address: "Sheet1!Sales",
    changedCount: 1,
    cellCount: 1,
    beforeValues: [],
    beforeFormulas: [],
    snapshotKind: "chart_state",
    chartState: createChartState("Sales"),
    workbookId: "url_sha256:charts-workbook",
    workbookLabel: "Charts.xlsx",
  };

  const rangeSnapshot: WorkbookRecoverySnapshot = {
    id: "snap-range-persist",
    at: 1700000000301,
    toolName: "write_cells",
    toolCallId: "tc-write",
    address: "Sheet1!A1",
    changedCount: 1,
    cellCount: 1,
    beforeValues: [[1]],
    beforeFormulas: [["=A1"]],
    snapshotKind: "range_values",
    workbookId: "url_sha256:charts-workbook",
    workbookLabel: "Charts.xlsx",
  };

  const payload = createPersistedWorkbookRecoveryPayload([chartSnapshot, rangeSnapshot]);

  const persistedChart = payload.snapshots.find((item) => item.id === "snap-chart-persist");
  assert.ok(persistedChart);
  // Older codecs default unknown snapshot kinds to range_values and then
  // require grids; omitting them makes downgraded readers drop the entry
  // instead of misreading it as an empty range backup.
  assert.equal("beforeValues" in persistedChart, false);
  assert.equal("beforeFormulas" in persistedChart, false);

  const persistedRange = payload.snapshots.find((item) => item.id === "snap-range-persist");
  assert.ok(persistedRange);
  assert.deepEqual(persistedRange.beforeValues, [[1]]);
  assert.deepEqual(persistedRange.beforeFormulas, [["=A1"]]);

  // The current codec must still round-trip the stripped chart snapshot.
  const reparsed = parsePersistedSnapshots(payload, { maxEntries: 10 });
  const chartAgain = findSnapshotById(reparsed, "snap-chart-persist");
  assert.ok(chartAgain);
  assert.equal(chartAgain.snapshotKind, "chart_state");
  assert.deepEqual(withoutUndefined(chartAgain.chartState), withoutUndefined(createChartState("Sales")));
});
