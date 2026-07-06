/** Chart-state capture/apply for workbook recovery snapshots. */

import { excelRun } from "../../excel/helpers.js";
import { cloneRecoveryChartState } from "./clone.js";
import type {
  RecoveryChartApplyResult,
  RecoveryChartPositionState,
  RecoveryChartPresentState,
  RecoveryChartState,
  RecoveryChartTitleState,
} from "./types.js";

interface ChartTarget {
  sheet: Excel.Worksheet;
  chart: Excel.Chart;
}

interface QualifiedChartName {
  sheetName?: string;
  chartName: string;
}

function unquoteSheetName(raw: string): string {
  return raw.replace(/^'|'$/gu, "").replace(/''/gu, "'");
}

function findUnquotedBang(ref: string): number {
  let inQuote = false;

  for (let index = 0; index < ref.length; index += 1) {
    const ch = ref[index];

    if (ch === "'") {
      if (inQuote && ref[index + 1] === "'") {
        index += 1;
        continue;
      }

      inQuote = !inQuote;
      continue;
    }

    if (ch === "!" && !inQuote) {
      return index;
    }
  }

  return -1;
}

function parseQualifiedChartName(ref: string): QualifiedChartName {
  const bang = findUnquotedBang(ref);
  if (bang < 0) {
    return { chartName: ref };
  }

  return {
    sheetName: unquoteSheetName(ref.slice(0, bang)),
    chartName: ref.slice(bang + 1),
  };
}

function chartAddress(sheetName: string, chartName: string): string {
  const escaped = sheetName.replace(/'/gu, "''");
  const quoted = /[\s'!]/u.test(sheetName) ? `'${escaped}'` : sheetName;
  return `${quoted}!${chartName}`;
}

function makeTitleState(title: Excel.ChartTitle | Excel.ChartAxisTitle): RecoveryChartTitleState {
  return {
    text: typeof title.text === "string" ? title.text : "",
    visible: Boolean(title.visible),
  };
}

function makePositionState(chart: Excel.Chart): RecoveryChartPositionState {
  return {
    top: typeof chart.top === "number" ? chart.top : 0,
    left: typeof chart.left === "number" ? chart.left : 0,
    width: typeof chart.width === "number" ? chart.width : 0,
    height: typeof chart.height === "number" ? chart.height : 0,
  };
}

async function findChartsByName(
  context: Excel.RequestContext,
  name: string,
  sheetName?: string,
): Promise<ChartTarget[]> {
  if (sheetName) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name");
    const charts = sheet.charts;
    charts.load("items/name");
    await context.sync();

    return charts.items
      .filter((chart) => chart.name === name)
      .map((chart) => ({ sheet, chart }));
  }

  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  for (const sheet of sheets.items) {
    sheet.charts.load("items/name");
  }
  await context.sync();

  const matches: ChartTarget[] = [];
  for (const sheet of sheets.items) {
    for (const chart of sheet.charts.items) {
      if (chart.name === name) {
        matches.push({ sheet, chart });
      }
    }
  }

  return matches;
}

async function findChartByName(
  context: Excel.RequestContext,
  name: string,
  sheetName?: string,
): Promise<ChartTarget> {
  const matches = await findChartsByName(context, name, sheetName);

  if (matches.length === 0) {
    const location = sheetName ? ` on sheet "${sheetName}"` : "";
    throw new Error(`Chart "${name}" was not found${location}.`);
  }

  if (matches.length > 1) {
    throw new Error(`Multiple charts named "${name}" were found. Provide sheet to disambiguate.`);
  }

  const match = matches[0];
  if (!match) {
    throw new Error(`Chart "${name}" was not found.`);
  }

  return match;
}

/**
 * Locate a chart by its stable Office.js id across all sheets. Rename-proof:
 * the id survives chart renames, so this never confuses the target with an
 * unrelated chart that reused its name. Throws when the host cannot load
 * chart ids; callers must fail safely rather than degrade to name matching
 * (name matching is reserved for legacy states with no captured id).
 */
async function findChartById(
  context: Excel.RequestContext,
  chartId: string,
): Promise<ChartTarget | null> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  for (const sheet of sheets.items) {
    sheet.charts.load("items/id");
  }
  await context.sync();

  for (const sheet of sheets.items) {
    for (const chart of sheet.charts.items) {
      if (chart.id === chartId) {
        return { sheet, chart };
      }
    }
  }

  return null;
}

async function findChartByNameOrNull(
  context: Excel.RequestContext,
  name: string,
  sheetName?: string,
): Promise<ChartTarget | null> {
  const matches = await findChartsByName(context, name, sheetName);
  if (matches.length === 0) return null;
  if (matches.length > 1) {
    throw new Error(`Multiple charts named "${name}" were found. Provide sheet to disambiguate.`);
  }

  return matches[0] ?? null;
}

async function captureChartStateFromTarget(
  context: Excel.RequestContext,
  target: ChartTarget,
): Promise<RecoveryChartPresentState> {
  const { sheet, chart } = target;

  sheet.load("name");
  chart.load("name,chartType,top,left,width,height");
  chart.title.load("text,visible");
  chart.legend.load("position,visible");
  await context.sync();

  let xAxisTitle: RecoveryChartTitleState | undefined;
  let yAxisTitle: RecoveryChartTitleState | undefined;

  try {
    chart.axes.categoryAxis.title.load("text,visible");
    chart.axes.valueAxis.title.load("text,visible");
    await context.sync();
    xAxisTitle = makeTitleState(chart.axes.categoryAxis.title);
    yAxisTitle = makeTitleState(chart.axes.valueAxis.title);
  } catch {
    xAxisTitle = undefined;
    yAxisTitle = undefined;
  }

  const state: RecoveryChartPresentState = {
    kind: "chart_present",
    sheetName: sheet.name,
    name: chart.name,
    chartType: String(chart.chartType),
    title: makeTitleState(chart.title),
    legend: {
      position: typeof chart.legend.position === "string" ? chart.legend.position : "Right",
      visible: Boolean(chart.legend.visible),
    },
    position: makePositionState(chart),
  };

  if (xAxisTitle) {
    state.xAxisTitle = xAxisTitle;
  }
  if (yAxisTitle) {
    state.yAxisTitle = yAxisTitle;
  }

  return state;
}

function applyTitleState(target: Excel.ChartTitle | Excel.ChartAxisTitle, state: RecoveryChartTitleState): void {
  target.text = state.text;
  target.visible = state.visible;
}

function applyPositionState(chart: Excel.Chart, position: RecoveryChartPositionState): void {
  chart.top = position.top;
  chart.left = position.left;
  chart.width = position.width;
  chart.height = position.height;
}

function applyChartPresentState(chart: Excel.Chart, state: RecoveryChartPresentState): void {
  chart.chartType = state.chartType as Excel.ChartType;
  applyTitleState(chart.title, state.title);
  chart.legend.position = state.legend.position as Excel.ChartLegendPosition;
  chart.legend.visible = state.legend.visible;

  if (state.xAxisTitle) {
    applyTitleState(chart.axes.categoryAxis.title, state.xAxisTitle);
  }

  if (state.yAxisTitle) {
    applyTitleState(chart.axes.valueAxis.title, state.yAxisTitle);
  }

  applyPositionState(chart, state.position);
  chart.name = state.name;
}

export async function captureChartPresentState(
  name: string,
  sheetName?: string,
): Promise<RecoveryChartPresentState> {
  return excelRun<RecoveryChartPresentState>(async (context) => {
    const target = await findChartByName(context, name, sheetName);
    const state = await captureChartStateFromTarget(context, target);
    return cloneRecoveryChartState(state) as RecoveryChartPresentState;
  });
}

export async function applyChartState(
  address: string,
  targetState: RecoveryChartState,
): Promise<RecoveryChartApplyResult> {
  return excelRun<RecoveryChartApplyResult>(async (context) => {
    const parsedAddress = parseQualifiedChartName(address);
    const currentSheetName = parsedAddress.sheetName ?? targetState.sheetName;
    const currentChartName = parsedAddress.chartName || targetState.name;

    if (targetState.kind === "chart_absent") {
      if (targetState.chartId) {
        let target: ChartTarget | null;
        try {
          target = await findChartById(context, targetState.chartId);
        } catch {
          // The id was captured at create time, so a host that cannot read
          // ids now is anomalous. Never degrade to mutable-name deletion —
          // that is exactly the wrong-chart risk the id exists to prevent.
          throw new Error(
            `Could not verify the created chart's identity by id; nothing was deleted. ` +
            `Delete chart "${currentChartName}" manually if it still exists.`,
          );
        }

        if (target) {
          target.chart.delete();
          await context.sync();
        }

        return { state: null, address };
      }

      // Legacy states without a captured id: unique-name match only.
      const target = await findChartByNameOrNull(context, currentChartName, currentSheetName);
      if (target) {
        target.chart.delete();
        await context.sync();
      }

      return { state: null, address };
    }

    let target = await findChartByNameOrNull(context, currentChartName, currentSheetName);
    if (!target && currentChartName !== targetState.name) {
      target = await findChartByNameOrNull(context, targetState.name, targetState.sheetName);
    }

    if (!target) {
      throw new Error(`Chart "${currentChartName}" was not found for restore.`);
    }

    const currentState = await captureChartStateFromTarget(context, target);
    applyChartPresentState(target.chart, targetState);
    await context.sync();

    // The restore may have renamed the chart back to targetState.name, so the
    // inverse snapshot must be addressed at the post-restore identity.
    return {
      state: cloneRecoveryChartState(currentState),
      address: chartAddress(currentState.sheetName, targetState.name),
    };
  });
}

export function formatRecoveryChartAddress(sheetName: string, name: string): string {
  return chartAddress(sheetName, name);
}
