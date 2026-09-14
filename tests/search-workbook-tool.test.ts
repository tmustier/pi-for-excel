import assert from "node:assert/strict";
import { test } from "node:test";

import { createSearchWorkbookTool } from "../src/tools/search-workbook.ts";

interface SearchSheetFixture {
  name: string;
  visibility?: "Visible" | "Hidden";
  address: string;
  values: unknown[][];
  formulas?: unknown[][];
}

function firstText(result: { content: Array<{ type: string; text?: string }> }): string {
  const block = result.content[0];
  if (!block || block.type !== "text" || typeof block.text !== "string") {
    throw new Error("Expected a text tool result.");
  }
  return block.text;
}

async function withSearchWorkbook<T>(
  fixtures: SearchSheetFixture[],
  action: () => Promise<T>,
): Promise<T> {
  const sheets = fixtures.map((fixture) => ({
    name: fixture.name,
    visibility: fixture.visibility ?? "Visible",
    getUsedRangeOrNullObject: () => ({
      isNullObject: false,
      address: fixture.address,
      values: fixture.values,
      formulas: fixture.formulas ?? fixture.values.map((row) => row.map(() => "")),
      load: (_properties: string) => undefined,
    }),
  }));

  const context = {
    workbook: {
      worksheets: {
        items: sheets,
        load: (_properties: string) => undefined,
      },
    },
    sync: () => Promise.resolve(),
  };

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

void test("search finds values on a worksheet whose quoted name contains an exclamation mark", async () => {
  const result = await withSearchWorkbook([
    {
      name: "Q1!Ops",
      address: "'Q1!Ops'!B2:C3",
      values: [["Owner", "Status"], ["Alice", "Ready"]],
    },
  ], () => createSearchWorkbookTool().execute("search-quoted-sheet", {
    query: "ready",
    sheet: "Q1!Ops",
  }));

  assert.match(firstText(result), /\*\*'Q1!Ops'!C3\*\*: Ready/u);
});

void test("search reports a limit only when another matching cell exists", async () => {
  const exactPage = await withSearchWorkbook([
    {
      name: "Data",
      address: "Data!A1:A2",
      values: [["match one"], ["match two"]],
    },
  ], () => createSearchWorkbookTool().execute("search-exact-page", {
    query: "match",
    max_results: 2,
  }));

  assert.doesNotMatch(firstText(exactPage), /limit reached/u);

  const truncatedPage = await withSearchWorkbook([
    {
      name: "Data",
      address: "Data!A1:A3",
      values: [["match one"], ["match two"], ["match three"]],
    },
  ], () => createSearchWorkbookTool().execute("search-truncated-page", {
    query: "match",
    max_results: 2,
  }));

  assert.match(firstText(truncatedPage), /2 match\(es\).*limit reached/u);
  assert.doesNotMatch(firstText(truncatedPage), /match three/u);
});

void test("formula search matches formulas rather than string constants", async () => {
  const result = await withSearchWorkbook([
    {
      name: "Data",
      address: "Data!A1:B1",
      values: [["Ready", 30]],
      formulas: [["Ready", "=SUM(B2:B3)"]],
    },
  ], () => createSearchWorkbookTool().execute("search-formulas", {
    query: "Ready|SUM",
    search_formulas: true,
    use_regex: true,
  }));

  const text = firstText(result);
  assert.match(text, /\*\*Data!B1\*\*: 30 ← =SUM\(B2:B3\)/u);
  assert.doesNotMatch(text, /Data!A1/u);
});

void test("search skips hidden worksheets unless one is explicitly requested", async () => {
  const fixtures: SearchSheetFixture[] = [
    { name: "Visible", address: "Visible!A1", values: [["ordinary"]] },
    { name: "Hidden", visibility: "Hidden", address: "Hidden!A1", values: [["needle"]] },
  ];

  const defaultResult = await withSearchWorkbook(fixtures, () => createSearchWorkbookTool().execute(
    "search-visible-sheets",
    { query: "needle" },
  ));
  assert.match(firstText(defaultResult), /^No matches/u);

  const targetedResult = await withSearchWorkbook(fixtures, () => createSearchWorkbookTool().execute(
    "search-hidden-sheet",
    { query: "needle", sheet: "Hidden" },
  ));
  assert.match(firstText(targetedResult), /\*\*Hidden!A1\*\*: needle/u);
});

void test("search context shows surrounding rows and escapes table separators", async () => {
  const result = await withSearchWorkbook([
    {
      name: "Data",
      address: "Data!B4:C6",
      values: [["before", 1], ["target\\|value", 2], ["after", 3]],
    },
  ], () => createSearchWorkbookTool().execute("search-with-context", {
    query: "target",
    context_rows: 1,
  }));

  const text = firstText(result);
  assert.match(text, /\| 4 \| before \| 1 \|/u);
  assert.ok(text.includes(`| 5 | target${"\\".repeat(3)}|value | 2 | ◀`));
  assert.match(text, /\| 6 \| after \| 3 \|/u);
});

void test("search returns an early match without traversing every loaded cell during validation", async () => {
  let cellReads = 0;
  const row = new Array<string>(1_000);
  for (let column = 0; column < row.length; column += 1) {
    Object.defineProperty(row, column, {
      configurable: true,
      enumerable: true,
      get: () => {
        cellReads += 1;
        return column < 2 ? "needle" : "";
      },
    });
  }

  const values = Array.from({ length: 1_000 }, () => row);
  const formulaRow = Array.from({ length: 1_000 }, () => "");
  const formulas = Array.from({ length: 1_000 }, () => formulaRow);

  await withSearchWorkbook([
    {
      name: "Large",
      address: "Large!A1:ALL1000",
      values,
      formulas,
    },
  ], async () => {
    const result = await createSearchWorkbookTool().execute("search-large-grid", {
      query: "needle",
      max_results: 1,
    });

    assert.match(firstText(result), /1 match\(es\).*limit reached/u);
    assert.ok(cellReads < 10, `Search read ${cellReads} cells before returning the first result.`);
  });
});
