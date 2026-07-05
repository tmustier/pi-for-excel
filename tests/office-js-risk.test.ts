import assert from "node:assert/strict";
import { test } from "node:test";

import { assessOfficeJsCodeRisk } from "../src/tools/experimental-tool-gates/office-js-risk.ts";

void test("pure Excel API code is not flagged", () => {
  const code = [
    "const sheet = context.workbook.worksheets.getActiveWorksheet();",
    "const range = sheet.getRange(\"A1:B12\");",
    "range.load(\"values\");",
    "await context.sync();",
    "const chart = sheet.charts.add(Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.auto);",
    "chart.title.text = \"Sales\";",
    "return { rows: range.values.length };",
  ].join("\n");

  const assessment = assessOfficeJsCodeRisk(code);
  assert.equal(assessment.flagged, false);
  assert.deepEqual(assessment.identifiers, []);
});

void test("member access on Excel objects is not flagged (chart.top, comment.parent)", () => {
  const code = [
    "const chart = context.workbook.worksheets.getActiveWorksheet().charts.getItem(\"Sales\");",
    "chart.top = 50;",
    "chart.left = 10;",
    "const reply = comment.parent;",
    "await context.sync();",
    "return chart.top;",
  ].join("\n");

  const assessment = assessOfficeJsCodeRisk(code);
  assert.equal(assessment.flagged, false);
});

void test("network egress identifiers are flagged", () => {
  const assessment = assessOfficeJsCodeRisk(
    "const response = await fetch(\"https://example.com\"); return response.status;",
  );

  assert.equal(assessment.flagged, true);
  assert.deepEqual(assessment.identifiers, ["fetch"]);
});

void test("storage access is flagged", () => {
  const assessment = assessOfficeJsCodeRisk(
    "return localStorage.getItem(\"connections.store.v1\");",
  );

  assert.equal(assessment.flagged, true);
  assert.ok(assessment.identifiers.includes("localStorage"));
});

void test("global handles are flagged", () => {
  for (const snippet of [
    "return window.location.href;",
    "return globalThis[\"fe\" + \"tch\"];",
    "return document.cookie;",
    "navigator.sendBeacon(url, data);",
  ]) {
    const assessment = assessOfficeJsCodeRisk(snippet);
    assert.equal(assessment.flagged, true, `expected flagged: ${snippet}`);
  }
});

void test("dynamic evaluation and imports are flagged", () => {
  for (const snippet of [
    "eval(payload);",
    "new Function(\"return 1\")();",
    "const mod = await import(url);",
  ]) {
    const assessment = assessOfficeJsCodeRisk(snippet);
    assert.equal(assessment.flagged, true, `expected flagged: ${snippet}`);
  }
});

void test("constructor realm escape is flagged even as member access", () => {
  const assessment = assessOfficeJsCodeRisk(
    "return ({}).constructor.constructor(\"return 1\")();",
  );

  assert.equal(assessment.flagged, true);
  assert.deepEqual(assessment.identifiers, ["constructor"]);
});

void test("Office host surface is flagged, Excel namespace is not", () => {
  const flagged = assessOfficeJsCodeRisk("return Office.context.document.url;");
  assert.equal(flagged.flagged, true);
  assert.deepEqual(flagged.identifiers, ["Office"]);

  const clean = assessOfficeJsCodeRisk("return Excel.ChartType.columnClustered;");
  assert.equal(clean.flagged, false);
});

void test("identifiers inside strings and comments are flagged (documented over-approximation)", () => {
  const assessment = assessOfficeJsCodeRisk(
    "// stage payload for eval via fetch\nconst hint = \"fetch\";\nreturn hint;",
  );

  assert.equal(assessment.flagged, true);
  assert.ok(assessment.identifiers.includes("fetch"));
  assert.ok(assessment.identifiers.includes("eval"));
});

void test("longer identifiers containing risk names are not flagged", () => {
  const assessment = assessOfficeJsCodeRisk(
    "const refetchCount = 1; const windowless = 2; const selfie = 3; return refetchCount;",
  );

  assert.equal(assessment.flagged, false);
});

void test("identifiers are deduplicated", () => {
  const assessment = assessOfficeJsCodeRisk(
    "await fetch(a); await fetch(b); await fetch(c);",
  );

  assert.deepEqual(assessment.identifiers, ["fetch"]);
});

void test("unicode-escaped identifiers are flagged (reviewer P1 bypass)", () => {
  for (const snippet of [
    "await \\u0066etch(\"https://example.com\");",
    "await \\u{66}etch(url);",
    "return fe\\u0074ch(url);",
    "return \\u0077indow.location.href;",
    "return globalTh\\u0069s;",
  ]) {
    const assessment = assessOfficeJsCodeRisk(snippet);
    assert.equal(assessment.flagged, true, `expected flagged: ${snippet}`);
  }
});

void test("nested escape encodings are decoded to a fixpoint", () => {
  // \u005c decodes to a backslash, yielding \u0066etch, which decodes to fetch.
  const assessment = assessOfficeJsCodeRisk("\\u005cu0066etch(url);");
  assert.equal(assessment.flagged, true);
  assert.ok(assessment.identifiers.includes("fetch"));
});

void test("string-literal unicode escapes for plain text are not flagged", () => {
  const assessment = assessOfficeJsCodeRisk(
    "range.values = [[\"caf\\u00e9\", \"\\u00fcber\"]]; await context.sync();",
  );

  assert.equal(assessment.flagged, false);
});

void test("empty code is not flagged", () => {
  const assessment = assessOfficeJsCodeRisk("");
  assert.equal(assessment.flagged, false);
  assert.deepEqual(assessment.identifiers, []);
});
