import assert from "node:assert/strict";
import { test } from "node:test";

import { getWorkbookContext } from "../src/workbook/context.ts";

function setOfficeDocumentUrl(url: string | null): () => void {
  const hadOffice = Reflect.has(globalThis, "Office");
  const previousOffice = Reflect.get(globalThis, "Office");

  if (url === null) {
    Reflect.deleteProperty(globalThis, "Office");
  } else {
    Reflect.set(globalThis, "Office", {
      context: {
        document: {
          url,
        },
      },
    });
  }

  return () => {
    if (hadOffice) {
      Reflect.set(globalThis, "Office", previousOffice);
      return;
    }

    Reflect.deleteProperty(globalThis, "Office");
  };
}

void test("extracts workbook basename from Windows local-path URLs", async () => {
  const restore = setOfficeDocumentUrl("C:\\Users\\alice\\Budget.xlsx");

  try {
    const context = await getWorkbookContext();
    assert.equal(context.workbookName, "Budget.xlsx");
  } finally {
    restore();
  }
});

void test("extracts workbook basename from document URL path", async () => {
  const restore = setOfficeDocumentUrl("https://example.com/reports/Forecast%20Q1.xlsx?x=1");

  try {
    const context = await getWorkbookContext();
    assert.equal(context.workbookName, "Forecast Q1.xlsx");
  } finally {
    restore();
  }
});

void test("normalizes workbook identity by ignoring query/hash URL parts", async () => {
  const restoreFirst = setOfficeDocumentUrl("https://example.com/reports/Forecast.xlsx?session=abc#sheet=1");
  let firstId: string | null = null;

  try {
    const first = await getWorkbookContext();
    firstId = first.workbookId;
    assert.equal(first.source, "document.url");
  } finally {
    restoreFirst();
  }

  const restoreSecond = setOfficeDocumentUrl("https://example.com/reports/Forecast.xlsx?session=xyz#sheet=2");
  try {
    const second = await getWorkbookContext();
    assert.equal(second.source, "document.url");
    assert.equal(firstId, second.workbookId);
  } finally {
    restoreSecond();
  }
});

void test("the workbook context carries a hash of the document URL, never the URL or its path", async () => {
  const restore = setOfficeDocumentUrl("https://tenant.sharepoint.com/sites/finance/Shared%20Documents/Forecast.xlsx?web=1");

  try {
    const context = await getWorkbookContext();
    assert.match(context.workbookId ?? "", /^url_sha256:[0-9a-f]+$/u);
    assert.deepEqual(Object.keys(context).sort(), ["source", "workbookId", "workbookName"]);
    assert.doesNotMatch(JSON.stringify(context), /sharepoint|finance|Shared/u);
  } finally {
    restore();
  }
});

void test("produces different workbook identities for different workbook paths", async () => {
  const restoreFirst = setOfficeDocumentUrl("https://example.com/reports/Forecast-A.xlsx?x=1");
  let firstId: string | null = null;

  try {
    const first = await getWorkbookContext();
    firstId = first.workbookId;
  } finally {
    restoreFirst();
  }

  const restoreSecond = setOfficeDocumentUrl("https://example.com/reports/Forecast-B.xlsx?x=1");
  try {
    const second = await getWorkbookContext();
    assert.notEqual(firstId, second.workbookId);
  } finally {
    restoreSecond();
  }
});

void test("returns null workbook name when Office document URL is unavailable", async () => {
  const restore = setOfficeDocumentUrl(null);

  try {
    const context = await getWorkbookContext();
    assert.equal(context.workbookName, null);
    assert.equal(context.source, "unknown");
  } finally {
    restore();
  }
});
