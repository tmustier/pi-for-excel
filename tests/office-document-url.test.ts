import assert from "node:assert/strict";
import { test } from "node:test";

import { createOfficeDocumentUrlReader } from "../src/host/office-document-url.ts";

const PATH = "/Users/someone/Documents/Book.xlsx";

void test("the live file properties URL wins over the static document URL", async () => {
  const reader = createOfficeDocumentUrlReader({
    readStaticUrl: () => "",
    readLiveUrl: () => Promise.resolve(PATH),
  });

  assert.equal(await reader.read(), PATH);
});

void test("a never-saved workbook has no URL from either source", async () => {
  const reader = createOfficeDocumentUrlReader({
    readStaticUrl: () => "",
    readLiveUrl: () => Promise.resolve(""),
  });

  assert.equal(await reader.read(), null);
});

void test("a Save As mid-session is picked up on the next read without a reload", async () => {
  let livePath = "";
  const reader = createOfficeDocumentUrlReader({
    readStaticUrl: () => "",
    readLiveUrl: () => Promise.resolve(livePath),
  });

  assert.equal(await reader.read(), null);
  livePath = PATH;
  assert.equal(await reader.read(), PATH);
});

void test("the static URL is the fallback when the live query fails and nothing was known", async () => {
  const reader = createOfficeDocumentUrlReader({
    readStaticUrl: () => PATH,
    readLiveUrl: () => Promise.reject(new Error("host error")),
  });

  assert.equal(await reader.read(), PATH);
});

void test("a stalled host answers with the last known live URL, so identity does not flap", async () => {
  let stalled = false;
  const reader = createOfficeDocumentUrlReader({
    readStaticUrl: () => "",
    readLiveUrl: () => (stalled ? new Promise<string | null>(() => {}) : Promise.resolve(PATH)),
    timeoutMs: 20,
  });

  assert.equal(await reader.read(), PATH);
  stalled = true;
  assert.equal(await reader.read(), PATH);
});

void test("a stalled host with nothing known falls back to the static URL within the timeout", async () => {
  const reader = createOfficeDocumentUrlReader({
    readStaticUrl: () => null,
    readLiveUrl: () => new Promise<string | null>(() => {}),
    timeoutMs: 20,
  });

  const startedAt = Date.now();
  assert.equal(await reader.read(), null);
  assert.ok(Date.now() - startedAt < 1000, "did not wait on the stalled host");
});
