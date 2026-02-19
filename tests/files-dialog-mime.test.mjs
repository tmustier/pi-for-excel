import assert from "node:assert/strict";
import { test } from "node:test";

import { resolveSafeBlobUrlMimeType } from "../src/files/blob-url-safety.ts";

test("resolveSafeBlobUrlMimeType downgrades active-content types", () => {
  assert.equal(resolveSafeBlobUrlMimeType("text/html"), "application/octet-stream");
  assert.equal(resolveSafeBlobUrlMimeType("image/svg+xml"), "application/octet-stream");
  assert.equal(resolveSafeBlobUrlMimeType("text/javascript; charset=utf-8"), "application/octet-stream");
});

test("resolveSafeBlobUrlMimeType preserves safe types", () => {
  assert.equal(resolveSafeBlobUrlMimeType("text/plain"), "text/plain");
  assert.equal(resolveSafeBlobUrlMimeType("application/pdf"), "application/pdf");
  assert.equal(resolveSafeBlobUrlMimeType(""), "application/octet-stream");
});
