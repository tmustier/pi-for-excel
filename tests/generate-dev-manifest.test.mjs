import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
  DEV_BASE_URL,
  DEFAULT_DEV_PROXY_HOST,
  renderDevManifest,
  resolveDevOrigin,
} from "../scripts/generate-dev-manifest.mjs";

// ── resolveDevOrigin ────────────────────────────────────────────────────────

test("defaults to pi-excel.localhost (matches npm run dev:portless)", () => {
  const resolved = resolveDevOrigin({});
  assert.equal(resolved.origin, `https://${DEFAULT_DEV_PROXY_HOST}`);
  assert.equal(resolved.source, "default");
});

test("explicit argument wins over env", () => {
  const resolved = resolveDevOrigin({
    arg: "my-addin.localhost",
    env: { DEV_HOST: "other.localhost", PORTLESS_URL: "https://third.localhost" },
  });

  assert.equal(resolved.origin, "https://my-addin.localhost");
  assert.equal(resolved.source, "argument");
});

test("DEV_HOST wins over PORTLESS_URL", () => {
  const resolved = resolveDevOrigin({
    env: { DEV_HOST: "pi-excel.localhost", PORTLESS_URL: "https://other.localhost" },
  });

  assert.equal(resolved.origin, "https://pi-excel.localhost");
  assert.equal(resolved.source, "DEV_HOST");
});

test("PORTLESS_URL is used when DEV_HOST is unset", () => {
  const resolved = resolveDevOrigin({ env: { PORTLESS_URL: "https://pi-excel.localhost" } });
  assert.equal(resolved.origin, "https://pi-excel.localhost");
  assert.equal(resolved.source, "PORTLESS_URL");
});

test("non-443 proxy ports are preserved", () => {
  const resolved = resolveDevOrigin({ env: { PORTLESS_URL: "https://pi-excel.localhost:1355" } });
  assert.equal(resolved.origin, "https://pi-excel.localhost:1355");
});

test("accepts full https URL in DEV_HOST", () => {
  const resolved = resolveDevOrigin({ env: { DEV_HOST: "https://pi-excel.localhost" } });
  assert.equal(resolved.origin, "https://pi-excel.localhost");
});

test("blank values fall through to the default", () => {
  const resolved = resolveDevOrigin({ arg: "  ", env: { DEV_HOST: "", PORTLESS_URL: "  " } });
  assert.equal(resolved.origin, `https://${DEFAULT_DEV_PROXY_HOST}`);
  assert.equal(resolved.source, "default");
});

test("rejects non-https schemes", () => {
  assert.throws(() => resolveDevOrigin({ arg: "http://pi-excel.localhost" }), /https/);
});

test("rejects origins with a path", () => {
  assert.throws(() => resolveDevOrigin({ arg: "https://pi-excel.localhost/taskpane" }), /bare https origin/);
});

test("rejects origins with credentials", () => {
  assert.throws(() => resolveDevOrigin({ arg: "https://user:pass@pi-excel.localhost" }), /bare https origin/);
});

test("rejects unparseable input", () => {
  assert.throws(() => resolveDevOrigin({ arg: "not a host" }), /Invalid dev proxy host/);
});

test("rejects the default dev URL itself", () => {
  assert.throws(() => resolveDevOrigin({ arg: "localhost:3000" }), /already the default/);
});

// ── renderDevManifest ───────────────────────────────────────────────────────

test("replaces every dev base URL occurrence", () => {
  const template = `<a>${DEV_BASE_URL}/src/taskpane.html</a><b>${DEV_BASE_URL}/assets/icon-32.png</b>`;
  const rendered = renderDevManifest(template, "https://pi-excel.localhost");

  assert.equal(rendered.includes(DEV_BASE_URL), false);
  assert.equal(
    rendered,
    "<a>https://pi-excel.localhost/src/taskpane.html</a><b>https://pi-excel.localhost/assets/icon-32.png</b>",
  );
});

test("throws when the template lacks the dev base URL", () => {
  assert.throws(() => renderDevManifest("<xml></xml>", "https://pi-excel.localhost"), /expected dev base URL/);
});

test("real manifest.xml renders with no dev base URLs left over", async () => {
  const xml = await readFile(new URL("../manifest.xml", import.meta.url), "utf8");
  const originalCount = xml.split(DEV_BASE_URL).length - 1;
  assert.ok(originalCount > 0, "manifest.xml should reference the dev base URL");

  const rendered = renderDevManifest(xml, "https://pi-excel.localhost");

  assert.equal(rendered.includes(DEV_BASE_URL), false);
  assert.equal(rendered.split("https://pi-excel.localhost").length - 1, originalCount);
  // Same add-in Id: the dev-proxy manifest replaces the default sideload.
  assert.match(rendered, /<Id>[0-9a-f-]+<\/Id>/i);
});
