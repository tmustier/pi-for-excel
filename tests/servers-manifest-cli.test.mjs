// Static config-parser contract: generated manifest values and package-script wiring are build artifacts.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
  DEV_BASE_URL,
  DEFAULT_DEV_PROXY_HOST,
  renderDevManifest,
  resolveDevOrigin,
} from "../scripts/generate-dev-manifest.mjs";

test("dev origin resolver applies precedence and validates every supported input form", () => {
  const cases = [
    { name: "default", input: {}, expected: { origin: `https://${DEFAULT_DEV_PROXY_HOST}`, source: "default" } },
    { name: "argument wins", input: { arg: "my-addin.localhost", env: { DEV_HOST: "other.localhost", PORTLESS_URL: "https://third.localhost" } }, expected: { origin: "https://my-addin.localhost", source: "argument" } },
    { name: "DEV_HOST wins", input: { env: { DEV_HOST: "pi-excel.localhost", PORTLESS_URL: "https://other.localhost" } }, expected: { origin: "https://pi-excel.localhost", source: "DEV_HOST" } },
    { name: "PORTLESS_URL fallback", input: { env: { PORTLESS_URL: "https://pi-excel.localhost" } }, expected: { origin: "https://pi-excel.localhost", source: "PORTLESS_URL" } },
    { name: "non-443 port", input: { env: { PORTLESS_URL: "https://pi-excel.localhost:1355" } }, expected: { origin: "https://pi-excel.localhost:1355", source: "PORTLESS_URL" } },
    { name: "full HTTPS URL", input: { env: { DEV_HOST: "https://pi-excel.localhost" } }, expected: { origin: "https://pi-excel.localhost", source: "DEV_HOST" } },
    { name: "blank fallback", input: { arg: "  ", env: { DEV_HOST: "", PORTLESS_URL: "  " } }, expected: { origin: `https://${DEFAULT_DEV_PROXY_HOST}`, source: "default" } },
    { name: "retired port accepted", input: { arg: "https://localhost:3000" }, expected: { origin: "https://localhost:3000", source: "argument" } },
    { name: "HTTP rejected", input: { arg: "http://pi-excel.localhost" }, error: /https/ },
    { name: "path rejected", input: { arg: "https://pi-excel.localhost/taskpane" }, error: /bare https origin/ },
    { name: "credentials rejected", input: { arg: "https://user:pass@pi-excel.localhost" }, error: /bare https origin/ },
    { name: "unparseable rejected", input: { arg: "not a host" }, error: /Invalid dev proxy host/ },
    { name: "default URL rejected", input: { arg: "localhost:3141" }, error: /already the default/ },
  ];

  for (const entry of cases) {
    if (entry.error) assert.throws(() => resolveDevOrigin(entry.input), entry.error, entry.name);
    else assert.deepEqual(resolveDevOrigin(entry.input), entry.expected, entry.name);
  }
});

test("manifest renderer handles every template validity row", () => {
  const cases = [
    {
      name: "replaces every occurrence",
      template: `<a>${DEV_BASE_URL}/src/taskpane.html</a><b>${DEV_BASE_URL}/assets/icon-32.png</b>`,
      expected: "<a>https://pi-excel.localhost/src/taskpane.html</a><b>https://pi-excel.localhost/assets/icon-32.png</b>",
    },
    { name: "rejects a template without the base URL", template: "<xml></xml>", error: /expected dev base URL/ },
  ];

  for (const entry of cases) {
    if (entry.error) assert.throws(() => renderDevManifest(entry.template, "https://pi-excel.localhost"), entry.error, entry.name);
    else assert.equal(renderDevManifest(entry.template, "https://pi-excel.localhost"), entry.expected, entry.name);
  }
});

test("real manifest.xml renders with no dev base URLs left over", async () => {
  const xml = await readFile(new URL("../manifest.xml", import.meta.url), "utf8");
  const originalCount = xml.split(DEV_BASE_URL).length - 1;
  assert.ok(originalCount > 0, "manifest.xml should reference the dev base URL");
  const rendered = renderDevManifest(xml, "https://pi-excel.localhost");
  assert.equal(rendered.includes(DEV_BASE_URL), false);
  assert.equal(rendered.split("https://pi-excel.localhost").length - 1, originalCount);
  assert.match(rendered, /<Id>[0-9a-f-]+<\/Id>/i);
});
