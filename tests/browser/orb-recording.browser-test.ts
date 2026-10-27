import assert from "node:assert/strict";
import { spawnSync, spawn } from "node:child_process";
import { once } from "node:events";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";
import { chromium } from "playwright";

await test("recorder preserves idle duration, child failure and the existing browser", { timeout: 30_000 }, async () => {
  const scratch = await fs.mkdtemp(path.join(os.tmpdir(), "orb-recording-test-"));
  const profile = path.join(scratch, "profile");
  const context = await chromium.launchPersistentContext(profile, {
    headless: true,
    args: ["--remote-debugging-port=0"],
  });
  try {
    const [port] = (await fs.readFile(path.join(profile, "DevToolsActivePort"), "utf8")).split("\n");
    const page = await context.newPage();
    await page.route("http://127.0.0.1/**", (route) => route.fulfill({
      contentType: "text/html",
      body: "<title>Idle recording</title><h1>Idle recording</h1>",
    }));
    await page.goto("http://127.0.0.1/_layouts/15/Doc.aspx");
    const output = path.join(scratch, "idle.mp4");
    const child = spawn(process.execPath, [
      "scripts/orb/record-excel-web.mjs", output, "--",
      process.execPath, "-e", "setTimeout(() => process.exit(7), 2200)",
    ], { env: { ...process.env, PI_EXCEL_WEB_CDP_PORT: port, TMPDIR: scratch } });
    let diagnostics = "";
    child.stderr.setEncoding("utf8").on("data", (data: string) => { diagnostics += data; });
    child.stdout.resume();
    await once(child, "exit");
    assert.equal(child.signalCode, null, diagnostics);
    assert.equal(child.exitCode, 7, diagnostics);
    const video = spawnSync("ffprobe", [
      "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", output,
    ], { encoding: "utf8" });
    assert.equal(video.status, 0, video.stderr);
    assert.ok(Math.abs(Number(video.stdout.trim()) - 2.2) < 0.2, `Video duration: ${video.stdout}`);
    assert.equal(await page.title(), "Idle recording");
    assert.deepEqual((await fs.readdir(scratch)).filter((file) => file.startsWith("excel-recording-")), []);
  } finally {
    await context.close();
    await fs.rm(scratch, { recursive: true, force: true });
  }
});
