// Record the already-open Excel workbook while a command runs. See docs/orb-excel-web.md.
import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { chromium } from "playwright";

const [output, separator, command, ...args] = process.argv.slice(2);
if (!output?.endsWith(".mp4") || separator !== "--" || !command) {
  console.error("Usage: node scripts/orb/record-excel-web.mjs <out.mp4> -- <command ...>");
  process.exit(1);
}
if (spawnSync("ffmpeg", ["-version"], { stdio: "ignore" }).status !== 0) {
  throw new Error("System ffmpeg is required (install ffmpeg before recording).");
}
const destination = path.resolve(output);
if (fs.existsSync(destination)) throw new Error(`Output already exists: ${destination}`);
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "excel-recording-"));
let browser;
let child;
let interrupted;
const forwardSignal = (signal) => {
  interrupted = signal;
  child?.kill(signal);
};
process.on("SIGINT", forwardSignal);
process.on("SIGTERM", forwardSignal);

try {
  browser = await chromium.connectOverCDP(`http://127.0.0.1:${process.env.PI_EXCEL_WEB_CDP_PORT ?? "9223"}`);
  const pages = browser.contexts().flatMap((context) => context.pages());
  const workbooks = pages.filter((page) => new URL(page.url()).pathname.endsWith("/_layouts/15/Doc.aspx"));
  if (workbooks.length !== 1) throw new Error("Open exactly one Excel workbook with excel-web-session.mjs before recording.");
  const page = workbooks[0];
  const session = await page.context().newCDPSession(page);
  await page.bringToFront();
  // agent-default (not a user rule): 2026-10-27 — 1600×1000, JPEG 70;
  // silent H.264 at 15 fps, CRF 28, 2 Mb/s ceiling (about 15 MB/minute).
  const frames = [];
  let recordingError;
  session.on("Page.screencastFrame", ({ data, metadata, sessionId }) => {
    try {
      const file = `${frames.length}.jpg`;
      fs.writeFileSync(path.join(temporary, file), Buffer.from(data, "base64"));
      frames.push({ file, timestamp: metadata.timestamp });
    } catch (error) {
      recordingError = error;
    }
    session.send("Page.screencastFrameAck", { sessionId }).catch((error) => { recordingError = error; });
  });
  await session.send("Page.startScreencast", { format: "jpeg", quality: 70, maxWidth: 1600, maxHeight: 1000 });
  // A screenshot also covers commands that finish before the first screencast event.
  const initial = await session.send("Page.captureScreenshot", { format: "jpeg", quality: 70 });
  fs.writeFileSync(path.join(temporary, "initial.jpg"), Buffer.from(initial.data, "base64"));
  const started = Date.now() / 1000;
  const result = await new Promise((resolve, reject) => {
    child = spawn(command, args, { stdio: "inherit" });
    child.once("error", reject);
    child.once("exit", (code, signal) => resolve({ code, signal }));
    if (interrupted) child.kill(interrupted);
  });
  const ended = Date.now() / 1000;
  await session.send("Page.stopScreencast");
  if (recordingError) throw recordingError;
  const interval = [
    { file: "initial.jpg", timestamp: started },
    ...frames.filter((frame) => frame.timestamp > started && frame.timestamp < ended),
  ];
  const timeline = interval.map((frame, index) => {
    const next = interval[index + 1]?.timestamp ?? ended;
    return `file '${frame.file}'\nduration ${(next - frame.timestamp).toFixed(6)}`;
  });
  timeline.push(`file '${interval.at(-1).file}'`);
  fs.writeFileSync(path.join(temporary, "frames.txt"), timeline.join("\n"));
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  const encoded = spawnSync("ffmpeg", [
    "-hide_banner", "-loglevel", "error", "-n", "-f", "concat", "-safe", "0",
    "-i", path.join(temporary, "frames.txt"),
    "-vf", "scale=w='min(1600,iw)':h='min(1000,ih)':force_original_aspect_ratio=decrease:force_divisible_by=2,fps=15",
    "-c:v", "libx264", "-preset", "fast", "-crf", "28", "-maxrate", "2M", "-bufsize", "4M",
    "-pix_fmt", "yuv420p", "-movflags", "+faststart", "-an", destination,
  ], { stdio: "inherit" });
  if (encoded.error) throw encoded.error;
  if (encoded.status !== 0) throw new Error(`ffmpeg exited ${encoded.status}`);
  console.log(`Recorded ${frames.length} source frames; command took ${(ended - started).toFixed(2)}s: ${destination}`);
  process.exitCode = result.signal ? 128 + os.constants.signals[result.signal] : result.code;
} finally {
  process.removeListener("SIGINT", forwardSignal);
  process.removeListener("SIGTERM", forwardSignal);
  // close() disconnects a connectOverCDP client; the existing Chrome stays open.
  await browser?.close();
  fs.rmSync(temporary, { recursive: true, force: true });
}
