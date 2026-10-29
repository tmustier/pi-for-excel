import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const repo = path.resolve(import.meta.dirname, "../..");
const args = process.argv.slice(2);
const flag = (name) => args.includes(name);
const option = (name, fallback) => {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] ? args[i + 1] : fallback;
};

const folderLink = process.env.PI_EXCEL_WEB_FOLDER_LINK?.trim();
if (!folderLink) fail("PI_EXCEL_WEB_FOLDER_LINK is not set (Amp project secret; run `amp orb restart-processes` after adding it).");
const workbookFile = path.resolve(option("--workbook", path.join(repo, "scripts/orb/blank.xlsx")));
const manifestFile = path.join(repo, "manifest.xml");
const port = process.env.PI_EXCEL_WEB_CDP_PORT ?? "9223";
const cdpBase = `http://127.0.0.1:${port}`;
const orbId = (process.env.AMP_THREAD_ID ?? os.hostname()).replace(/[^A-Za-z0-9-]/g, "").slice(-12);
const stamp = new Date().toISOString().replace(/\D/g, "").slice(0, 14);
const workbookName = `orb-${orbId}-${stamp}.xlsx`;

function fail(message) {
  console.error(`excel-web-session: ${message}`);
  process.exit(1);
}
const log = (message) => console.log(`==> ${message}`);
const ONEDRIVE = "https://onedrive.live.com/";
const EXCEL_FRAME = "https://excel.officeapps.live.com/";
const OFFICE_DIALOG = "https://res.public.onecdn.static.microsoft/";
const TASKPANE = "https://localhost:3141/src/taskpane.html";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function targets() {
  return (await fetch(`${cdpBase}/json`)).json();
}

async function connect(webSocketDebuggerUrl) {
  const ws = new WebSocket(webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { ws.onopen = resolve; ws.onerror = reject; });
  let nextId = 0;
  const pending = new Map();
  ws.onmessage = (event) => {
    const message = JSON.parse(event.data);
    const resolve = pending.get(message.id);
    if (typeof resolve !== "function") return;
    pending.delete(message.id);
    resolve(message);
  };
  return {
    send: (method, params = {}) => new Promise((resolve) => {
      const id = ++nextId;
      pending.set(id, resolve);
      ws.send(JSON.stringify({ id, method, params }));
    }),
    close: () => ws.close(),
  };
}

async function evaluateIn(prefix, expression) {
  for (const target of await targets()) {
    if (!["page", "iframe"].includes(target.type) || !target.url.startsWith(prefix)) continue;
    const session = await connect(target.webSocketDebuggerUrl);
    try {
      const reply = await session.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true, userGesture: true });
      const exception = reply.result?.exceptionDetails;
      if (exception) throw new Error(exception.exception?.description ?? exception.text);
      const value = reply.result?.result?.value;
      if (value !== undefined && value !== null) return value;
    } finally {
      session.close();
    }
  }
  return null;
}

async function waitFor(what, check, timeoutMs = 60_000) {
  const deadline = Date.now() + timeoutMs;
  let lastError = "";
  for (;;) {
    try {
      const value = await check();
      if (value) return value;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
    if (Date.now() > deadline) fail(`timed out waiting for ${what}${lastError ? ` (last error: ${lastError})` : ""}`);
    await sleep(1000);
  }
}

const clickExpression = (label) => `(() => {
  const els = [...document.querySelectorAll('button,[role=button],[role=menuitem],[role=tab],a,input[type=button],input[type=submit]')];
  const el = els.find((e) => e.getAttribute('aria-label') === ${JSON.stringify(label)})
    || els.find((e) => e.textContent.trim() === ${JSON.stringify(label)} || e.value === ${JSON.stringify(label)});
  if (!el || el.disabled) return null;
  el.click();
  return true;
})()`;
const click = (match, label, timeoutMs) => waitFor(`"${label}"`, () => evaluateIn(match, clickExpression(label)), timeoutMs);

async function acceptDeveloperMode() {
  const accepted = await evaluateIn(EXCEL_FRAME, `(() => {
    const box = [...document.querySelectorAll('input[type=checkbox]')].find((b) => /Developer Mode/i.test(b.closest('label, div')?.textContent ?? ''));
    if (!box) return null;
    if (!box.checked) box.click();
    return true;
  })()`);
  if (accepted) {
    await evaluateIn(EXCEL_FRAME, clickExpression("OK"));
    log("Accepted the Developer Mode prompt");
  }
}

const authFile = path.join(os.homedir(), ".pi/agent/auth.json");
if (process.env.PI_EXCEL_WEB_AUTH_JSON?.trim() && !fs.existsSync(authFile)) {
  fs.mkdirSync(path.dirname(authFile), { recursive: true, mode: 0o700 });
  fs.writeFileSync(authFile, process.env.PI_EXCEL_WEB_AUTH_JSON.trim(), { mode: 0o600 });
  log(`Wrote model credentials from PI_EXCEL_WEB_AUTH_JSON to ${authFile}`);
}

log("Ensuring dev server and proxy");
execFileSync("amp", ["orb", "services", "ensure"], { cwd: repo, stdio: "inherit" });

const cdpUp = () => fetch(`${cdpBase}/json/version`).then((r) => r.ok, () => false);
if (!(await cdpUp())) {
  const service = port === "9223" ? "excel-web" : `excel-web-${port}`;
  log(`Starting Chrome (${service}, CDP ${port})`);
  const passThrough = ["PI_EXCEL_WEB_CDP_PORT", "PI_EXCEL_WEB_PROFILE", "PI_EXCEL_WEB_HEADLESS"]
    .filter((name) => process.env[name])
    .map((name) => `${name}='${process.env[name].replace(/'/g, "")}'`);
  const command = [...passThrough, "scripts/orb/excel-web.sh", "about:blank"].join(" ");
  execFileSync("amp", ["orb", "service", "start", service, "--command", command], { cwd: repo, stdio: "inherit" });
  await waitFor("Chrome CDP", cdpUp, 30_000);
}

const reused = (await targets()).some((t) => t.url.startsWith(TASKPANE))
  ? (await targets()).find((t) => t.type === "page" && t.url.includes("/_layouts/15/Doc.aspx"))
  : undefined;
if (reused) log("Pi is already open; reusing its workbook");
const workbookUrl = reused ? reused.url : await openNewSession();

async function openNewSession() {
const pageTarget = (await targets()).find((t) => t.type === "page");
if (!pageTarget) fail("no page target in Chrome");
const page = await connect(pageTarget.webSocketDebuggerUrl);
const navigate = (url) => page.send("Page.navigate", { url });
const pageUrl = async () => (await page.send("Runtime.evaluate", { expression: "location.href", returnByValue: true })).result?.result?.value ?? "";

log("Opening the shared folder");
await navigate(folderLink);
const folderUrl = await waitFor("the OneDrive folder", async () => {
  const href = await pageUrl();
  return href.startsWith(ONEDRIVE) && new URL(href).searchParams.get("id")?.startsWith("/") ? href : null;
});
const folderPath = new URL(folderUrl).searchParams.get("id");
const sitePath = folderPath.split("/").slice(0, 3).join("/");

log(`Uploading ${workbookName}`);
const bytes = fs.readFileSync(workbookFile).toString("base64");
const upload = await waitFor("the guest upload", () => evaluateIn(ONEDRIVE, `(async () => {
  const h = { accept: 'application/json;odata=nometadata' };
  const json = async (step, r) => {
    const text = await r.text();
    if (!r.ok || !text.startsWith('{')) throw new Error(step + ' ' + r.status + ' ' + r.url + ' ' + text.slice(0, 120));
    return JSON.parse(text);
  };
  const digest = (await json('contextinfo', await fetch(${JSON.stringify(sitePath)} + '/_api/contextinfo', { method: 'POST', headers: h }))).FormDigestValue;
  const user = await json('currentuser', await fetch(${JSON.stringify(sitePath)} + '/_api/web/currentuser', { headers: h }));
  const body = Uint8Array.from(atob(${JSON.stringify(bytes)}), (c) => c.charCodeAt(0));
  const folder = ${JSON.stringify(folderPath)}.replace(/'/g, "''");
  const r = await fetch(${JSON.stringify(sitePath)} + "/_api/web/GetFolderByServerRelativeUrl('" + folder + "')/Files/add(url='${workbookName}',overwrite=false)",
    { method: 'POST', headers: { ...h, 'X-RequestDigest': digest }, body });
  const file = await json('upload', r);
  return { id: file.UniqueId, user: user.Title };
})()`), 30_000);
log(`Uploaded as "${upload.user}" (file ${upload.id})`);

const workbookUrl = `https://onedrive.live.com${sitePath}/_layouts/15/Doc.aspx?sourcedoc={${upload.id}}&action=edit`;
log("Opening the workbook in Excel");
await navigate(workbookUrl);
await waitFor("the Office Add-ins dialog", async () => {
  if ((await targets()).some((t) => t.url.startsWith(OFFICE_DIALOG))) return true;
  if (!(await evaluateIn(EXCEL_FRAME, clickExpression("More Add-ins")))) {
    await evaluateIn(EXCEL_FRAME, clickExpression("Add-ins"));
    await sleep(2000);
  } else {
    await sleep(5000);
  }
  return null;
}, 180_000);

log("Uploading manifest.xml");
await click(OFFICE_DIALOG, "MY ADD-INS");
await click(OFFICE_DIALOG, "Manage My Add-ins");
await click(OFFICE_DIALOG, "Upload My Add-in");
await waitFor("the manifest file input", async () => {
  for (const target of await targets()) {
    if (!target.url.startsWith(OFFICE_DIALOG)) continue;
    const session = await connect(target.webSocketDebuggerUrl);
    try {
      const doc = await session.send("DOM.getDocument", { depth: -1, pierce: true });
      const found = await session.send("DOM.querySelectorAll", { nodeId: doc.result.root.nodeId, selector: "input[type=file]" });
      const nodeId = found.result?.nodeIds?.[0];
      if (nodeId) {
        await session.send("DOM.setFileInputFiles", { nodeId, files: [manifestFile] });
        return true;
      }
    } finally {
      session.close();
    }
  }
  return null;
});
await click(OFFICE_DIALOG, "Upload");

log("Opening Pi");
await waitFor("the Open Pi button", async () => {
  await acceptDeveloperMode();
  if (await evaluateIn(EXCEL_FRAME, clickExpression("Open Pi"))) return true;
  await evaluateIn(EXCEL_FRAME, clickExpression("More Options"));
  return evaluateIn(EXCEL_FRAME, clickExpression("Open Pi"));
}, 90_000);
await waitFor("the Pi taskpane", async () => (await targets()).some((t) => t.url.startsWith(TASKPANE)), 60_000);
page.close();
log("Pi taskpane is open");
return workbookUrl;
}

let check = null;
if (flag("--check")) {
  log("Checking Office.js in the taskpane");
  check = await waitFor("Office.js", () => evaluateIn(TASKPANE, `(async () => {
    await Office.onReady();
    const marker = 'orb-check-' + Date.now();
    const readBack = await Excel.run(async (ctx) => {
      const cell = ctx.workbook.worksheets.getActiveWorksheet().getRange('Z1');
      cell.values = [[marker]];
      await ctx.sync();
      cell.load('values');
      await ctx.sync();
      const value = cell.values[0][0];
      cell.clear();
      await ctx.sync();
      return value;
    });
    return { host: Office.context.host, platform: Office.context.platform, wrote: marker, readBack, ok: readBack === marker };
  })()`), 60_000);
  if (!check.ok) fail(`Office.js read-back mismatch: ${JSON.stringify(check)}`);
}

let turn = null;
const prompt = option("--prompt", "");
if (prompt) {
  const isRunning = `[...document.querySelectorAll('button')].some((b) => /^Stop/.test(b.getAttribute('aria-label') ?? b.textContent.trim()))`;
  const ready = await waitFor("a logged-in composer", () => evaluateIn(TASKPANE, `(() => {
    const model = [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith('π'));
    const box = [...document.querySelectorAll('textarea')].find((e) => e.offsetParent !== null);
    return model && box ? model.textContent.trim().replace(/^π\\s*/, '') : null;
  })()`), 60_000);
  const loginOverlay = `(() => {
    const el = ['pi-api-key-overlay', 'pi-welcome-login-overlay'].map((id) => document.getElementById(id))
      .find((e) => e && e.getBoundingClientRect().width > 0);
    return el ? (el.querySelector('h2')?.textContent.trim() || el.id) : null;
  })()`;
  const blocked = await evaluateIn(TASKPANE, loginOverlay);
  if (blocked) fail(`Pi is asking for a login ("${blocked}") before the prompt: model credentials are missing or were rejected.`);
  log(`Sending a prompt to Pi (${ready})`);
  await waitFor("the prompt to send", () => evaluateIn(TASKPANE, `(async () => {
    const box = [...document.querySelectorAll('textarea')].find((e) => e.offsetParent !== null);
    box.focus();
    Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(box, ${JSON.stringify(prompt)});
    box.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
    await new Promise((r) => setTimeout(r, 300));
    const send = [...document.querySelectorAll('button')].find((b) => (b.getAttribute('aria-label') ?? b.textContent.trim()) === 'Send');
    if (!send || send.disabled) return null;
    send.click();
    return true;
  })()`), 30_000);
  const started = Date.now();
  await waitFor("the turn to start", () => evaluateIn(TASKPANE, isRunning), 30_000);
  await waitFor("the turn to finish", async () => {
    const blocker = await evaluateIn(TASKPANE, loginOverlay);
    if (blocker) fail(`Pi is asking for a login ("${blocker}"): its model credentials are missing or were rejected. Check the taskpane console for [auth] lines.`);
    return !(await evaluateIn(TASKPANE, isRunning));
  }, 15 * 60_000);
  turn = await evaluateIn(TASKPANE, `(() => {
    const all = document.body.innerText;
    const text = all.slice(all.lastIndexOf(${JSON.stringify(prompt.slice(0, 80))})).replace(/\\n{2,}/g, '\\n');
    const tools = [...text.matchAll(/\\n([a-z][a-z_]*)\\nINPUT\\n/g)].map((m) => m[1]);
    return { model: ${JSON.stringify(ready)}, tools, replyTail: text.slice(-1200) };
  })()`);
  turn.seconds = Math.round((Date.now() - started) / 1000);
  log(`Turn finished in ${turn.seconds}s with ${turn.tools.length} tool call(s)`);
}

console.log(JSON.stringify({ workbook: reused ? null : workbookName, reused: Boolean(reused), workbookUrl, cdpPort: Number(port), check, turn }, null, 2));
