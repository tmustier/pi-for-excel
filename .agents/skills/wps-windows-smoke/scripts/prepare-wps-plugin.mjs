#!/usr/bin/env node
import { spawnSync, spawn } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

function usage() {
  console.log(`Usage: scripts/prepare-wps-plugin.mjs [options]

Build a WPS smoke plugin root from this repo's wps/ skeleton, patched for a
Windows QEMU guest. The generated directory can be served over HTTP and used by
wpsjs publish or an enterprise jsplugins.xml server.

Options:
  --repo <path>          pi-for-excel repo root (default: cwd or nearest parent with wps/main.js)
  --out <path>           output plugin dir (default: ~/VMs/wps-win11/wps-smoke-plugin)
  --name <name>          WPS plugin name (default: PiForExcelSmoke)
  --taskpane-url <url>   taskpane URL seen by guest (default: http://10.0.2.2:3141/src/taskpane.html)
  --plugin-url <url>     plugin base URL seen by guest (default: http://10.0.2.2:3889/)
  --publish              run npx wpsjs publish and copy publish.html into --out
  --serve [port]         start python3 -m http.server from --out (default port: 3889)
  --help                 show this help

Examples:
  npm run dev
  .agents/skills/wps-windows-smoke/scripts/prepare-wps-plugin.mjs --publish --serve
  open http://10.0.2.2:3889/publish.html inside the Windows guest
`);
}

function parseArgs(argv) {
  const opts = {
    repo: null,
    out: path.join(os.homedir(), "VMs/wps-win11/wps-smoke-plugin"),
    name: "PiForExcelSmoke",
    taskpaneUrl: "http://10.0.2.2:3141/src/taskpane.html",
    pluginUrl: "http://10.0.2.2:3889/",
    publish: false,
    serve: false,
    servePort: 3889,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      usage();
      process.exit(0);
    } else if (arg === "--repo") {
      opts.repo = argv[++i];
    } else if (arg === "--out") {
      opts.out = argv[++i];
    } else if (arg === "--name") {
      opts.name = argv[++i];
    } else if (arg === "--taskpane-url") {
      opts.taskpaneUrl = argv[++i];
    } else if (arg === "--plugin-url") {
      opts.pluginUrl = argv[++i];
    } else if (arg === "--publish") {
      opts.publish = true;
    } else if (arg === "--serve") {
      opts.serve = true;
      const maybePort = argv[i + 1];
      if (maybePort && /^\d+$/.test(maybePort)) {
        opts.servePort = Number(maybePort);
        i += 1;
      }
    } else {
      throw new Error(`unknown argument: ${arg}`);
    }
  }
  return opts;
}

function findRepo(start) {
  let dir = path.resolve(start);
  while (true) {
    if (fs.existsSync(path.join(dir, "wps/main.js")) && fs.existsSync(path.join(dir, "wps/ribbon.xml"))) {
      return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  throw new Error("could not find repo root containing wps/main.js and wps/ribbon.xml; pass --repo");
}

function ensureTrailingSlash(url) {
  return url.endsWith("/") ? url : `${url}/`;
}

function patchMainJs(source, taskpaneUrl) {
  let out = source.replace(
    /const DEV_TASKPANE_URL = "[^"]+";/,
    `const DEV_TASKPANE_URL = ${JSON.stringify(taskpaneUrl)};`,
  );

  // WPS templates spell the ribbon callback OnAddinLoad. Keep the repo's
  // OnAddInLoad spelling too so either ribbon form works.
  if (!out.includes("globalScope.OnAddinLoad")) {
    out = out.replace(
      /globalScope\.OnAddInLoad = function OnAddInLoad\(\) \{\n    storageGet\(TASKPANE_ID_STORAGE_KEY\);\n  \};/,
      `globalScope.OnAddInLoad = function OnAddInLoad() {\n    storageGet(TASKPANE_ID_STORAGE_KEY);\n  };\n\n  globalScope.OnAddinLoad = globalScope.OnAddInLoad;`,
    );
  }
  return out;
}

function patchRibbonXml(source) {
  // Official wpsjs ET templates use OnAddinLoad. Also make the copied ribbon
  // start with <customUI because WPS's publish-page validator treats an XML
  // declaration prefix as invalid.
  return source
    .replace(/^<\?xml[^>]*>\s*/u, "")
    .replace(/onLoad="OnAddInLoad"/g, 'onLoad="OnAddinLoad"');
}

function buildIndexHtml() {
  return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <script type="text/javascript" src="./main.js"></script>
  </head>
  <body></body>
</html>
`;
}

function writeSmokeFiles(opts) {
  const repo = opts.repo ? path.resolve(opts.repo) : findRepo(process.cwd());
  const out = path.resolve(opts.out);
  const pluginUrl = ensureTrailingSlash(opts.pluginUrl);
  fs.rmSync(out, { recursive: true, force: true });
  fs.mkdirSync(out, { recursive: true });

  const mainJs = patchMainJs(fs.readFileSync(path.join(repo, "wps/main.js"), "utf8"), opts.taskpaneUrl);
  const ribbonXml = patchRibbonXml(fs.readFileSync(path.join(repo, "wps/ribbon.xml"), "utf8"));
  fs.writeFileSync(path.join(out, "index.html"), buildIndexHtml());
  fs.writeFileSync(path.join(out, "main.js"), mainJs);
  fs.writeFileSync(path.join(out, "ribbon.xml"), ribbonXml);
  fs.writeFileSync(path.join(out, "manifest.xml"), `<?xml version="1.0" encoding="UTF-8"?>
<JsPlugin>
  <ApiVersion>1.0.0</ApiVersion>
  <Name>${opts.name}</Name>
  <Description>pi-for-excel WPS smoke plugin</Description>
</JsPlugin>
`);
  fs.writeFileSync(path.join(out, "jsplugins.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<jsplugins>
  <jspluginonline name="${opts.name}" type="et" url="${pluginUrl}" debug="code" version="1.0.0" enable="enable_dev" install="null" customDomain="" />
</jsplugins>
`);
  fs.writeFileSync(path.join(out, "package.json"), JSON.stringify({
    name: opts.name,
    addonType: "et",
    version: "1.0.0",
    customDomain: "",
  }, null, 2));

  return { repo, out, pluginUrl };
}

function runPublish(opts, generated) {
  const project = path.join(generated.out, ".publish-project");
  fs.rmSync(project, { recursive: true, force: true });
  fs.mkdirSync(path.join(project, "wps-addon-build"), { recursive: true });
  for (const file of ["index.html", "ribbon.xml", "main.js", "manifest.xml"]) {
    fs.copyFileSync(path.join(generated.out, file), path.join(project, "wps-addon-build", file));
  }
  fs.writeFileSync(path.join(project, "package.json"), JSON.stringify({
    name: opts.name,
    addonType: "et",
    version: "1.0.0",
    customDomain: "",
  }, null, 2));

  const result = spawnSync("npx", ["--yes", "wpsjs", "publish", "-s", generated.pluginUrl], {
    cwd: project,
    stdio: "inherit",
  });
  if (result.status !== 0) {
    throw new Error(`wpsjs publish failed with status ${result.status}`);
  }
  fs.copyFileSync(path.join(project, "wps-addon-publish/publish.html"), path.join(generated.out, "publish.html"));
}

function startServer(out, port) {
  const pidFile = path.join(os.tmpdir(), `pi-wps-smoke-plugin-${port}.pid`);
  try {
    const existing = fs.readFileSync(pidFile, "utf8").trim();
    if (existing && process.kill(Number(existing), 0)) {
      console.log(`HTTP server already recorded as PID ${existing}; pid file: ${pidFile}`);
      return;
    }
  } catch {
    // no live recorded server
  }
  const child = spawn("python3", ["-m", "http.server", String(port), "--bind", "0.0.0.0"], {
    cwd: out,
    detached: true,
    stdio: ["ignore", "ignore", "ignore"],
  });
  child.unref();
  fs.writeFileSync(pidFile, String(child.pid));
  console.log(`Started HTTP server PID ${child.pid} at http://127.0.0.1:${port}/ (pid file: ${pidFile})`);
}

try {
  const opts = parseArgs(process.argv.slice(2));
  opts.pluginUrl = ensureTrailingSlash(opts.pluginUrl);
  const generated = writeSmokeFiles(opts);
  if (opts.publish) runPublish(opts, generated);
  if (opts.serve) startServer(generated.out, opts.servePort);

  console.log(`WPS smoke plugin written to ${generated.out}`);
  console.log(`Taskpane URL: ${opts.taskpaneUrl}`);
  console.log(`Plugin URL:   ${generated.pluginUrl}`);
  console.log("Files: index.html, ribbon.xml, main.js, manifest.xml, jsplugins.xml" + (opts.publish ? ", publish.html" : ""));
  console.log("Next: open the publish page inside Windows, e.g. http://10.0.2.2:3889/publish.html");
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
}
