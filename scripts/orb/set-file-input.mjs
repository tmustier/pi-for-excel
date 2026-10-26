// Sets a file on the first <input type=file> found in any page or iframe of
// the Excel-on-the-web Chrome (CDP 127.0.0.1:$PI_EXCEL_WEB_CDP_PORT, default 9223). The "Upload My Add-in"
// dialog lives in a cross-origin iframe that agent-browser refs cannot reach.
//
//   node scripts/orb/set-file-input.mjs manifest.xml
//
// Use 127.0.0.1, not localhost: Node resolves localhost to ::1, where another
// browser may be listening.
import path from "node:path";

const file = path.resolve(process.argv[2] ?? "manifest.xml");
const port = process.env.PI_EXCEL_WEB_CDP_PORT ?? "9223";
const targets = await (await fetch(`http://127.0.0.1:${port}/json`)).json();

for (const target of targets.filter((t) => t.type === "page" || t.type === "iframe")) {
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve) => { ws.onopen = resolve; });
  let nextId = 0;
  const pending = new Map();
  ws.onmessage = (event) => {
    const message = JSON.parse(event.data);
    pending.get(message.id)?.(message);
  };
  const send = (method, params = {}) => new Promise((resolve) => {
    const id = ++nextId;
    pending.set(id, resolve);
    ws.send(JSON.stringify({ id, method, params }));
  });

  const doc = await send("DOM.getDocument", { depth: -1, pierce: true });
  const found = await send("DOM.querySelectorAll", { nodeId: doc.result.root.nodeId, selector: "input[type=file]" });
  const nodeId = found.result?.nodeIds?.[0];
  if (nodeId) {
    const set = await send("DOM.setFileInputFiles", { nodeId, files: [file] });
    ws.close();
    if (set.error) throw new Error(set.error.message);
    console.log(`Set ${file} in ${target.url.slice(0, 80)}`);
    process.exit(0);
  }
  ws.close();
}

console.error("No file input found. Is the Upload Add-in dialog open?");
process.exit(1);
