import { readFile } from "node:fs/promises";

function isHelpersTaskpaneCspPayloadShape(value) {
  return typeof value === "object" && value !== null;
}

export async function readTaskpaneCspDirectiveTokens(directiveName) {
  const raw = await readFile(new URL("../../vercel.json", import.meta.url), "utf8");
  const parsed = JSON.parse(raw);
  if (!isHelpersTaskpaneCspPayloadShape(parsed)) {
    throw new Error("Invalid vercel.json root structure");
  }

  const headersRaw = parsed.headers;
  if (!Array.isArray(headersRaw)) {
    throw new Error("vercel.json is missing top-level headers array");
  }

  const taskpaneEntry = headersRaw.find((entry) => isHelpersTaskpaneCspPayloadShape(entry) && entry.source === "/src/taskpane.html");
  if (!isHelpersTaskpaneCspPayloadShape(taskpaneEntry)) {
    throw new Error("vercel.json is missing /src/taskpane.html header configuration");
  }

  const headerListRaw = taskpaneEntry.headers;
  if (!Array.isArray(headerListRaw)) {
    throw new Error("/src/taskpane.html entry has no headers array");
  }

  const cspEntry = headerListRaw.find((entry) => isHelpersTaskpaneCspPayloadShape(entry) && entry.key === "Content-Security-Policy");
  if (!isHelpersTaskpaneCspPayloadShape(cspEntry) || typeof cspEntry.value !== "string") {
    throw new Error("Missing Content-Security-Policy value for /src/taskpane.html");
  }

  const directive = cspEntry.value
    .split(";")
    .map((entry) => entry.trim())
    .find((entry) => entry.startsWith(`${directiveName} `));

  if (!directive) {
    throw new Error(`CSP missing ${directiveName} directive`);
  }

  const tokens = directive
    .split(/\s+/)
    .slice(1)
    .filter((token) => token.length > 0);

  return new Set(tokens);
}

export async function readTaskpaneConnectSrcTokens() {
  return readTaskpaneCspDirectiveTokens("connect-src");
}

export async function readTaskpaneScriptSrcTokens() {
  return readTaskpaneCspDirectiveTokens("script-src");
}
