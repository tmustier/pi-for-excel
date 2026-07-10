import os from "node:os";

import WebSocket from "ws";

export const CODEX_WEBSOCKET_BRIDGE_HEADER = "X-Pi-For-Excel-Codex-WebSocket-Bridge";
export const CODEX_WEBSOCKET_BRIDGE_TRANSPORT = "codex-websocket";

const MAX_REQUEST_BODY_BYTES = 32 * 1024 * 1024;
const CODEX_WEBSOCKET_TARGET_URL = "wss://chatgpt.com/backend-api/codex/responses";
const OPENAI_WEBSOCKET_BETA = "responses_websockets=2026-02-06";
const TERMINAL_EVENT_TYPES = new Set([
  "error",
  "response.cancelled",
  "response.completed",
  "response.failed",
  "response.incomplete",
]);

export function isCodexWebSocketBridgeTarget(targetUrl) {
  return targetUrl.protocol === "https:"
    && targetUrl.hostname.toLowerCase() === "chatgpt.com"
    && targetUrl.port === ""
    && targetUrl.username === ""
    && targetUrl.password === ""
    && targetUrl.pathname === "/backend-api/codex/responses"
    && targetUrl.search === "";
}

function websocketHeaders(outboundHeaders) {
  const headers = Object.fromEntries(outboundHeaders.entries());
  delete headers.accept;
  delete headers.cookie;
  delete headers["content-encoding"];
  delete headers["content-length"];
  delete headers["content-type"];
  delete headers.origin;
  delete headers.referer;
  for (const name of Object.keys(headers)) {
    if (name.startsWith("sec-websocket-")) {
      delete headers[name];
    }
  }
  headers["openai-beta"] = OPENAI_WEBSOCKET_BETA;
  headers.originator = "pi";
  // Match native pi-ai exactly. ChatGPT's Codex model router uses this header;
  // forwarding a browser/proxy UA can select an unavailable rollout alias.
  headers["user-agent"] = `pi (${os.platform()} ${os.release()}; ${os.arch()})`;
  return headers;
}

async function readRequestBody(req) {
  const chunks = [];
  let totalBytes = 0;

  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    totalBytes += buffer.byteLength;
    if (totalBytes > MAX_REQUEST_BODY_BYTES) {
      const error = new Error("Codex WebSocket bridge request body is too large");
      error.code = "request_body_too_large";
      throw error;
    }
    chunks.push(buffer);
  }

  return Buffer.concat(chunks).toString("utf8");
}

function endResponse(res, statusCode, message) {
  if (res.writableEnded) return;
  if (!res.headersSent) {
    res.statusCode = statusCode;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
  }
  res.end(message);
}

function endBridgeFailure(res, statusCode, message) {
  if (res.writableEnded) return;
  if (!res.headersSent) {
    endResponse(res, statusCode, message);
    return;
  }

  const payload = {
    type: "error",
    error: { type: "proxy_stream_error", message },
    status: statusCode,
  };
  res.write(encodeSseData("", payload));
  res.end();
}

function parseWebSocketEvent(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function isTerminalWebSocketEvent(payload) {
  return typeof payload?.type === "string" && TERMINAL_EVENT_TYPES.has(payload.type);
}

function encodeSseData(text, payload) {
  const normalized = payload === null ? text : JSON.stringify(payload);
  return normalized
    .split(/\r?\n/u)
    .map((line) => `data: ${line}`)
    .join("\n") + "\n\n";
}

function buildWebSocketRequestBody(requestBody) {
  const payload = JSON.parse(requestBody);
  if (typeof payload !== "object" || payload === null || Array.isArray(payload)) {
    throw new Error("Codex WebSocket bridge request body must be a JSON object");
  }
  return JSON.stringify({ ...payload, type: "response.create" });
}

/**
 * Convert one authenticated HTTP POST from the browser into an upstream Codex
 * WebSocket request, then expose each upstream JSON frame as an SSE data event.
 *
 * The caller must apply origin/client/target policy before invoking this helper.
 */
export async function bridgeCodexWebSocketToSse({
  req,
  res,
  targetUrl,
  outboundHeaders,
  WebSocketConstructor = WebSocket,
}) {
  if (req.method !== "POST") {
    endResponse(res, 405, "Codex WebSocket bridge requires POST");
    return;
  }
  if (!isCodexWebSocketBridgeTarget(targetUrl)) {
    endResponse(res, 400, "Codex WebSocket bridge target is not allowed");
    return;
  }

  let requestBody;
  try {
    requestBody = buildWebSocketRequestBody(await readRequestBody(req));
  } catch (error) {
    const tooLarge = error instanceof Error && error.code === "request_body_too_large";
    const message = error instanceof Error ? error.message : "Invalid request body";
    endResponse(res, tooLarge ? 413 : 400, message);
    return;
  }

  await new Promise((resolve) => {
    let settled = false;
    let handlingUnexpectedResponse = false;
    let receivedTerminalEvent = false;
    // Never connect to the request URL directly: it is validated above for
    // request semantics, while the network sink stays a compile-time constant.
    const socket = new WebSocketConstructor(CODEX_WEBSOCKET_TARGET_URL, {
      headers: websocketHeaders(outboundHeaders),
      perMessageDeflate: false,
    });

    const settle = () => {
      if (settled) return;
      settled = true;
      resolve(undefined);
    };

    const closeUpstream = (reason) => {
      try {
        socket.close(1000, reason);
      } catch {
        // Best-effort cleanup only.
      }
    };

    req.once("aborted", () => {
      closeUpstream("client_aborted");
      settle();
    });

    res.once("close", () => {
      if (!res.writableEnded) {
        closeUpstream("client_closed");
      }
      settle();
    });

    socket.once("open", () => {
      if (settled) {
        closeUpstream("client_closed");
        return;
      }

      res.statusCode = 200;
      res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
      res.setHeader("Cache-Control", "no-store");
      res.setHeader(CODEX_WEBSOCKET_BRIDGE_HEADER, "1");
      res.flushHeaders?.();
      socket.send(requestBody);
    });

    socket.on("message", (data) => {
      if (settled || res.writableEnded) return;
      const text = typeof data === "string" ? data : Buffer.from(data).toString("utf8");
      const payload = parseWebSocketEvent(text);
      res.write(encodeSseData(text, payload));
      if (isTerminalWebSocketEvent(payload)) {
        receivedTerminalEvent = true;
        res.end();
        closeUpstream("response_complete");
        settle();
      }
    });

    socket.once("unexpected-response", (_request, upstreamResponse) => {
      handlingUnexpectedResponse = true;
      if (res.headersSent) {
        upstreamResponse.resume();
        endBridgeFailure(res, 502, "Codex WebSocket bridge rejected after response start");
        settle();
        return;
      }

      res.statusCode = upstreamResponse.statusCode ?? 502;
      const contentType = upstreamResponse.headers["content-type"];
      if (typeof contentType === "string") {
        res.setHeader("Content-Type", contentType);
      } else {
        res.setHeader("Content-Type", "text/plain; charset=utf-8");
      }
      res.setHeader(CODEX_WEBSOCKET_BRIDGE_HEADER, "1");
      upstreamResponse.pipe(res);
      upstreamResponse.once("end", settle);
      upstreamResponse.once("error", () => {
        if (res.headersSent) {
          res.destroy(new Error("Codex WebSocket bridge upstream response failed"));
        } else {
          endResponse(res, 502, "Codex WebSocket bridge upstream response failed");
        }
        settle();
      });
    });

    socket.once("error", () => {
      if (handlingUnexpectedResponse || settled) return;
      endBridgeFailure(res, 502, "Codex WebSocket bridge connection failed");
      closeUpstream("connection_failed");
      settle();
    });

    socket.once("close", () => {
      if (handlingUnexpectedResponse || settled) return;
      if (!res.headersSent) {
        endResponse(res, 502, "Codex WebSocket bridge closed before connecting");
      } else if (!receivedTerminalEvent) {
        endBridgeFailure(res, 502, "Codex WebSocket bridge closed before a terminal response event");
      }
      settle();
    });
  });
}
