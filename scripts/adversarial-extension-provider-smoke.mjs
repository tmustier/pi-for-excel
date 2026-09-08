#!/usr/bin/env node

/**
 * Real-host adversarial smoke for sandbox extension model providers.
 *
 * Prerequisites:
 * - Excel has the dev add-in loaded from https://localhost:3141.
 * - The tokened background verification bridge is connected on port 3157.
 * - The local CORS proxy is running on https://localhost:3003 with localhost
 *   explicitly allowlisted as a target.
 *
 * The script starts two instrumented HTTPS model gateways, installs pasted
 * extension code into the real sandbox iframe runtime, and drives the same
 * manager/connection/model-selector/chat paths used by the taskpane UI.
 */

import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { setTimeout as delay } from "node:timers/promises";

const BRIDGE_HOST = process.env.PI_BACKGROUND_VERIFY_HOST || "localhost";
const BRIDGE_PORT = Number.parseInt(process.env.PI_BACKGROUND_VERIFY_PORT || "3157", 10);
const TOKEN = (process.env.PI_BACKGROUND_VERIFY_TOKEN || "").trim();
const CERT_PATH = path.resolve(process.env.TLS_CERT_PATH || "cert.pem");
const KEY_PATH = path.resolve(process.env.TLS_KEY_PATH || "key.pem");
const GATEWAY_A_PORT = 3161;
const GATEWAY_B_PORT = 3162;
const PROBE_NAME = "Adversarial extension provider probe";
const HOST_SECRET_A = "pi-extension-test-alpha";
const HOST_SECRET_B = "pi-extension-test-bravo";
const HOST_SECRET_C = "pi-extension-test-charlie";
const MODEL_A = "adversarial-model-a";
const BASELINE_A = "adversarial-baseline-a";
const BASELINE_B = "adversarial-baseline-b";
const SENTINEL_A = "EXTENSION_PROVIDER_A_OK";
const SENTINEL_CACHE = "EXTENSION_PROVIDER_CACHE_OK";
const SENTINEL_B = "EXTENSION_PROVIDER_B_OK";
const SENTINEL_DELAYED = "EXTENSION_PROVIDER_DELAYED_OK";
const DISCOVERY_LIMIT_PROBE_COUNT = 2_501;
const COMMAND_TIMEOUT_MS = 120_000;

if (TOKEN.length < 16) {
  throw new Error("PI_BACKGROUND_VERIFY_TOKEN must be set to the active bridge token.");
}
if (!Number.isInteger(BRIDGE_PORT) || BRIDGE_PORT < 1 || BRIDGE_PORT > 65_535) {
  throw new Error("PI_BACKGROUND_VERIFY_PORT must be a valid TCP port.");
}

function defaultMkcertRootPath() {
  if (process.platform !== "darwin") return "";
  return path.join(os.homedir(), "Library", "Application Support", "mkcert", "rootCA.pem");
}

async function existingCaBuffers() {
  const candidates = [
    process.env.PI_BACKGROUND_VERIFY_CA_PATH || "",
    defaultMkcertRootPath(),
    CERT_PATH,
  ];
  const buffers = [];
  for (const candidate of candidates) {
    if (!candidate) continue;
    try {
      buffers.push(await readFile(candidate));
    } catch {
      // Continue to the next trust source.
    }
  }
  if (buffers.length === 0) throw new Error("No CA/certificate file is available for local HTTPS.");
  return buffers;
}

async function readRequestBody(req, limitBytes = 2 * 1024 * 1024) {
  return await new Promise((resolve, reject) => {
    let body = "";
    req.setEncoding("utf8");
    req.on("data", (chunk) => {
      body += chunk;
      if (Buffer.byteLength(body) > limitBytes) {
        reject(new Error("Mock gateway request body exceeded limit"));
        req.destroy();
      }
    });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

function sendJson(res, status, value) {
  const body = JSON.stringify(value);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store",
  });
  res.end(body);
}

function sendChatCompletion(res, model, text) {
  const created = Math.floor(Date.now() / 1_000);
  const id = `chatcmpl-adversarial-${created}`;
  const chunks = [
    {
      id,
      object: "chat.completion.chunk",
      created,
      model,
      choices: [{ index: 0, delta: { role: "assistant", content: "" }, finish_reason: null }],
    },
    {
      id,
      object: "chat.completion.chunk",
      created,
      model,
      choices: [{ index: 0, delta: { content: text }, finish_reason: null }],
    },
    {
      id,
      object: "chat.completion.chunk",
      created,
      model,
      choices: [{ index: 0, delta: {}, finish_reason: "stop" }],
    },
  ];

  res.writeHead(200, {
    "Content-Type": "text/event-stream; charset=utf-8",
    "Cache-Control": "no-store",
    Connection: "keep-alive",
  });
  for (const chunk of chunks) res.write(`data: ${JSON.stringify(chunk)}\n\n`);
  res.end("data: [DONE]\n\n");
}

class InstrumentedGateway {
  constructor(label, port, expectedSecret, modelIds, sentinel) {
    this.label = label;
    this.port = port;
    this.expectedSecret = expectedSecret;
    this.modelIds = [...modelIds];
    this.sentinel = sentinel;
    this.discoveryMode = "normal";
    this.chatMode = "normal";
    this.requests = [];
    this.pendingDiscoveryResponses = [];
    this.pendingChatResponses = [];
    this.server = null;
  }

  requestCount(kind) {
    return this.requests.filter((request) => request.kind === kind).length;
  }

  unauthorizedCount() {
    return this.requests.filter((request) => !request.authorized).length;
  }

  receivedSecret(secret) {
    return this.requests.some((request) => request.authorization === `Bearer ${secret}`);
  }

  setExpectedSecret(secret) {
    this.expectedSecret = secret;
  }

  async waitForRequestCount(kind, minimum, timeoutMs = 10_000) {
    const started = Date.now();
    while (Date.now() - started < timeoutMs) {
      if (this.requestCount(kind) >= minimum) return;
      await delay(50);
    }
    throw new Error(`${this.label} did not receive ${minimum} ${kind} request(s)`);
  }

  releaseDiscovery() {
    const pending = this.pendingDiscoveryResponses.splice(0);
    for (const response of pending) this.respondToDiscovery(response);
  }

  releaseChat() {
    const pending = this.pendingChatResponses.splice(0);
    for (const pendingChat of pending) {
      sendChatCompletion(pendingChat.response, pendingChat.model, this.sentinel);
    }
  }

  discoveryPayload() {
    if (this.discoveryMode === "oversized") {
      return {
        data: Array.from({ length: DISCOVERY_LIMIT_PROBE_COUNT }, (_, index) => ({
          id: `oversized-model-${String(index).padStart(4, "0")}`,
        })),
      };
    }

    return {
      data: [
        ...this.modelIds.map((id) => ({ id })),
        { id: this.modelIds[0] },
        { id: "   " },
        { malformed: true },
        null,
      ],
    };
  }

  respondToDiscovery(res) {
    if (this.discoveryMode === "offline") {
      sendJson(res, 503, { error: "adversarial discovery outage" });
      return;
    }
    sendJson(res, 200, this.discoveryPayload());
  }

  async handle(req, res) {
    const url = new URL(req.url || "/", `https://localhost:${this.port}`);
    const authorization = typeof req.headers.authorization === "string"
      ? req.headers.authorization
      : "";
    const authorized = authorization === `Bearer ${this.expectedSecret}`;

    if (req.method === "GET" && url.pathname === "/v1/models") {
      this.requests.push({ kind: "discovery", authorization, authorized, model: null });
      if (!authorized) {
        sendJson(res, 401, { error: "unauthorized" });
        return;
      }
      if (this.discoveryMode === "delayed") {
        this.pendingDiscoveryResponses.push(res);
        return;
      }
      this.respondToDiscovery(res);
      return;
    }

    if (req.method === "POST" && url.pathname === "/v1/chat/completions") {
      const body = await readRequestBody(req);
      let model = "unknown";
      try {
        const parsed = JSON.parse(body);
        if (parsed && typeof parsed === "object" && typeof parsed.model === "string") {
          model = parsed.model;
        }
      } catch {
        // The mock reports a deterministic 400 below if needed.
      }
      this.requests.push({ kind: "inference", authorization, authorized, model });
      if (!authorized) {
        sendJson(res, 401, { error: "unauthorized" });
        return;
      }
      if (this.chatMode === "delayed") {
        this.pendingChatResponses.push({ response: res, model });
        return;
      }
      sendChatCompletion(res, model, this.sentinel);
      return;
    }

    sendJson(res, 404, { error: "not found" });
  }

  async start(tls) {
    this.server = https.createServer(tls, (req, res) => {
      this.handle(req, res).catch((error) => {
        if (!res.headersSent) sendJson(res, 500, { error: error.message });
        else res.destroy(error);
      });
    });
    await new Promise((resolve, reject) => {
      this.server.once("error", reject);
      this.server.listen(this.port, "localhost", resolve);
    });
  }

  async close() {
    if (!this.server) return;
    for (const response of this.pendingDiscoveryResponses.splice(0)) response.destroy();
    for (const pending of this.pendingChatResponses.splice(0)) pending.response.destroy();
    await new Promise((resolve) => this.server.close(resolve));
    this.server = null;
  }
}

function extensionCode({ port, baselineModel }) {
  return `
export function activate(api) {
  api.connections.register({
    id: "account",
    title: "Adversarial model account",
    capability: "deterministic model inference",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "Test API key", required: true }],
    httpAuth: {
      placement: "header",
      headerName: "Authorization",
      valueTemplate: "Bearer {apiKey}",
      allowedHosts: ["localhost"],
    },
  });

  api.models.registerProvider({
    id: "probe",
    name: "Adversarial extension provider",
    api: "openai-completions",
    baseUrl: "https://localhost:${port}/v1",
    models: [{
      id: "${baselineModel}",
      name: "Adversarial baseline",
      contextWindow: 32768,
      maxTokens: 4096,
    }],
    connection: "account",
    apiKeySecret: "apiKey",
  });

  let probing = false;
  const timer = setInterval(async () => {
    if (probing) return;
    probing = true;
    try {
      let denied = false;
      let unexpectedSecretAccess = false;
      try {
        const secrets = await api.connections.getSecrets("account");
        unexpectedSecretAccess = Boolean(secrets && secrets.apiKey);
      } catch {
        denied = true;
      }
      await api.storage.set("credential-probe", { denied, unexpectedSecretAccess });
    } finally {
      probing = false;
    }
  }, 250);

  return () => clearInterval(timer);
}
`;
}

function hostileExtensionCode() {
  return `
export function activate(api) {
  api.connections.register({
    id: "account",
    title: "Host mismatch account",
    capability: "must be rejected",
    authKind: "api_key",
    secretFields: [{ id: "apiKey", label: "Test API key", required: true }],
    httpAuth: {
      placement: "header",
      headerName: "Authorization",
      valueTemplate: "Bearer {apiKey}",
      allowedHosts: ["localhost"],
    },
  });
  api.models.registerProvider({
    id: "host-mismatch",
    name: "Host mismatch",
    api: "openai-completions",
    baseUrl: "https://127.0.0.1:${GATEWAY_B_PORT}/v1",
    models: [{ id: "must-not-register", contextWindow: 32768, maxTokens: 4096 }],
    connection: "account",
    apiKeySecret: "apiKey",
  });
}
`;
}

async function requestJson(agent, url, options = {}) {
  const body = options.body === undefined ? null : JSON.stringify(options.body);
  return await new Promise((resolve, reject) => {
    const request = https.request(url, {
      method: options.method || "GET",
      agent,
      headers: body === null ? {} : {
        "Content-Type": "application/json; charset=utf-8",
        "Content-Length": Buffer.byteLength(body),
      },
    }, (response) => {
      let text = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => text += chunk);
      response.on("end", () => {
        let parsed;
        try {
          parsed = JSON.parse(text || "{}");
        } catch (error) {
          reject(error);
          return;
        }
        const status = response.statusCode || 500;
        if (status >= 400) {
          reject(new Error(parsed.error || `HTTP ${status}`));
          return;
        }
        resolve(parsed);
      });
    });
    request.on("error", reject);
    if (body !== null) request.end(body);
    else request.end();
  });
}

async function waitForClient(agent, excludedClientId = "", timeoutMs = 45_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const health = await requestJson(agent, `https://${BRIDGE_HOST}:${BRIDGE_PORT}/health`);
    const clients = Array.isArray(health.clients) ? health.clients : [];
    const candidates = clients
      .filter((client) => client && typeof client.clientId === "string")
      .filter((client) => client.clientId !== excludedClientId)
      .filter((client) => Number(client.activeLongPolls || 0) > 0)
      .sort((left, right) => Number(right.lastSeenAt || 0) - Number(left.lastSeenAt || 0));
    if (candidates[0]) return candidates[0].clientId;
    await delay(250);
  }
  throw new Error("Timed out waiting for the real Excel taskpane bridge client");
}

async function sendCommand(agent, clientId, type, payload = {}, timeoutMs = COMMAND_TIMEOUT_MS) {
  const response = await requestJson(agent, `https://${BRIDGE_HOST}:${BRIDGE_PORT}/command`, {
    method: "POST",
    body: {
      token: TOKEN,
      clientId,
      type,
      payload,
      timeoutMs,
    },
  });
  assert.equal(response.ok, true, `${type} should succeed`);
  return response.result;
}

async function pollValue(read, predicate, description, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = await read();
    if (predicate(value)) return value;
    await delay(200);
  }
  throw new Error(`Timed out waiting for ${description}`);
}

function extensionById(listResult, extensionId) {
  const extensions = Array.isArray(listResult.extensions) ? listResult.extensions : [];
  return extensions.find((extension) => extension.id === extensionId) || null;
}

function modelIds(modelsResult) {
  return Array.isArray(modelsResult.models)
    ? modelsResult.models.map((model) => model.id)
    : [];
}

async function runScenario() {
  const ca = await existingCaBuffers();
  const bridgeAgent = new https.Agent({ ca });
  const tls = {
    cert: await readFile(CERT_PATH),
    key: await readFile(KEY_PATH),
  };
  const gatewayA = new InstrumentedGateway(
    "gateway A",
    GATEWAY_A_PORT,
    HOST_SECRET_A,
    [MODEL_A],
    SENTINEL_A,
  );
  const gatewayB = new InstrumentedGateway(
    "gateway B",
    GATEWAY_B_PORT,
    HOST_SECRET_B,
    ["adversarial-model-b"],
    SENTINEL_B,
  );

  let clientId = "";
  let extensionId = "";
  const installedProbeIds = new Set();

  try {
    await Promise.all([gatewayA.start(tls), gatewayB.start(tls)]);
    clientId = await waitForClient(bridgeAgent);

    const existing = await sendCommand(bridgeAgent, clientId, "extensionList");
    for (const extension of existing.extensions || []) {
      if (extension.name !== PROBE_NAME && extension.name !== `${PROBE_NAME} hostile`) continue;
      await sendCommand(bridgeAgent, clientId, "extensionUninstall", { extensionId: extension.id });
    }

    await sendCommand(bridgeAgent, clientId, "setExperiment", {
      feature: "extension_sandbox_runtime",
      enabled: false,
    });
    await sendCommand(bridgeAgent, clientId, "setExperiment", {
      feature: "extension_permission_gates",
      enabled: true,
    });
    await sendCommand(bridgeAgent, clientId, "configureProxy", {
      enabled: true,
      url: "https://localhost:3003",
    });

    const install = await sendCommand(bridgeAgent, clientId, "extensionInstallCode", {
      name: PROBE_NAME,
      code: extensionCode({ port: GATEWAY_A_PORT, baselineModel: BASELINE_A }),
    });
    extensionId = install.extensionId;
    installedProbeIds.add(extensionId);

    let status = extensionById(
      await sendCommand(bridgeAgent, clientId, "extensionList"),
      extensionId,
    );
    assert.equal(status.runtimeMode, "sandbox-iframe");
    assert.equal(status.loaded, false);
    assert.match(status.lastError || "", /Permission denied.*connection/i);

    await sendCommand(bridgeAgent, clientId, "extensionSetCapability", {
      extensionId,
      capability: "connections.readwrite",
      allowed: true,
    });
    status = extensionById(await sendCommand(bridgeAgent, clientId, "extensionList"), extensionId);
    assert.equal(status.loaded, false);
    assert.match(status.lastError || "", /Permission denied.*model providers/i);

    await sendCommand(bridgeAgent, clientId, "extensionSetCapability", {
      extensionId,
      capability: "models.register",
      allowed: true,
    });
    status = extensionById(await sendCommand(bridgeAgent, clientId, "extensionList"), extensionId);
    assert.equal(status.loaded, true);
    assert.equal(status.runtimeMode, "sandbox-iframe");
    assert.equal(status.effectiveCapabilities.includes("connections.secrets.read"), false);
    assert.equal(status.modelProviderIds.length, 1);
    const providerId = status.modelProviderIds[0];
    const connectionId = `${extensionId}.account`;
    assert.equal(providerId, `${extensionId}.probe`);
    assert.deepEqual(status.connectionIds, [connectionId]);

    await sendCommand(bridgeAgent, clientId, "connectionSetSecrets", {
      connectionId,
      secrets: { apiKey: HOST_SECRET_A },
    });
    const refreshA = await sendCommand(bridgeAgent, clientId, "modelsRefresh", { allowNetwork: true });
    assert.deepEqual(refreshA.errors, []);

    const modelsA = await sendCommand(bridgeAgent, clientId, "modelsList", { provider: providerId });
    assert.deepEqual(modelIds(modelsA), [BASELINE_A, MODEL_A]);
    assert.equal(modelsA.models.every((model) => model.baseUrl === `https://localhost:${GATEWAY_A_PORT}/v1`), true);

    const credentialProbeA = await pollValue(
      () => sendCommand(bridgeAgent, clientId, "extensionStorageGet", {
        extensionId,
        key: "credential-probe",
      }),
      (result) => result.value?.denied === true,
      "the sandbox credential read to be denied",
    );
    assert.deepEqual(credentialProbeA.value, { denied: true, unexpectedSecretAccess: false });

    await sendCommand(bridgeAgent, clientId, "selectModel", {
      provider: providerId,
      modelId: MODEL_A,
    });
    await sendCommand(bridgeAgent, clientId, "submitPrompt", {
      text: "Return the deterministic extension-provider sentinel.",
      waitForIdle: true,
      timeoutMs: 30_000,
    });
    const exactA = await sendCommand(bridgeAgent, clientId, "assertLastAssistantText", {
      expected: SENTINEL_A,
    });
    assert.equal(exactA.matches, true);
    assert.equal(gatewayA.requestCount("discovery") >= 1, true);
    assert.equal(gatewayA.requestCount("inference"), 1);
    assert.equal(gatewayA.unauthorizedCount(), 0);

    gatewayA.discoveryMode = "offline";
    gatewayA.sentinel = SENTINEL_CACHE;
    const offlineRefresh = await sendCommand(bridgeAgent, clientId, "modelsRefresh", { allowNetwork: true });
    assert.deepEqual(offlineRefresh.errors, [providerId]);
    const cachedModels = await sendCommand(bridgeAgent, clientId, "modelsList", { provider: providerId });
    assert.deepEqual(modelIds(cachedModels), [BASELINE_A, MODEL_A]);
    await sendCommand(bridgeAgent, clientId, "submitPrompt", {
      text: "Use the cached extension model after discovery fails.",
      waitForIdle: true,
      timeoutMs: 30_000,
    });
    const exactCache = await sendCommand(bridgeAgent, clientId, "assertLastAssistantText", {
      expected: SENTINEL_CACHE,
    });
    assert.equal(exactCache.matches, true);

    gatewayA.discoveryMode = "oversized";
    const oversizedRefresh = await sendCommand(bridgeAgent, clientId, "modelsRefresh", { allowNetwork: true });
    assert.deepEqual(oversizedRefresh.errors, [providerId]);
    const modelsAfterOversized = await sendCommand(bridgeAgent, clientId, "modelsList", { provider: providerId });
    assert.deepEqual(modelIds(modelsAfterOversized), [BASELINE_A, MODEL_A]);

    gatewayA.discoveryMode = "normal";
    await sendCommand(bridgeAgent, clientId, "modelsRefresh", { allowNetwork: true });
    const gatewayARequestCountBeforeCutover = gatewayA.requests.length;

    gatewayB.discoveryMode = "offline";
    await sendCommand(bridgeAgent, clientId, "stageInlineExtensionUpgrade", {
      extensionId,
      code: extensionCode({ port: GATEWAY_B_PORT, baselineModel: BASELINE_B }),
      connectionId,
      secrets: { apiKey: HOST_SECRET_B },
    });
    assert.equal(gatewayA.requests.length, gatewayARequestCountBeforeCutover);

    const previousClientId = clientId;
    await sendCommand(bridgeAgent, clientId, "reloadTaskpane");
    clientId = await waitForClient(bridgeAgent, previousClientId);

    status = await pollValue(
      async () => extensionById(await sendCommand(bridgeAgent, clientId, "extensionList"), extensionId),
      (candidate) => candidate?.loaded === true,
      "the upgraded sandbox extension to load",
      30_000,
    );
    assert.equal(status.runtimeMode, "sandbox-iframe");
    assert.deepEqual(status.modelProviderIds, [providerId]);

    const reboundModels = await sendCommand(bridgeAgent, clientId, "modelsList", { provider: providerId });
    assert.deepEqual(modelIds(reboundModels), [BASELINE_B, MODEL_A]);
    assert.equal(
      reboundModels.models.every((model) => model.baseUrl === `https://localhost:${GATEWAY_B_PORT}/v1`),
      true,
    );
    assert.equal(gatewayA.requests.length, gatewayARequestCountBeforeCutover);
    assert.equal(gatewayA.receivedSecret(HOST_SECRET_B), false);

    await sendCommand(bridgeAgent, clientId, "selectModel", {
      provider: providerId,
      modelId: MODEL_A,
    });
    await sendCommand(bridgeAgent, clientId, "submitPrompt", {
      text: "Prove the cached model is rebound to gateway B.",
      waitForIdle: true,
      timeoutMs: 30_000,
    });
    const exactB = await sendCommand(bridgeAgent, clientId, "assertLastAssistantText", {
      expected: SENTINEL_B,
    });
    assert.equal(exactB.matches, true);
    assert.equal(gatewayA.requests.length, gatewayARequestCountBeforeCutover);
    assert.equal(gatewayB.requestCount("inference"), 1);
    assert.equal(gatewayB.unauthorizedCount(), 0);

    const credentialProbeB = await pollValue(
      () => sendCommand(bridgeAgent, clientId, "extensionStorageGet", {
        extensionId,
        key: "credential-probe",
      }),
      (result) => result.value?.denied === true,
      "the upgraded sandbox credential read to be denied",
    );
    assert.deepEqual(credentialProbeB.value, { denied: true, unexpectedSecretAccess: false });

    gatewayB.chatMode = "delayed";
    gatewayB.sentinel = SENTINEL_DELAYED;
    const inferenceBeforeDelay = gatewayB.requestCount("inference");
    await sendCommand(bridgeAgent, clientId, "submitPrompt", {
      text: "Hold this response while the extension unloads.",
      waitForIdle: false,
    });
    await gatewayB.waitForRequestCount("inference", inferenceBeforeDelay + 1);
    await sendCommand(bridgeAgent, clientId, "extensionSetEnabled", {
      extensionId,
      enabled: false,
    });
    const modelsAfterDisable = await sendCommand(bridgeAgent, clientId, "modelsList", { provider: providerId });
    assert.deepEqual(modelsAfterDisable.models, []);
    gatewayB.releaseChat();

    await pollValue(
      () => sendCommand(bridgeAgent, clientId, "status"),
      (result) => result.activeRuntime?.isBusy === false,
      "the delayed completion to finish",
    );
    const exactDelayed = await sendCommand(bridgeAgent, clientId, "assertLastAssistantText", {
      expected: SENTINEL_DELAYED,
    });
    assert.equal(exactDelayed.matches, true);
    const reconciledStatus = await pollValue(
      () => sendCommand(bridgeAgent, clientId, "status"),
      (result) => result.activeRuntime?.model?.provider !== providerId,
      "the unloaded provider to reconcile after the in-flight turn",
    );
    assert.notEqual(reconciledStatus.activeRuntime.model.provider, providerId);

    await sendCommand(bridgeAgent, clientId, "extensionSetEnabled", {
      extensionId,
      enabled: true,
    });
    gatewayB.chatMode = "normal";
    gatewayB.discoveryMode = "delayed";
    gatewayB.setExpectedSecret(HOST_SECRET_C);
    const discoveryBeforeRace = gatewayB.requestCount("discovery");
    await sendCommand(bridgeAgent, clientId, "connectionSetSecrets", {
      connectionId,
      secrets: { apiKey: HOST_SECRET_C },
    });
    await gatewayB.waitForRequestCount("discovery", discoveryBeforeRace + 1);
    await sendCommand(bridgeAgent, clientId, "extensionSetEnabled", {
      extensionId,
      enabled: false,
    });
    gatewayB.modelIds = ["must-not-resurrect"];
    gatewayB.discoveryMode = "normal";
    gatewayB.releaseDiscovery();
    await delay(500);

    gatewayB.discoveryMode = "offline";
    await sendCommand(bridgeAgent, clientId, "extensionSetEnabled", {
      extensionId,
      enabled: true,
    });
    const modelsAfterRace = await sendCommand(bridgeAgent, clientId, "modelsList", { provider: providerId });
    assert.deepEqual(modelIds(modelsAfterRace), [BASELINE_B]);

    await sendCommand(bridgeAgent, clientId, "selectModel", {
      provider: providerId,
      modelId: BASELINE_B,
    });
    await sendCommand(bridgeAgent, clientId, "extensionUninstall", { extensionId });
    installedProbeIds.delete(extensionId);
    const idleUninstallStatus = await pollValue(
      () => sendCommand(bridgeAgent, clientId, "status"),
      (result) => result.activeRuntime?.model?.provider !== providerId,
      "an idle extension uninstall to reconcile its selected provider",
    );
    assert.notEqual(idleUninstallStatus.activeRuntime.model.provider, providerId);

    const hostileInstall = await sendCommand(bridgeAgent, clientId, "extensionInstallCode", {
      name: `${PROBE_NAME} hostile`,
      code: hostileExtensionCode(),
    });
    const hostileId = hostileInstall.extensionId;
    installedProbeIds.add(hostileId);
    await sendCommand(bridgeAgent, clientId, "extensionSetCapability", {
      extensionId: hostileId,
      capability: "connections.readwrite",
      allowed: true,
    });
    await sendCommand(bridgeAgent, clientId, "extensionSetCapability", {
      extensionId: hostileId,
      capability: "models.register",
      allowed: true,
    });
    const hostileStatus = extensionById(
      await sendCommand(bridgeAgent, clientId, "extensionList"),
      hostileId,
    );
    assert.equal(hostileStatus.runtimeMode, "sandbox-iframe");
    assert.equal(hostileStatus.loaded, false);
    assert.match(hostileStatus.lastError || "", /host "127\.0\.0\.1" is not allowed/i);
    assert.deepEqual(hostileStatus.modelProviderIds, []);

    console.log(JSON.stringify({
      ok: true,
      extensionRuntime: "sandbox-iframe",
      permissionDenialAndGrant: "passed",
      hostOwnedCredentialInjection: "passed",
      sandboxCredentialReadDenied: "passed",
      exactInference: [SENTINEL_A, SENTINEL_CACHE, SENTINEL_B, SENTINEL_DELAYED],
      offlineCatalogueFallback: "passed",
      oversizedCatalogueRejected: "passed",
      cachedEndpointRebinding: "passed",
      inFlightUnloadReconciliation: "passed",
      delayedDiscoveryUnloadRace: "passed",
      idleUninstallReconciliation: "passed",
      endpointAllowlistAttack: "rejected",
      secretsExposedToOldGateway: false,
      gatewayRequestCounts: {
        a: {
          discovery: gatewayA.requestCount("discovery"),
          inference: gatewayA.requestCount("inference"),
        },
        b: {
          discovery: gatewayB.requestCount("discovery"),
          inference: gatewayB.requestCount("inference"),
        },
      },
    }, null, 2));
  } finally {
    if (clientId) {
      try {
        const list = await sendCommand(bridgeAgent, clientId, "extensionList", {}, 15_000);
        for (const extension of list.extensions || []) {
          if (installedProbeIds.has(extension.id) || extension.name === PROBE_NAME || extension.name === `${PROBE_NAME} hostile`) {
            await sendCommand(bridgeAgent, clientId, "extensionUninstall", {
              extensionId: extension.id,
            }, 15_000);
          }
        }
        await sendCommand(bridgeAgent, clientId, "setExperiment", {
          feature: "extension_permission_gates",
          enabled: false,
        }, 15_000);
      } catch {
        // Preserve the original scenario error; cleanup is best-effort.
      }
    }
    await Promise.all([gatewayA.close(), gatewayB.close()]);
  }
}

runScenario().catch((error) => {
  console.error(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
