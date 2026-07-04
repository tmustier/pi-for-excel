/**
 * Stub for Amazon Bedrock provider.
 *
 * Pi for Excel runs in a browser webview. The Bedrock implementation pulls in
 * AWS SDK Node-only transports (`@smithy/node-http-handler`, Node `http`/
 * `https`, etc.), which cannot run in the Excel WebView.
 *
 * pi-ai 0.80 loads the Bedrock implementation lazily and accepts an override
 * module with the uniform `ProviderStreams` shape (`stream` / `streamSimple`),
 * installed via `setBedrockProviderModule()` (see
 * src/compat/bedrock-provider-stub.ts). Both exports throw with a clear
 * unsupported message.
 *
 * If/when we want Bedrock support, we should add a browser-safe implementation
 * (SigV4 + fetch) or load this provider dynamically only in Node environments.
 */

export function stream(): never {
  throw new Error("Amazon Bedrock is not supported in the Excel add-in build.");
}

export function streamSimple(): never {
  throw new Error("Amazon Bedrock is not supported in the Excel add-in build.");
}
