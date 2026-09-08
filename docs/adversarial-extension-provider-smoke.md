# Adversarial extension-provider smoke

This smoke verifies the complete model-provider path exposed to pasted extensions. It does not use a built-in provider.

The scenario runs inside a real Excel taskpane. It installs inline code through `ExtensionRuntimeManager.installFromCode()`, which routes the extension into the sandbox iframe runtime. It then uses the same host connection store, dynamic model runtime, selector and chat stream as the taskpane UI.

## Coverage

The runner starts two instrumented HTTPS OpenAI-compatible gateways and verifies:

- default permission denial, followed by explicit `connections.readwrite` and `models.register` grants
- real sandbox iframe activation from pasted extension code
- owner-qualified connection and provider IDs
- host-owned credential injection for discovery and inference
- denial of `connections.getSecrets()` inside the sandbox
- dynamic `/models` discovery and exact streamed completions
- cached operation when `/models` returns an outage
- rejection of oversized catalogues without replacing the last safe cache
- endpoint upgrades with the same extension/provider ID, including cache rebinding from gateway A to gateway B
- no new credential or request reaches gateway A after the staged cutover
- safe provider fallback after the extension unloads during an in-flight completion
- cancellation of delayed discovery during unload, so a late response cannot resurrect deleted catalogue data
- rejection when a provider endpoint host is outside the owning connection's exact allowlist

The gateways compare credentials internally but never print them. The taskpane bridge compares exact assistant text and returns only the match result and lengths.

## Run it

The worktree needs trusted local `cert.pem` and `key.pem` files. Start the normal dev taskpane and tokened background bridge:

```bash
export PI_BACKGROUND_VERIFY_TOKEN="$(openssl rand -hex 24)"
export PI_BACKGROUND_VERIFY_HOST=localhost

npm run background:verify:bridge
```

In another terminal, start Vite with the same token:

```bash
VITE_PI_BACKGROUND_VERIFY_URL=https://localhost:3157 \
VITE_PI_BACKGROUND_VERIFY_TOKEN="$PI_BACKGROUND_VERIFY_TOKEN" \
npm run dev
```

Start the local proxy with an explicit localhost-only target policy. These loopback overrides are for this local smoke only; do not use them on a shared proxy:

```bash
ALLOWED_TARGET_HOSTS=localhost \
ALLOW_LOOPBACK_TARGETS=1 \
NODE_EXTRA_CA_CERTS="$HOME/Library/Application Support/mkcert/rootCA.pem" \
npm run proxy:https
```

Sideload/open the add-in in Excel, wait for the background bridge client, then run:

```bash
PI_BACKGROUND_VERIFY_TOKEN="$PI_BACKGROUND_VERIFY_TOKEN" \
PI_BACKGROUND_VERIFY_HOST=localhost \
NODE_EXTRA_CA_CERTS="$HOME/Library/Application Support/mkcert/rootCA.pem" \
npm run smoke:extension-provider
```

The mock gateways use ports `3161` and `3162`. The runner removes its extensions and resets the permission-gate experiment in a `finally` block.

A successful run prints one bounded JSON summary with `"ok": true` and no credential values.
