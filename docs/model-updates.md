# Model / dependency update playbook

New models reach users without a release. Pi for Excel ships a pinned, tested Pi runtime
(`@earendil-works/pi-ai`, `@earendil-works/pi-agent-core`) and refreshes the model catalogue
from pi.dev at runtime, the way `pi update --models` does in the Pi CLI. Upgrade the runtime
only when you want its code changes (new transports, streaming, auth or agent fixes).

This doc covers:
- how the model catalogue stays current
- upgrading the **Pi dependency versions** we ship
- the **model ordering/default-selection behavior** in the add-in (`src/models/model-ordering.ts`, `src/models/featured-models.ts`, `src/taskpane/default-model.ts`)
- the **thinking-level UI** that reflects registry capabilities (`src/models/thinking-levels.ts`, `src/taskpane/thinking-display.ts`)

## How the model catalogue stays current

```
earendil-works/pi ──publish──▶ pi.dev ──/model-catalog/<version>/<provider>──▶ add-in ──▶ IndexedDB
 (catalogue workflow)                  (vercel.json rewrite,                     (validated overlay
                                        30 min edge cache)                        on bundled models)
```

- **Bundled models:** `node_modules/@earendil-works/pi-ai/dist/providers/data/*.json`, exposed through `builtinProviders()`. They are the fallback when there is no network, no relay or no newer catalogue.
- **Catalogue overlay:** `withRemoteCatalog()` (`src/models/remote-model-catalog.ts`) wraps each built-in provider, ported from Pi coding-agent's `remote-catalog-provider.ts`. It fetches `/model-catalog/<pi-ai version>/<provider>` on the add-in's own origin. `vercel.json` rewrites that to `https://pi.dev/api/models/providers/<provider>?types=chat&pi-version=<version>` (pi.dev serves the newest catalogue that version supports) with `CDN-Cache-Control: max-age=1800`; the Vite dev server proxies the same path.
- **When it refreshes:** whenever `ModelRefreshOwner` refreshes with network access: startup, opening the model picker, and connection or extension-provider changes. Only providers with configured credentials are fetched. Cached overlays are restored offline at startup. Requests revalidate with the stored ETag.
- **What remote data can change:** model IDs and descriptive metadata (name, limits, cost, reasoning, thinking levels, input types, prompt cache, compat). An entry must use an API and base URL that the bundled provider already uses, and it takes that provider's headers, so the catalogue cannot add a transport this runtime cannot stream or send credentials to a new host. Malformed entries are dropped.
- **Fail-safe:** a 404 (self-hosted without the relay), a failed request, a catalogue older than the bundled data, or a `x-pi-model-catalog-minimum-version` above the installed `pi-ai` leaves the last published list in place. The selected model is never switched; an idle session picks up changed metadata for its current model.
- **Self-hosting:** forward `/model-catalog/<version>/<provider>` the same way to get updates, for example with nginx:

  ```nginx
  location ~ ^/model-catalog/([0-9][0-9a-z.+-]*)/([a-z0-9-]+)$ {
    proxy_pass https://pi.dev/api/models/providers/$2?types=chat&pi-version=$1;
  }
  ```

  Without it, the add-in uses the models bundled with its `pi-ai` version.

Custom gateways and extension providers keep their own discovery:

- **Dynamic catalogue cache:** IndexedDB store `model-catalogs`, accessed through `ModelCatalogsStore`; restored entries are rebound to the provider's current API/base URL before use.
- **Discovery bounds:** responses are limited to 2 MiB, 2,000 entries and 256 characters per model ID. An invalid/oversized refresh leaves the last safe cache and configured baseline intact.
- **Manual gateways:** models remain in `CustomProvidersStore` and register through the same collection without `/models` discovery. An exact endpoint, model ID and API match supplies metadata from the shared registry; an explicit context limit takes precedence. Unknown models retain the existing defaults. Previously discovered additions are no longer loaded for manual-form entries.
- **Extension providers:** `api.models.registerProvider()` declarations are runtime-owned and unload with their extension. Unregistering aborts in-flight discovery before deleting its cache so late responses cannot resurrect stale entries.

Do not use Pi coding-agent's Node/file `ModelRuntime` directly in the Office WebView. Pi for Excel uses the same Pi AI provider primitives with browser storage, OAuth and proxy policy. Never infer aliases or metadata from marketing names.

Manual gateway metadata is a save-time snapshot. Existing entries retain saved limits until edited; clear the context field when re-saving to use current registry defaults. Other manual gateway snapshots are excluded from metadata lookup.

## Runtime upgrades

- Dependabot checks npm dependencies **daily**. The `pi-stack` group keeps `@earendil-works/pi-ai` and `@earendil-works/pi-agent-core` in one PR. Review and merge it by hand after an Excel smoke test; it is no longer needed for new models.
- `npm run check` includes `scripts/check-pi-deps-lockstep.mjs`, which enforces the dependency policy below (`pi-ai` === `pi-agent-core`, exact pins, single shared `pi-ai` copy in the lockfile, `typebox` pinned to the version `pi-ai` depends on).
- Dependabot ignores `typebox`: it is bumped by hand as part of a pi-stack bump, to whatever `npm view @earendil-works/pi-ai@<version> dependencies.typebox` says. A Dependabot PR that bumps `typebox` on its own fails the lockstep check by design.
- Upgrade when a model needs a transport (`api`) the add-in does not have yet; the catalogue drops such entries until then.

## Step-by-step

### 1) Check current installed versions

```bash
node -p "require('./node_modules/@earendil-works/pi-ai/package.json').version"
node -p "require('./node_modules/@earendil-works/pi-agent-core/package.json').version"
```

### 2) Check latest published versions

```bash
npm view @earendil-works/pi-ai version
npm view @earendil-works/pi-agent-core version
```

Also inspect the version lists before choosing a target:

```bash
npm view @earendil-works/pi-ai versions --json
npm view @earendil-works/pi-agent-core versions --json
```

**Dependency policy:**

- `@earendil-works/pi-ai` and `@earendil-works/pi-agent-core` move **together** to the newest common version. The browser runtime migration was verified on `0.83.0`.
- Both must be **exact-pinned** in `package.json`.
- The lockfile must resolve **exactly one** copy of `pi-ai`. Two copies = two model registries = the model selector and the app disagree about available models.
- `scripts/check-pi-deps-lockstep.mjs` (run via `npm run check:pi-lockstep`) enforces all of this.
- `@earendil-works/pi-web-ui` was **removed entirely** on 2026-07-07 (the UI layer is first-party; see `docs/ui-ownership.md`). Do not re-add it.
- Check the pi changelog for breaking OAuth/streaming surface changes when bumping (e.g. 0.79.x made `onDeviceCode` / `onSelect` required in `OAuthLoginCallbacks`).

### 3) Bump dependencies in `package.json`

```bash
npm install @earendil-works/pi-ai@<version> @earendil-works/pi-agent-core@<version> --save-exact
npm install typebox@"$(npm view @earendil-works/pi-ai@<version> dependencies.typebox)" --save-exact
```

### 4) Verify the new model IDs exist in the registry

Search the generated catalogue and query the runtime-facing built-in provider collection:

```bash
rg -n "gpt-5\\.6-(sol|terra|luna)" node_modules/@earendil-works/pi-ai/dist/providers/data -S
rg -n "claude-fable-5"             node_modules/@earendil-works/pi-ai/dist/providers/data -S
rg -n "claude-opus-(4-8|5)"        node_modules/@earendil-works/pi-ai/dist/providers/data -S
rg -n "gemini-3\\.1-pro-preview"   node_modules/@earendil-works/pi-ai/dist/providers/data -S
node --input-type=module -e 'import { builtinModels } from "@earendil-works/pi-ai/providers/all"; const m=builtinModels(); console.log(m.getModel("openai", "gpt-5.6-sol"))'
npm run test:models
```

These checks cover the bundled fallback. Models published to pi.dev after this `pi-ai` release arrive through the catalogue overlay instead.

### 5) Update model ordering + default selection logic (avoid hardcoding exact IDs)

Files:
- `src/models/model-ordering.ts` (provider/family priority + version/recency scoring)
- `src/models/featured-models.ts` (featured-model ordering used by the model picker)
- `src/taskpane/default-model.ts` (default-model selection rules)
- `src/models/thinking-levels.ts` (registry-driven thinking capabilities)
- `src/taskpane/thinking-display.ts` (localized thinking labels/hints/colors)
- `src/ui/model-selector-dialog.ts` (first-party model picker UI)
- `tests/model-ordering.test.ts` (selection and thinking-control regressions; run `npm run test:models` — requires Node 22.19+)

We intentionally avoid pinning exact versioned IDs now. Instead we:

- In the model picker, show:
  1) current model first
  2) **featured models** (pattern-based “latest” picks)
  3) then the rest sorted deterministically

  Featured rules (current desired behavior):
  - **Anthropic:** latest **Fable** first (post-4.x flagship family, e.g. `claude-fable-5`), then latest **Sonnet** *if* its version >= latest **Opus**, then latest **Opus**
    - This is picker ordering only; default selection currently skips Fable because it is in the registry but unavailable for normal Anthropic use.
    - Version compare uses `parseLeadingMajorMinor()` where `claude-opus-4-6` → `46`, `claude-opus-4-7` → `47`, `claude-fable-5` → `50`.
    - Important: IDs like `claude-opus-4-20250514` are treated as **major only** (`40`) and the `YYYYMMDD` part is considered a separate date suffix by `modelRecencyScore()`.
  - **OpenAI (`openai` + `openai-codex`):** latest general GPT (`gpt-5.x`, `gpt-6.x`, and later majors) *if* its version >= latest Codex, then latest Codex
    - GPT-5.6 and GPT-6 have named tiers and no bare alias. Same-version tier order is Sol, Astra, Terra, Luna, so GPT-5.6 orders as `gpt-5.6-sol` → `gpt-5.6-terra` → `gpt-5.6-luna`.
    - `gpt-6.1-sol` scores as `61`; `gpt-6-sol` scores as `60`; `gpt-5.6-sol` scores as `56`; `gpt-5.5` scores as `55`; `gpt-5.3-codex` scores as `53`.
    - Major-only GPT-5 IDs are also handled (`gpt-5`, `gpt-5-pro`, `gpt-5-codex`).
    - Where a plain ID exists, it beats same-version suffixed variants (`gpt-5.5` before `gpt-5.5-pro`, `gpt-5` before `gpt-5-pro`).
  - **Google (API key):** latest `gemini-*-pro*` (regex: `/^gemini-.*-pro/i`)
  - **Google OAuth providers (`google-gemini-cli`, `google-antigravity`):** prefer stable Gemini before previews

  The ordering logic is driven by:
  - `providerPriority()` (Anthropic → OpenAI Codex → OpenAI → Google → …)
  - `familyPriority()` / `openAiFamilyPriority()` (Opus/Sonnet/Haiku, GPT vs Codex, etc.)
  - `openAiVariantPriority()` (explicit same-version GPT tier order)
  - `parseLeadingMajorMinor()` + `modelRecencyScore()` (treats `4-6` / `4.6` as `46`, `5.6` as `56`, keeps embedded date suffixes such as `YYYYMMDD` separate, and ignores later date-like suffixes such as `gpt-4o-2024-11-20` or `gemini-2.5-pro-preview-06-05` when extracting the family version)
  - `compareModels()` (provider + family + recency tie-breaks; deterministic sorting)

  UI: the model picker is opened from the footer status bar (click the π model button).

- Pick the default model via provider-aware rules:
  - Anthropic is a small special-case: latest Opus by default while Fable is in the registry but unavailable; Sonnet and Fable remain fallbacks if Opus is absent.
  - OpenAI (`openai` + `openai-codex`) prefers the newest general GPT when it is at least as new as Codex, with Codex as fallback; the current default is GPT-6.1 Sol
  - otherwise `DEFAULT_MODEL_RULES` filters and sorts the injected `Models` catalogue by recency

- Populate thinking controls from `getSupportedThinkingLevels()` instead of provider-specific hardcoded lists. This keeps model-level maps authoritative and ensures `xhigh` and `max` remain distinct.

When new models ship, this usually “just works” as long as naming stays consistent. You only need to update these rules if:
- a provider changes their naming scheme, or
- you want different provider/family preferences.

Reminder: **`openai-codex` is NOT `openai`** (different base URL). See `src/auth/provider-map.ts`.

### 6) Run it in Excel (dev vs build)

**Important:** our `manifest.xml` currently points at the **dev server**:

- `https://localhost:3141/src/taskpane.html`

That means:
- `npm run build` is a *sanity check* (TypeScript + bundling), but it does **not** change what Excel loads.
- To test changes in Excel, you need a dev server running on **port 3141**.

Recommended local loop:

```bash
# 1) Start dev server (must be :3141 because manifest hardcodes it)
npm run dev

# 2) (Re)register / launch Excel with the add-in
npm run sideload
```

If `npm run dev` fails with “Port 3141 is already in use” (the config sets
`strictPort`, so Vite exits rather than silently picking another port),
**stop the old server** — Excel will keep loading whatever is on
`https://localhost:3141/`.

```bash
lsof -nP -iTCP:3141 -sTCP:LISTEN
# then kill the PID, or just stop the process in the terminal running it
```

#### Sideload troubleshooting

If `npm run sideload` fails with `EEXIST: file already exists, link 'manifest.xml' -> ...`:

```bash
npx office-addin-debugging stop manifest.xml desktop
npm run sideload
```

#### “I updated models but they don’t show up” checklist

1) **Provider filter:** the model picker only shows models for connected providers (saved API key/OAuth, keyless configured custom providers, or connected extension providers).
2) **Catalogue:** built-in providers add pi.dev models on startup and when the picker opens; check the console for `[models]` warnings and the `/model-catalog/<version>/<provider>` request. Manual-form gateways expose only configured models. Other discovery-enabled registrations request `<baseUrl>/models` and retain their baseline when discovery fails.
3) **Excel caching:** quit Excel completely (Cmd+Q) and reopen.
4) **Hot reload note:** taskpane JS/CSS is served from Vite; edits to model-selection files (`src/models/model-ordering.ts`, `src/models/featured-models.ts`, `src/taskpane/default-model.ts`) should apply via HMR without needing to re-sideload, as long as Excel is pointed at the same running dev server.
5) **Vite optimized deps:** after dependency bumps, clear and restart:

```bash
rm -rf node_modules/.vite
npm run dev
```
