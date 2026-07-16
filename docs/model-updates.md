# Model / dependency update playbook

**Last verified:** 2026-07-16

This repo hardcodes a small set of "featured" and "preferred" model patterns for sorting and default selection. Static built-in models come from Pi AI, while custom and extension providers can add cached, dynamically discovered catalogues at runtime.

This doc describes how to update:
- the **Pi dependency versions** we ship (`@earendil-works/pi-ai`, `@earendil-works/pi-agent-core`)
- the **model ordering/default-selection behavior** in the add-in (`src/models/model-ordering.ts`, `src/models/featured-models.ts`, `src/taskpane/default-model.ts`)
- the **thinking-level UI** that reflects registry capabilities (`src/models/thinking-levels.ts`, `src/taskpane/thinking-display.ts`)
- the browser-native runtime and dynamic catalogue path (`src/models/browser-model-runtime.ts`, `src/storage/local/model-catalogs-store.ts`)

## Sources of truth

- **Built-in model IDs:** `node_modules/@earendil-works/pi-ai/dist/models.generated.js`, exposed through `builtinProviders()`.
- **Runtime lookup and streaming:** the taskpane-owned `BrowserModelRuntime`, backed by Pi AI's `createModels()` collection.
- **Dynamic catalogue cache:** IndexedDB store `model-catalogs`, accessed through `ModelCatalogsStore`; restored entries are rebound to the provider's current API/base URL before use.
- **Discovery bounds:** responses are limited to 2 MiB, 2,000 entries and 256 characters per model ID. An invalid/oversized refresh leaves the last safe cache and configured baseline intact.
- **Custom gateways:** baseline models remain in `CustomProvidersStore`; `/models` discovery overlays them without deleting the configured fallback model.
- **Extension providers:** `api.models.registerProvider()` declarations are runtime-owned and unload with their extension. Unregistering aborts in-flight discovery before deleting its cache so late responses cannot resurrect stale entries.

Do not use Pi coding-agent's Node/file `ModelRuntime` directly in the Office WebView. Pi for Excel uses the same Pi AI provider primitives with browser storage, OAuth and proxy policy. Cross-check the installed Pi package and changelog when the generated registry changes. Never infer aliases or metadata from marketing names.

### Current GPT-5.6 registry snapshot (`pi-ai` 0.80.8)

Upstream exposes exactly three IDs on both `openai` and `openai-codex`; there is deliberately no bare `gpt-5.6` alias:

| ID | Display name | Standard input / output | Cache read / write | Above 272k input / output | Above 272k cache read / write |
|---|---|---:|---:|---:|---:|
| `gpt-5.6-sol` | GPT-5.6 Sol | $5 / $30 | $0.50 / $6.25 | $10 / $45 | $1 / $12.50 |
| `gpt-5.6-terra` | GPT-5.6 Terra | $2.50 / $15 | $0.25 / $3.125 | $5 / $22.50 | $0.50 / $6.25 |
| `gpt-5.6-luna` | GPT-5.6 Luna | $1 / $6 | $0.10 / $1.25 | $2 / $9 | $0.20 / $2.50 |

Prices are registry values in USD per million tokens. All three models:

- accept text and image input, support reasoning, and expose `off`, `minimal`, `low`, `medium`, `high`, `xhigh`, and `max` through `getSupportedThinkingLevels()`
- have a 128,000-token maximum output
- use a 272,000-token context window through the OpenAI API (`openai-responses`, `https://api.openai.com/v1`)
- use a 372,000-token context window through ChatGPT (`openai-codex-responses`, `https://chatgpt.com/backend-api`)
- preserve distinct native `xhigh` and `max` efforts; the ChatGPT map also maps the add-in's `minimal` level to upstream `low`

`tests/model-ordering.test.ts` pins this metadata so a future registry change is reviewed rather than silently changing the UI/runtime contract.

## When to run this

- If you want to add newly-released models and they’re missing from our add-in.
- **If it’s been > 1 week since the last verification date above**, refresh deps + re-check model IDs.

## What is now automated

- Dependabot checks npm dependencies **daily**.
- A dedicated Dependabot group (`pi-stack`) keeps these packages in one PR:
  - `@earendil-works/pi-ai`
  - `@earendil-works/pi-agent-core`
- `.github/workflows/dependabot-pi-automerge.yml` auto-approves + enables auto-merge for that Dependabot group (merge still waits for green checks).
- `npm run check` includes `scripts/check-pi-deps-lockstep.mjs`, which enforces the dependency policy below (`pi-ai` === `pi-agent-core`, exact pins, single shared `pi-ai` copy in the lockfile).

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

- `@earendil-works/pi-ai` and `@earendil-works/pi-agent-core` move **together** to the newest common version. The browser runtime migration was verified on `0.80.8`.
- Both must be **exact-pinned** in `package.json`.
- The lockfile must resolve **exactly one** copy of `pi-ai`. Two copies = two model registries = the model selector and the app disagree about available models.
- `scripts/check-pi-deps-lockstep.mjs` (run via `npm run check:pi-lockstep`) enforces all of this.
- `@earendil-works/pi-web-ui` was **removed entirely** on 2026-07-07 (the UI layer is first-party; see `docs/ui-ownership.md`). Do not re-add it.
- Check the pi changelog for breaking OAuth/streaming surface changes when bumping (e.g. 0.79.x made `onDeviceCode` / `onSelect` required in `OAuthLoginCallbacks`).

### 3) Bump dependencies in `package.json`

```bash
npm install @earendil-works/pi-ai@<version> @earendil-works/pi-agent-core@<version> --save-exact
```

### 4) Verify the new model IDs exist in the registry

Search the generated catalogue and query the runtime-facing built-in provider collection:

```bash
rg -n "gpt-5\\.6-(sol|terra|luna)" node_modules/@earendil-works/pi-ai/dist/models.generated.js -S
rg -n "claude-fable-5"             node_modules/@earendil-works/pi-ai/dist/models.generated.js -S
rg -n "claude-opus-4-8"  node_modules/@earendil-works/pi-ai/dist/models.generated.js -S
rg -n "gemini-3\\.1-pro-preview" node_modules/@earendil-works/pi-ai/dist/models.generated.js -S
node --input-type=module -e 'import { builtinModels } from "@earendil-works/pi-ai/providers/all"; const m=builtinModels(); console.log(m.getModel("openai", "gpt-5.6-sol"))'
npm run test:models
```

If an ID doesn’t appear there, **don’t** add it to the add-in yet—either:
- bump `@earendil-works/pi-ai` further, or
- use an older/fallback ID, or
- configure a custom gateway baseline model in Pi for Excel.

### 5) Update model ordering + default selection logic (avoid hardcoding exact IDs)

Files:
- `src/models/model-ordering.ts` (provider/family priority + version/recency scoring)
- `src/models/featured-models.ts` (featured-model ordering used by the model picker)
- `src/taskpane/default-model.ts` (default-model selection rules)
- `src/models/thinking-levels.ts` (registry-driven thinking capabilities)
- `src/taskpane/thinking-display.ts` (localized thinking labels/hints/colors)
- `src/ui/model-selector-dialog.ts` (first-party model picker UI)
- `tests/model-ordering.test.ts` (metadata + behavior regressions; run `npm run test:models` — requires Node 22.19+)

We intentionally avoid pinning exact versioned IDs now. Instead we:

- In the model picker, show:
  1) current model first
  2) **featured models** (pattern-based “latest” picks)
  3) then the rest sorted deterministically

  Featured rules (current desired behavior):
  - **Anthropic:** latest **Fable** first (post-4.x flagship family, e.g. `claude-fable-5`), then latest **Sonnet** *if* its version >= latest **Opus**, then latest **Opus**
    - This is picker ordering only; default selection currently skips Fable because it is in the registry but unavailable for normal Anthropic use.
    - Version compare uses `parseMajorMinor()` where `claude-opus-4-6` → `46`, `claude-opus-4-7` → `47`, `claude-fable-5` → `50`.
    - Important: IDs like `claude-opus-4-20250514` are treated as **major only** (`40`) and the `YYYYMMDD` part is considered a separate date suffix by `modelRecencyScore()`.
  - **OpenAI (`openai` + `openai-codex`):** latest general `gpt-5.x` *if* its version >= latest `gpt-5.x-codex`, then latest Codex
    - GPT-5.6 has three explicit tiers and no bare alias. Same-version ordering is `gpt-5.6-sol` → `gpt-5.6-terra` → `gpt-5.6-luna`.
    - `gpt-5.6-sol` scores as `56`; `gpt-5.5` scores as `55`; `gpt-5.3-codex` scores as `53`.
    - Major-only GPT-5 IDs are also handled (`gpt-5`, `gpt-5-pro`, `gpt-5-codex`).
    - Where a plain ID exists, it beats same-version suffixed variants (`gpt-5.5` before `gpt-5.5-pro`, `gpt-5` before `gpt-5-pro`).
  - **Google (API key):** latest `gemini-*-pro*` (regex: `/^gemini-.*-pro/i`)
  - **Google OAuth providers (`google-gemini-cli`, `google-antigravity`):** prefer stable Gemini before previews

  The ordering logic is driven by:
  - `providerPriority()` (Anthropic → OpenAI Codex → OpenAI → Google → …)
  - `familyPriority()` / `openAiFamilyPriority()` (Opus/Sonnet/Haiku, GPT vs Codex, etc.)
  - `openAiVariantPriority()` (explicit same-version GPT-5.6 tier order)
  - `parseMajorMinor()` + `modelRecencyScore()` (treats `4-6` / `4.6` as `46`, `5.6` as `56`, keeps embedded date suffixes such as `YYYYMMDD` separate, and ignores later date-like suffixes such as `gpt-4o-2024-11-20` or `gemini-2.5-pro-preview-06-05` when extracting the family version)
  - `compareModels()` (provider + family + recency tie-breaks; deterministic sorting)

  UI: the model picker is opened from the footer status bar (click the π model button).

- Pick the default model via provider-aware rules:
  - Anthropic is a small special-case: latest Opus by default while Fable is in the registry but unavailable; Sonnet and Fable remain fallbacks if Opus is absent.
  - OpenAI (`openai` + `openai-codex`) prefers the newest general GPT-5 when it is at least as new as Codex, with Codex as fallback; current GPT-5.6 default is Sol
  - otherwise `DEFAULT_MODEL_RULES` + `pickLatestMatchingModel()` (uses the injected `Models` runtime to find the newest available ID)

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
2) **Dynamic catalogue:** custom OpenAI-compatible gateways request `<baseUrl>/models` in the background. The configured model remains as a baseline if discovery is unsupported. Check the local proxy when CORS blocks discovery.
3) **Excel caching:** quit Excel completely (Cmd+Q) and reopen.
4) **Hot reload note:** taskpane JS/CSS is served from Vite; edits to model-selection files (`src/models/model-ordering.ts`, `src/models/featured-models.ts`, `src/taskpane/default-model.ts`) should apply via HMR without needing to re-sideload, as long as Excel is pointed at the same running dev server.
5) **Vite optimized deps:** after dependency bumps, clear and restart:

```bash
rm -rf node_modules/.vite
npm run dev
```

### 7) Update this doc’s date

Bump `Last verified:` at the top to today’s date when you finish.
