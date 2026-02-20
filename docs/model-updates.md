# Model / dependency update playbook

**Last verified:** 2026-02-20

This repo hardcodes a small set of "featured" and "preferred" model IDs (for sorting + default selection). Those IDs come from Pi’s model registry (`@mariozechner/pi-ai`) and will drift as new models ship (e.g. `gpt-5.3-codex`, `claude-opus-4-6`).

This doc describes how to update:
- the **Pi dependency versions** we ship (`@mariozechner/pi-ai`, `@mariozechner/pi-web-ui`, `@mariozechner/pi-agent-core`)
- the **model ordering/default-selection behavior** in the add-in (`src/models/model-ordering.ts`, `src/taskpane/default-model.ts`, `src/compat/model-selector-patch.ts`)

## Source of truth

- **Built-in model IDs:** `node_modules/@mariozechner/pi-ai/dist/models.generated.js`
  - This file is auto-generated upstream and is what `getModel(provider, id)` resolves against.
- Don’t rely on Pi’s `docs/models.md` for built-in IDs — that doc is about **custom models** via `~/.pi/agent/models.json`.

## When to run this

- If you want to add newly-released models and they’re missing from our add-in.
- **If it’s been > 1 week since the last verification date above**, refresh deps + re-check model IDs.

## Step-by-step

### 1) Check current installed versions

```bash
node -p "require('./node_modules/@mariozechner/pi-ai/package.json').version"
node -p "require('./node_modules/@mariozechner/pi-web-ui/package.json').version"
node -p "require('./node_modules/@mariozechner/pi-agent-core/package.json').version"
```

### 2) Check latest published versions

```bash
npm view @mariozechner/pi-ai version
npm view @mariozechner/pi-web-ui version
npm view @mariozechner/pi-agent-core version
```

### 3) Bump dependencies in `package.json`

Update these to the same latest version (keep them in lockstep unless you *know* otherwise):
- `@mariozechner/pi-ai`
- `@mariozechner/pi-web-ui`
- `@mariozechner/pi-agent-core`

Then:

```bash
npm install
```

### 4) Verify the new model IDs exist in the registry

Search the local registry:

```bash
rg -n "gpt-5\\.3-codex" node_modules/@mariozechner/pi-ai/dist/models.generated.js -S
rg -n "claude-opus-4-6"  node_modules/@mariozechner/pi-ai/dist/models.generated.js -S
```

If an ID doesn’t appear there, **don’t** add it to the add-in yet—either:
- bump `@mariozechner/pi-ai` further, or
- use an older/fallback ID, or
- define a custom model via `~/.pi/agent/models.json`.

### 5) Update model ordering + default selection logic (avoid hardcoding exact IDs)

Files:
- `src/models/model-ordering.ts` (provider/family priority + version/recency scoring)
- `src/taskpane/default-model.ts` (default-model selection rules)
- `src/compat/model-selector-patch.ts` (ModelSelector ordering/featured-model behavior)
- `tests/model-ordering.test.ts` (sanity tests; run `npm run test:models` — requires Node 22+)

We intentionally avoid pinning exact versioned IDs now. Instead we:

- In the model picker, show:
  1) current model first
  2) **featured models** (pattern-based “latest” picks)
  3) then the rest sorted deterministically

  Featured rules (current desired behavior):
  - **Anthropic:** latest **Sonnet** *if* its version > latest **Opus**, then latest **Opus**
    - Version compare uses `parseMajorMinor()` where `claude-opus-4-5` → `45`, `claude-opus-4-6` → `46`.
    - Important: IDs like `claude-opus-4-20250514` are treated as **major only** (`40`) and the `YYYYMMDD` part is considered a separate date suffix by `modelRecencyScore()`.
  - **OpenAI Codex:** latest `gpt-5.x-codex`, then latest `gpt-5.x`
    - `gpt-5.3-codex` scores as `53`.
  - **Google (API key):** latest `gemini-*-pro*` (regex: `/^gemini-.*-pro/i`)
  - **Google OAuth providers (`google-gemini-cli`, `google-antigravity`):** prefer stable Gemini before previews

  The ordering logic is driven by:
  - `providerPriority()` (Anthropic → OpenAI Codex → OpenAI → Google → …)
  - `familyPriority()` (Opus/Sonnet/Haiku, Codex vs non-Codex, etc.)
  - `parseMajorMinor()` + `modelRecencyScore()` (treats `4-6` as `46`, `5.3` as `53`, and keeps `YYYYMMDD` as a separate date suffix)
  - `compareModels()` (provider + family + recency tie-breaks; deterministic sorting)

  UI: the model picker is opened from the footer status bar (click the π model button).

- Pick the default model via pattern rules:
  - Anthropic is a small special-case (Opus unless a newer-version Sonnet exists; version compare uses `parseMajorMinor`)
  - otherwise `DEFAULT_MODEL_RULES` + `pickLatestMatchingModel()` (uses `getModels(provider)` to find the newest available ID)

When new models ship, this usually “just works” as long as naming stays consistent. You only need to update these rules if:
- a provider changes their naming scheme, or
- you want different provider/family preferences.

Reminder: **`openai-codex` is NOT `openai`** (different base URL). See `src/auth/provider-map.ts`.

### 6) Run it in Excel (dev vs build)

**Important:** our `manifest.xml` currently points at the **dev server**:

- `https://localhost:3000/src/taskpane.html`

That means:
- `npm run build` is a *sanity check* (TypeScript + bundling), but it does **not** change what Excel loads.
- To test changes in Excel, you need a dev server running on **port 3000**.

Recommended local loop:

```bash
# 1) Start dev server (must be :3000 because manifest hardcodes it)
npm run dev

# 2) (Re)register / launch Excel with the add-in
npm run sideload
```

If `npm run dev` says “Port 3000 is in use, trying another one…”, **stop the old server**.
Excel will keep loading whatever is on `https://localhost:3000/`.

```bash
lsof -nP -iTCP:3000 -sTCP:LISTEN
# then kill the PID, or just stop the process in the terminal running it
```

#### Sideload troubleshooting

If `npm run sideload` fails with `EEXIST: file already exists, link 'manifest.xml' -> ...`:

```bash
npx office-addin-debugging stop manifest.xml desktop
npm run sideload
```

#### “I updated models but they don’t show up” checklist

1) **Provider filter:** the model picker only shows models for **connected providers** (saved API key/OAuth). Make sure the provider is connected.
2) **Excel caching:** quit Excel completely (Cmd+Q) and reopen.
3) **Hot reload note:** taskpane JS/CSS is served from Vite; edits to model-selection files (`src/models/model-ordering.ts`, `src/taskpane/default-model.ts`, `src/compat/model-selector-patch.ts`) should apply via HMR without needing to re-sideload, as long as Excel is pointed at the same running dev server.
4) **Vite optimized deps:** after dependency bumps, clear and restart:

```bash
rm -rf node_modules/.vite
npm run dev
```

### 7) Update this doc’s date

Bump `Last verified:` at the top to today’s date when you finish.
