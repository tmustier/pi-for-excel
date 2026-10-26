# Verification: setup

How to run this file: bring up Pi in Excel on the web with `node scripts/orb/excel-web-session.mjs --check` (see [docs/orb-excel-web.md](../../docs/orb-excel-web.md)). It starts logged in to OpenAI (ChatGPT) through the shared `PI_EXCEL_WEB_AUTH_JSON` credential. Never log out of that credential, disconnect it, or make Pi refresh it: refreshing rotates the refresh token and logs out every other orb. For items that need no provider at all, use the standalone taskpane (`https://localhost:3141/src/taskpane.html` after `amp orb services ensure`) in a fresh browser profile. Block `/__pi-auth` there (for example `agent-browser network route "**/__pi-auth*" --body '{}'`), because the dev server otherwise hands every page the same credential. API-key and gateway items use made-up keys and endpoints. Keep model turns few and tiny; they share one rate limit. Before each section: close Settings and the picker, keep "Fork model switch into new tab" off, and keep Settings → Proxy on at `https://localhost:3003` unless the item says otherwise. Restore both afterwards.

Device column: `web` is the Excel on the web session; `standalone` is the standalone taskpane in a fresh profile with no provider; `desktop` is Excel desktop (not available in an orb).

## foundations/models-and-providers.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| MODEL-01 | P1 | web | A new tab starts on the active tab's model, with thinking "High" for a reasoning model ([How a tab gets its model](../foundations/models-and-providers.md#how-a-tab-gets-its-model)). | Active tab on GPT-6.1 Sol. | 1. Click "+". | The new tab shows "GPT-6.1 Sol" and "High". | pass |
| MODEL-02 | P2 | web | A restored tab gets the model saved with its session ([How a tab gets its model](../foundations/models-and-providers.md#how-a-tab-gets-its-model)). | A tab switched to a non-default model. | 1. Reload the taskpane. | The tab comes back on the chosen model. | blocked: reloading the Excel session repeats credential restore; not risked. |
| MODEL-03 | P1 | standalone | Saving a key does not test it; a wrong key fails on the next turn ([Connecting a provider](../foundations/models-and-providers.md#connecting-a-provider)). | No provider. | 1. Save `sk-fake-docs-test-key` for OpenAI (API).<br>2. Send "hello". | The key is accepted; the turn fails with "LLM error: OpenAI API error (401): …". | pass |
| MODEL-04 | P2 | web | The provider rows appear in the documented order ([Connecting a provider](../foundations/models-and-providers.md#connecting-a-provider)). | — | 1. `/login`. | Anthropic, OpenAI (ChatGPT), Google Code Assist, Google Antigravity, GitHub Copilot, OpenAI (API), Google Gemini (API), DeepSeek, Amazon Bedrock, Mistral, Groq, xAI / Grok. | pass |
| MODEL-05 | P1 | web | OpenAI (ChatGPT) model requests need the proxy ([Connecting a provider](../foundations/models-and-providers.md#connecting-a-provider)). | Proxy switched off. | 1. Send "hi" on GPT-6.1 Sol. | Record what happens. | fail: the turn succeeded with the proxy off. Document was wrong and has been corrected; desktop not tried. Same run as PROXY-18. |
| MODEL-06 | P2 | web | "Disconnect {label}" removes the credential with the toast "{label} disconnected" ([Where credentials live](../foundations/models-and-providers.md#where-credentials-live)). | Google Gemini (API) connected with a made-up key. | 1. `/login`, expand the row.<br>2. Click "Disconnect Google Gemini (API)". | Toast "Google Gemini (API) disconnected"; the row shows "set up →". | pass |
| MODEL-07 | P1 | web | With only OpenAI (ChatGPT) connected, the default model is its newest general GPT ([The default model](../foundations/models-and-providers.md#the-default-model)). | — | 1. Disconnect a second provider that a tab uses (MODEL-11). | The tab moves to "GPT-6.1 Sol". | pass |
| MODEL-08 | P1 | web | With the fork setting on, choosing a model in an empty tab still forks (suspected bug) ([Changing the model](../foundations/models-and-providers.md#changing-the-model)). | Fork setting on; empty "Chat 1". | 1. Pick "gpt-6-luna". | Record what happens. | pass: suspected bug confirmed. "Chat 1 (gpt-6-luna)" opened with "Opened Chat 1 (gpt-6-luna) in a new tab". |
| MODEL-09 | P1 | web | Choosing a different model while busy is refused ([Changing the model](../foundations/models-and-providers.md#changing-the-model)). | A turn responding. | 1. Click the model control.<br>2. Choose another model. | Toast "Wait for this tab to finish before changing models"; the model is unchanged. | pass |
| MODEL-10 | P1 | web | An in-place switch keeps the thinking level even when the new model does not reason ([Changing the model](../foundations/models-and-providers.md#changing-the-model)). | Tab on "High"; a gateway with an unknown model. | 1. Pick the gateway model. | Status bar shows the gateway model with "High". | pass |
| MODEL-11 | P1 | web | An idle tab whose provider is disconnected moves to the default model silently ([When a provider is disconnected](../foundations/models-and-providers.md#when-a-provider-is-disconnected)). | Idle tab on Gemini 2.5 Flash-Lite. | 1. Disconnect Google Gemini (API).<br>2. Close Settings. | The tab shows "GPT-6.1 Sol" and "High"; no toast about the model. | pass |
| MODEL-12 | P2 | web | A busy tab is not swapped until it is idle ([When a provider is disconnected](../foundations/models-and-providers.md#when-a-provider-is-disconnected)). | A turn on a gateway model. | 1. Delete the gateway while the turn runs. | The tab keeps its model until the turn ends. | blocked: not run (would need a working gateway turn). |
| MODEL-13 | P2 | standalone | With no provider, the status bar names GPT-6.1 Sol and the first send asks for OpenAI (API) ([The default model](../foundations/models-and-providers.md#the-default-model)). | No provider; welcome dismissed. | 1. Send "hi". | "Connect OpenAI (API)" opens. | pass |
| MODEL-14 | P2 | web | Switching to a smaller model changes the context meter at once ([Context window per model](../foundations/models-and-providers.md#context-window-per-model)). | Tab with some history. | 1. Pick a 16k gateway model. | The meter jumps (here to "60% / 16k"). | pass |
| MODEL-15 | P1 | web | An OAuth login is refreshed only when the taskpane opens (suspected bug) ([Open questions](../foundations/models-and-providers.md#open-questions-and-verification)). | — | — | — | blocked: refreshing the shared credential is not allowed. |
| MODEL-16 | P2 | web | The same model from two providers looks identical in the status bar ([Edge cases](../foundations/models-and-providers.md#edge-cases)). | OpenAI (API) and OpenAI (ChatGPT) connected. | — | — | blocked: only one GPT provider available. |
| MODEL-17 | P3 | web | An unknown gateway model defaults to 16,384 context and 4,096 output ([Context window per model](../foundations/models-and-providers.md#context-window-per-model)). | — | 1. Save a gateway with blank limits. | Card shows "Max context: 16,384 tokens · Max output: 4,096 tokens". | pass |
| MODEL-18 | P1 | web | With OpenAI (API) and OpenAI (ChatGPT) both connected, the default is the API one (product call) ([The default model](../foundations/models-and-providers.md#the-default-model)). | — | — | — | blocked: the default is only observable on a fresh taskpane or a swap; not set up. |
| MODEL-19 | P2 | web | An automatic swap is not saved until the session next saves ([When a provider is disconnected](../foundations/models-and-providers.md#when-a-provider-is-disconnected)). | — | — | — | blocked: needs a reload. |
| MODEL-20 | P1 | web | Renaming a gateway orphans tabs that use it (suspected bug) ([Edge cases](../foundations/models-and-providers.md#edge-cases)). | See GATEWAY-14. | — | Record what happens. | pass: suspected bug confirmed, see GATEWAY-14. |

Not checkable by hand:

- Whether subscription users expect the API-key model as default when both OpenAI providers are connected ([Open questions](../foundations/models-and-providers.md#open-questions-and-verification)). A product call.
- The order in which other providers are tried for the default model.

## setup/welcome-and-login.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| LOGIN-01 | P1 | standalone | The welcome overlay appears when no provider is connected after startup ([Arriving](../setup/welcome-and-login.md#arriving-at-the-welcome-overlay)). | Fresh profile, `/__pi-auth` blocked. | 1. Open the taskpane; wait 15 s. | The "Pi for Excel" overlay covers the empty tab. | pass (about 13 s after load) |
| LOGIN-02 | P2 | standalone | The card shows the language switch, title, subtitle, introduction, provider rows, gateway link and proxy toggle ([Arriving](../setup/welcome-and-login.md#arriving-at-the-welcome-overlay)). | As LOGIN-01. | 1. Read the card. | All present, in that order. | pass |
| LOGIN-03 | P1 | standalone | Escape dismisses the overlay and stores nothing ([Ends at once](../setup/welcome-and-login.md#ends-at-once)). | As LOGIN-01. | 1. Press Escape.<br>2. Press Escape again. | Record how many presses close it and where focus lands. | pass after correction: focus is in the composer behind the overlay, so the first Escape only leaves it; the second closes the overlay; focus returns to the composer. Document updated. |
| LOGIN-04 | P2 | standalone | A click on the backdrop dismisses the overlay ([Ends at once](../setup/welcome-and-login.md#ends-at-once)). | As LOGIN-01. | 1. Click outside the card. | The overlay closes. | pass (seen when a scripted click landed outside the card) |
| LOGIN-05 | P2 | standalone | The welcome overlay turns the proxy on when a local proxy answers ([Arriving](../setup/welcome-and-login.md#arriving-at-the-welcome-overlay)). | Proxy service running. | 1. Open "Having login trouble? Configure local proxy". | "Enabled" is ticked. | pass |
| LOGIN-06 | P2 | standalone | Disconnecting the last provider does not bring the overlay back ([Arriving](../setup/welcome-and-login.md#arriving-at-the-welcome-overlay)). | One API key connected. | 1. Disconnect it in Settings. | No welcome overlay. | blocked: not run |
| LOGIN-07 | P2 | standalone | A slow credential restore closes the overlay by itself ([Edge cases](../setup/welcome-and-login.md#edge-cases)). | — | — | — | blocked: cannot stage a slow restore |
| LOGIN-08 | P2 | standalone | Clicking a row expands it and focuses the key field ([Expanding a provider row](../setup/welcome-and-login.md#expanding-a-provider-row)). | As LOGIN-01. | 1. Click "Anthropic". | "Login with Anthropic", "or enter API key", key field and "Save"; focus in the key field. | pass |
| LOGIN-09 | P3 | standalone | Key placeholders depend on the provider ([Expanding a provider row](../setup/welcome-and-login.md#expanding-a-provider-row)). | — | 1. Expand Anthropic, then OpenAI (API). | "sk-ant-api… or sk-ant-oat…", "Enter API key". | pass |
| LOGIN-10 | P2 | standalone | "Save" with an empty field does nothing ([Ends at once](../setup/welcome-and-login.md#ends-at-once)). | Anthropic expanded. | 1. Press Enter in the empty field. | No error, nothing stored. | pass |
| LOGIN-11 | P2 | standalone | An Anthropic `code#state` is refused with the OAuth-code message ([Ends at once](../setup/welcome-and-login.md#ends-at-once)). | Anthropic expanded. | 1. Type `abc123#state456`, Enter. | "That looks like an OAuth authorization code (code#state). …"; row stays open. | pass |
| LOGIN-12 | P1 | standalone | "Testing…" does not test the key (suspected bug) ([Committing a credential](../setup/welcome-and-login.md#committing-a-credential)). | — | See MODEL-03. | Record what happens. | pass: suspected bug confirmed |
| LOGIN-13 | P2 | web | A redirect URL pasted into Google Gemini (API) is refused ([Ends at once](../setup/welcome-and-login.md#ends-at-once)). | `/login`, Google Gemini (API) expanded. | 1. Type `https://accounts.google.com/o?code=abc`, Enter. | "That looks like an OAuth redirect URL/code. Use Google API key auth here, …". | pass |
| LOGIN-14 | P2 | web | A leading "Bearer " is removed silently ([Ends at once](../setup/welcome-and-login.md#ends-at-once)). | As LOGIN-13. | 1. Type `Bearer AIzaFAKE…`, Enter. | Saved and connected; no error. | pass (stored value not inspected) |
| LOGIN-15 | P2 | standalone | A key saved on the welcome overlay shows the welcome toast, closes the overlay and focuses the composer ([Connected, or failed](../setup/welcome-and-login.md#connected-or-failed)). | As LOGIN-01. | 1. Save a made-up OpenAI (API) key. | "OpenAI (API) connected — try "Explain this workbook""; overlay closed; focus in the composer. | pass |
| LOGIN-16 | P2 | web | A key saved in Settings shows "{label} connected" and the page stays open ([Connected, or failed](../setup/welcome-and-login.md#connected-or-failed)). | `/login`. | 1. Save a Google Gemini (API) key. | "Google Gemini (API) connected"; "Model providers" still open. | pass |
| LOGIN-17 | P2 | web | Disconnect toast ([Connected, or failed](../setup/welcome-and-login.md#connected-or-failed)). | See MODEL-06. | — | "{label} disconnected". | pass |
| LOGIN-18 | P1 | standalone | "Login with {label}" opens the sign-in page and the login prompt ([Committing a credential](../setup/welcome-and-login.md#committing-a-credential)). | Anthropic expanded. | 1. Click "Login with Anthropic". | A new browser tab on claude.ai; dialog "Login with Anthropic" with the helper text, "Open login page", "Copy login link", the URL field, "Cancel" and "Continue"; the button shows progress. | pass (button read "Waiting for authorization code…") |
| LOGIN-19 | P2 | standalone | With the proxy on, the prompt says "Waiting for the browser callback via the local proxy…" ([Committing a credential](../setup/welcome-and-login.md#committing-a-credential)). | As LOGIN-18. | 1. Read the status line. | That text. | pass |
| LOGIN-20 | P1 | standalone | Escape cancels the login prompt; the row returns with no message; the browser tab stays ([Ends at once](../setup/welcome-and-login.md#ends-at-once)). | As LOGIN-18. | 1. Escape twice. | Prompt closed; "Login with Anthropic"; no error; claude.ai tab still open. | pass (first Escape leaves the paste field) |
| LOGIN-21 | P1 | standalone | GitHub Copilot shows the device-code dialog ([Committing a credential](../setup/welcome-and-login.md#committing-a-credential)). | Proxy on and answering. | 1. Expand GitHub Copilot, click "Login with GitHub Copilot". | Device-code dialog with the code. | fail: "Opening login…", then "Login couldn't connect — this provider needs a helper running on your Mac. …" under the row; no window, no dialog. Possible product bug; not traced. |
| LOGIN-22 | P1 | — | The proxy gate appears when the proxy is not reachable ([Committing a credential](../setup/welcome-and-login.md#committing-a-credential)). | Production build. | — | — | blocked: development builds skip the gate |
| LOGIN-23 | P2 | standalone | The proxy panel shows "Local HTTPS proxy", "Enabled", the URL, "Save", the hint, and "Hide proxy settings" ([The proxy panel](../setup/welcome-and-login.md#the-proxy-panel-on-the-welcome-overlay)). | As LOGIN-01. | 1. Click "Having login trouble? Configure local proxy". | As stated. | pass |
| LOGIN-24 | P1 | standalone | The welcome proxy panel saves an unvalidated URL (suspected bug) ([Open questions](../setup/welcome-and-login.md#open-questions-and-verification)). | Panel open. | 1. Enter `http://example.com:3003`, Save. | Record what happens. | pass: suspected bug confirmed ("Proxy settings saved") |
| LOGIN-25 | P2 | standalone | "Use a custom OpenAI-compatible gateway" closes the overlay and opens Settings on the gateway page with the root underneath ([Arriving](../setup/welcome-and-login.md#arriving-at-the-welcome-overlay)). | As LOGIN-01. | 1. Click it.<br>2. Escape, Escape. | Gateway page with Back; Escape goes to "Settings", then closes. | pass |
| LOGIN-26 | P2 | standalone | "English" / "中文" switch language and reload ([Arriving](../setup/welcome-and-login.md#arriving-at-the-welcome-overlay)). | — | — | — | blocked: not run |
| LOGIN-27 | P1 | standalone | The connection dialog opens with the user message in the transcript, the tab busy and the row expanded ([The connection dialog](../setup/welcome-and-login.md#the-connection-dialog)). | No provider, welcome dismissed. | 1. Send "hi". | "Connect OpenAI (API)"; "hi" in the transcript; Stop button shown. | pass |
| LOGIN-28 | P2 | standalone | The connection dialog focuses the key field ([The connection dialog](../setup/welcome-and-login.md#the-connection-dialog)). | As LOGIN-27. | 1. Check focus. | Focus in "Enter API key". | pass |
| LOGIN-29 | P1 | standalone | Connecting in the dialog shows no toast and the turn continues (suspected bug) ([The connection dialog](../setup/welcome-and-login.md#the-connection-dialog)). | As LOGIN-27. | 1. Type a key, Enter. | Record toasts and what the turn does. | pass: suspected bug confirmed. No toast; the turn ran (and failed with a 401 for the made-up key). |
| LOGIN-30 | P2 | standalone | With no provider, the first send asks for OpenAI (API) whatever the user meant ([Edge cases](../setup/welcome-and-login.md#edge-cases)). | See MODEL-13. | — | — | pass |
| LOGIN-31 | P1 | standalone | Dismissing the connection dialog fails the turn; the banner text ([The connection dialog](../setup/welcome-and-login.md#the-connection-dialog)). | As LOGIN-27. | 1. Escape twice. | Record the banner. | fail against [turns](../foundations/turns.md), which says "API key required for {provider}.": the banner is "LLM error: Provider is not configured: openai" and the transcript "Error: Provider is not configured: openai". This document now says so; turns.md needs the same fix. |
| LOGIN-32 | P2 | standalone | A CORS-like OAuth failure shows the helper message with "Step-by-step guide →" ([Connected, or failed](../setup/welcome-and-login.md#connected-or-failed)). | See LOGIN-21. | — | — | pass |
| LOGIN-33 | P3 | standalone | The welcome toast lasts a little over three seconds ([Connected, or failed](../setup/welcome-and-login.md#connected-or-failed)). | — | — | — | blocked: not timed |
| LOGIN-34 | P2 | standalone | "Hide proxy settings" stays English in Chinese ([Interactions](../setup/welcome-and-login.md#interactions-with-other-systems)). | — | — | — | blocked: Chinese not tried |

Not checkable by hand:

- Where the login window opens on Excel desktop (system browser) — needs desktop.
- The device-code dialog's polling after Escape, while LOGIN-21 fails.
- Whether "on your Mac" wording should change on Windows (a copy decision).

## setup/the-model-picker.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| PICKER-01 | P1 | web | The status-bar model control opens "Select model" with focus in the search ([Opening the picker](../setup/the-model-picker.md#opening-the-picker)). | Idle tab. | 1. Click the model control. | Picker open; focus in "Search models…". | pass |
| PICKER-02 | P2 | web | Rows show id, check, provider tag and limits ([Opening the picker](../setup/the-model-picker.md#opening-the-picker)). | — | 1. Read the first row. | "gpt-6.1-sol ✓ openai-codex 500k ctx · 128k out". | pass |
| PICKER-03 | P2 | web | Default order: current, featured, rest ([Opening the picker](../setup/the-model-picker.md#opening-the-picker)). | OpenAI (ChatGPT) and Google Gemini (API) connected. | 1. Read the list. | Current; then gpt-6.1-sol, gpt-5.3-codex-spark, gemini-3.1-pro-preview; then the rest by provider. | pass |
| PICKER-04 | P2 | web | The search is a fuzzy subsequence match and resets the highlight to the first row ([Searching](../setup/the-model-picker.md#searching)). | — | 1. Type "luna", then "gemflash". | Two Luna rows, first highlighted; gemini flash rows. | pass |
| PICKER-05 | P2 | standalone | With no provider the list says "No matching models" ([Edge cases](../setup/the-model-picker.md#edge-cases)). | No provider. | 1. Open the picker. | Only "No matching models". | pass |
| PICKER-06 | P1 | web | Escape leaves the search, then closes; focus returns to the composer ([Ends at once](../setup/the-model-picker.md#ends-at-once)). | Picker open. | 1. Escape.<br>2. Escape. | Open after one, closed after two; focus in the composer. | pass |
| PICKER-07 | P2 | web | Down stops at the last row and Enter chooses the highlight ([Choosing a model](../setup/the-model-picker.md#choosing-a-model)). | Search "luna". | 1. Down twice, Enter. | Second row chosen. | pass |
| PICKER-08 | P1 | web | A pointer resting over the list does not move the highlight until the mouse moves ([Choosing a model](../setup/the-model-picker.md#choosing-a-model)). | Pointer left over the middle of the taskpane by an earlier click. | 1. Type `/default-models`, Enter.<br>2. Enter. | Highlight on the current model; nothing changes. | fail: the row under the pointer ("gpt-5.6-terra") was highlighted and Enter switched to it. Seen with CDP input; repeat with a physical mouse. Document updated. |
| PICKER-09 | P1 | web | Choosing the current model does nothing ([Ends at once](../setup/the-model-picker.md#ends-at-once)). | — | 1. Click the checked row. | Picker closes; no toast; no change. | pass |
| PICKER-10 | P1 | web | An in-place switch updates the status bar with no toast ([Choosing a model](../setup/the-model-picker.md#choosing-a-model)). | Fork off. | 1. Choose "gpt-6.1-sol" from another model. | Status bar "GPT-6.1 Sol"; no toast; same tab. | pass |
| PICKER-11 | P2 | web | × closes the picker ([Ends at once](../setup/the-model-picker.md#ends-at-once)). | — | — | — | blocked: not run |
| PICKER-12 | P1 | web | `/default-models` opens the ordinary picker though it says "Switch models with Ctrl+P" (suspected bug) ([Ctrl+P and /default-models](../setup/the-model-picker.md#ctrlp-and-default-models)). | — | 1. Type `/default-models`; read the menu; Enter. | Record what happens. | pass: suspected bug confirmed |
| PICKER-13 | P1 | web | Fork happens in an empty tab (suspected bug). | See MODEL-08. | — | — | pass: suspected bug confirmed |
| PICKER-14 | P2 | web | A fork opens a new active tab "{title} ({model id})" with the toast "Opened {title} in a new tab" ([Choosing a model](../setup/the-model-picker.md#choosing-a-model)). | Fork on. | 1. Choose another model. | New tab active, toast shown. | pass |
| PICKER-15 | P1 | web | While busy the status-bar control opens the picker but a different choice is refused ([Choosing a model](../setup/the-model-picker.md#choosing-a-model)). | A turn responding. | 1. Click the control.<br>2. Choose another model. | Picker opens; toast "Wait for this tab to finish before changing models". | pass |
| PICKER-16 | P1 | web | `/model` is refused while busy ([Ends at once](../setup/the-model-picker.md#ends-at-once)). | A turn responding. | 1. Type `/model`, Enter. | "Can't run /model while Pi is busy". | pass (`/default-models` while busy not tried) |
| PICKER-17 | P2 | web | `/model` with text ignores the text ([Edge cases](../setup/the-model-picker.md#edge-cases)). | — | 1. `/model luna`, Enter. | Full list, empty search. | pass |
| PICKER-18 | P2 | web | Newly connected providers' models appear in the picker ([Opening the picker](../setup/the-model-picker.md#opening-the-picker)). | Connect Google Gemini (API). | 1. Open the picker. | Gemini rows tagged "google". | pass |
| PICKER-19 | P1 | web | A model swapped by the opening refresh is not checked (suspected bug) ([Open questions](../setup/the-model-picker.md#open-questions-and-verification)). | — | — | Record what happens. | blocked: could not stage the race; not reproduced |
| PICKER-20 | P3 | web | Fork titles stack ([Edge cases](../setup/the-model-picker.md#edge-cases)). | Fork on, in a forked tab. | 1. Choose another model. | "Chat 1 (gpt-6-luna) (gpt-6-sol)". | pass |
| PICKER-21 | P2 | web | Ctrl/Cmd+P does nothing in Pi ([Ctrl+P and /default-models](../setup/the-model-picker.md#ctrlp-and-default-models)). | — | — | — | blocked: would open the browser's print dialog over Excel |

Not checkable by hand:

- Whether the status-bar control should refuse before opening while busy, like `/model` (a product call).

## setup/settings.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| SETTINGS-01 | P2 | web | The root page's groups, rows and values ([The root page](../setup/settings.md#the-root-page)). | — | 1. Utilities menu → "Settings". | Five groups; "1 connected", "None", "On", Plugins "1 installed". | pass |
| SETTINGS-02 | P1 | web | Settings moves focus to the page body ([Opening Settings](../setup/settings.md#opening-settings)). | — | 1. Open Settings; check focus. | Focus on the page body. | pass |
| SETTINGS-03 | P1 | web | Escape goes back one page and closes on the root ([Leaving a page and closing](../setup/settings.md#leaving-a-page-and-closing)). | On "Proxy". | 1. Escape.<br>2. Escape. | Root page, then closed; focus in the composer. | pass |
| SETTINGS-04 | P1 | web | The utilities menu opens Settings while the tab is busy ([Ends at once](../setup/settings.md#ends-at-once)). | A turn responding. | 1. Utilities menu → "Settings". | Settings opens. | pass |
| SETTINGS-05 | P1 | web | `/settings` and `/login` are refused while busy ([Ends at once](../setup/settings.md#ends-at-once)). | A turn responding. | 1. `/settings`, Enter.<br>2. `/login`, Enter. | "Can't run /settings while Pi is busy"; "Can't run /login while Pi is busy". | pass |
| SETTINGS-06 | P2 | web | `/login` opens "Model providers" with the root underneath ([Opening Settings](../setup/settings.md#opening-settings)). | — | 1. `/login`.<br>2. Click Back. | Back visible; Back shows "Settings". | pass |
| SETTINGS-07 | P2 | web | × closes Settings from a page ([Leaving a page and closing](../setup/settings.md#leaving-a-page-and-closing)). | On "Model providers". | 1. "Close settings". | Closed. | pass |
| SETTINGS-08 | P1 | web | Only Rules guards unsaved edits, on Escape and × ([Leaving a page and closing](../setup/settings.md#leaving-a-page-and-closing)). | `/rules`. | 1. Type in "All my files"; Escape twice.<br>2. "Keep editing".<br>3. "Close settings"; "Discard". | "Discard changes?" each time; Keep editing stays; Discard closes. | pass |
| SETTINGS-09 | P1 | web | A half-filled gateway form is dropped silently (suspected missing guard) ([Leaving a page and closing](../setup/settings.md#leaving-a-page-and-closing)). | Gateway page with Name and Endpoint typed. | 1. Back.<br>2. Reopen "Custom gateway". | Record what happens. | pass: no question; form empty |
| SETTINGS-10 | P2 | web | The fork switch toasts ([The root page](../setup/settings.md#the-root-page)). | — | 1. Toggle on, then off. | "Model switching will open a new tab for non-empty chats." / "Model switching will stay in the current tab.". | pass |
| SETTINGS-11 | P2 | web | "Auto mode" toggles execution mode with a toast ([The root page](../setup/settings.md#the-root-page)). | — | — | — | blocked: not run |
| SETTINGS-12 | P2 | web | Changing language reloads after about a second ([The root page](../setup/settings.md#the-root-page)). | — | — | — | blocked: not run (reload) |
| SETTINGS-13 | P2 | web | The "Model providers" count leaves out gateways ([Edge cases](../setup/settings.md#edge-cases)). | Three gateways, one OAuth provider. | 1. Read the root page. | "1 connected" and gateway row "3". | pass |
| SETTINGS-14 | P2 | web | `/settings` while Settings is open closes it ([Opening Settings](../setup/settings.md#opening-settings)). | — | — | — | blocked: the composer and header cannot be reached while Settings is open |
| SETTINGS-15 | P2 | web | Saving Rules returns to the root page ([Leaving a page and closing](../setup/settings.md#leaving-a-page-and-closing)). | — | — | — | blocked: not run (would store rules) |
| SETTINGS-16 | P3 | web | Pages have a title and subtitle ([The pages](../setup/settings.md#the-pages)). | — | 1. Open "Proxy". | Record. | pass: "Proxy" has no subtitle (noted) |

## setup/custom-gateways.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| GATEWAY-01 | P2 | web | Fields and placeholders ([Filling the form](../setup/custom-gateways.md#filling-the-form)). | Gateway page. | 1. Read the form. | Name, Endpoint, Model, two limits with hints, API key (password), Connect directly. | pass |
| GATEWAY-02 | P1 | web | Each validation message ([Ends at once](../setup/custom-gateways.md#ends-at-once)). | — | 1. Save empty; `ftp://x`; `not a url`; no model; context 500; context 2000.5; output 0. | The seven messages in the document, in warning colour, form kept. | pass |
| GATEWAY-03 | P2 | web | Connect directly with a remote `http://` endpoint is refused ([Ends at once](../setup/custom-gateways.md#ends-at-once)). | — | 1. `http://example.com/v1`, Connect directly on, Save. | "Direct connections need an https:// endpoint. …". | pass |
| GATEWAY-04 | P2 | web | The endpoint loses its trailing slash and fragment ([Saving](../setup/custom-gateways.md#saving)). | — | 1. Save `https://llm.example.com/v1/#frag`. | Card "Endpoint: https://llm.example.com/v1". | pass |
| GATEWAY-05 | P2 | web | Output is clamped to context ([Saving](../setup/custom-gateways.md#saving)). | — | 1. Output 99999, context blank. | "Max output: 16,384 tokens". | pass |
| GATEWAY-06 | P2 | web | A blank name becomes host:port ([Filling the form](../setup/custom-gateways.md#filling-the-form)). | — | 1. Save `http://localhost:9090/v1` with no name. | "Gateway · localhost:9090". | pass |
| GATEWAY-07 | P2 | web | A duplicate name gets " (2)" ([Saving](../setup/custom-gateways.md#saving)). | "Warehouse" exists. | 1. Save "warehouse". | "Gateway · warehouse (2)". | pass (toast names it "warehouse" without the suffix) |
| GATEWAY-08 | P2 | web | The card lines ([The simple case](../setup/custom-gateways.md#the-simple-case)). | — | 1. Read a card. | Endpoint, Model, Max context, Max output, API key, Connection. | pass |
| GATEWAY-09 | P2 | web | A keyless gateway is connected and listed in the picker with its tag ([How its model appears](../setup/custom-gateways.md#how-its-model-appears-and-is-used)). | Keyless gateways saved. | 1. Open the picker. | Rows tagged "Gateway · {name}". | pass |
| GATEWAY-10 | P2 | web | Only the configured model is offered ([How its model appears](../setup/custom-gateways.md#how-its-model-appears-and-is-used)). | — | 1. Open the picker. | One row per gateway. | pass |
| GATEWAY-11 | P2 | web | Edit prefills every field, focuses Name and shows Cancel / Update ([Editing and deleting](../setup/custom-gateways.md#editing-and-deleting)). | — | 1. Click "Edit". | As stated; key hidden. | pass (Cancel not clicked) |
| GATEWAY-12 | P2 | web | "Updated gateway {name}." ([Saving](../setup/custom-gateways.md#saving)). | Editing. | 1. Rename, Update. | Toast. | pass |
| GATEWAY-13 | P2 | web | Delete asks, Escape cancels, Delete removes with a toast ([Editing and deleting](../setup/custom-gateways.md#editing-and-deleting)). | — | 1. Delete; Escape.<br>2. Delete; Delete. | "Delete custom gateway?" text; nothing removed; then "Deleted gateway {name}.". | pass |
| GATEWAY-14 | P1 | web | Renaming a gateway orphans the tabs using it (suspected bug) ([Open questions](../setup/custom-gateways.md#open-questions-and-verification)). | Idle tab on "gpt-oss-120b" from "Warehouse". | 1. Edit "Warehouse" to "Depot", Update.<br>2. Close Settings. | Record the tab's model. | pass: suspected bug confirmed; tab moved to "GPT-6.1 Sol" silently |
| GATEWAY-15 | P1 | web | Deleting a gateway moves idle tabs on it to the default model ([Modifiers](../setup/custom-gateways.md#modifiers)). | Idle tab on "dead-model". | 1. Delete the gateway. | Tab on "GPT-6.1 Sol". | pass |
| GATEWAY-16 | P2 | web | An unreachable endpoint fails the turn with "Connection error." ([How its model appears](../setup/custom-gateways.md#how-its-model-appears-and-is-used)). | Proxy off; gateway at `https://localhost:3999/v1`. | 1. Send "hi". | Transcript "Error: Connection error."; banner "LLM error: Connection error.". | pass |
| GATEWAY-17 | P2 | web | A half-filled form is dropped on Back. | See SETTINGS-09. | — | — | pass |
| GATEWAY-18 | P2 | web | A proxy refusal of a loopback target reaches the error banner ([How its model appears](../setup/custom-gateways.md#how-its-model-appears-and-is-used)). | — | — | — | blocked: not run |
| GATEWAY-19 | P3 | web | Enter in a field does not save ([Interactions](../setup/custom-gateways.md#interactions-with-other-systems)). | — | — | — | blocked: not run |

## setup/the-proxy.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| PROXY-01 | P2 | web | The Proxy page: switch, URL, status line, hint ([Changing the setting](../setup/the-proxy.md#changing-the-setting)). | Proxy on and running. | 1. Settings → "Proxy". | "✓ Proxy connected at https://localhost:3003". | pass |
| PROXY-02 | P2 | web | The switch toasts and the "Proxy disabled." status ([Changing the setting](../setup/the-proxy.md#changing-the-setting)). | — | 1. Turn off, then on. | "Proxy disabled." / "Proxy enabled.". | pass |
| PROXY-03 | P2 | web | A new URL is saved with "Proxy URL saved." and loses its trailing slash ([Changing the setting](../setup/the-proxy.md#changing-the-setting)). | — | 1. Enter `https://localhost:3999/`, Enter. | Toast; field shows no slash. | pass |
| PROXY-04 | P2 | web | An `http://` URL is rejected with the mixed-content message and toast ([Changing the setting](../setup/the-proxy.md#changing-the-setting)). | — | 1. Enter `http://example.com`, Enter. | Status line and "Proxy URL not saved: …". | pass |
| PROXY-05 | P2 | web | The root page shows "On" or "Off" ([Interactions](../setup/the-proxy.md#interactions-with-other-systems)). | — | 1. Read the root row. | "On" while on. | pass ("Off" not read) |
| PROXY-06 | P1 | web | The banner appears when a check fails ([The proxy banner](../setup/the-proxy.md#the-proxy-banner)). | URL set to `https://localhost:3999`. | 1. Wait up to 30 s. | Banner "Proxy not running · Some features won't work.". | pass (12 s) |
| PROXY-07 | P1 | web | The banner shows even with the switch off (suspected bug) ([The proxy banner](../setup/the-proxy.md#the-proxy-banner)). | As PROXY-06. | 1. Turn the switch off; close Settings. | Record. | pass: suspected bug confirmed |
| PROXY-08 | P2 | web | "How to fix →" expands the steps and becomes "Hide steps" ([The proxy banner](../setup/the-proxy.md#the-proxy-banner)). | Banner shown. | 1. Click it. | The command, copy button, "Open Terminal · …", install link. | pass |
| PROXY-09 | P1 | web | The banner hides after a successful check ([The proxy banner](../setup/the-proxy.md#the-proxy-banner)). | Banner shown. | 1. Restore `https://localhost:3003`; wait. | Banner gone within 30 s. | pass (25 s) |
| PROXY-10 | P3 | web | Copy turns the icon to a check for 1.4 s ([The proxy banner](../setup/the-proxy.md#the-proxy-banner)). | — | — | — | blocked: not run |
| PROXY-11 | P1 | web | After a URL change the page claims the old check's result (suspected bug) ([Edge cases](../setup/the-proxy.md#edge-cases)). | Proxy at :3003 detected. | 1. Change URL to `https://localhost:3999`. | Record the status line. | pass: suspected bug confirmed, "✓ Proxy connected at https://localhost:3999" |
| PROXY-12 | P1 | web | The reverse: after switching back, the page says unreachable ([Edge cases](../setup/the-proxy.md#edge-cases)). | Not detected at :3999. | 1. Set `https://localhost:3003`; turn on. | Record. | pass: "⚠ Proxy is enabled at https://localhost:3003 but is not currently reachable." |
| PROXY-13 | P2 | standalone | The welcome overlay turns the switch on when a proxy answers. | See LOGIN-05. | — | — | pass |
| PROXY-14 | P1 | web | A remote proxy URL shows the security toast at startup ([What is decided](../setup/the-proxy.md#what-is-decided-when-the-taskpane-opens)). | — | — | — | blocked: needs a reload |
| PROXY-15 | P2 | — | The gate before OAuth. | See LOGIN-22. | — | — | blocked |
| PROXY-16 | P2 | web | The GPT-5.6 Luna messages ([When the proxy is missing](../setup/the-proxy.md#when-the-proxy-is-missing-at-the-moment-it-is-needed)). | — | — | — | blocked: needs a Luna turn with the proxy off; not spent |
| PROXY-17 | P2 | web | A default URL is stored ([What is decided](../setup/the-proxy.md#what-is-decided-when-the-taskpane-opens)). | — | 1. Read the URL field. | `https://localhost:3003`. | pass |
| PROXY-18 | P1 | web | OpenAI (ChatGPT) model requests fail without the proxy ([Summary](../setup/the-proxy.md#summary)). | Switch off. | 1. Send "hi". | Record. | fail: the turn succeeded. Document corrected. |
| PROXY-19 | P2 | web | Turning the switch off applies from the next request ([Modifiers](../setup/the-proxy.md#modifiers)). | As PROXY-18. | — | The next request went direct. | pass |
| PROXY-20 | P1 | web | A network failure shows the CORS-worded banner ([When the proxy is missing](../setup/the-proxy.md#when-the-proxy-is-missing-at-the-moment-it-is-needed)). | See GATEWAY-16. | — | "Network error (likely CORS). …". | fail: "LLM error: Connection error.". Document corrected: the CORS text depends on the error wording. |
| PROXY-21 | P2 | web | Escape on the Proxy page saves a pending URL edit ([Cancel and interrupt](../setup/the-proxy.md#cancel-and-interrupt)). | — | — | — | blocked: not run |

Not checkable by hand:

- Whether the banner should consider the switch and the provider (a product call).

## Results so far

One pass on 2026-10-26 by an agent in an Amp orb, at pi-for-excel `ec453df`. It drove Pi in Excel on the web (Chrome, OfficeOnline) through CDP input on the taskpane, and the standalone taskpane through agent-browser in fresh profiles with the pre-login blocked. What was on screen was read from the DOM, with a few screenshots; timings are approximate. Nothing on Excel desktop was run.

131 items: 99 pass, 6 fail, 26 blocked. Ten of the passes confirm suspected bugs (MODEL-08/PICKER-13, MODEL-20/GATEWAY-14, LOGIN-12, LOGIN-24, LOGIN-29, PICKER-12, PROXY-07, PROXY-11), and SETTINGS-09 confirms the missing guard. LOGIN-03 passed after the document was corrected. The fails:

- PICKER-08: a pointer resting over the list decides what Enter chooses. Possible product bug.
- LOGIN-21: GitHub Copilot's login fails with the CORS message although the proxy runs. Possible product bug.
- LOGIN-31: the banner after dismissing the connection dialog is "LLM error: Provider is not configured: openai". The welcome document now says so; [turns](../foundations/turns.md) still says "API key required for {provider}.".
- MODEL-05, PROXY-18: ChatGPT turns work without the proxy in Excel on the web. The documents were wrong and are corrected.
- PROXY-20: an unreachable gateway shows "LLM error: Connection error.", not the CORS banner. The document was wrong and is corrected.

Most blocked items need a reload of the Excel session (which repeats credential restore), a second real provider, a production build (the proxy gate), or would touch the shared OAuth credential. No document in this cluster is marked verified: P1 and P2 items are still blocked, and the fails are not yet filed in bug-triage.md.
