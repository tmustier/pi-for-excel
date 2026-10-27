# Review: setup, guidance and extensibility

This review asks whether the behavior is the best behavior for users. It does not check whether the documents describe it correctly. It covers first run, login and providers, the proxy, model choice, rules, conventions, skills, the Files workspace, extensions, web search and MCP, and consent. The target is `main` as documented plus the fixes in flight ([#786](https://github.com/tmustier/pi-for-excel/pull/786)–[#795](https://github.com/tmustier/pi-for-excel/pull/795)), read against the PR descriptions on 27 October 2026. Known defects in [bug-triage.md](../bug-triage.md) are not raised again unless a fix PR's decision is itself the problem.

The users in mind are analysts and finance professionals installing Pi for the first time, in Excel desktop or on the web, who are often not technical about proxies or API keys.

## Summary

Once a user is connected, the fix PRs make the product coherent. The two biggest gaps are at the edges.

- **Getting in.** The subscription logins that most of these users would pick need a Terminal command, Node.js and possibly an admin password. Then they need a helper running on later days too. A user who dismisses the welcome screen is asked for an OpenAI API key they never chose.
- **Changes Pi makes to its own future behavior.** Pi can change rules and conventions, and install skills, with no consent and no undo. After #790 it must ask before writing a scratch file in Confirm mode, but it can still rewrite the user's standing instructions without asking. MCP tool calls, which can act in the outside world, never ask.

The other proposals remove complexity: the scope matrix for integrations, the Experimental page and the fork setting. They also bring setup in line with what users of Copilot, Claude for Excel and Cursor already expect: rules that travel with the file, attachments that attach, and gateways that list their models.

## Proposals, ranked

### SE-01: Make a subscription login work without Terminal

- **Today / planned.** OAuth for ChatGPT, Anthropic, the Google sign-ins and Copilot first hits the proxy gate: "This login method needs a small helper running on your Mac. Open the Terminal app and paste this command" ([welcome and login](../setup/welcome-and-login.md#committing-a-credential)). The install guide warns that the helper may ask for the Mac password to install a certificate, and that a non-admin should "ask your IT team" ([docs/install.md](../../docs/install.md#oauth-logins-and-cors-proxy)). P-60 suspects that current Chrome and Edge block a hosted page from reaching `localhost` at all, which would affect Excel on the web. #795 makes Mac/Terminal wording platform-aware, but the step itself stays.
- **Proposed.**
  1. Do the OAuth code and refresh-token exchange through a small hosted endpoint on the add-in's own origin (the add-in is already served from Vercel). The local helper is then needed only for model traffic that really is CORS-blocked. Excel on the web already shows that ChatGPT model requests work with no proxy (PROXY-18).
  2. Until that ships, label each provider row before the click with what it needs: "Paste an API key: no setup", or "Sign in: needs a helper app on this computer". Put the no-setup options first in the welcome list.
- **Why.** The user Pi most wants, a finance professional with a ChatGPT Plus seat, cannot finish the default path on a locked-down corporate laptop. Comparable products need none of this. Claude for Excel installs from AppSource and the user then "sign[s] in with your Claude account" ([Claude docs](https://claude.com/docs/office-agents/excel)). Copilot in Excel uses the Microsoft 365 sign-in that is already there.
- **Cost / risk.** A hosted exchange endpoint sees the refresh tokens in transit. It must store nothing, accept only the token endpoints and be covered by the threat model. It also adds a server dependency to what is today a static add-in. Org builds can keep the central proxy instead ([docs/central-proxy.md](../../docs/central-proxy.md)).
- **PR.** #795 (wording), fix-host-web (P-60).
- **Decision.** Product call.

### SE-02: Never show a model the user cannot use, and say when Pi changes one

- **Today / planned.** With no provider, the status bar names "GPT-6.1 Sol" from OpenAI (API), and the first send opens "Connect OpenAI (API)", whatever the user meant to use ([models and providers](../foundations/models-and-providers.md#the-default-model)). When a provider disconnects, idle tabs move to the default model without a word ([When a provider is disconnected](../foundations/models-and-providers.md#when-a-provider-is-disconnected)). The control does not name the provider, so two tabs labelled "GPT-6.1 Sol" can bill different accounts (P-89). #795 fixes the subscription preference and the picker, but not these three things.
- **Proposed.**
  1. With no provider connected, the model control reads "Connect a provider". Sending opens the provider list (the welcome card) with the user's message kept, not a dialog for one fixed provider.
  2. An automatic swap shows a toast once per change, for example "Chat 2 now uses Claude Opus 5.5 (OpenAI (API) disconnected)".
  3. When two connected providers offer a model with the same name, the control adds the provider ("GPT-6.1 Sol · ChatGPT").
- **Why.** A user who dismisses the welcome card hits a dead end on the first message: a request for an API key from a provider they don't have. Silent swaps also change cost and context size without telling the user.
- **Cost / risk.** Small. These are UI changes in the status bar and `model-refresh-owner.ts`. Point 3 takes a little status-bar width only in the ambiguous case.
- **PR.** #795.
- **Decision.** Fix.

### SE-03: Treat changes to Pi's own standing instructions as consent-worthy

- **Today / planned.** Pi can append to or replace user and workbook rules (`instructions`), change or reset conventions (`conventions`), and install or uninstall skills (`skills`). None of these ask, in either mode, and none shows a toast. The only trace is a tool card ([rules](../guidance/rules.md#when-pi-changes-the-rules), [conventions](../guidance/conventions.md#when-pi-changes-the-conventions), [skills](../guidance/skills.md#when-pi-uses-skills)). #790's rule says so on purpose: "Not gated: `instructions`, `conventions`, `skills` (prompt guidance, visible on the Rules/Skills pages)". `skills` is even reclassified as a read, although its install action writes a `SKILL.md` into Files. The same PR makes Confirm mode ask before every `files` write.
- **Proposed.**
  1. Any change Pi makes to rules, conventions or installed skills shows a toast with Undo, in both modes, for example "Pi updated your rules for all files · Undo".
  2. In Confirm mode it asks first and shows the new text.
  3. A workbook-rules change by Pi always asks. Pi's prompt already says to ask in chat, but nothing enforces it.
- **Why.** These are the most durable changes Pi can make. They apply to every future conversation in every workbook, and the user cannot see them happen. They are also the obvious prompt-injection target: a page read with `fetch_page`, or a pasted document, only has to convince Pi to "remember" an instruction once. Today's design gates a scratch file in Confirm mode but not a permanent instruction. Cursor and Claude Code keep rules in files the user edits and reviews. They are not changed silently by the agent.
- **Cost / risk.** Low. It reuses #790's `requestRuntimeToolApproval` and #793's toasts. Undo needs the previous value, which the tool already reads. A toast per rule append could feel noisy, but appends are rare.
- **PR.** #790 (amend its consent rule).
- **Decision.** Product call.

### SE-04: Ask before MCP tools that can change things

- **Today / planned.** "External tools … never ask for approval" ([glossary](../glossary.md#the-workbook), [web search and MCP](../extensions/web-search-and-mcp.md#what-the-user-sees-when-pi-uses-them): "Whether a remote tool changes anything is up to the server; Pi asks nothing first."). #790 makes MCP usable for the first time (the "Let Pi use MCP servers" card), so this becomes real.
- **Proposed.** The first `MCP call` to each tool asks, naming the server, the tool and its inputs, with "Allow once" and "Always allow this tool". The exception is tools the server marks `readOnlyHint: true`. In Confirm mode every call asks. The server card lists the remembered tools and lets the user revoke them.
- **Why.** An MCP server can send email, post to Slack or write to a CRM. Analysts will connect exactly those systems. In ChatGPT, "write actions by default require confirmation" ([OpenAI developer mode](https://developers.openai.com/api/docs/guides/developer-mode)). Claude's apps ask per tool too. Pi asks before a scratch file is overwritten but not before an email is sent.
- **Cost / risk.** Medium. The decision is per tool and a few dialogs are needed. Relying on the server's hints means trusting the server, so default to asking when there is no hint.
- **PR.** #790.
- **Decision.** Product call.

### SE-05: Store workbook rules in the workbook

- **Today / planned.** Workbook rules ("This file") are kept in the browser on this computer and keyed by a hash of the file's address. They do not travel with the file, they are lost on rename, move or Save As, and they cannot exist for an unidentified workbook ([rules](../guidance/rules.md#after-saving)). #788 disables the box there, with "Try saving the file first". That cuts off every signed-out web workbook and every new workbook.
- **Proposed.** Save workbook rules inside the file, in Office document settings or a custom XML part, labelled on the page "Saved in this file. Anyone you share it with gets them." Keep "All my files" in the browser.
- **Why.** Workbook rules describe the workbook ("Revenue assumptions in Inputs!B5:B15"), so they belong with it, the way Cursor project rules "are stored in `.cursor/rules/` inside your project folder" ([Cursor](https://cursor.com/help/customization/rules)). A model handed to a colleague should arrive with its notes. Renaming `Model v3.xlsx` to `Model v4.xlsx` should not wipe them. It also removes the identity dependency that makes "This file" unusable on the web.
- **Cost / risk.**
  - Rules become visible to everyone who gets the file. That is a privacy change, so the label must say so.
  - A one-time migration is needed from the stored copy to the document.
  - Document settings are saved only when the user saves the workbook. The page should say so.
- **PR.** #788.
- **Decision.** Product call.

### SE-06: Drop or flag the Claude Pro/Max sign-in

- **Today / planned.** The first provider row is "Anthropic ("Claude Pro/Max")", OAuth through the proxy ([models and providers](../foundations/models-and-providers.md#connecting-a-provider)).
- **Proposed.** Remove the subscription login from the Anthropic row and keep the API key, or show a clear notice before login. Bring the row back if Anthropic offers an approved third-party sign-in.
- **Why.** Anthropic's terms say: "Anthropic does not permit third-party developers to offer Claude.ai login into their own applications, or to route requests through Free, Pro, or Max plan credentials on behalf of their users" ([Claude Code legal and compliance](https://code.claude.com/docs/en/legal-and-compliance)). Other third-party tools report the token being refused ("This credential is only authorized for use with Claude Code", [developersdigest](https://www.developersdigest.tech/errors/opencode/claude-code-credential-only)). A finance user who picks the first row risks a refused login, or an account action, and cannot tell why. I did not log in with a real Claude account, so whether Pi's flow is refused today is unverified.
- **Cost / risk.** It loses a path some users rely on. The ChatGPT subscription row stays: OpenAI documents Codex sign-in for third-party clients, and #795 already prefers it.
- **PR.** None.
- **Decision.** Product call.

### SE-07: Don't offer web search until it can work

- **Today / planned.** Web search is on by default, but every provider needs a key. The first search fails, and then a setup card asks the user to get a Jina key ([web search and MCP](../extensions/web-search-and-mcp.md#the-simple-case)). #790 fixes the copy ("every provider needs a key") but keeps the default.
- **Proposed.**
  1. Leave `web_search` out of Pi's toolset until a key exists. Pi then says it cannot search instead of failing a tool call mid-turn, and the disclosure and Connections page say "Add a search key to let Pi search".
  2. Better: when the connected model provider has its own web search (the OpenAI Responses and Anthropic APIs both do), use it with no extra key, and keep the third-party providers for users who want them.
- **Why.** A failed tool call, then a card asking an analyst to sign up for a search API in the middle of a question, is a poor first experience. Users of ChatGPT or Claude expect search to just work under the account they already pay for.
- **Cost / risk.** Point 1 is small. Point 2 is medium: the tools differ per provider, and native search results bypass Pi's citation formatting.
- **PR.** #790.
- **Decision.** Point 1 fix, point 2 product call.

### SE-08: In Confirm mode, don't ask about Pi's own scratch files

- **Today / planned.** #790: "Confirm mode: every `files` write/delete … asks." That includes `notes/`, `scratch/` and files Pi created a moment ago in the same turn.
- **Proposed.** Use #790's Auto rule in both modes for `files`: ask before deleting or overwriting a file Pi did not create in this chat. Confirm mode additionally asks before every extension change, as #790 already does.
- **Why.** Users choose Confirm to control changes to the workbook. A dialog for "Write file scratch/pivot-plan.md" teaches them to click Allow without reading, which weakens the dialogs that matter. This also frees the attention budget SE-03 and SE-04 need.
- **Cost / risk.** Very small, since the rule already exists. Users who want every file write gated lose that option. I doubt anyone wants it.
- **PR.** #790.
- **Decision.** Product call.

### SE-09: A file dropped on the composer should be attached to the message

- **Today / planned.** A drop imports the file into Files with "Imported 1 file into Files.". Nothing is attached, the draft is untouched, and the file does not count as relevant for auto-context ([the Files workspace](../files/the-files-workspace.md#the-simple-case), [importing](../files/the-files-workspace.md#importing)).
- **Proposed.** A drop on the composer also adds a chip to the draft ("budget.csv", removable). When sent, Pi is told the file is part of this message and reads it. Files stays the library. Upload from the Files overlay keeps today's behavior.
- **Why.** In ChatGPT, Claude and Copilot, a dropped file is an attachment. A user who drops `budget.csv` and types "compare this with Sheet1" expects "this" to mean the file. Today Pi has to guess from the Files summary.
- **Cost / risk.** Small to medium: a chip in the composer and one line in the user message. Per-tab drafts in #789 make the chip safe to hold in the draft.
- **PR.** #789 (composer), #792 (context).
- **Decision.** Product call.

### SE-10: One switch per integration instead of a scope matrix

- **Today / planned.** Web search has "Enable for this session" and "Enable for workbook ({label})", an "Availability" line with five values, and a rule that a never-configured workbook counts as on while the session switch starts off ([web search and MCP](../extensions/web-search-and-mcp.md#the-page)). On top of that sits the "External tools" master switch. #790's new MCP card copies the same session and workbook switches.
- **Proposed.** One switch per integration ("Web search", "MCP servers") that applies everywhere, under the master switch. If per-workbook control is needed, add one "Off for this workbook" exception. Drop the session scope: sessions are invisible to users, and a switch that resets on a new tab is surprising.
- **Why.** Nobody in this audience will understand a choice between session and workbook scope. The current defaults already contradict each other: on for an identified workbook, off for a new session, on again for an unidentified workbook. Copilot and Claude for Excel have one on/off per connector.
- **Cost / risk.** Small. Existing stored scopes need a migration rule (on if any scope was on).
- **PR.** #790.
- **Decision.** Fix.

### SE-11: Delete the Experimental page and the fork setting

- **Today / planned.** After #790 removes `extension-permissions`, four flags remain ([experimental](../extensions/experimental.md#the-flags)):
  - "Dark mode", off. #793 makes it follow Office.
  - "Extension widget API v2", which only affects extensions.
  - "Remote extension URLs", a security opt-in.
  - "Extension sandbox rollback", a kill switch that gives pasted code full access to the page.

  Settings also has "Fork model switch into new tab" ([settings](../setup/settings.md#the-root-page)), and the Connections page shows the Python and tmux bridge cards to everyone.
- **Proposed.**
  - Dark mode follows Office's theme by default, with no flag.
  - Widget API v2 is always on.
  - The remote-URL opt-in moves to the Plugins page's URL install dialog, as an explicit "Allow code from this address" checkbox with the warning.
  - Sandbox rollback stays reachable only by command, for support.
  - Delete the Experimental row and page.
  - Remove the fork setting. Users who want a copy can use "Duplicate tab" and then change the model.
  - Hide the bridge cards until a bridge has been seen or a command opened it.
- **Why.** Every row in Settings is a decision the user has to understand. These are developer controls, and one of them (rollback) is dangerous behind a friendly switch. The fork setting covers a rare need that "Duplicate tab" already meets.
- **Cost / risk.** Small. Developers lose a GUI for the flags, but the commands remain.
- **PR.** #790 (flags), #793 (dark mode), #795 (fork).
- **Decision.** Fix.

### SE-12: Show the proxy only when a request actually needs it, and warn before renewal fails

- **Today / planned.** #793 shows the proxy banner only when the proxy switch is on. But the switch is turned on by itself: the welcome overlay turns it on when a helper answers, and the gate's "Retry" turns it on ([welcome and login](../setup/welcome-and-login.md#arriving-at-the-welcome-overlay)). So a ChatGPT user who ran the helper once to log in has it on for good. They then see "Proxy not running" on every day they don't start Terminal, although ChatGPT model requests work without it on the web. Separately, #795 refreshes OAuth tokens before requests, and the `auth.openai.com/oauth` refresh is routed through the proxy (`src/auth/cors-proxy.ts`). If that refresh is CORS-blocked without the helper, the login will fail on the first expiry after the helper stops. This is a hypothesis: I did not check whether the refresh endpoint blocks direct browser calls.
- **Proposed.**
  1. Show the banner only when a request the active tab needs would go through the proxy and the proxy is down.
  2. Send ChatGPT model traffic direct when that works, as on the web today.
  3. If renewing a login needs the helper, say so before it fails: "Your ChatGPT login renews in 2 days and needs the helper running." Do not end the turn with a CORS error.
- **Why.** A warning shown on most days for something the user doesn't need trains them to ignore warnings. An expired login that fails as "Network error (likely CORS)" is a dead end. SE-01 removes most of this. This is the fallback if SE-01 is not done.
- **Cost / risk.** Small to medium. Routing by need is already the design for gateways.
- **PR.** #793, #795.
- **Decision.** Fix.

### SE-13: Let a gateway list its models

- **Today / planned.** One gateway holds one model, and "Pi does not ask the gateway for its list of models" ([custom gateways](../setup/custom-gateways.md#how-its-model-appears-and-is-used)). Unknown models default to 16,384 tokens of context and no reasoning.
- **Proposed.** After the endpoint and key are entered, "Find models" calls `GET /v1/models` and shows a checklist. The user ticks the models to offer, and they become one provider with several models, with context taken from the response where the gateway gives it.
- **Why.** Corporate LLM gateways (LiteLLM, Azure APIM, internal OpenAI-compatible fronts) expose dozens of models, and these are exactly how finance organisations roll Pi out. Adding one gateway entry per model, and guessing each context size, is clerical work. A 16k default makes long workbooks fail early. Claude for Microsoft 365 already has an "Enterprise gateway" sign-in with a gateway URL and token ([Claude help](https://support.claude.com/en/articles/13945233-)).
- **Cost / risk.** Medium. Stored gateways need a data-model change from one model to several, with a migration. Some gateways don't implement `/v1/models`, so manual entry stays.
- **PR.** None.
- **Decision.** Fix.

## Decisions for Thomas

- **SE-01.** A hosted token-exchange endpoint in place of the local helper for login.
- **SE-03.** Gate Pi's changes to rules, conventions and skills. This reverses #790's "Not gated" line.
- **SE-04.** Consent for MCP tools that can change things.
- **SE-05.** Workbook rules stored in the file, so they are shared with recipients.
- **SE-06.** Remove or flag the Claude Pro/Max login.
- **SE-07, point 2.** Use the model provider's native web search.
- **SE-08.** Narrow #790's Confirm-mode files gate.
- **SE-09.** Composer drops become attachments.
- On the questions the brief raised:
  - **#795's subscription-first default.** I agree (see below).
  - **Removing `/default-models`.** I agree (see below).
  - **#790's consent rule.** The direction is right, but it is inverted at both ends: too much in Confirm (SE-08), too little for durable instructions and MCP (SE-03, SE-04).

## Right, and should not change

- **#795 prefers the ChatGPT subscription over an OpenAI API key.** The subscription is already paid for, and API usage bills every token. Billing someone's API account should be an explicit choice. SE-02's provider label makes the choice visible.
- **Removing `/default-models` and having no default-model setting.** A new tab copies the active tab's model. That matches what users just chose, and it saves a settings page nobody would find.
- **#795's key check on save (401/403 rejected, anything else saved as unverified).** A wrong key now fails at the moment the user can fix it, not in the middle of their first question.
- **Credentials never live in the workbook.** Sharing a file never shares a login. That stays true even if SE-05 moves workbook rules into the file.
- **Skills are read on demand from a one-line catalogue.** This keeps prompts small and matches how agent skills work elsewhere. #788's removal confirmation is the right amount of friction.
- **Conventions are applied by the product, not the model.** "Currency" gives the same format string with every model, which is what a finance house style needs.
- **#788's Rules save.** Writing only the scopes that changed, and asking only on a real conflict, avoids both lost edits and needless prompts.
- **#788's "keep both" for same-name imports.** It never destroys anything and needs no dialog.
- **#790's extension security defaults.** Permissions are always enforced, remote URLs are off, and new code never runs silently, even in Auto mode.
- **#793's dismissible proxy banner and "Review in Settings".** Dead switches are replaced by a path to the real controls.

Reviewed against `main` at `ec453df` and PR descriptions #786–#795 as of 27 October 2026.
