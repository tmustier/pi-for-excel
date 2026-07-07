# Settings & nested screens — UX audit and redesign

**Status:** Implementation in progress (branch `settings-ux-redesign`)
**Last reviewed:** 2026-07-07

## 1. Why

The settings/config surfaces were built screen-by-screen early in the project. Each
one is individually fine, but together they grew into seven sibling overlays with
three hand-rolled tab systems, inconsistent naming, and no navigational model. This
doc records the audit of what exists, the problems, and the target design.

## 2. Audit — what exists today (pre-redesign)

### Entry points

| Entry | Opens |
|---|---|
| Gear button → utilities menu | 7 items: Setup, Extensions, Files, Rules, Resume session, Backups, Keyboard shortcuts |
| `/settings`, `/login` | Settings overlay (tab: Providers) |
| `/extensions`, `/tools`, `/skills`, `/plugins` | Extensions hub (per-tab deep links) |
| `/rules` | Rules overlay |
| `/history` | Backups overlay |
| `/shortcuts`, `/help` | Shortcuts overlay |
| `/resume`, `/resume-here` | Resume overlay |
| `/files` | Files workspace |
| Disclosure bar "Customize", proxy banner "How to fix", welcome login | Settings deep links |

### Screens

1. **Settings** (`settings-overlay.ts`, menu label "Setup") — tabs *Providers | More*.
   Providers tab: Proxy card, provider list, custom gateway form. More tab:
   Execution mode, "Advanced" (fork-on-model-switch toggle, launcher buttons for
   Rules/Backups/Shortcuts, language select), Experimental toggles.
2. **Extensions hub** (`extensions-hub-overlay.ts`) — tabs *Connections | Plugins |
   Skills* with live refresh subscriptions.
3. **Rules** (`rules-overlay.ts`, 1011 lines) — tabs *All my files | This file |
   Formats* with draft + Save/Cancel footer.
4. **Backups** (`recovery-overlay.ts`) — search/filter/sort toolbar, retention,
   restore/delete.
5. **Shortcuts** (`shortcuts-overlay.ts`) — static list.
6. **Resume session** (`resume-overlay.ts`) — session picker; tab-styled buttons
   used as a *mode selector* (not navigation).
7. **Files workspace** (`files-dialog.ts`) — document workspace.

### Problems found

**Information architecture**

- **No navigational model.** Seven sibling modals; opening Rules from inside
  Settings *stacks* a second overlay on top with no back affordance. Escape pops
  overlays in accidental stack order.
- **Junk-drawer "More" tab** mixing safety-critical (execution mode), preferences
  (language), launchers, and dev-facing experimental toggles with raw slash-command
  hints.
- **Arbitrary split of credentials**: provider API keys live in Settings, web
  search keys and MCP tokens live in Extensions → Connections. Users must learn
  the split.
- **Proxy shown above Providers** although providers are the #1 first-run need and
  proxy is an edge-case corporate feature.
- **Naming drift**: menu says "Setup", dialog says "Settings"; disclosure bar says
  "Change anytime in Settings".
- **Three duplicated tab implementations** (`data-settings-tab`, `data-hub-tab`,
  rules' own record of buttons), each re-implementing ARIA and activation.

**Polish bugs (screenshot-verified)**

- Rules → Formats renders raw i18n keys (`rules.preset.number` …) — the
  `rules.preset.*` keys don't exist in the locales.
- Extensions → Plugins header says "1 skill" — `createSectionHeader` hardcodes
  the "skill(s)" unit for every consumer.
- Backups empty state: `recovery.emptyState` contains `\n` rendered via
  `textContent` into a collapsed line ("No backups yet Pi will save snapshots…").
- Expanded connection card at taskpane width clips its Validate/Save buttons.
- Long connection names truncate hard ("Jina Search (defa…").

## 3. Target design

### 3.1 One navigable Settings surface

Replace the seven-sibling-modal model with **one overlay containing a page
stack** (the pattern used by iOS Settings / narrow-viewport settings UIs — the
taskpane is ~350 px wide, which is exactly the geometry this pattern is for):

```
┌───────────────────────────────┐
│ ‹ Back   Page title         × │   header: back (when depth>0), title, close
│───────────────────────────────│
│  scrollable page content      │   grouped rows / page-specific content
│───────────────────────────────│
│  optional sticky footer       │   e.g. Rules Save/Cancel
└───────────────────────────────┘
```

- **Back** pops the stack. **Escape** = back (close at root). **×** always closes.
- Pages declare `parentId`; deep links reconstruct the stack so Back always works.
- Pages can install a `beforeLeave` guard (Rules uses it for a discard-changes
  confirm when dirty).
- One shared implementation of header/ARIA/cleanup/transitions replaces the three
  tab systems.

### 3.2 Information architecture

Root page (`Settings`), grouped rows with value previews:

| Group | Rows |
|---|---|
| **AI providers** | Model providers → (`N connected`) · Custom gateway → · Proxy → (`On/Off`) |
| **Behavior** | Auto-apply changes (toggle, inline) · Fork model switch into new tab (toggle, inline) · Language (select, inline) |
| **Rules & data** | Rules & conventions → · Backups → |
| **Extensions** | Connections → · Plugins → (`N installed`) · Skills → (`N active`) |
| **Help** | Keyboard shortcuts → · Experimental features → |

Decisions:

- **"More" and "Advanced" are gone.** The three settings they held are ordinary
  toggles/selects on the root page; the launcher buttons become real navigation.
- **Execution mode (Auto-apply)** gets top billing in Behavior — it is the single
  most consequential setting.
- **Extensions stay three separate pages** (Connections / Plugins / Skills) but as
  siblings under root with descriptive sublabels, instead of a second hub with its
  own tab bar. The credentials story is now one hop apart (Model providers vs
  Connections) under the same roof.
- **Files and Resume session are not settings.** They remain their own overlays
  (workspaces), still launched from the utilities menu.
- **Naming standardizes on "Settings"** everywhere (menu, title, banners).

### 3.3 Component architecture

```
src/ui/settings-shell.ts        overlay + page stack + header/footer chrome
src/ui/settings-rows.ts         group cards, nav rows, value previews, segmented control
src/ui/theme/overlays/settings-shell.css
src/commands/builtins/settings-pages/
  index.ts                      page registry + deps wiring + openSettings(pageId)
  root-page.ts                  grouped root
  providers-page.ts             provider list (from provider-login.ts)
  gateway-page.ts               custom gateway (reuses buildCustomGatewaySection)
  proxy-page.ts                 proxy card (moved from settings-overlay)
  rules-page.ts                 rules editor + conventions (moved from rules-overlay)
  extensions-pages.ts           connections/plugins/skills (reuse tab renderers)
  backups-page.ts               recovery list (moved from recovery-overlay)
  shortcuts-page.ts             static list (moved from shortcuts-overlay)
  experimental-page.ts          experimental toggles (reuses experimental-overlay builders)
```

Contracts:

```ts
interface SettingsPage {
  id: SettingsPageId;
  parentId?: SettingsPageId;
  title(): string;
  render(ctx: SettingsPageContext): void | Promise<void>;
}
interface SettingsPageContext {
  body: HTMLElement;                       // scrollable content region
  setFooter(el: HTMLElement | null): void; // sticky footer slot
  navigate(id: SettingsPageId): void;
  back(): void;  close(): void;
  addCleanup(fn: () => void): void;        // run on page leave/overlay close
  setBeforeLeave(guard: (() => Promise<boolean>) | null): void;
}
```

- Tab renderers from the extensions hub (`renderConnectionsTab`, etc.) are reused
  as-is — they already take a container. The hub's live-refresh subscription logic
  moves into the pages' `addCleanup` lifecycle.
- `createSectionHeader` count becomes a caller-supplied localized label (fixes the
  "1 skill" bug for plugins).

### 3.4 Compatibility

- `showSettingsDialog({ section })` remains as a shim mapping old section names to
  page ids (`logins`→`providers`, `more`→root, `execution-mode`→root, etc.).
- Slash commands deep-link: `/settings`→root, `/login`→providers, `/rules`→rules,
  `/history`→backups, `/shortcuts`→shortcuts, `/tools`→connections,
  `/extensions`→plugins, `/skills`→skills.
- `SETTINGS_OVERLAY_ID` is reused for the shell; the retired per-overlay ids
  (`ADDONS/RULES/RECOVERY/SHORTCUTS`) are removed with their overlays.
- Types used elsewhere (`RecoveryCheckpointSummary`, etc.) keep their exports from
  the new page modules.

## 4. Out of scope (follow-ups)

- Resume overlay: replace the tab-styled mode selector with an explicit control
  (radio/segmented "open in" choice) using the shared segmented component.
- Files workspace visual alignment with the new row primitives.
- Welcome/onboarding overlay convergence on the same primitives.
- Settings search (not worth it until content grows further).

## 5. Verification

- `npm run check`, `npm run build`, `npm run test:context` (i18n parity),
  targeted test updates in `tests/builtins-registry.test.ts`.
- Visual: taskpane in browser (`scripts/ui-verify.sh` flow) — screenshot every
  page at taskpane width, plus deep-link and back-stack behavior.
- Manual Excel smoke for gear menu, slash commands, and dirty-guard on Rules.

## 6. Implementation status (2026-07-07)

Implemented on branch `settings-ux-redesign`:

- **Shell + primitives**: `src/ui/settings-shell.ts` (page stack, back/close,
  Escape-goes-back, deep links via `parentId` chains, footer slot, cleanup +
  `beforeLeave` guard), `src/ui/settings-rows.ts` (groups, nav/toggle/select
  rows, segmented control), `src/ui/theme/overlays/settings-shell.css`.
- **Pages** under `src/commands/builtins/settings-pages/`: root, providers,
  gateway, proxy, rules (drafts + Save/Cancel footer + discard guard), backups,
  connections/plugins/skills (reusing hub tab renderers with page-scoped
  lifecycle), shortcuts, experimental. Runtime wiring via
  `configureSettingsPages` from `taskpane/init.ts`.
- **Removed**: `settings-overlay.ts`, `extensions-hub-overlay.ts`,
  `rules-overlay.ts`, `recovery-overlay.ts`, `shortcuts-overlay.ts`, retired
  overlay ids, dead CSS (old tab system, settings panels, dialog cards).
- **Compat**: `showSettingsDialog({ section })` shim; `openSettings(pageId)` is
  the new API; slash commands and sidebar menu deep-link to pages.
- **Polish fixes shipped**: `rules.preset.*` locale keys (raw-key bug),
  numeric section-header count (the "1 skill" plugins bug),
  `recovery.emptyState` line break (`white-space: pre-line`), API-key row
  overflow at taskpane width (flex wrap), menu "Setup"→"Settings",
  experimental page duplicate heading, nav-row value truncation.
- **Verification**: `npm run check`, full `npm test` (889 tests), `npm run
  build` green; browser walkthrough at 380px with screenshots of every page,
  back-stack, Escape semantics, and the rules dirty-guard.

Remaining follow-ups: section 4 items (resume mode selector, files workspace
alignment, welcome overlay convergence) plus a manual Excel smoke test before
merge.
