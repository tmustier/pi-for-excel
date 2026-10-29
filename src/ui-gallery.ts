import "./ui/register-components.js";

import { requestApproval } from "./ui/approval-dialog.js";
import { createBanner } from "./ui/banner.js";
import { humanizeToolInput } from "./ui/humanize-params.js";
import { describePendingToolCall } from "./ui/tool-renderers.js";
import { createCopyableCommand } from "./ui/command-copy.js";
import { mountTranscriptHarness } from "./ui-gallery-transcript.js";
import { setSafeInnerHTML } from "./utils/html.js";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { JsonObject } from "@earendil-works/pi-ai";
import type { TurnChangesMessage } from "./messages/turn-changes.js";

const galleryRoot = document.getElementById("gallery-root");
if (!galleryRoot) throw new Error("Missing #gallery-root");
const root: HTMLElement = galleryRoot;
root.style.cssText = `
  max-width: 380px;
  margin: 0 auto;
  padding: 16px;
  font-family: var(--font-sans);
  background: var(--background);
  color: var(--foreground);
  min-height: 100vh;
`;

function section(id: string, title: string): HTMLDivElement {
  const el = document.createElement("div");
  el.setAttribute("data-gallery", id);
  el.style.cssText = "margin-bottom: 32px;";

  const heading = document.createElement("h3");
  heading.textContent = title;
  heading.style.cssText = `
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--muted-foreground);
    margin: 0 0 12px;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--alpha-8);
  `;

  el.appendChild(heading);
  root.appendChild(el);
  return el;
}

const badgeSection = section("badges", "Overlay Badges");
const badgeRow = document.createElement("div");
badgeRow.style.cssText = "display: flex; gap: 8px; flex-wrap: wrap;";

for (const tone of ["muted", "ok", "warn", "info"]) {
  const badge = document.createElement("span");
  badge.className = `pi-overlay-badge pi-overlay-badge--${tone}`;
  badge.textContent = tone === "muted" ? "Read only" : tone === "ok" ? "Connected" : tone === "warn" ? "Warning" : "Info";
  badgeRow.appendChild(badge);
}
badgeSection.appendChild(badgeRow);

const fileItemSection = section("file-items", "File List Items");

function createMockFileItem(name: string, meta: string, badgeLabel?: string): HTMLButtonElement {
  const row = document.createElement("button");
  row.type = "button";
  row.className = "pi-files-item pi-files-item--muted";
  row.style.cssText = "width: 100%;";

  const icon = document.createElement("span");
  icon.className = "pi-files-item__icon";
  icon.textContent = "📄";

  const info = document.createElement("div");
  info.className = "pi-files-item__info";

  const nameRow = document.createElement("div");
  nameRow.className = "pi-files-item__name-row";

  const nameEl = document.createElement("span");
  nameEl.className = "pi-files-item__name";
  nameEl.textContent = name;
  nameRow.appendChild(nameEl);

  const metaEl = document.createElement("span");
  metaEl.className = "pi-files-item__meta";
  metaEl.textContent = meta;

  info.append(nameRow, metaEl);
  row.append(icon, info);

  if (badgeLabel) {
    const badge = document.createElement("span");
    badge.className = "pi-overlay-badge pi-overlay-badge--muted";
    badge.textContent = badgeLabel;
    row.appendChild(badge);
  }

  const arrow = document.createElement("span");
  arrow.className = "pi-files-item__arrow";
  arrow.textContent = "›";

  row.appendChild(arrow);
  return row;
}

fileItemSection.appendChild(createMockFileItem("cache-observability-baselines.md", "Pi documentation · 2.98 KB", "Read only"));
fileItemSection.appendChild(createMockFileItem("context-management-policy.md", "Pi documentation · 12.0 KB", "Read only"));
fileItemSection.appendChild(createMockFileItem("quarterly-report.xlsx", "1.2 MB · Uploaded · 2h ago"));

const turnSection = section("turn-group", "Turn group");

const fixtureUsage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } };
const fixtureMeta = { api: "openai-responses", provider: "openai", model: "gallery", usage: fixtureUsage } as const;

interface FixtureCall {
  id: string;
  name: string;
  args: JsonObject;
  result?: { text: string; details?: JsonObject; isError?: boolean };
}

function fixtureTurn(tag: string, options: {
  lead?: string;
  calls: FixtureCall[];
  answer?: string;
  stopReason?: "stop" | "toolUse" | "aborted";
  changes?: TurnChangesMessage["changes"];
}): AgentMessage[] {
  const base = options.answer ? 1_000 : Date.now() - 23_000;
  const calls = options.calls.map((call) => ({ type: "toolCall" as const, id: `${tag}-${call.id}`, name: call.name, arguments: call.args }));
  const messages: AgentMessage[] = [
    { role: "user", content: "Fill the Q3 rows from the Inputs sheet and flag anything that doesn't reconcile.", timestamp: base },
    {
      role: "assistant",
      content: [...(options.lead ? [{ type: "text" as const, text: options.lead }] : []), ...calls],
      ...fixtureMeta, stopReason: options.answer ? "toolUse" : options.stopReason ?? "toolUse", timestamp: base + 1,
    },
    ...options.calls.flatMap((call, index): AgentMessage[] => call.result ? [{
      role: "toolResult", toolCallId: `${tag}-${call.id}`, toolName: call.name,
      content: [{ type: "text", text: call.result.text }], ...(call.result.details === undefined ? {} : { details: call.result.details }),
      isError: call.result.isError ?? false, timestamp: base + 2_000 + index * 1_500,
    }] : []),
  ];
  if (options.answer) {
    messages.push({ role: "assistant", content: [{ type: "text", text: options.answer }], ...fixtureMeta, stopReason: "stop", timestamp: base + 8_200 });
  }
  if (options.changes) {
    messages.push({ role: "turnChanges", turnId: `${tag}-turn`, timestamp: base + 8_300, outcome: "finished", changes: options.changes });
  }
  return messages;
}

const fillResult = (address: string, count: number) => ({
  text: `Filled **${address}** — ${count} changed`,
  details: { kind: "fill_formula", blocked: false, address, changes: { changedCount: count, truncated: false, sample: [
    { address: address.replace(/:.*$/u, ""), beforeValue: "125,000", afterValue: "130,000", beforeFormula: "=C19+C24", afterFormula: "=Inputs!C4" },
  ] } },
});
const readResult = (address: string, rows: number, cols: number) => ({ text: `**${address}** (${rows}×${cols})\n\n| | A | B |\n| --- | --- | --- |\n| 1 | Jan | 130 |` });

const turnFixtures: Array<[string, AgentMessage[]]> = [
  ["Working (live rows)", fixtureTurn("live", {
    lead: "Inputs has monthly values in `Inputs!C4:K9`. I'll fill rows 10–15 with formulas, then check the totals.",
    calls: [
      { id: "r1", name: "read_range", args: { range: "Inputs!C4:K9" }, result: readResult("Inputs!C4:K9", 6, 9) },
      { id: "f1", name: "fill_formula", args: { range: "'Cash Flow'!D10:L10", formula: "=Inputs!C4" }, result: fillResult("'Cash Flow'!D10:L10", 9) },
      { id: "f2", name: "fill_formula", args: { range: "'Cash Flow'!D11:L11", formula: "=Inputs!C5" } },
    ],
  })],
  ["Failed step, then recovery", fixtureTurn("fail", {
    calls: [
      { id: "r1", name: "read_range", args: { range: "Inputs!C4:K9" }, result: readResult("Inputs!C4:K9", 6, 9) },
      { id: "f1", name: "fill_formula", args: { range: "'Cash Flow'!D12:L12", formula: "=Inputs!C6" },
        result: { text: "Error: Inputs!C6 no longer exists, so the formula returns #REF!.", isError: true } },
      { id: "f2", name: "fill_formula", args: { range: "'Cash Flow'!D12:L12", formula: "=Inputs!C7" }, result: fillResult("'Cash Flow'!D12:L12", 9) },
      { id: "p1", name: "python_transform_range", args: { range: "'Cash Flow'!D10:L15", code: "df.sum()" },
        result: { text: "Written to **Summary!B2:B7**", details: { kind: "python_transform_range", blocked: false, outputAddress: "Summary!B2:B7", changes: { changedCount: 6, truncated: false, sample: [] } } } },
      { id: "w1", name: "web_search", args: { query: "UK CPI Q3 2026" },
        result: { text: "1. ONS: CPI rose 3.1%", details: { kind: "web_search", ok: true, provider: "brave", query: "UK CPI Q3 2026", sentQuery: "UK CPI Q3 2026", maxResults: 10, resultCount: 6 } } },
    ],
    answer: "Filled rows 10–15. Five rows reconcile to Summary. **Row 12 (Opex)** returned `#REF!` because `Inputs!C6` was deleted, so I used C7 instead.",
    changes: [{ address: "'Cash Flow'!D10:L15", changedCount: 54, backupIds: ["gallery-backup-1"] }, { address: "Summary!B2:B7", changedCount: 6, backupIds: ["gallery-backup-2"] }],
  })],
];

turnFixtures.push(["Waiting for approval", fixtureTurn("approve", {
  calls: [
    { id: "f1", name: "fill_formula", args: { range: "'Cash Flow'!D10:L15", formula: "=Inputs!C4" }, result: fillResult("'Cash Flow'!D10:L15", 54) },
    { id: "d1", name: "modify_structure", args: { action: "delete_rows", sheet: "Summary", start_row: 40, count: 13 } },
  ],
})]);

for (const [label, messages] of turnFixtures) {
  const caption = document.createElement("div");
  caption.style.cssText = "font-size: 10px; color: var(--muted-foreground); margin: 12px 0 4px;";
  caption.textContent = label;
  const list = document.createElement("message-list");
  list.messages = messages;
  list.isStreaming = label.startsWith("Working") || label.startsWith("Waiting");
  list.pendingToolCalls = new Set(["live-f2", "approve-d1"]);
  turnSection.append(caption, list);
}

const approvalAbort = new AbortController();
const approvalParams = { action: "delete_rows", sheet: "Summary", start_row: 40, count: 13 };
const approvalDetails = humanizeToolInput("modify_structure", approvalParams);
void requestApproval({
  title: "Allow this change to the workbook?",
  toolCallId: "approve-d1",
  headline: describePendingToolCall("modify_structure", approvalParams) ?? "Delete rows",
  ...(approvalDetails ? { details: approvalDetails } : {}),
  note: "Confirm mode: Pi asks before each change to the workbook.",
  confirmLabel: "Delete",
  allowForTurn: true,
  signal: approvalAbort.signal,
}).then((approved) => { console.log(`[ui-gallery] approval answered: ${approved}`); });

const commandCopyWrap = document.createElement("div");
commandCopyWrap.style.cssText = "margin-top: 12px; display: flex; flex-direction: column; gap: 8px;";
commandCopyWrap.append(
  createCopyableCommand("npx pi-for-excel-proxy"),
  createCopyableCommand("npx pi-for-excel-python-bridge"),
);
turnSection.appendChild(commandCopyWrap);

const diffSection = section("diff-table", "Cell Changes Diff Table");

const diffWrap = document.createElement("div");
diffWrap.className = "pi-tool-card__section";
setSafeInnerHTML(
  diffWrap,
  `
  <div class="pi-tool-card__section-label">Changes (9)</div>
  <div class="pi-tool-card__diff">
    <table class="pi-tool-card__diff-table">
      <thead>
        <tr><th>Cell</th><th>Before</th><th>After</th></tr>
      </thead>
      <tbody>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">D10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$125,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =C19+C24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$130,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =D19+D24</div>
          </td>
        </tr>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">E10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$130,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =D19+D24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$135,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =E19+E24</div>
          </td>
        </tr>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">F10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$135,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =E19+E24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$140,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =F19+F24</div>
          </td>
        </tr>
      </tbody>
    </table>
  </div>
`,
  "static UI gallery diff-table fixture markup",
);
diffSection.appendChild(diffWrap);

const previewSection = section("text-preview", "File Text Preview");

const sampleMarkdown = `# Context Management Policy

**Status:** Active policy (2026-02-12)
**Scope:** How Pi for Excel builds and manages context

## Why this exists

We optimize for **answer quality and reliability** across multi-turn sessions. In practice, quality drops when we blindly stuff context or let it grow unbounded.

## Core principles

1. **Minimal viable context** — include only what improves this turn.
2. **Freshness over volume** — recent state > historical state.
3. **Bounded growth** — auto-compact before hitting limits.`;

const previewCaptionMd = document.createElement("div");
previewCaptionMd.style.cssText = "font-size: 10px; color: var(--muted-foreground); margin: 0 0 4px; opacity: 0.8;";
previewCaptionMd.textContent = "Markdown file — rendered";
previewSection.appendChild(previewCaptionMd);

const mdPreview = document.createElement("div");
mdPreview.className = "pi-files-detail-preview pi-files-detail-preview--markdown";
const mdPreviewBlock = document.createElement("markdown-block") as HTMLElement & { content: string };
mdPreviewBlock.content = sampleMarkdown;
mdPreview.appendChild(mdPreviewBlock);
const mdFade = document.createElement("div");
mdFade.className = "pi-files-detail-preview__fade";
mdPreview.appendChild(mdFade);
previewSection.appendChild(mdPreview);

const previewCaptionText = document.createElement("div");
previewCaptionText.style.cssText = "font-size: 10px; color: var(--muted-foreground); margin: 12px 0 4px; opacity: 0.8;";
previewCaptionText.textContent = "Plain text file — wrapped, no line numbers";
previewSection.appendChild(previewCaptionText);

const textPreview = document.createElement("div");
textPreview.className = "pi-files-detail-preview pi-files-detail-preview--text";
const textPreviewBody = document.createElement("div");
textPreviewBody.className = "pi-files-detail-preview__body";
textPreviewBody.textContent = `region,quarter,revenue,cost,margin
EMEA,Q1 2026,1204000,861000,0.285
EMEA,Q2 2026,1287500,903200,0.298
APAC,Q1 2026,644800,517900,0.197
APAC,Q2 2026,701300,548100,0.218`;
textPreview.appendChild(textPreviewBody);
previewSection.appendChild(textPreview);

const buttonsSection = section("buttons", "Overlay Buttons");

const btnRow = document.createElement("div");
btnRow.className = "pi-files-detail-actions";

for (const [label, cls] of [
  ["Open ↗", "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"],
  ["Download", "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"],
  ["Delete", "pi-overlay-btn pi-overlay-btn--danger pi-overlay-btn--compact"],
] as const) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = cls;
  btn.textContent = label;
  btnRow.appendChild(btn);
}
buttonsSection.appendChild(btnRow);

const toastSection = section("toasts", "Toast Notifications");

const toastStack = document.createElement("div");
toastStack.className = "pi-toast-region";
toastStack.style.cssText = "position: static; transform: none; width: auto;";
for (const [msg, variant] of [
  ["Closed “Chat 3”", "action"],
  ["Renamed tab to “Forecast”", "info"],
  ["Couldn’t save the session. Try again.", "error"],
] as const) {
  const toast = document.createElement("div");
  toast.className = `pi-toast visible${variant === "error" ? " pi-toast--error" : ""}${variant === "action" ? " pi-toast--action" : ""}`;

  const message = document.createElement("span");
  message.className = "pi-toast__message";
  message.textContent = msg;
  toast.appendChild(message);

  if (variant === "action") {
    const action = document.createElement("button");
    action.type = "button";
    action.className = "pi-toast__action";
    action.textContent = "Undo";
    toast.appendChild(action);
  }

  if (variant === "error") {
    const dismiss = document.createElement("button");
    dismiss.type = "button";
    dismiss.className = "pi-toast__dismiss";
    dismiss.setAttribute("aria-label", "Dismiss");
    dismiss.textContent = "×";
    toast.appendChild(dismiss);
  }

  toastStack.appendChild(toast);
}
toastSection.appendChild(toastStack);

const bannerSection = section("banners", "Banners (notice area)");
const bannerStack = document.createElement("div");
bannerStack.className = "pi-notices";
bannerStack.append(
  createBanner({ id: "error", tone: "error", message: "The last request failed: the provider returned 429 (rate limited).", actions: [{ label: "Retry", primary: true, onClick: () => {} }], onDismiss: () => {} }).root,
  createBanner({ id: "proxy", tone: "warning", message: "The proxy is turned on but isn’t running, so some providers and web tools won’t work.", actions: [{ label: "Show how to start it", onClick: () => {} }], onDismiss: () => {} }).root,
  createBanner({ id: "disclosure", tone: "info", message: "Pi can search the web, use extensions and connect to services you add.", actions: [{ label: "Got it", primary: true, onClick: () => {} }, { label: "Review in Settings", onClick: () => {} }], onDismiss: () => {} }).root,
);
bannerSection.appendChild(bannerStack);

const mdSection = section("markdown", "Markdown Rendering (font consistency)");

const mdBlock = document.createElement("markdown-block") as HTMLElement & { content: string };
mdBlock.content = `The formula is \`=IF(C$4-Assumptions!$B$10+1=Assumptions!$B$49,...)\` — C4 = calendaryear (2025 for Year 1, 2031 for Year 7).

Assumptions!B10 = 2025 (start year) − Assumptions!B$49 = 7

So for Year 7 (column I, calendar year 2031): 2031 – 2025 + 1 = 7 ✓`;

mdSection.appendChild(mdBlock);

const answerSection = section("answer", "Answer actions");
const answer = document.createElement("assistant-message");
answer.message = {
  role: "assistant",
  content: [{ type: "text", text: "Wrote and verified the data in Sheet1!B2:C4.\n\n| Region | Revenue |\n| --- | --- |\n| East | 17 |\n| West | 29 |" }],
  api: "openai-responses", provider: "openai", model: "gallery",
  usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
  stopReason: "stop", timestamp: 1,
};
answerSection.append(answer);

const transcriptSection = section("transcript", "Live turn (scripted stream)");
const transcriptControls = document.createElement("div");
transcriptControls.style.cssText = "display: flex; gap: 8px; margin-bottom: 8px;";
const transcriptHost = document.createElement("div");
transcriptHost.style.cssText = "height: 640px; border: 1px solid var(--alpha-8); border-radius: 12px; overflow: hidden;";
transcriptSection.append(transcriptControls, transcriptHost);

const transcriptHarness = mountTranscriptHarness(transcriptHost);
window.piTranscriptHarness = transcriptHarness;

for (const [label, action] of [
  ["Play turn", () => { transcriptHarness.reset(); void transcriptHarness.start({ autoRelease: true }); }],
  ["Reset", () => transcriptHarness.reset()],
] as const) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "pi-overlay-btn pi-overlay-btn--ghost";
  button.textContent = label;
  button.addEventListener("click", action);
  transcriptControls.appendChild(button);
}

const turnChangesSection = section("turn-changes", "Turn changes");
const turnChangesList = document.createElement("message-list");
turnChangesList.messages = (["finished", "stopped", "failed"] as const).map((outcome, index): TurnChangesMessage => ({
  role: "turnChanges", turnId: `gallery-turn-${index}`, timestamp: index, outcome,
  changes: [
    { address: "Sheet1!B2:B40", changedCount: 39, backupIds: [`gallery-backup-${index}`] },
    { address: "Summary!A1", changedCount: 1, backupIds: outcome === "failed" ? [] : [`gallery-summary-${index}`] },
  ],
}));
turnChangesSection.append(turnChangesList);

const turnCoverageSection = section("turn-coverage", "Turn undo coverage");
const turnCoverageList = document.createElement("message-list");
turnCoverageList.messages = [false, true].map((mixed): TurnChangesMessage => ({
  role: "turnChanges", turnId: `gallery-coverage-${mixed}`, timestamp: 0, outcome: "finished", ranOfficeJs: true,
  changes: mixed ? [{ address: "Sheet1!D7", changedCount: 1, backupIds: ["core-backup"] }] : [],
  ...(mixed ? { officeJsRanges: ["Sheet1!X12", "'Q3 Plan'!B2:C3"] } : {}),
}));
turnCoverageSection.append(turnCoverageList);

console.log("[ui-gallery] Rendered all sections");
