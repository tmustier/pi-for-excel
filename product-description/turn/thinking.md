# Thinking

## Summary

The [thinking level](../glossary.md#models-and-providers) sets how much the model reasons before it answers, trading speed for thoroughness. Each tab has its own level, saved with its session. The level is shown in the status bar next to the model, with a brain icon ("High"); the user changes it by clicking that control and picking from a popover, or by pressing Shift+Tab to step to the next level. What the model reasons, when the provider returns it, appears in the transcript as a [thinking block](../glossary.md#the-turn): a collapsible card labelled "Thinking…" while the model response streams. Changing the level is allowed at any time, including while the tab is busy and with no provider connected; it takes effect from the next turn. Models that do not reason offer only "Off"; some reasoning models do not offer "Off" at all. The control and the blocks behave the same on every host.

## The simple case

The status bar reads "High". The user clicks it. A small popover titled "Thinking level" opens above it, explaining "How deeply Pi reasons before answering — higher is slower but more thorough. Click to select, or press ⇧Tab to cycle." and listing the levels the model supports, each with a hint ("Off · Fastest — no reasoning step", "Low · Fast — moderate reasoning", …), the current one ticked. The user picks "Low". The popover closes and the status bar reads "Low". The next message the user sends is answered at that level.

From the keyboard, Shift+Tab does the same thing one step at a time: the status-bar control flashes in the level's colour, a thin coloured line sweeps along the bottom of the taskpane, and a toast says "Thinking: Low (next turn)".

When the model reasons, its answer starts with a block labelled "Thinking…" with a spinner. The block stays collapsed; clicking it shows the reasoning as it streams. Once the model response is finished, the block reads "Thought".

## The interaction, event by event

```mermaid
stateDiagram-v2
    [*] --> set : tab created (High, or Off for a non-reasoning model)
    set --> popover : click the thinking control
    popover --> set : pick a level (level stored for next turn)
    popover --> set : Escape or click elsewhere (nothing changed)
    set --> set : Shift+Tab (next level, toast)
    set --> streaming_block : turn starts, model reasons ("Thinking…")
    streaming_block --> finished_block : model response ends ("Thought")
    finished_block --> finished_block : click header or Ctrl/Cmd+O (expand or collapse)
```

### Choosing a level

There are seven levels. Their labels and popover hints are:

| Level | Label | Hint |
| --- | --- | --- |
| off | "Off" | "Fastest — no reasoning step" |
| minimal | "Min" | "Quick — light reasoning" |
| low | "Low" | "Fast — moderate reasoning" |
| medium | "Medium" | "Balanced — solid reasoning" |
| high | "High" | "Slow — thorough reasoning" |
| xhigh | "Extra high" | "Very slow — extra-high reasoning" |
| max | "Max" | "Slowest — maximum reasoning" |

Which levels a model offers comes from the model registry bundled with Pi:

- A model that does not reason offers only "Off".
- A reasoning model offers "Off", "Min", "Low", "Medium" and "High", minus any level the registry marks as unavailable for that model; "Off" can be one of them. Confirmed in Excel on the web: GPT-6.1 Sol (ChatGPT login) offered "Min", "Low", "Medium", "High", "Extra high" and "Max", with no "Off".
- "Extra high" and "Max" are offered only by models the registry says support them.

The popover and Shift+Tab both act on the active tab only. Shift+Tab moves to the next offered level and wraps from the last back to the first offered one ("Off" where it is offered; confirmed in Excel on the web as "Max" → "Min" on GPT-6.1 Sol, each press with the toast "Thinking: {level} (next turn)"). If the tab's current level is not one the model offers (for example the tab was on "High" and the model was switched to one that does not reason), Shift+Tab goes to the first offered level and the popover shows no tick.

A new tab starts at "High" if its model reasons and "Off" if not; it does not copy the active tab's level (confirmed in Excel on the web: "+" from a tab on "Min" opened a tab on "High"). A [duplicated or forked](../sessions/tabs.md) tab copies its source's level. A tab whose model is replaced because it disappeared from the provider gets "High" or "Off" by the same rule.

### Ends at once

- **Picking the level that is already set** closes the popover and changes nothing.
- **Escape**, or a click anywhere outside the popover, closes it without a change. The popover claims Escape, so Escape does not also stop a responding tab.
- **No active tab** (only during startup or teardown): clicking the control shows the toast "No active session". Shift+Tab does nothing and moves focus as it normally would.

Nothing is saved in these cases.

### The level is set

Picking a level or pressing Shift+Tab changes the tab's level immediately: the status bar shows the new label at once. The popover gives no other feedback; Shift+Tab adds the flash and the toast "Thinking: {level} (next turn)" for 1.5 seconds.

The new level is used from the next turn: the next message the user sends, or the next queued prompt. A turn that is already running keeps the level it started with for all of its model responses, including those after steers and follow-ups, as the toast's "(next turn)" says.

> Technical note: pi-agent-core reads the model and thinking level once when a run starts (`createLoopConfig` in `agent.js`) and Pi's between-response hook does not change them. Confirmed in Excel on the web: Shift+Tab to "Extra high" during the first model response of a turn with two tool calls left all three requests at "high".

The level is passed to the provider as a request for that much reasoning. Providers map it onto their own settings and, where a model does not support the requested level, use the nearest one it does; the status bar keeps showing the stored level.

### While the model reasons

When readable reasoning starts, a thinking block says "Thinking…" with a spinner. It stops saying that as soon as reasoning ends, even while text or tools continue. Its button header expands the reasoning. Initial expansion follows the remembered details mode; an individual toggle overrides that block until the next global toggle.

The thinking level control stays usable throughout. The working indicator's rotating hints include "⇧Tab to adjust reasoning depth".

Some providers return no readable reasoning even at a high level (only a summary, or nothing). An empty reasoning block is not shown at all; the user sees the answer with no thinking block.

### The block after the model response

After reasoning ends, the header says "Thought for {duration}" if its start was observed in this taskpane, otherwise "Thought". The duration and disclosure choice survive the streaming-to-transcript handoff and tab switches. Ctrl/Cmd+O controls thinking blocks, cards and groups, including content that appears later. Collapsed blocks keep their normal card style.

Restored sessions whose reasoning start was not observed read "Thought"; saved text does not invent a timing measurement.

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | Decides supported levels and readable reasoning. A model switch clamps the stored level to one the new model supports. | Busy picker is read-only. |
| Thinking level | This is the feature. | Changing it is allowed; the running turn keeps its level, the next turn uses the new one. |
| Execution mode | No effect. | No effect. |
| Files in the workspace, or an overlay open | Shift+Tab cycles only in the composer with no modal open. In dialogs it moves focus backwards normally. | Same. |
| Tab state (new, restored, another tab busy) | New tab: "High" or "Off". Restored tab: its saved level. Another tab being busy does not matter. | Switching tabs shows the new active tab's level; Shift+Tab and the popover act on the active tab only. |
| Host (desktop or web) | No difference. | No difference. |

A level change never affects a model response or turn already running.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Escape closes the popover if it is open. | Stop ends the turn; the partial thinking block stays, labelled "Thought". With the popover open, Escape closes the popover and does not stop. |
| The user sends something else | The next turn uses whatever level is set when it starts. | Steers and follow-ups join the running turn at its original level. |
| Tab switched, created, closed or reopened | Each tab keeps its own level. Clicking another tab closes the popover. Switching tabs from the keyboard leaves it open, still listing the previous tab's levels; picking one then sets it on the tab where the popover was opened. | The running turn keeps its level in its own tab. |
| An overlay or menu opens | Opening the other status popover closes this one. | No effect on the running turn. |
| Model or thinking level changes | Applies to the next turn. | Model refused; level applies to the next turn. |
| The provider fails | No effect on the stored level. | The turn fails; a partial thinking block stays in the transcript. |
| The taskpane goes away | A level changed since the last saved message is lost; the session reopens with its saved level. | The streaming thinking block is lost with its model response. |
| The workbook changes underneath | No effect. | No effect. |
| Focus moves to the grid or another window | Shift+Tab in the grid is Excel's; the popover stays open until the next click in the taskpane. | No effect. |
| The context window fills | No effect. Compaction always summarizes with the "high" reasoning setting, whatever the tab's level. | Reasoning text counts towards the context window like any other text. |

After any interrupt the tab keeps whatever level it had.

## Interactions with other systems

**Persistence.** The level is saved with the session each time a message completes, and restored when the session is reopened or resumed. Changing the level does not save the session by itself; see [persistence and restore](../cross-cutting/persistence-and-restore.md). Thinking blocks are saved as part of their model response.

**Workbook changes and recovery.** No interaction.

**Other tabs and sessions.** Each tab has its own level. Two tabs on the same model can run at different levels.

**Execution mode and approvals.** Cycling thinking preserves the actual mode and lock badges.

**Models and providers.** The offered levels come from the model; see [models and providers](../foundations/models-and-providers.md). Higher levels make responses slower and use more tokens.

**Offline and provider errors.** No interaction.

**Notifications.** Shift+Tab: toast and flash. The popover: none.

**Settings and preferences.** There is no default thinking level in Settings.

**Host differences.** None.

**Language.** Level labels, hints, the popover title, the toast, "Thinking…", "Thought for {duration}" and "Thought" are translated.

**Accessibility and keyboard.** Status control, levels and thinking headers are buttons. Thinking headers expose `aria-expanded` and can be toggled with Enter/Space. Shift+Tab cycles only from the composer, consuming the key there; elsewhere it retains normal reverse-focus navigation. Modal focus stays trapped inside the modal.

## Edge cases

- With a model that offers only "Off", the popover lists one item and Shift+Tab keeps showing "Thinking: Off (next turn)".
- A non-reasoning model switch clamps the level to "Off".
- Several thinking blocks can appear in one turn, one per model response.
- Pressing Shift+Tab rapidly cycles through levels; each press replaces the previous toast.
- The thinking flash uses one colour per level (grey for Off through magenta for Max); the colours are decorative and carry no other meaning.

## Open questions and verification

- Merged code uses shared transcript state for duration/disclosure and a button header. The old duplicate rendering, lost timing and global Shift+Tab defects are historical verification findings, not current behavior.
- A level change made with no message afterwards is not saved; read from `src/taskpane/sessions.ts:344-352`, not confirmed by hand.
- Which providers return readable reasoning, and at which levels, was not checked.

Verified against pi-for-excel commit `1f51cf8` (code review).
