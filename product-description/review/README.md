# Behavior review

Four reviews asked whether the documented behavior (`main` plus the fixes in flight) is the best behavior for users:

- [core-loop.md](core-loop.md) (CL)
- [workbook-trust.md](workbook-trust.md) (WT)
- [setup-extensibility.md](setup-extensibility.md) (SE)
- [chrome-cross-cutting.md](chrome-cross-cutting.md) (CC)

Plain fixes were sent to the owning pull requests:

| Pull request | Fixes |
| --- | --- |
| #789 | CL-04 |
| #792 | WT-06 |
| #791 | WT-08 |
| #793 | CC-06, CC-09 |
| #794 | WT-02, WT-04, WT-07, WT-09 |
| #795 | CL-02 |

The calls below wait for the product owner.

## Product calls

Related proposals are grouped. Each line gives the proposal, then the reviewer's recommendation.

### Undo and consent

- **CL-01:** after each turn, show "Changed X, Y · Review · Undo this turn". Recommended: yes.
- **WT-01:** Auto asks before workbook changes that cannot be undone (structure deletes too large to back up, chart deletes, destructive Office.js). Recommended: yes.
- **WT-03:** backups also expire by age. Recommended: yes; the age is the call.
- **WT-05:** Confirm stops asking twice (dialog and chat) and gains "Allow for the rest of this turn". Recommended: yes.
- **SE-03:** changes Pi makes to rules, conventions and skills show an Undo toast, and ask in Confirm mode. This reverses #790's "not gated".
- **SE-08:** in Confirm mode, Files changes ask only where Auto would.
- **SE-04:** MCP tool calls ask on the first call per tool, and on every call in Confirm mode.
- **WT-10:** mutating tools reject addresses that do not name a sheet (option A).

### Turn and composer

- **CL-03:** one way to message Pi while it works: Enter, plus a Send button next to Stop. Drop Alt+Enter follow-up.
- **CL-05:** the working line says what Pi is doing instead of playful phrases.
- **CL-09 / CC-07:** give Shift+Tab back to focus movement and move thinking cycling elsewhere. This reverses #793 Q4.
- **CL-11:** a model chosen while busy applies from the next turn (instead of #795's read-only picker); remove the fork setting.
- **CL-10:** drop "Replace current tab" and `/resume-here`.
- **CL-12:** prune the command menu; refuse a command while busy only when it conflicts with the turn.
- **CC-01:** a failed turn shows once, inline, with Retry and a fix action; the run-error banner goes.
- **CC-02:** retry the way providers fail: honour Retry-After over about 60 s, and tell a transient 429 from a usage limit. The call is what happens to a partial answer the user already read.
- **CC-03:** with auto-compaction on, the meter stops telling users to `/compact`.

### Chrome

- **CL-08:** tabs show a model-written conversation title instead of "Chat N". This costs one small request per session.
- **CC-04:** cell references in Pi's answer text become clickable.
- **CC-05:** a Copy action on each answer; tables copy as tab-separated text.
- **CC-08:** one tab-shortcut set that works on every host; drop Ctrl/Cmd+Z as "reopen tab".
- **CC-10:** follow the Office theme by default, so dark mode graduates.

### Setup

- **SE-01:** a hosted endpoint for the subscription token exchange; until then, list API-key options first and label which ones need the helper.
- **SE-02:** no phantom model; the control reads "Connect a provider".
- **SE-05:** store workbook rules in the file.
- **SE-06:** remove or flag the Claude Pro/Max login (Anthropic's terms; not verified).
- **SE-07:** use the provider's native web search.
- **SE-09:** files dropped on the composer attach to the message.
