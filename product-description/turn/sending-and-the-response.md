# Sending and the response

Behaviour as of [PR #803](https://github.com/tmustier/pi-for-excel/pull/803), implementing CL-03, CL-05, CC-01 and CC-02. Historical review observations about rotating hints and duplicate error banners no longer describe this version.

## Summary

An accepted send puts the user's message into the tab's turn. Pi gathers workbook context, obtains credentials, asks the model, runs requested tools, and streams the answer until done, stopped or failed. The [composer](the-composer.md) owns drafting; [steering and the queue](steering-and-the-queue.md) owns messages sent while work is running.

## The simple case

Type a question and press Enter or click Send. The draft clears, its user message appears right-aligned, and Pi's answer streams below it as markdown. The tab shows its busy dot and the working line appears above the composer. Stop is available; typing another draft also reveals Send beside Stop.

When Pi finishes, the working line and Stop disappear. The answer stands in the transcript, the idle placeholder resumes, and the composer is ready for another message. Completed messages are saved with the session.

## While the turn runs

- **Enter steers**, as does the visible Send button. A steer is delivered after the current response and its tool batch; it does not interrupt or abort them.
- **Alt+Enter** queues a follow-up on Windows/Linux; **Option+Enter** does so on Mac. It is delivered when Pi would otherwise finish.
- **Shift+Enter** inserts a newline; **Shift+Tab** moves focus backwards, without changing thinking.
- Stop or an unclaimed Escape requests cancellation. [Stopping](stopping.md) describes retained text and undelivered messages.
- The busy placeholder says “Enter to steer · Alt+Enter for a follow-up” (Option on Mac), matching the Shortcuts page.
- Changing thinking is allowed, for the next turn. Changing the model while busy is not.

The existing submission path handles the accepted draft for keyboard and button sends alike. Slash commands are handled locally before ordinary message submission.

## The working indicator

The line says what Pi is doing, not just that it is alive:

| Activity | Left label | Hint |
| --- | --- | --- |
| Preparing/writing | Working… | Press Esc to stop |
| Reasoning with readable thinking | A playful thinking phrase, changing every 6 seconds | Press Esc to stop |
| Running a tool | Actual tool activity and target, for example Read Sheet1!A1:A3 | Press Esc to stop |
| Compacting | Compacting context… | Send messages and Pi will see them after compaction |
| Waiting for a transient retry | Retrying in {seconds}s (attempt {attempt} of {max})… | Provider hiccup · Esc to stop |
| Cancelling | Stopping… | Press Esc to stop |

Playful phrases are restricted to thinking; the hint no longer rotates. A queued/busy activity label takes precedence over model/thinking labels, including during mid-turn compaction. Stop is disabled after cancellation is requested, until the work settles. A tool already submitted to Excel may still finish.

## Preparing and streaming

Before each model response Pi reads [auto-context](auto-context.md), including workbook structure, selection and recent changes. It can compact older history before sending or between responses. A credential wait may open the provider connection dialog; connecting resumes the waiting turn. Cancelling that connection can fail the turn.

Text updates at most once per screen frame. Thinking blocks and tool cards appear alongside it. When a response asks for tools, those tools finish before the next model response. Steers appear as user messages between those responses; follow-ups come when the turn would otherwise end.

The transcript follows new content unless the user deliberately scrolls up. Returning near the bottom or sending a new prompt resumes following. Composer growth preserves the existing follow choice; expanded card bodies can scroll independently.

## Finishes

- **Done:** no further tool calls or waiting messages; the streaming cursor, working line and busy dot disappear.
- **Aborted:** the partial answer remains with “Request aborted”; no failure banner is raised.
- **Failed:** one inline plain-language explanation, Retry, a relevant fix action and collapsed diagnostic Details. No separate run-error banner. A failure belongs to its tab even if it happened in the background.

Transient provider failures are retried visibly and cancellably. Retry-After is honoured up to 60 seconds, usage-limit errors fail without waiting, and already-read partial answers are preserved and supplied back as continuation context. See [provider errors and offline](../cross-cutting/provider-errors-and-offline.md).

## Persistence, safety and host scope

Completed messages and stopped/failed responses are saved. Closing the taskpane can lose an answer still streaming. A plain answer changes no cells; completed mutations retain their [automatic backups](../workbook/backups-and-revert.md), and neither failure nor Stop undoes them.

Tabs stream independently. Switching tabs changes which transcript and working activity are visible, without pausing background work. Overlays generally leave the turn running; credential and approval dialogs are waits the turn may depend on.

UI labels and platform-specific hints are translated. Send and Stop have accessible button names. The composer and response semantics are shared by desktop and web hosts; proxy requirements and workbook batch cancellation differ by host.

## Verification

PR #803's composer browser tests exercise Enter, Alt+Enter, visible Send, Stop/Stopping and inline Retry. Real Excel-web turns with GPT-6 Luna steered input values to 4 and 7, kept the SUM formula at 11, delivered a follow-up and a visible Send, and stopped a real turn. Failure and Retry-After acceptance used substituted HTTP failures followed by real GPT-6 Luna responses. A short held-provider recording captures keyboard and cancelling states without depending on model latency.
