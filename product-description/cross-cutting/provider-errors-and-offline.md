# Provider errors and offline

Behaviour as of [PR #803](https://github.com/tmustier/pi-for-excel/pull/803), implementing CC-01 and CC-02. This supersedes the review's historical no-retry/error-banner observations.

## Summary

A failed turn belongs to its tab and appears once in that tab's transcript. The last failed response shows a plain-language explanation, **Retry**, a relevant fix action, and collapsed **Details** containing the provider diagnostic. There is no separate run-error banner. Proxy reachability and other persistent-condition banners are separate from a failed turn.

## The simple case

A provider rejects the login. The turn becomes idle and its inline explanation says “Your provider rejected the login or API key. Update it, then retry.” The user can choose **Update login or API key**, fix the credential, return, and press **Retry**. Retry continues the existing exchange; it does not add another copy of the user's message. A successful retry removes the failed response's error UI.

### Before the request

Pi resolves the tab's provider credential before sending. A missing credential can open the connection dialog; closing it ends that attempt with inline recovery guidance.

## The request fails

Pi owns one visible, cancellable retry loop rather than hiding retries inside each provider adapter. It retries transient rate limits (429), overload/server failures (including 529), and errors classified as retryable by the provider runtime. Ordinary authentication, permission, model and invalid-request failures do not automatically retry.

- Six retries, with fallback waits of 2, 4, 8, 16, 16 and 16 seconds.
- A valid **Retry-After**, in seconds or HTTP-date form, replaces the fallback delay and is capped at 60 seconds. A past date means no additional wait.
- The working line says “Retrying in {seconds}s (attempt {attempt} of {max})…” and “Provider hiccup · Esc to stop”. Stop cancels the wait and prevents the next request.
- A usage/billing/quota limit is not a momentary rate limit: codes such as `usage_limit_reached` or `insufficient_quota` fail without waiting. A short-term `rate_limit_exceeded` code takes precedence over the ChatGPT adapter's generic “usage limit” wording.
- Raw HTTP status, error code/type and Retry-After are captured before an adapter rewrites the error into friendly text. Without those facts Pi uses the provider runtime's classifier and the fallback delay.

When retries are exhausted, the turn shows its one inline failure and manual Retry remains available. Retry is disabled by the running state; clicking it twice cannot start two runs.

## The connection drops while streaming

Text the user already saw is kept as an interrupted response, not deleted when retry starts. Pi sends that text back as a text-only assistant response plus a continuation request at the message tail, so the provider can continue from it instead of silently starting the answer again. Incomplete tool calls are not replayed as executable calls. System prompt and tool-schema prefixes are unchanged.

A final failure still keeps the partial response and its diagnostic. Workbook edits already completed remain applied; Retry is not Revert. Use [Backups](../workbook/backups-and-revert.md) to reverse a completed mutation.

## Fix actions

| Failure | Next action |
| --- | --- |
| Missing, expired or rejected credential | Update login or API key |
| Unknown/unavailable model, quota or general failure | Choose another model |
| Connection, proxy or offline failure | Open proxy settings |

The summary is translated; raw provider text in Details remains exactly what the provider returned. Browser failures do not establish that CORS is the cause: a stopped proxy and a dropped network can look the same.

### The error banner

The old taskpane-wide run-error banner is removed. This heading remains for links to the historical review: the current replacement is the one inline failure described above, not a second notification.

## The context window overflows

With auto-compaction on, Pi attempts one compact-and-continue recovery for an overflow from the current model. If it cannot free context or the request overflows again, the final inline explanation offers the appropriate recovery. With auto-compaction off there is no automatic overflow recovery. See [context and compaction](../turn/context-and-compaction.md).

## Interactions with other systems

- A deliberate Stop is aborted, not a provider failure; the partial answer remains with “Request aborted”. A provider error merely containing “cancel” is not proof that the user stopped it.
- A hanging request has no taskpane-owned timeout; Stop can end it.
- A text-only response that reaches its output limit can simply end. Truncated tool calls are rejected instead of executed, and the model can reissue them.
- Web search, fetch and MCP errors are tool-result failures, not failed turns. Pi can keep working after those results.
- Background failures stay in their own transcripts. Switching tabs does not carry another tab's error over the active conversation.
- Local persisted sessions render offline. An uncompleted stream may be lost when the taskpane closes; completed partial responses and diagnostics are saved with the session.
- Other taskpanes and tabs using the same account still share the provider's external limits.

## Open questions and verification

PR #803 runs public Agent/action-queue tests for retry classification, 60-second cap, partial preservation and manual Retry; browser tests cover one inline failure, Details, fix action and successful Retry. Excel on the web with GPT-6 Luna verified a forced 401 followed by Retry and a successful real answer. A stubbed transient 429 with Retry-After: 7 waited 7.166 seconds before the next real request and returned READY. Desktop provider failure acceptance was not repeated for this PR.
