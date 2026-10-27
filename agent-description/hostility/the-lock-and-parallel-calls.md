# The lock and parallel calls

## Summary

Tool calls in a model response start in parallel, but workbook mutations serialize through one taskpane coordinator. Within one runtime, a read waits for mutations registered before it. It does not wait for later calls or another runtime's writes, so this is not snapshot isolation.

## Parallel calls in one batch

- Write followed by read: the read waits for the earlier mutation to settle and then reads Excel.
- Read followed by write: the read need not wait and can see the old value.
- Write beside write: mutations serialize in lock-acquisition order.
- Results retain call order even when execution completes in another order.

## Other tabs

Mutations share the taskpane write lock. Reads in another runtime can overlap a multi-batch write and observe an intermediate state. Separate taskpanes do not share this in-memory lock. A stopped waiter checks cancellation before applying anything.

## Racing approvals

Confirm approvals queue rather than replacing one another. Each identifies its tab/action, is asked before lock acquisition and has a cancellation signal. Stop withdraws outstanding requests. Decline is an explicit user action, not a side effect of another dialog opening. A pending approval does not hold the write lock.

## Classification quirks

- `execute_office_js` takes the write lock even for code that only reads.
- Recovery `delete` and `clear` do not need a workbook mutation lock, but always ask approval because they destroy recovery data.
- Checked-write preflight can refuse a blocked target before showing a mutation approval.

## Open questions and verification

Read from merged wrapper/coordinator and approval queue at `1f51cf8`. Cross-runtime and remote co-authoring are deliberately not serialized by this guarantee. Historical old-value and dialog-replacement observations remain in verification logs.

Verified against pi-for-excel commit `1f51cf8`.
