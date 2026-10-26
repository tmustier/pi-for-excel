# Recalculation

## Summary

In Manual calculation mode, formulas Pi writes and the cells that depend on what Pi changes keep their last calculated values until something recalculates. Every tool reads values, and none checks or reports the calculation mode. Verification "verifies" stale numbers, and the agent can report them as results.

## What the agent receives (confirmed)

With `calculationMode` set to Manual, `write_cells Calc!A3 = "=A1+A2"` (A1 = 1, A2 = 2) returned:

```
Written to **Calc!A3** (1×1)

**Verified values:**
| 0.0000e+0 |
| --- |

Changed cell(s): 1.
```

After A1 was changed to 100, `read_range Calc!A1:A3` still showed `0.0000e+0` for A3. Nothing in any result, in the blueprint or in the system prompt mentions the calculation mode.

## What the agent does

In a real turn under Manual ("put a formula in B1 that adds A1 and A2, then tell me the value"), the agent saw 0, distrusted it, used `execute_office_js` to read `calculationMode: "Manual"`, called `range.calculate()` on B1, and correctly reported 102. That took two extra calls. A less careful agent, or one that did not expect a non-zero result, would have reported 0. The zero rendered as `0.0000e+0` made the oddness more visible here.

## Other pending-recalculation cases

- `Automatic except tables` leaves data tables stale.
- Volatile functions (`NOW`, `RAND`) change on every recalculation, so read-back differs from later reads.
- Formulas referencing external workbooks show cached values.

## Proposed fix

[B-03](../bug-triage.md#b-03-verification-checks-only-the-written-cells-so-broken-spills-stale-formulas-and-invalid-entries-pass): load the calculation mode with every write; recalculate the written range or say plainly that the values are not current. Add the mode to the overview's state line ([B-11](../bug-triage.md#b-11-the-agent-cannot-see-workbook-state-that-changes-what-its-calls-do)).

## Open questions and verification

- Confirmed: RECALC-01, RECALC-02, one real turn.

Verified against pi-for-excel commit `ec453df`.
