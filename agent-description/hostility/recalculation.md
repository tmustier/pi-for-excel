# Recalculation

## Summary

Read results report non-automatic calculation state and warn that cached values may be stale. Checked cell writes and formula fills calculate the formulas they wrote in Manual mode before read-back. They do not recalculate every dependent elsewhere in the workbook.

## What the agent receives (confirmed)

In ordinary checked-write results, "Verified values" refers to the written range. Under Manual mode the receipt additionally states the calculation scope and warns about stale dependents. A later explicit read reports Manual state rather than presenting cached values without qualification.

## What the agent does

It can report the calculated written formula, but must not infer that unrelated formulas, external links or all dependent cells have refreshed. For broader calculations it should request an explicit calculation with the appropriate tool and verify the result.

## Other pending-recalculation cases

- Automatic-except-tables can leave data tables stale.
- Volatile formulas can change between verification and a later read.
- External workbook references can retain cached values.
- Error scans cover the target, not every dependent or spill elsewhere.

## Proposed fix

Merged: workbook-state reporting and target calculation for checked writes. Global recalculation is not silently added.

## Open questions and verification

Read from merged read-state and range-write/fill implementations at `1f51cf8`. The historical Manual-mode zero and extra model calculation are pre-fix observations, not the present receipt contract. Native desktop calculation was not exercised here.

Verified against pi-for-excel commit `1f51cf8`.
