# Phase 5 - DDCFullSummary

Produces: the consultant-facing expansion - every individual point with its field device and model,
alongside the controller list, per panel.

## Run

```
DDCSummaryModule.GenerateFullDDCSummary
```

Save first. See `references/workbook-mechanics.md`.

## Gate

**DDCFullSummary point totals == DDCSummary == IOSummary "All Systems".**

If the three disagree, trace back to the source phase. **Never patch this sheet to make it match** -
it is the audit trail, and a patched total hides the real error upstream.

## Present

The three-way reconciliation as a table, plus the total point count per panel.
