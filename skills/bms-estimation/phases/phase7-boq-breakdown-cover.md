# Phase 7 - BOQ, Breakdown, Cover Page

Produces: the priced offer.

## BOQ

Run `BOQModule.StartBOQGeneration` (it clears and rebuilds; take a backup first, and allow a few
seconds - reading the sheet mid-run shows it empty).

Columns: `Item | Category | Model number | DESCRIPTION | QTY | Manufacturer | Origin | unit list |
total list | unit net | margin | total net`. Quantities come **from the summary sheets, never
retyped**. Margin factor sits in `M2`; pricing resolves through `Product_Finder` / `Listprice` /
`Pricelist` - **never hardcode a price over a formula**.

Totals: `J110` total, `J112` total cost, `J113` Siemens, `J114` Al Fanar, `J115` local.

## Breakdown

Project header (Project Name, Location, Reference, Division `Automation`, Customer, System `BMS`,
Total Value, SOW, Status, Discount, Net Value) and cost sections A–H. Values go in **column D**, not C
(C is part of the merged label). `Cover Page!H5` is the single source for the project name and feeds
`Breakdown!D6`.

**The margin trap - check this before quoting any number:**
`Breakdown!D23` (Siemens selling cost) reads `=Product_Finder_...!K14`, **not** the BOQ. It does not
follow a BOQ regeneration. After adding scope, the Product Finder must be refreshed with the new items
or the Breakdown understates cost and overstates margin. No macro does this - it is an estimator step.
**Cross-check `BOQ!J112` against `Breakdown!I91`; if they differ, the margin shown is wrong.**
Never rewire D23 to hide the gap.

**Sections E and F generate nothing - fill them from the nearest comparable project.**
GTS rates seen so far: accommodation 3,000/month, rental car 2,500/month, car fuel 600/month,
air ticket 1,200 each, T&C Engineer 13,000, Technician 6,000. Set durations from the contract period
and site distance, and say which numbers are historical rates and which are your own judgement.

## Cover Page

From `German Technical Services`, To (contact + company), Date, Ref as `BMS-R01-<Mon><Year>`, subject,
the standard 30-day-validity Siemens Desigo/BACnet paragraph, headline prices in **SAR excluding VAT**,
then notes, inclusions, exclusions, payment/delivery/warranty terms.

**Two checks that catch real contract defects:**

1. **Every note must match what the BOQ actually prices.** A note excluding what you charge for - or
   promising what you omit - is a contract defect, not a typo. Re-read all notes after any BOQ
   regeneration.
2. **Check for template carry-over.** Priced option lines survive from project to project (a
   "redundancy server" line appeared in three workbooks at an identical hardcoded value). Verify each
   against the tender; an identical figure in another project's workbook is the tell. Delete what this
   tender does not ask for.

Where the offer deliberately excludes something the tender asks for, state it as a **Deviations from
tender BOQ** list rather than leaving the client to discover it at evaluation.

## Exit criteria - the full final audit

1. `EquipmentList` qty == `DDC List` row 3 `Total` == row 4 `Assigned`, per column.
2. IOSummary "All Systems" == DDCSummary == DDCFullSummary.
3. Every panel's capacity >= demand including spare; no panel over 250 points.
4. Workstation licence tier matches total BA points; "Other" = 0.
5. Every BOQ quantity traces to a summary sheet; no orphan or hand-typed lines.
6. **`BOQ!J112` reconciles with `Breakdown!I91`** before any margin is quoted.
7. Cover Page total == Breakdown Net Value == BOQ total.
8. Notes and exclusions match the BOQ; deviations from the tender are declared.

Present the audit as a pass/fail checklist, then the headline price.
