# Phase 7 - BOQ, Breakdown, Cover Page

Produces: the priced offer.

## BOQ

Run `BOQModule.StartBOQGeneration` (it clears and rebuilds; take a backup first, and allow a few
seconds - reading the sheet mid-run shows it empty).

Columns: `Item | Category | Model number | DESCRIPTION | QTY | Manufacturer | Origin | unit list |
total list | unit net | margin | total net`. Quantities come **from the summary sheets, never
retyped**. Margin factor sits in `M2`; pricing resolves through `Product_Finder` / `Listprice` /
`Pricelist` - **never hardcode a price over a formula**.

Totals sit just below the last item row (`J110..J115` on a ~100-line BOQ; the row moves with the
line count - find the TOTAL label, and check `Breakdown!D9` points at the selling total).
Every cost cell K must stay `=VLOOKUP($D<r>,Pricelist!$A$2:$E$1000,3,FALSE)`; new models need a
Pricelist row (net price) first; items priced 0 on purpose need a Pricelist row at UP 0.

**After generating, check nothing was dropped.** Compare the BOQ against the selection sheets - a
valve or actuator selected in Phase 6 can fail to reach the BOQ. Reconcile part by part, not by
eyeballing the total.

## Product Finder - carry every Siemens part across

**This step comes right after the BOQ and before the Breakdown means anything.** The Product Finder is
the Siemens quotation: its total `K14` is what `Breakdown!D23` reads, so until it holds the current
BOQ the cost is stale and every margin is wrong.

Take **every Siemens part code and quantity from the BOQ** into the item table:

- The item table starts at **row 17**. Write **only two columns**:
  - **C = `Art.Type`** - the part code exactly as the BOQ's `Model number` (e.g. `VVF42.65-50`,
    `G120P-45/32A`, `TXM1.16D`)
  - **E = `Qty`**
- `D` (product number), `F` (list price), `G` (customs %), `J`, `K` and `L` (line total) are **lookup
  formulas against `Listprice_...`** and fill themselves. Never type over them.
- A code that resolves to a product number is valid; `#N/A` in D means the code is not in the price
  list - fix the code rather than forcing a price.
- **Only Siemens parts.** Non-Siemens lines (AX field devices, Onicon meters, Al Fanar enclosures,
  local items) are not in the Siemens quotation and belong in the Breakdown's other supplier rows.
- Rows may already be part-filled from an earlier run, with **gaps** - scan from row 17 for the first
  empty `C` rather than appending at the end, and never leave a duplicate code.
- Use `scripts/pf_refill.ps1`: clears C17:E480 and writes every Siemens BOQ line as a Value2 array
  (a single-cell integer write can fail with `Unable to cast Int32 to String`).
- **`G11` (Discount %) stays empty / 0.** The Breakdown applies the Siemens discount (E23). Both set
  = double discount and a fake ~55% margin.

Check afterwards: `G13` = number of line items, `K14` = Siemens list total. `G13` should equal the
count of distinct Siemens models in the BOQ.

**Price chain** (answer "why is this price different?" from this): Pricelist UP = net (~list x 45%) ->
BOQ cost; Product Finder K = list x (1 - G11) x 1.08 customs; Breakdown B.1 = K14 x (1 - E23); BOQ
selling = Pricelist UP / 0.6. BOQ Siemens cost vs Breakdown B.1 differ by roughly the customs.

## Breakdown

Project header (Project Name, Location, Reference, Division `Automation`, Customer, System `BMS`,
Total Value, SOW, Status, Discount, Net Value) and cost sections A–H. Values go in **column D**, not C
(C is part of the merged label). `Cover Page!H5` is the single source for the project name and feeds
`Breakdown!D6`.

**The margin trap - check this before quoting any number:**
`Breakdown!D23` (Siemens selling cost) reads `=Product_Finder_...!K14`, **not** the BOQ. It does not
follow a BOQ regeneration. After adding scope, the Product Finder must be refreshed with the new items
or the Breakdown understates cost and overstates margin. No macro does this - it is an estimator step.
**Cross-check: Product Finder lines == BOQ Siemens lines (part + qty), G11 = 0, and Breakdown B.1 is
within ~10% of the BOQ Siemens cost (customs). Anything else and the margin shown is wrong.**
Never rewire D23 to hide the gap.

**Sections E and F generate nothing - fill them from the nearest comparable project.**
Rates: LEARNED G "Breakdown expenses". Durations are the estimator's call - ask, and say which numbers
are historical rates and which are your own judgement.

## Cover Page

From `German Technical Services`, To (contact + company), Date, Ref as `BMS-R00-<Mon><Year>`, subject,
the standard 30-day-validity Siemens Desigo/BACnet paragraph, headline prices in **SAR excluding VAT**,
then notes, inclusions, exclusions, payment/delivery/warranty terms.

**Two checks that catch real contract defects:**

1. **Every note must match what the BOQ actually prices.** A note excluding what you charge for - or
   promising what you omit - is a contract defect, not a typo. Re-read all notes after any BOQ
   regeneration.
2. **Check for template carry-over.** Priced option lines survive from project to project (a
   "redundancy server" line appeared in three workbooks at an identical hardcoded value). Verify each
   against the tender; an identical figure in another project's workbook is the tell. Delete what this
   tender does not ask for - but a server the tender DOES ask for stays (one fault-tolerant server per
   campus unless the BOQ says per building; priced on its own Cover Page line).

Where the offer deliberately excludes something the tender asks for, state it as a **Deviations from
tender BOQ** list rather than leaving the client to discover it at evaluation.

## Exit criteria - the full final audit

1. `EquipmentList` qty == `DDC List` row 3 `Total` == row 4 `Assigned`, per column.
2. IOSummary "All Systems" == DDCSummary == DDCFullSummary.
3. Every panel's capacity >= demand including spare; no panel over 250 points.
4. Workstation licence tier matches total BA points; "Other" = 0.
5. Every BOQ quantity traces to a summary sheet; no orphan or hand-typed lines.
6. **Product Finder == BOQ Siemens lines, PF G11 = 0, Breakdown B.1 ~ BOQ Siemens cost** before any
   margin is quoted.
7. Cover Page total == Breakdown Net Value == BOQ total.
8. Notes and exclusions match the BOQ; deviations from the tender are declared.

9. No #N/A / #REF in BOQ, Breakdown, Cover Page, Product Finder; every BOQ cost cell a VLOOKUP.
10. Side files (IO Summary Sources, DDC List Sources, state file) regenerated from the final workbook.

Present the audit as a pass/fail checklist, then the headline price. For a final revision the
estimator asked for, go read-only and ask about each finding one at a time (LEARNED A).
