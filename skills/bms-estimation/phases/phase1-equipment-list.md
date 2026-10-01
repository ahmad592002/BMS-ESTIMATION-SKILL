# Phase 1 - EquipmentList

Produces: `EquipmentList` filled. Columns `SN | Qty | Equipment | Function | Type`.

This sheet is the single source of truth. Every later phase derives from it, so an error here is
embedded in the BOQ by Phase 7.

## Rules

- **One row per equipment family, not per tag.** `Exhaust Fan` qty 23, not EF-01…EF-23.
- **`Qty` from the schedule/drawings**, verified against two sources where possible. Any disagreement
  between the mechanical schedule and the client BOQ is a question, not a judgement call.
- **`Function` is exactly `Monitor` or `Monitor&Control`.** Monitoring-only equipment gets status and
  alarm DIs and no outputs - this one word decides the IO count.
- **`Type` must come from the `TemplateLists` picklist**: `Chilled Water System`, `Air Handling Unit`,
  `Fan Coil Unit`, `VAV Box`, `Fan`, `Pump`, `Electrical Panel`, `Tank`, `Low Current Systems`,
  `Meters`, `Other`. Wrong Type = wrong IO template = wrong controller.
  - Estimators often use short forms in practice (`AHU`, `VAV`, `Chiller`, `CWS`,
    `Mechanical Equipment`). Follow the existing sheet's convention rather than rewriting it, but keep
    it consistent within the project.
- **Keep duty/standby notation** from the schedule (`(2D/1S)`) - it drives DI/DO counts.
- **Carry the identifying data in the name** where it drives later selection: CHW pipe size and VFD
  rating (`AHU-12B (CHW DN125 / VFD 37 kW)`). Phase 6 reads these back for valve and VFD sizing.
- **Cover all disciplines, not just HVAC**: tanks, water treatment/RO, boilers, diesel, LPG, CO/gas
  detection, lighting, MDB/SMDB/PFC, generator, ATS, UPS, low current, CCTV, fire alarm interface -
  **but only where the tender asks for them.** If the client BOQ covers HVAC and mechanical only, say
  so explicitly rather than adding electrical scope nobody requested.
- **Monitoring-only system interfaces** (CCTV, FA, low current) are normally qty `1` per system.

Ask before inventing equipment that is implied but not scheduled.

## Exit criteria

- Every scheduled item is either a row or a listed open question.
- No `Type` outside the picklist / project convention; no blank `Function`.
- Quantities reconciled against the client BOQ, with every difference listed as a question.

## Present

The full table (SN, Qty, Equipment, Function, Type) plus a subtotal per discipline. This is the phase
the estimator checks hardest - make it easy to scan.

## Proven workflow (Al Moosa University, 2026-10 - estimator: "very good way")

1. **Read the client BOQ first** (`dump_sheets.ps1`). It often prices only controllers, field
   devices and LS items - no equipment counts. Equipment then comes from the drawings.
2. **The BMS system-architecture risers (B-92) are the equipment source.** Render each A0 riser
   as a 4x3 grid of tiles at ~140 dpi (`pdftoppm -x -y -W -H`) and read them visually - extracted
   text scrambles which item hangs on which panel. Use `pdftotext -raw` + `grep -o` only to
   cross-check repeated symbols (FM200, FCS, LIFT, "xNOS").
3. **Write one EquipmentList row per family per building** (`ATC - ...`, `EEC - ...`), tags in the
   name. Head-end interfaces (FA, LCS, ELV, VT, CBS) go under the building that holds the servers -
   never a 4th "Campus" building.
4. **Build the per-tag takeoff workbook** (`scripts/build_takeoff.ps1` from `panels.txt` +
   `items.txt`): every tag -> DDC/RIO/PLC -> room -> interface. Untagged groups ("AHU 5NOS") become
   `AHU R01-07 #1..5`; duplicate symbols get qty 0 (grey) so they stay visible but uncounted.
   Sheets: Equipment Register, Per DDC Panel, Summary (incl. drawing vs BOQ), Drawing Issues.
5. **Sync the EquipmentList to the takeoff** and check the totals match per building.
6. **Follow the workbook's own Type list** (`Lists!F`: AHU, CWS, FCU, VAV, Fan, Pump, ...), not the
   long names above, when the dropdown uses it.
7. **Yellow** the qty cell (B) or name cell (C) of every row whose quantity or identity is not 100%
   from the data, reason in EquipmentList column F (see LEARNED "Mark in yellow").
