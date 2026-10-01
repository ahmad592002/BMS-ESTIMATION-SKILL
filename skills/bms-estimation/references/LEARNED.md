# Learned rules - BMS estimation

Read this file at the start of **every** invocation of the `bms-estimation` skill. Rules here
**override** the generic guidance in `SKILL.md` wherever the two conflict.

When the estimator changes something you produced: **apply it, then ask why** (spec / GTS standard /
client preference / project condition / your error / commercial), and let their answer set the `Scope`.
No answer given -> record with `Reason: not given` and `Scope: this project`, and never ask twice.
See the "Self-upgrade: ask WHY, then record" section of `SKILL.md` for the full protocol.

Edit an existing entry rather than adding a near-duplicate; a newer reason overrides an older one.
Confirm in one line: `Learned: <title> (scope: <scope>).`

Entry format:

```markdown
## <short rule title>
- **Phase:** <number/name>   **Date:** <YYYY-MM-DD>   **Project:** <name>
- **Rule:** <the rule, imperative, one or two lines>
- **Change:** <what was changed - from -> to>
- **Reason:** <their answer, in their terms> | not given
- **Reason type:** spec | GTS standard | client preference | project condition | my error | commercial
- **Scope:** always | this client | this project type | this project
- **Also changed in SKILL.md:** <section, or "no">
```

---

## Estimation runs one phase per invocation
- **Phase:** all   **Date:** 2026-09-28   **Project:** (skill setup)
- **Rule:** Complete one phase, present it for review, then stop. Never chain two phases in one turn,
  and never mark a phase DONE without the estimator's approval or an explicit "next".
- **Change:** skill rewritten from a single end-to-end run -> gated phase-by-phase run
- **Reason:** "i want the skill like every time i use it it finish one phase per time like first get
  all equipement i check if i find thing not good i ask"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** whole structure (phase gates, state file)

## Ask why a correction was made, then scope the rule accordingly
- **Phase:** all   **Date:** 2026-09-28   **Project:** (skill setup)
- **Rule:** After applying any substantive correction, ask once whether it came from the spec, GTS
  standard practice, this client, this project, your error, or a commercial call - and set the recorded
  rule's `Scope` from that answer. Never generalise an unexplained edit.
- **Change:** silent recording of corrections -> ask the reason, then record with a reason type
- **Reason:** "and it ask me if i update why i update and based on my answer update the skill"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** "Self-upgrade: ask WHY, then record"

## Self-upgrades are published as pull requests, never pushed to main
- **Phase:** all   **Date:** 2026-09-28   **Project:** (skill setup)
- **Rule:** After recording a rule in this file, batch that phase's learned rules onto a
  `learn/<phase>-<slug>` branch and open a PR against `main` of
  ahmad592002/BMS-ESTIMATION-SKILL. Never commit to `main` directly and never self-merge - the
  estimator reviews. The rule applies locally at once; a failed push never blocks the estimation.
- **Change:** rules saved only to the local skill folder -> local save plus a PR for review
- **Reason:** "the skills in it update should pull request to the repo"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** "Self-upgrade: ask WHY, then record" -> new subsection
  "Then publish the change as a pull request"

<!-- New entries go below this line -->

## No DDC panel may exceed 250 points
- **Phase:** 3 DDC List / 4 DDCSummary   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Cap every DDC at 250 points - physical (DI+AI+AO+DO) plus software points - and test the
  cap with the `Options` spare factor applied, not on raw counts. Split the panel rather than exceed
  it. Report each panel's loading against 250 whenever Phase 3 or 4 is presented.
- **Change:** panel sizing judged only by module count and enclosure practicality -> explicit 250-point
  ceiling checked including spare
- **Reason:** "you should know every ddc should not have more then 250 point"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** Phase 3 - DDC List (panel loading rules)

## Standardise field devices on parts already used in other GTS projects
- **Phase:** 2 IOSummary   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Before pricing, compare every field device against the other BMS_PROJECTS workbooks. A
  model appearing in only one project is a candidate for substitution by the house-standard part for
  the same duty. Match by FUNCTION: status/proof points take D.P.S. switches (QBM81-x), measured
  values take transmitters (QBM3020-x, QBE3000-D16). Never swap where the SOW or physics dictates the
  part - raise it as a question instead.
- **Change:** SDA-only devices -> house standard: 28x QBM3020-25 -> QBM81-5 (fan airflow proof),
  44x QBM3020-10 -> QBM81-10 (filter DP), 5x PL-FD113 -> QBE3000-D16; kept 8x QBM3020-5 for duct
  static pressure under VFD control
- **Reason:** "in the filed device in the io summary i think some device are not used in any other
  project then change them to another thing used if is possible"
- **Reason type:** (pending - GTS standard or this tender?)
- **Scope:** always (provisional - confirm)
- **Also changed in SKILL.md:** no

## The model column in IOSummary is a VLOOKUP - edit the description, not the model
- **Phase:** 2 IOSummary   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Column O is `IFNA(VLOOKUP($N<row>, FieldDevices!A:B, 2, FALSE), ...)`. To change a device,
  write the exact `FieldDevices!A` description into column N and let O resolve; FieldDevices quantity
  counts then update themselves. Never overwrite O - it destroys the formula. Changing a point between
  status and measurement also means moving its 1 between column C (DI) and D (AI); H:L are formulas
  (`C:G * $A<block>`) and must never be typed.
- **Change:** n/a - workbook mechanics learned while editing
- **Reason:** discovered when swapping field devices; writing O directly would have broken the lookup
- **Reason type:** my error (avoided)
- **Scope:** always
- **Also changed in SKILL.md:** no

## Write to the workbook via Excel COM, with events disabled
- **Phase:** all   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Excel COM is available on this machine. Back the workbook up first, attach to the running
  instance if the file is already open (`GetActiveObject("Excel.Application")`), set `EnableEvents =
  $false` so no `Worksheet_Change` macro fires mid-write, edit, `CalculateFullRebuild()`, verify the
  gate, then Save. Never write these workbooks with a plain XLSX library.
- **Change:** content handed over as tables to paste -> written directly into the workbook
- **Reason:** Excel 16.0 present; the workbook was open in a live session during the edit
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## Drive the workbook's own macros - do not hand-build the generated sheets
- **Phase:** 4,5,6,7   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** The workbook carries purpose-built VBA for most phases. Run these instead of writing cells:
  `DDCSummaryModule.GenerateDDCSummary` then `SelectControllersForSheet` (Phase 4);
  `GenerateFullDDCSummary` (Phase 5); `DamperActuatorSelectionModule.LoadDamperActuators` +
  `DamperActuatorDefaultSelection`; `VASelectionModule.LoadValves` + `VADefaultSelection`;
  `VFDSelectionModule.LoadVFDs` + `SelectVFDs`; `WorstationModule.GenerateWorstation`;
  `BOQModule.StartBOQGeneration`. Save before each macro - some crash Excel (SelectVFDs did), and a
  saved file loses nothing. `SelectControllers` needs an argument; use `SelectControllersForSheet`.
  A Load* macro that fails once via COM often succeeds on a retry with the workbook freshly opened.
- **Change:** hand-building DDCSummary/BOQ -> running the workbook's macros
- **Reason:** the macros hold GTS's own selection logic; hand-built output would diverge from it
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## LoadVFDs defaults every drive to 0.75 kW - fill the real ratings before selecting
- **Phase:** 6 Ancillaries   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** `LoadVFDs` cannot parse kW from the equipment name, so every row loads as 0.75 kW / 1 HP.
  Write the true kW and HP into columns D and E from the equipment names or the SOW before running
  `SelectVFDs`, or every drive is selected one frame size. Check the same defaulting on any Load*
  macro: damper actuators load at the `Options` default duct size (2.5 m2), and valves at the default
  AHU valve size, so both need real duct areas / flow data to size correctly.
- **Change:** 9 VFD rows at 0.75 kW -> 7.5/15/7.5/11/37/30/15/30/45 kW, then selected as G120P
- **Reason:** discovered when the selection returned one frame size for 11 different drives
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** no

## Tag software points with their protocol in IOSummary column P
- **Phase:** 2 IOSummary / 6 Workstation   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** The Workstation sheet classifies software points with
  `SUMIF(IOSummary!P:P, <protocol>, IOSummary!L:L)`, where the key must match its own labels exactly:
  `MODBUS`, `BACNET/IP`, `BACNET/MSTP`, `MBUS`, `KNX`. An untagged SP row falls into "Other", so a
  whole project can show 0 on every protocol and the licence mix is then unverifiable. Tag every SP
  row in column P, and check "Other" reads 0 before pricing the workstation.
- **Change:** 879 SP all in "Other" -> 804 BACNET/MSTP (FCU, zone thermostats, VAV) + 75 BACNET/IP
  (chillers via Chiller Plant Manager)
- **Reason:** "you don't say the software point is mode bass or other thing check them"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## Every exclusion note must match what the BOQ actually prices
- **Phase:** 7 Cover Page   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** After any BOQ regeneration, re-read the Cover Page inclusions, exclusions and notes against
  the BOQ categories. A note that excludes what the BOQ charges for (or promises what it omits) is a
  contract defect, not a typo. Check the template's carry-over notes especially - they arrive from the
  previous project and describe its scope, not this one.
- **Change:** notes excluded damper actuators, valves and VFDs while the BOQ priced all three
  (26,986 + 105,058 + 117,720); rewritten to match
- **Reason:** "okay change the note", "see notes and other thing"
- **Scope:** always
- **Reason type:** GTS standard
- **Also changed in SKILL.md:** no

## Check the Cover Page for template carry-over from the previous project
- **Phase:** 7 Cover Page   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Priced option lines survive from project to project. Verify each against the tender before
  quoting it; an identical figure in another project's workbook is the tell. Delete what this tender
  does not ask for.
- **Change:** removed "Total Price for Redundancy server" SAR 52,857.14 (`=37000/0.7`) - absent from
  all 11 SOW pages and 43 BOQ items, and carrying the identical value in YALJ and RAPEH
- **Reason:** "the redundancy server needed or is telled in the data ? if not delete it"
- **Reason type:** spec
- **Scope:** always
- **Also changed in SKILL.md:** no

## Breakdown Siemens cost comes from the Product Finder, not the BOQ
- **Phase:** 7 Breakdown   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** `Breakdown!D23` is `=Product_Finder_...!K14`, so it does NOT follow a BOQ regeneration.
  After adding scope to the BOQ, the Product Finder must be refreshed with the new items or the
  Breakdown understates cost and overstates margin. No macro does this - it is an estimator step.
  Always cross-check `BOQ!J112` (total cost) against `Breakdown!I91` before quoting a margin; if they
  differ, the margin shown is wrong. Never rewire D23 to hide the gap.
- **Change:** flagged - BOQ cost 513,547.77 vs Breakdown 359,815.51 after valves and VFDs were added
- **Reason:** found while auditing; margin displayed 54.46% against a materially higher real cost
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** Phase 7 final audit

## Take project expenses and resource rates from previous projects
- **Phase:** 7 Breakdown   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Sections E (General Expenses) and F (Resources) generate nothing - fill them from the
  nearest comparable project. GTS rates seen so far: accommodation 3,000/month, rental car
  2,500/month, car fuel 600/month, air ticket 1,200 each, T&C Engineer 13,000 (in all three reference
  projects), Technician 6,000. Set durations from the contract period and site distance, and say which
  numbers are rates from history and which are your own judgement.
- **Change:** E and F all zero -> E 17,000 + F 19,000; project cost 323,816 -> 359,816
- **Reason:** "i think need to add something like the boq some thing like the the car ticket if needed
  take price of them from other project"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## Verify BA licence extensions - none below 2000 points, and the macro over-adds
- **Phase:** 6 Ancillaries (Workstation)   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** `CCA-CMPXL-BA` (compact) already includes 2000 BA points. Test `Total BA * (1 + spare)`:
  at or below 2000 the BOQ gets **no** `CCA-*-BA` extension line at all; above 2000, deduct the
  included 2000 and cover the remainder with `CCA-1000-BA` / `CCA-500-BA` / `CCA-100-BA`
  (`CCA-5000-BA` only on the non-compact route). Always read the parts list after
  `GenerateWorstation` and delete extensions that are not called for.
- **Change:** workstation licence accepted as the macro generated it -> extension lines verified
  against the 2000-point allowance every time
- **Reason:** "the extention is not needed if the number of point +20% less then 2000 if more we add
  item 100 point 500 200 1000 extra but the auto somtimes add one not needed" / "no i put nothing in
  this project because is less then 2000"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** phases/phase6-ancillaries.md - new "BA licence extensions" section

## Why SelectBALicenses over-adds (two VBA defects)
- **Phase:** 6 Ancillaries   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Two defects in `WorstationModule.SelectBALicenses` make it add unneeded extensions:
  (1) the deduction is guarded by `If (Compact And TotalBA >= 2000)`, so below 2000 - exactly when the
  compact licence already covers everything - the 2000 included points are never deducted and the full
  count falls into the extension ladder; (2) `BA1000 = RemainingBA / 1000` assigns a Double to an
  Integer, so VBA rounds rather than truncates (1.656 -> 2). Spare IS applied upstream at line 16
  (`TotalBA = TotalBA * (1 + Spare)`), so the ladder always works on the spared figure.
- **Change:** n/a - root cause of the over-add, found by reading the VBA
- **Reason:** traced after the estimator reported "the auto somtimes add one not needed"
- **Reason type:** my error (avoided)
- **Scope:** always
- **Also changed in SKILL.md:** no

## After the BOQ, carry every Siemens part into the Product Finder
- **Phase:** 7 BOQ   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** The Product Finder is the Siemens quotation and its `K14` feeds `Breakdown!D23`, so it must
  be repopulated from the BOQ after every BOQ generation. Item table starts **row 17**; write only
  **C = Art.Type** (the BOQ `Model number`) and **E = Qty** - D/F/G/J/K/L are lookups against
  `Listprice_...` and fill themselves. Siemens parts only. Scan from row 17 for the first empty C
  (earlier runs leave gaps); never duplicate a code. Through COM the qty must be written as a string.
  Verify `G13` (line count) equals the distinct Siemens models in the BOQ.
- **Change:** Product Finder left at 27 controller-era items while the BOQ held 38 Siemens models ->
  all 38 carried across; K14 433,649 -> 981,127, Siemens cost 216,825 -> 490,564, gross margin
  57.96% (fiction) -> 19.81% (real)
- **Reason:** "other then the workspace we generate the boq take all field code to the sheet of
  product finder and put all siemens product"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** phases/phase7 - new "Product Finder" section

## Check the BOQ against the selection sheets - it silently drops lines
- **Phase:** 7 BOQ   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** `StartBOQGeneration` does not necessarily carry every selected item. After generating,
  reconcile the BOQ part by part against `ValvesAndActuators`, `VFDs`, `DamperActuators` and
  `Workstation` - not by comparing totals.
- **Change:** found `VXF42.65-50` x2 and `SQL36E65` x3 selected on the valve sheet but absent from the
  BOQ, therefore unpriced
- **Reason:** spotted while reconciling Siemens parts into the Product Finder
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** phases/phase7 - BOQ section

## Build IOSummary with the sheet's macros, sort by type, one blank row per block
- **Phase:** 2 IOSummary   **Date:** 2026-09-28   **Project:** Riyadh Air Premium Hub
- **Rule:** Never write raw cells into IOSummary - it destroys the block design. Use
  `ClearIOSummary`, then `AddEquipmentFromEquipmentList(<EquipmentList Qty cell>, <row>)` per
  equipment, `AddIO` per point, and `InsertIOLine` when a block needs more than the template's 22
  point rows; insert and fill bottom-up. Then lay the block out: **sort the points by type**
  (DI, AI, AO, DO, SP) and **leave exactly one blank row** before TOTAL, deleting the rest.
  When trimming, never `ClearContents` across H:L - those are the `C:G * $A<qty row>` all-systems
  formulas and wiping them makes every block total zero.
- **Change:** hand-written cells, source order, ~150 surplus blank rows -> macro-built formatted
  blocks, points grouped by type, one spare row each (sheet 691 -> 544 rows)
- **Reason:** "you break the designe of the iosummary you need to just full it not break designe" /
  "delete the free row and keep one in every box like pump box one space and seperate io point by
  types"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** phases/phase2 - two new sections

## A duty/standby set with package tags is one equipment row, not one per pump
- **Phase:** 1 EquipmentList / 2 IOSummary   **Date:** 2026-09-28   **Project:** Riyadh Air Premium Hub
- **Rule:** Where the IO list carries a set-level tag alongside the individual units (`BSP-01`,
  `BSP-02`, `BSP-SET-01`), enter **one row at qty 1** covering the whole assembly and put all its
  points in that block. Entering qty 2 multiplies the shared package points.
- **Change:** booster / circulation / submersible pumps at qty 2 -> qty 1 "... Set"; Total BA
  591 -> 556
- **Reason:** caught when IOSummary totals exceeded the IO list's own count by exactly the set points
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** phases/phase2 - "Lay the block out cleanly"

## Group IOSummary points by component, not by point type
- **Phase:** 2 IOSummary   **Date:** 2026-10-01   **Project:** Al Moosa University (supersedes RAPEH point-type titles)
- **Rule:** Inside each block, put a shaded title row (`Interior.ColorIndex = 15`, bold) per COMPONENT -
  Supply Fan, Exhaust Fan, Dampers, Filters, Coils & Valves, Heat Recovery, Electric Heater, Sensors,
  Pump, Chiller, Cooling Tower Fan, Breakers & Protection, Power Metering, Status & Alarms, Software
  Integration - and order the points DI, AI, AO, DO, SP inside each title. An airflow-proving DPS goes
  under the fan it proves. Title rows carry no IO value.
- **Change:** titles "Digital Inputs / Analog Inputs / ..." -> component titles (196 blocks rebuilt)
- **Reason:** "i like the way you categorize them but instead of these subcategory i want to separate
  them based on damper (control position) fan ... like category not based on type of point"
- **Reason type:** GTS standard
- **Scope:** always (provisional - confirm)
- **Also changed in SKILL.md:** phases/phase2 - "Lay the block out cleanly"

## ClearIOSummary only deletes rows 1-2000
- **Phase:** 2 IOSummary   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** `IOSummaryModule.ClearIOSummary` (called by the EquipmentList button) deletes `A1:A2000`
  only. A previous IOSummary longer than 2000 rows leaves old blocks below that mix with the new ones.
  Delete the whole used range of IOSummary before pressing the button, then check block count ==
  EquipmentList rows.
- **Change:** 199 blocks (old leftovers) -> full clear first, 196 blocks
- **Reason:** block-count check failed on rebuild
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** no

## Build the IOSummary layout in the fill pass, never retrofit into filled blocks
- **Phase:** 2 IOSummary   **Date:** 2026-09-28   **Project:** Riyadh Air Premium Hub
- **Rule:** Decide titles, sort order and blank rows before writing, and apply them in the same pass
  that fills each block. Retrofitting layout into populated blocks fails: repeated `InsertIOLine`
  calls crashed Excel into AutoRecover (file reopened as `.xlsb`); a batched `Rows.Insert` plus
  clear-and-rewrite dropped point names and left BA at 554 instead of 556. To change the layout of a
  filled sheet, `ClearIOSummary` and rebuild. Save after every good state; on a failed write close
  **without saving** and reopen from disk rather than repairing a half-written sheet.
- **Change:** three failed retrofit attempts -> clean rebuild with titles written during the fill
- **Reason:** observed across three consecutive failures in one session
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** phases/phase2 - new section

## Equipment quantities come from the drawings when the client BOQ has no equipment count
- **Phase:** 1 EquipmentList   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** Build the EquipmentList from the drawings (BMS risers / schematics / schedules) - the client
  BOQ often prices only controllers, field devices and LS items and carries no count for chillers, AHUs,
  fans, pumps or boards. Where the BOQ does quantify a line that the drawing also counts (e.g. FCU / VAV
  unitary controllers), follow the drawing in the EquipmentList, note the BOQ figure in the row text, and
  re-check it in Phase 2. Also build a per-tag takeoff workbook (tag -> DDC/RIO/PLC -> room) when the
  estimator wants to see "every equipment and where it connects".
- **Change:** MAB FCU 348 -> 139 and VAV 449 -> 691 (BOQ -> drawing); chillers 3 -> 5; +3 refrigerant
  detectors; per-tag takeoff workbook produced
- **Reason:** "in the boq no count see it if you want"
- **Reason type:** project condition
- **Scope:** this project
- **Also changed in SKILL.md:** no

## The EquipmentList "Generate IO" button stops at row 120 - add the rest yourself
- **Phase:** 2 IOSummary   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** `EquipmentListModule.GenerateIOPointsFromEquipmentList` (the EquipmentList button) loops
  only `Do While equipment.Row <= 120` and steps `currentRow += 29`. Run the button first (as the
  estimator asks), then append blocks for rows 121+ with the same steps the macro uses: write IOTemplate
  A4 = qty, B2 = Type, B4 = name; copy IOTemplate rows 1:30; insert at the next currentRow; clear
  B2, A4:G26, M4:N26. Check block count == EquipmentList rows before filling points.
- **Change:** 119 blocks from the button -> 196 blocks (77 appended)
- **Reason:** "generate the io summary in the button in the equipment list then complete points"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## IOSummary H:L formulas are row-relative - rewrite them after inserting rows
- **Phase:** 2 IOSummary   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** Template formulas read `=IF(C5="","",IF(C5*$A4>0,C5*$A4,""))` - the qty reference `$A4` is
  column-absolute but ROW-RELATIVE, so copying a row to enlarge a block makes it drift ($A5, $A6...).
  After any row insert, rewrite H:L of every point row to the block's qty row (what `InsertIOLine` does),
  then verify per block: TOTAL C:G x qty == TOTAL H:L. A run that skipped this under-counted DI by 492.
- **Change:** drifting formulas in 16 large blocks -> all 2065 rows rewritten; totals reconciled
- **Reason:** totals check failed (DI 5137 vs 5629 expected)
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** no

## Calling workbook macros with arguments through COM can raise a VBA debug dialog
- **Phase:** 2 IOSummary   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** `$x.Run(...AddEquipmentFromEquipmentList, <Range>, 1)` raised a VBA error dialog that froze
  Excel automation. Run only parameterless button macros through COM; replicate argument macros with
  direct COM steps. Cast every value written to a cell (`[string]`, `[double]`) - passing a PowerShell
  wrapped value raises "Specified cast is not valid".
- **Change:** macro-with-arguments calls -> direct COM steps
- **Reason:** estimator reported "error debub"; build blocked until the dialog was closed
- **Reason type:** my error (caught)
- **Scope:** always
- **Also changed in SKILL.md:** no

## Field devices: take them from the drawn instruments, models as previous projects, M = total qty
- **Phase:** 2 IOSummary   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** Give a field device only to points whose instrument is drawn on the schematic (DPS, TT/TW,
  T/RH, DPT, PIT/PT, FS, LS, AQ, CO2, current transducer) and pick the model GTS used before: fan DPS
  QBM81-5, filter DPS QBM81-10, duct T&RH QFM2120, duct T QAM2112.040, duct static QBM3020-5, duct AQ
  QPM2100, water temp QAE2120.010, water DPT QBE3000-D16, water DPS PL-FD113, pressure QBE2003-P16, tank
  hi/lo AX-LS-FL-1HM / -1LM, level AX-UL-SEP380-2, room unit RDF440BN, CT CTD-C3H00-1; dampers
  [Damper Actuator] [Spring Return], PICV / FCU valve / butterfly to the V.A. selection sheet.
  No device for duct smoke detectors (by FA) or sprinkler flow switches (by fire fighting).
  **Column M is the TOTAL for all units** (YALJ, Qiddiya, Ajyad) because FieldDevices counts with
  `SUMIFS(IOSummary!M:M, IOSummary!O:O, B)` - write it as `=<per unit>*$A<qty row>`.
- **Change:** 240 device rows; FieldDevices e.g. RDF440BN 3 -> 940 once M multiplied by block qty
- **Reason:** "field device are included in the drawing ... fill them in io summary but based on the
  previous project choice"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## Keep only final files in the project folder - everything else goes to "Old Versions"
- **Phase:** all   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** Save every backup (`SaveCopyAs`) and every superseded deliverable straight into an
  `Old Versions\` subfolder of the project. The project root holds only the inputs (BOQ, Drawings,
  Specifications), the live workbook, the current deliverables and `_ESTIMATION_STATE.md`. When a
  deliverable is replaced, move the old one there (rename it to say what it lacks) and give the new
  one the clean name.
- **Change:** 7 backups + 2 superseded files moved to `Old Versions\`; "(with devices)" file renamed
- **Reason:** "i want to add folder and any file not the final version put it in the folder"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## Mark in yellow everything not 100% from the project data
- **Phase:** 1-2 (all phases)   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** Any quantity, point, software-point count, schematic choice or device that is not taken
  directly from this project's drawings / spec / BOQ gets a yellow fill (`Interior.Color = 65535`) and
  a one-line reason. Mark ONLY THE CELL IN DOUBT, never the whole row: SP count -> G;
  inferred / assumed point -> its IO cell (C:F); device type -> N; device qty -> M; equipment qty -> A;
  schematic choice / typical points / borrowed schedule -> the equipment name cell B once per block.
  Reason in column Q "Check note" - never in P, which must stay the exact protocol for SP rows.
  Equipment-level doubts (qty, which schematic applies) go yellow on A:B of the equipment row and on
  EquipmentList B:C with the reason in column F. Assumed SP counts are ALWAYS yellow.
 whole rows -> 104 single cells + 92 qty/name cells ("can you specify the something need to check no every thing yellow like that")
- **Reason:** "if anything are not 100% true from the data you have in this project or if you assumed
  the number of software point you should mention them or label them in yellow color to be checked
  by myself"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no

## Name every point by its component as drawn - never a bare "FILTER STATUS"
- **Phase:** 2 IOSummary   **Date:** 2026-10-01   **Project:** Al Moosa University
- **Rule:** When a strip/schedule label is generic (FILTER STATUS), take the component name from the
  equipment symbol it is drawn under on the same sheet, left to right: AHU/MAHU -> PANEL FILTER, BAG
  FILTER; ERU -> FRESH AIR PRE-FILTER, FRESH AIR BAG FILTER, EXHAUST AIR BAG / PANEL FILTER; ECU -> PRE,
  ELECTROSTATIC PRECIPITATOR, BAG, HEPA, CARBON. Same for other repeated generic labels.
- **Change:** 33 IOSummary rows "FILTER STATUS" -> named by filter type
- **Reason:** "for the filter you write only filter status you should write his type like bag filter prefilter"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no
