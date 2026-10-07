# Learned rules - BMS estimation

Read this file at the start of **every** invocation of the `bms-estimation` skill. Rules here
**override** `SKILL.md` and the phase modules wherever they conflict.

When the estimator changes something you produced: **apply it, then ask why** (spec / GTS standard /
client preference / project condition / your error / commercial), and let the answer set the `Scope`.
No answer -> record with `Reason: not given`, `Scope: this project`, and never ask twice.
Edit an existing entry rather than adding a near-duplicate; a newer reason overrides an older one.
Confirm in one line: `Learned: <title> (scope: <scope>).`

Entry format (keep it short - the rule first, the evidence after):

```markdown
## <short rule title>
- **Rule:** <imperative, one to four lines>
- **Source:** <project> <YYYY-MM-DD> | <reason type> | scope: <always / this client / this project> | "<estimator's words>"
```

Cleaned up after Al Moosa University (2026-10-07): entries grouped by topic, duplicates merged, rules
the estimator confirmed as standing practice promoted to `always`. Full history: git log of the repo.

---

## A. Way of working

## One phase per invocation, then stop
- **Rule:** Complete one phase, present it, stop. Never chain phases; never mark DONE without approval
  or an explicit "next".
- **Source:** skill setup 2026-09-28 | GTS standard | scope: always | "finish one phase per time"

## Ask why a correction was made, then scope the rule
- **Rule:** After applying a substantive correction, ask once whether it came from the spec, GTS
  standard, this client, this project, your error or a commercial call; set `Scope` from the answer.
  Never generalise an unexplained edit.
- **Source:** skill setup 2026-09-28 | GTS standard | scope: always

## Self-upgrades go up as pull requests, never to main
- **Rule:** Batch a phase's rules on a `learn/<phase>-<slug>` branch and open a PR against `main` of
  ahmad592002/BMS-ESTIMATION-SKILL. Never commit to `main`, never self-merge. A failed push never
  blocks the estimation.
- **Source:** skill setup 2026-09-28 | GTS standard | scope: always

## Final review: read-only, one finding at a time
- **Rule:** When asked for a final check / full revision, go phase 0 -> 7 WITHOUT editing, and stop at
  each finding to ask (AskUserQuestion, recommended option first). Apply only what the estimator picks.
  Things the estimator says they will change by hand (Cover Page wording, SOW text) are theirs - list
  them, never edit them.
- **Source:** Al Moosa 2026-10-06 | GTS standard | scope: always | "ASK ME FOR EVERY ONE BY ONE"

## The estimator owns file names - never rename or move the live workbook
- **Rule:** Work on the file the estimator says is final, under the name they gave it. Do not rename,
  move or "promote" files on your own; set `$env:BMS_WB` to their path. Before a large edit, take a
  `SaveCopyAs` backup into `Old Versions\`.
- **Source:** Al Moosa 2026-10-06 | GTS standard | scope: always | "NO KEEP THE FINAL VERSION THE ONE I WANT"

## Only final files in the project root; everything else in "Old Versions"
- **Rule:** Backups (`SaveCopyAs`) and superseded deliverables go straight into `Old Versions\`. The
  root holds the inputs (BOQ, Drawings, Specifications), the live workbook, the current side files and
  `_ESTIMATION_STATE.md`.
- **Source:** Al Moosa 2026-10-01 | GTS standard | scope: always

## Every deliverable has a "where it came from" workbook - regenerate it at the end
- **Rule:** Keep `<project> - BMS Equipment Takeoff.xlsx`, `<project> - IO Summary Sources.xlsx`,
  `<project> - DDC List Sources.xlsx` and `_ESTIMATION_STATE.md` in step with the FINAL workbook. After
  the last edit, rebuild them (`refresh_sources.ps1`, `ddc_sources.ps1`) and check their totals equal
  the workbook (DI/AI/AO/DO/SP, panel count).
- **Source:** Al Moosa 2026-10-06 | GTS standard | scope: always | estimator chose "Regenerate for final"

## Mark in yellow everything not 100% from the project data - only the doubtful cell
- **Rule:** Any quantity, point, SP count, schematic choice, device or size not taken directly from
  this project's drawings / spec / BOQ gets `Interior.Color = 65535` + a one-line reason. Only the cell
  in doubt: SP -> G; inferred point -> its IO cell; device -> N; device qty -> M; equipment qty -> A;
  schematic choice -> name cell B once per block. Reason in IOSummary Q (never P - P is the protocol),
  EquipmentList F. Assumed SP counts and assumed valve/damper sizes are ALWAYS yellow.
- **Source:** Al Moosa 2026-10-01 | GTS standard | scope: always | "no every thing yellow like that"

---

## B. Quantities (Phase 1)

## Any quantity doubt -> take the HIGHER number, and say so
- **Rule:** Riser vs client BOQ, a group that may double count, a symbol drawn twice, a "duplicate"
  kept at qty 0 - always take the higher count. Yellow the qty with "riser X / BOQ Y - higher taken"
  and add one Cover Page note ("the riser quantity differs from the BOQ; we take the higher").
  Units above the riser count have no drawn location: the DDC builder puts them on the same
  panel/network after the drawn floors ("BOQ extra (location not drawn)").
- **Source:** Al Moosa 2026-10-05 | GTS standard (confirmed as standing rule) | scope: always |
  "take the highest number always" - e.g. MAB FCU 139 -> 348 (BOQ), VAV kept 691 (riser)

## Equipment comes from the drawings when the client BOQ has no equipment counts
- **Rule:** Client BOQs often price only controllers, devices and LS items. Build the EquipmentList
  from the BMS risers / schematics (per-tag takeoff workbook: tag -> panel -> room -> interface), then
  compare with any BOQ count using the "higher number" rule.
- **Source:** Al Moosa 2026-10-01 | project condition | scope: always (method)

## A duty/standby set with package tags is ONE equipment row at qty 1
- **Rule:** `BSP-01`, `BSP-02`, `BSP-SET-01` -> one row "... Set" qty 1 carrying all points; qty 2
  multiplies the shared package points.
- **Source:** RAPEH 2026-09-28 | my error | scope: always

## Client BOQ items not drawn on the risers still get points
- **Rule:** Read the client BOQ device lines against the IOSummary. Items the BOQ asks for but the
  risers do not draw get a new IOSummary line spread realistically over the panels that serve them
  (e.g. 32 room extra-low DP transmitters over the roof-AHU DDC panels). Devices that are by others
  (leak detection, security door contacts, HCHO) get monitoring points / integration only. Sensors the
  room thermostats or a multi-sensor already cover (zone T, T/RH, zone CO2, O2 via the AQ sensor) are
  not added twice. Declare each choice in a Cover Page note.
- **Source:** Al Moosa 2026-10-05 | GTS standard | scope: always | "ADD THEM IN THE IO SUMMARY ... DIVIDE THEM ON THE AHU TO BE REALISTIC"

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

## Group titles inside each IOSummary block
- **Phase:** 2 IOSummary   **Date:** 2026-09-28   **Project:** Riyadh Air Premium Hub
- **Rule:** Above each point group put a small shaded title row - "Digital Inputs", "Analog Inputs",
  "Analog Outputs", "Digital Outputs", "Software Points" - using `Interior.ColorIndex = 15` (the
  workbook's own `IsTitle` style). Title rows carry no IO value and no H:L formula.
- **Change:** sorted points with no headings -> 34 title rows across 23 blocks
- **Reason:** "not jyst like that en + add small title to every group"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** phases/phase2 - "Lay the block out cleanly"

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
