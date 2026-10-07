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

## C. IOSummary (Phase 2)

## Generate with the EquipmentList button, then complete rows 121+
- **Rule:** `GenerateIOPointsFromEquipmentList` stops at EquipmentList row 120 (`currentRow += 29`).
  Run the button, then append blocks for rows 121+ with the macro's own steps (IOTemplate A4 = qty,
  B2 = Type, B4 = name; copy rows 1:30; insert; clear B2, A4:G26, M4:N26). `ClearIOSummary` deletes
  only A1:A2000 - delete the whole used range first. Check block count == EquipmentList rows.
- **Source:** Al Moosa 2026-10-01 | GTS standard + my error | scope: always

## Never write raw cells into IOSummary blocks; never retrofit layout
- **Rule:** Use the block design (macro-built blocks, `InsertIOLine` for extra rows), decide titles /
  order / blank rows before writing and apply them in the same pass. Retrofitting crashed Excel and
  dropped names. To change a filled layout: clear and rebuild. On a failed write close WITHOUT saving.
- **Source:** RAPEH 2026-09-28 | GTS standard + my error | scope: always

## H:L are row-relative formulas - rewrite them after any row insert
- **Rule:** `=IF(C5="","",IF(C5*$A4>0,C5*$A4,""))` - `$A4` drifts when rows are copied. After inserts
  run `fix_hl.ps1` (points every row at its block's qty row) and verify TOTAL C:G x qty == TOTAL H:L.
  Never `ClearContents` across H:L.
- **Source:** Al Moosa 2026-10-01 | my error | scope: always

## Group points by component, name them as drawn, no duplicate names in a block
- **Rule:** Grey bold title row (`ColorIndex 15`) per COMPONENT (Supply Fan, Exhaust Fan, Dampers,
  Filters, Coils & Valves, Heat Recovery, Sensors, Pump, Chiller, Breakers & Protection, Power
  Metering, Status & Alarms, Software Integration); DI, AI, AO, DO, SP inside each; one blank row
  before TOTAL. Name every point by its component (PANEL / BAG / PRE / HEPA / CARBON filter; SUPPLY /
  EXHAUST FAN AIR FLOW STATUS; HEAT RECOVERY WHEEL 1 / 2). Two rows with the same name in one block
  are a defect - make them distinct.
- **Source:** Al Moosa 2026-10-01 / 10-06 | GTS standard | scope: always | "you should write his type like bag filter prefilter"

## Point-row conventions
- **Rule:** Modulating damper: command + feedback on ONE row (AI + AO) with the actuator device.
  Solenoid valve = status DI + command DO. Air separator = status DI + common alarm DI. Speed-control
  AO carries no VFD device.
- **Source:** Al Moosa 2026-10-02 | GTS standard | scope: always

## FCU and VAV are ALWAYS integration points - 7 SP per unit
- **Rule:** One "Integration" row, 7 SP, protocol in P; no DI/AI/AO/DO and no hardwired devices.
  Keeps network panels small (~29 units per panel at 250 incl. 20% spare) and stops TX-I/O pricing.
  VAV rows also carry the unitary controller + room unit as device rows (Al Moosa: DXR1.M09PLZ-112 +
  QMX1.M34H - confirm per project). FCU: controller through RoomUnits/FCUs (see F).
- **Source:** Al Moosa 2026-10-02 | GTS standard | scope: always | "we always take them as integration point"

## Field devices: only where drawn, house-standard models, M = total quantity
- **Rule:** Device only where the instrument is drawn. Models GTS uses: fan DPS QBM81-5, filter DPS
  QBM81-10, duct T&RH QFM2120, duct T QAM2112.040, duct static QBM3020-5, duct AQ QPM2100, water temp
  QAE2120.010, water DPT QBE3000-D16, water DPS PL-FD113 (Sontay - not Siemens, keep it out of the
  Product Finder), pressure QBE2003-P16, float AX-LS-FL-1HM/-1LM/-2LH, level AX-UL-SEP380-2, room
  thermostat RDF440BN, room AQ IAQRM5XC (Greystone), CO/NO2 Greystone. No device for duct smoke
  (by FA) or sprinkler flow switches. Column O is a VLOOKUP on N - write the FieldDevices!A text in N,
  never type O. Column M = `=<per unit>*$A<qty row>` (FieldDevices sums M). A one-project-only model
  is a candidate for the house standard for the same duty.
- **Source:** SDA 2026-09-28, Al Moosa 2026-10-01 | GTS standard | scope: always

## Valves: two priced rows, PICV where the tender says PICV
- **Rule:** Never leave a valve on a "V.A. Selection Sheet" device (no price). Each valve = command row
  `[Water] [Valve]` + feedback row `[Water] [Valve Actuator]`, M = 1*qty on both. GTS models
  (Al Moosa final):
  | Duty | Valve | Actuator |
  |---|---|---|
  | AHU / FAHU / ERU cooling coil (drawings or BOQ say PICV) | VPF44.65F25 | SAX61P03 |
  | Solenoid / run-around / heat-recovery coil (globe) | VVF42.65-50 | SKB62/F |
  | Butterfly: HEX, CHW header, isolation (ICV) | VFW41.150 | SQL341E100 |
  If the drawings or client BOQ say PICV, never price a globe valve for that duty. Models typed in O
  count only if listed in the Valves sheet direct list (rows 62+, `=SUMIFS(IOSummary!M:M,
  IOSummary!O:O, A<r>)`) and present in the Pricelist. FCU valves come with the FCU.
- **Source:** Al Moosa 2026-10-02 -> 10-07 | GTS standard | scope: always | estimator picked the models

## Valve / damper sizes are almost never in the tender - say so, don't invent precision
- **Rule:** Search the drawings (pdftotext: DN, mm, L/s, kW, TR, GPM, m3/h) and the client BOQ before
  assuming. Al Moosa had none (valves were LS items). Keep one middle size (DN65 coils, DN150
  butterfly; damper 2.5 m2 default), mark every model yellow "size not in drawings", and list
  "sizes assumed - to be confirmed" on the Cover Page. Real sizes need the mechanical schedules.
- **Source:** Al Moosa 2026-10-06 | project condition (recurs) | scope: always | "Keep default 2.5 m²"

## Tag every software point's protocol in column P
- **Rule:** Workstation counts SP with `SUMIF(IOSummary!P:P, <label>, L:L)`; labels exactly `MODBUS`,
  `BACNET/IP`, `BACNET/MSTP`, `MBUS`, `KNX`. "Other" must read 0.
- **Source:** SDA 2026-09-28 | GTS standard | scope: always

---

## D. DDC List (Phase 3)

## No DDC over 250 points incl. spare - network rows too
- **Rule:** Physical + SP, with the Options spare applied, <= 250 for EVERY DDC List row, including
  "virtual" rows for IP unitary controllers / VRF gateways (else the BOQ shows "Need To Seperate").
  Split overloaded panels into separate panels `<drawn name>-1`, `-2`... filled floor by floor.
- **Source:** SDA 2026-09-28, Al Moosa 2026-10-02 | GTS standard (250 cap) | scope: always; the
  `-1/-2` naming and network-panels-per-floor were Al Moosa choices - confirm on the next project

## GTS uses PXC, never PLC - every panel is a DDCP
- **Rule:** Panels drawn as PLC / RIO (chiller plant manager, MV/LV PLC, remote IO) are priced as PXC
  panels named `DDCP-<bldg>-<level>-<next nn>`; equipment names say "via DDC". Keep a drawn-tag ->
  DDCP map (`panel_rename.txt`, project data) so `ddc_build.ps1` applies it.
- **Source:** Al Moosa 2026-10-06 | GTS standard | scope: always | "we use just PXC not PLC"

## DDC List template limits - extend the design past them
- **Rule:** Formatted only B:DQ x 81 rows; `GenerateEquipmentForDDCList` scans IOSummary to row 3000;
  `ClearDDCList` clears B:DU; `FillDDCEquipmentInCell` reads 150 columns. Past these, write the list
  directly (`ddc_build.ps1 -Write`) and extend borders, vertical headers, row-2 fill 14806254, the
  crosshair CF and the three row-4 Assigned-vs-Total CFs to the last column (`full_chain.ps1` does it).
  Rebuild CFs with delete + Add (`ModifyAppliesToRange` crashed Excel).
- **Source:** Al Moosa 2026-10-02 | GTS standard | scope: always | "the design not applicable on all column"

---

## E. DDCSummary / DDCFullSummary (Phases 4-5)

## SelectControllersForSheet works on the ACTIVE sheet
- **Rule:** Activate `DDCSummary` (or DDCFullSummary) before `SelectControllersForSheet`, then check
  Controllers!L > 0. Run from another sheet it selects nothing and the BOQ silently has no
  controllers. (`run_macro.ps1 -Sheet DDCSummary` handles it.)
- **Source:** Al Moosa 2026-10-02 | my error | scope: always

---

## F. Ancillaries (Phase 6)

## Every Load* macro loads defaults - fix them before selecting
- **Rule:** `LoadVFDs` sets every drive to 0.75 kW / 1 HP - write real kW/HP first. Dampers load at
  the Options duct area (2.5 m2) and valves at the Options size - replace with real data when it
  exists, otherwise keep the default and declare it.
- **Source:** SDA 2026-09-28 | my error | scope: always

## BA licence: no extension at or below 2000 points incl. spare
- **Rule:** `CCA-CMPXL-BA` includes 2000 BA points. `Total BA * (1 + spare)` <= 2000 -> no `CCA-*-BA`
  line; above -> deduct 2000 and ladder 1000/500/100. `SelectBALicenses` over-adds (deduction guarded
  by `>= 2000`; Integer rounding) - read and correct its output every time.
- **Source:** SDA 2026-09-28 | GTS standard | scope: always

## Workstation hardware follows the client BOQ head-end lines
- **Rule:** The template macro sets licences only; PCs, monitors, printers, UPS default to 1 and there
  is no console / wall display. Take the client BOQ quantities (higher-number rule) - Al Moosa: 8
  workstations (Clients B28 = 8), 2 laser + 2 dot-matrix printer sets, 2 consoles, 2 wall displays
  (macro extended to read Workstation!B36:B39). Consoles / wall displays priced 0 on purpose need a
  Pricelist row at UP 0 (see G).
- **Source:** Al Moosa 2026-10-05 | GTS standard | scope: always | "I PUT ONLY 8 WORKSTATION AND I FORGET LASER"

## FCU controllers through RoomUnits/FCUs - never a hand-typed model
- **Rule:** FCU device = `[FCU Controller]`; set Options B13 to the real protocol; `LoadRoomUnits`,
  choose the controller from the dropdown. The FCUs sheet prices controllers only. Never price two
  controllers on one FCU (RDF440BN is complete; DXR2 needs a QMX3). Al Moosa: RDF440BN alone, with a
  Cover Page note that it replaces the BOQ's "IP unitary controller".
- **Source:** Al Moosa 2026-10-05 | GTS standard (path) / this project (RDF440BN) | scope: always (path)

## One head-end server for the whole campus unless the tender says otherwise
- **Rule:** Read the client BOQ and risers for the server line. Al Moosa: ONE fault-tolerant
  rack-mounted server (ATC BOQ item, qty 1) for all 3 buildings; MAB workstations are "remote users
  of the servers located in ATC"; EEC has none. "Fault tolerant" = one machine with redundant
  internals, not two servers. Price it separately on the Cover Page (cost / 0.6). If the tender has
  no server line, delete the template's carried-over server line (SDA).
- **Source:** SDA 2026-09-28 (deleted), Al Moosa 2026-10-07 (kept 1) | spec | scope: always

---

## G. BOQ, Product Finder, Breakdown, Cover Page (Phase 7)

## How a Siemens price is built - three different numbers
- **Rule:** Know which price you are looking at before answering "why is it different":
  - `Pricelist` UP (feeds BOQ cost K) = NET price, about list x 45% (55% off), no customs.
  - Product Finder K (unit) = list x (1 - G11) x (1 + customs 8%) + freight -> K14 total.
  - Breakdown B.1 = Product Finder K14 x (1 - E23 Siemens discount 55%) = the cost used for margin.
  - BOQ selling = Pricelist UP / (1 - margin 40%).
  So BOQ Siemens cost and Breakdown B.1 differ by about the customs (~6-8%) - that is normal.
  Example SAX61P03: list 1,994.06; PF 2,153.58; real cost 969.11; Pricelist 902.50; selling 1,504.
- **Source:** Al Moosa 2026-10-07 | GTS standard | scope: always

## Product Finder discount G11 stays 0 - discount once, in the Breakdown
- **Rule:** With G11 = 50/55 AND Breakdown E23 = 55%, Siemens material is discounted twice and the
  margin shows ~53-56% instead of ~28%. Keep G11 empty/0; E23 holds the discount. Check G11 in every
  audit.
- **Source:** Al Moosa 2026-10-07 | my error (caught twice) | scope: always

## Refill the Product Finder after every BOQ change
- **Rule:** Breakdown D23 = Product Finder K14, which does NOT follow the BOQ. After any BOQ change run
  `pf_refill.ps1` (C17 part no + E17 qty for every Siemens BOQ line; D/F/G/J/K/L are lookups), then
  confirm PF lines == BOQ Siemens lines part by part and no #N/A in D.
- **Source:** SDA 2026-09-28, Al Moosa 2026-10-07 | GTS standard | scope: always

## Every BOQ cost cell is a Pricelist VLOOKUP
- **Rule:** BOQ K = `=VLOOKUP($D<r>,Pricelist!$A$2:$E$1000,3,FALSE)`, E/G/H the same lookup on
  columns 2/5/4. A Product Finder formula pasted into K refers back to its own row and loops (wrong
  cost, wrong selling). New models need a Pricelist row (net price, manufacturer) first. Items priced
  0 on purpose (operator console, wall display) need a Pricelist row at UP 0 / "Others", or the next
  `StartBOQGeneration` shows #N/A.
- **Source:** Al Moosa 2026-10-06 / 10-07 | my error (caught) | scope: always

## Reconcile the BOQ part by part against its sources
- **Rule:** `StartBOQGeneration` can drop selected items. After generating, compare every BOQ line's
  qty with Controllers!L, FieldDevices C, Valves direct list, DamperActuators, Enclosures and
  Workstation (`audit.ps1` / the cross-check in `revision_facts.ps1`), and selling = cost x qty / 0.6.
- **Source:** SDA 2026-09-28 | my error | scope: always

## Breakdown expenses and resources from GTS history
- **Rule:** Sections E/F generate nothing. Rates (SAR/month): per diem 3,000 and accommodation 3,000
  per person; rental car 2,500; fuel 600; air ticket 1,200 each; T&C engineer 13,000; design engineer
  15,000; supervisor 7,000; technician 6,000; draftsman 6,000; contingency 1.5%. The DURATION is the
  estimator's call - ask (Al Moosa: 10 months site, edited by hand). Never change the Siemens discount
  (E23) yourself - commercial.
- **Source:** SDA 2026-09-28, Al Moosa 2026-10-06 | GTS standard (rates) | scope: always

## Cover Page: notes must match the BOQ; carry-over lines removed
- **Rule:** Re-read every note after each BOQ regeneration: nothing excluded that the BOQ prices, no
  template text from the last project. Ref `BMS-R00-<Mon><Year>`, Company = client. Standard notes
  for this kind of tender: valve sizes assumed; meters by others, integration included; IO estimated
  where no schematic; higher of riser/BOQ taken; BOQ items not on risers and how they are covered.
  Price lines kept separate when asked (server; Al Moosa also split valves & actuators onto its own
  line). Payment / warranty / SOW wording is the estimator's.
- **Source:** SDA 2026-09-28, Al Moosa 2026-10-06 | GTS standard | scope: always

---

## H. Excel / PowerShell automation

## Excel COM only, events off, back up first
- **Rule:** Never write these workbooks with an XLSX library. Back up, attach, `EnableEvents = $false`,
  edit, recalc, verify, Save. Save before every macro (some crash Excel).
- **Source:** SDA 2026-09-28 | GTS standard | scope: always

## Drive the workbook's macros instead of hand-building generated sheets
- **Rule:** GenerateDDCSummary -> SelectControllersForSheet; GenerateFullDDCSummary;
  GenerateWorstation; LoadDamperActuators + DamperActuatorDefaultSelection; LoadValves (+
  VADefaultSelection); LoadVFDs + SelectVFDs; LoadRoomUnits; StartBOQGeneration. Run through
  `run_macro.ps1` (save before/after, waits for Ready, retries 0x800A9C68 once). Only parameterless
  macros via `$x.Run` - argument macros raise a VBA debug dialog.
- **Source:** SDA 2026-09-28, Al Moosa 2026-10-06 | GTS standard | scope: always

## Any upstream edit -> regenerate everything downstream, in order
- **Rule:** DDC List -> GenerateDDCSummary -> SelectControllersForSheet -> GenerateFullDDCSummary ->
  GenerateWorstation -> dampers -> valves -> VFDs / RoomUnits -> StartBOQGeneration -> pf_refill ->
  Breakdown / Cover recalc -> audit (`full_chain.ps1`). Never quote a price while any step is older
  than the last upstream edit. A point-NAME-only change needs just GenerateFullDDCSummary.
- **Source:** Al Moosa 2026-10-02 | GTS standard | scope: always

## Open Excel with Start-Process; macros act on the ACTIVE workbook
- **Rule:** If Excel is closed, `Start-Process <file.xlsm>` and wait until `GetActiveObject` lists it.
  `BindToMoniker(path)` on a closed file starts a hidden instance that dies with the script ("RPC
  server unavailable" next call). Template VBA uses unqualified `Sheets()` = ActiveWorkbook: activate
  the workbook (and sheet) before every macro and verify `ActiveWorkbook.FullName`; keep only one
  estimation workbook open while macros run (a chain once regenerated the wrong file).
- **Source:** Al Moosa 2026-10-06 | my error | scope: always

## When Excel stops answering
- **Rule:** "Call was rejected by callee", "Unable to set Calculation", or `Workbooks.Count = 0` on a
  visible window = a cell is in edit mode or a dialog is open: ask the user to press Enter/Esc, don't
  loop. Wrap `$x.Calculation` changes in try/finally. A workbook window can be hidden
  (`Windows(1).Visible = $true`). For a read-only check while the live window is busy, open the SAVED
  file in a separate `New-Object Excel.Application` with `AutomationSecurity = 3`, ReadOnly.
- **Source:** Al Moosa 2026-10-06 / 10-07 | my error | scope: always

## PowerShell traps that already cost real damage
- **Rule:**
  - Variable names are case-insensitive: `$K`/`$k`, `$V`/`$v`, `$bq`/`$BQ`, `$yel`/`$YEL` are ONE
    variable (hit again 2026-10-07). Use distinct names.
  - `[ordered]@{228='x'}` indexed with an int returns by POSITION -> null -> blank cells written (20
    point names blanked, 2026-10-06). Use arrays of pairs or string keys; refuse to write empty values.
  - `$host` is reserved; `cat` is Get-Content; index as `$v[($r+2),2]`, not `$v[$r+2,2]`; `"$r:"`
    needs `"${r}:"`; `{0,>10}` is invalid; cast every value written to a cell (`[string]`/`[double]`).
- **Source:** Al Moosa 2026-10-01..07 | my error | scope: always
