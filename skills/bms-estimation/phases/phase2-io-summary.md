# Phase 2 - IOSummary

Produces: one block per equipment type, points counted, field devices selected, protocols tagged.

## Block layout

```
r(n)    A=QTY   B=<Type>   C="IO (1 System)"   H="IO (All Systems)"   M="Field Devices"
r(n+1)  B=Location:  C=DI D=AI E=AO F=DO G=SP   H=DI I=AI J=AO K=DO L=SP   M=QTY N=Field Device O=Model P=Comments
rows                 <point name> with a 1 in its column
r(last) B=TOTAL  + column sums
```

- `C:G` = one system. `H:L` = `C:G * $A<block>` - **formulas, never typed.**
- Generate points with **Ctrl+Shift+G** from `IOTemplate` / `Template` rather than typing them.
- `Template` drives parametric systems - Chilled Water System asks number of chillers, pipe size,
  primary/secondary pump counts, DPS/DPT, expansion tank, bypass + valve, valve type/size, supply and
  return temperatures. Fill from the schedule and let the points generate.

## Typical point patterns

| Equipment | Points |
|---|---|
| Pump | Run status, trip status (+ start/stop DO, auto/manual if `Monitor&Control`) |
| Tank | High level, low level, level sensor (AI) |
| AHU / FAHU | Supply & return temp, filter DP, fan status/trip/command, damper AO, valve AO, freeze stat |
| Electrical panel | Breaker status, trip, V/I via meter |

## Field devices (columns M:P)

`Model` in column O is **a formula**: `IFNA(VLOOKUP($N<row>, FieldDevices!A:B, 2, FALSE), ...)`.
**To change a device, write the exact `FieldDevices!A` description into column N** and let O resolve;
the `FieldDevices` quantity counts update themselves. Never overwrite O - it destroys the lookup.

Only points needing a physical sensor get a device; statuses and commands wired from an MCC do not.

**Match the device to the function, not just the medium:**

| Function | Device class | Point type |
|---|---|---|
| Proof / status (fan airflow proof, filter dirty) | D.P.S. **switch** - `QBM81-3/-5/-10` | DI |
| Measured value (duct static under VFD control) | **transmitter** - `QBM3020-x` | AI |
| Water DP | `QBE3000-D16` transmitter | AI |

Changing a point between status and measurement means **moving its `1` between column C (DI) and
column D (AI)**.

**Standardise on parts already used in other projects.** Before pricing, compare every model against
the other workbooks in `~/OneDrive/Desktop/GTS/BMS_PROJECTS/`. A model appearing in only one project is
a candidate for substitution by the house-standard part for the same duty. Never swap where the SOW or
physics dictates the part - raise it as a question instead.

House standard seen across YALJ / P.Mansour / RAPEH: `QFM2120` (duct H&T), `QAE2120.010` (water temp),
`AX-LS-FL-1HM`/`-1LM` (float switches), `QBM81-5` (fan DPS), `QBE2003-P16` (water pressure),
`QBM81-10` (filter DPS), `AX-UL-SEP380-2` (ultrasonic level), `QBE3000-D16` (water DP), `RDF440BN`
(modulating room thermostat). Full current list: LEARNED C "Field devices".

**Valves** get two priced rows each (command -> valve body, feedback -> actuator) - PICV where the
tender says PICV. Models and the Valves direct-list rule: LEARNED C "Valves". **FCU / VAV** are always
one Integration row of 7 SP - LEARNED C.


## Build IOSummary with the sheet's own macros - never write raw cells

Hand-writing values destroys the block design. Use the workbook's routines:

| Step | Macro |
|---|---|
| Clear the sheet | `IOSummaryModule.ClearIOSummary` |
| Create one formatted block per equipment | `AddEquipmentFromEquipmentList(<EquipmentList Qty cell>, <row>)` - it fills `IOTemplate` then `CopyIOTemplate` inserts the formatted 30-row block |
| Write a point | `AddIO(<column B cell>, <text>, BI, BO, AI, AO, SPBACNET, SPKNX, SPMODBUS)` |
| Point + device + comment in one call | `AddIOAndDeviceAnMoveNext(...)` |
| Add a point row beyond the template's capacity | `InsertIOLine(<cell>)` - copies the formatted row **and** rebuilds its H:L formulas |

- `AddEquipmentFromEquipmentList` takes the **Qty cell (column B)** of the EquipmentList row, not the
  name: it reads the name from `offset(0,1)` and the Type from `offset(0,3)`.
- A template block gives **22 usable point rows** (equipment row at header+2, points from header+3,
  TOTAL at header+25). Anything larger needs `InsertIOLine`.
- **Insert and fill blocks bottom-up**, so inserts never shift blocks already written.

## Lay the block out cleanly

- **Group the points by component** (Supply Fan, Exhaust Fan, Dampers, Filters, Coils & Valves, Sensors, Pump, ...) and **sort DI, AI, AO, DO, SP inside each component**. Never leave them in the
  order they came out of the source document.
- **Put a small shaded title row above each component** - e.g. "Supply Fan", "Dampers", "Filters",
  "Sensors", "Software Integration" - with `Interior.ColorIndex = 15`, the same
  style the workbook uses for its own `IsTitle` rows. A title row carries no IO value and no H:L
  formula.
- **Leave exactly one blank row** between the last point and the TOTAL row, in every block. Delete the
  rest of the template's unused rows - a 23-block sheet loses ~150 empty rows this way.
- When rewriting or trimming rows, **never ClearContents across H:L** - those are the
  `= C:G * $A<qty row>` all-systems formulas. Wipe them and every block silently totals zero. Clear
  only B:G and M:P, or restore H:L afterwards.
- One equipment row per **assembly**: a duty/standby set with its own package tags (`BSP-01`, `BSP-02`,
  `BSP-SET-01`) is **qty 1** carrying all the set's points, not qty 2 - otherwise the shared package
  points are multiplied.

## Software points - tag the protocol in column P

The Workstation sheet classifies software points with
`SUMIF(IOSummary!P:P, <protocol>, IOSummary!L:L)`. The key must match its labels **exactly**:
`MODBUS`, `BACNET/IP`, `BACNET/MSTP`, `MBUS`, `KNX`.

**An untagged SP row falls into "Other"**, so a whole project can show 0 on every protocol and the
licence mix becomes unverifiable. Tag every SP row, and confirm Workstation "Other" reads 0.

The protocol is usually stated in the point description or implied by the selected device (e.g.
`RDF440BN` is BACnet MS/TP; a chiller plant manager is normally BACnet/IP).

## Exit criteria

- Every EquipmentList row has a block; `H:L` are formulas; each block has a TOTAL row.
- Field devices selected via column N, with no one-off parts left unexplained.
- Every SP row tagged in column P; Workstation "Other" = 0.
- Grand totals DI/AI/AO/DO/SP presented.

## Present

Per-type point counts, the grand total, the field-device list, and the software-point split by
protocol.

## Build the layout in the fill pass - never retrofit rows into finished blocks

Titles, sorting and blank-row trimming must be decided **before** the points are written, and applied
in the same pass that fills the block. Do not try to insert title rows, re-sort, or delete rows in a
block that is already populated:

- repeated single-row `InsertIOLine` calls across many blocks destabilise Excel - in one run it
  crashed into AutoRecover and reopened the file as `.xlsb`;
- a batched `Rows(a:b).Insert` followed by clear-and-rewrite silently dropped point names and left the
  totals wrong (BA 554 instead of 556);
- `ClearContents` over H:L wipes the all-systems formulas.

If a populated block needs a different layout, **clear the sheet and rebuild it**: `ClearIOSummary`,
recreate the blocks, then write titles and sorted points together. That path is reliable and takes
seconds.

Save after each good state, and if a write fails, close the workbook **without saving** and reopen
from disk rather than trying to repair a half-written sheet.

## Proven workflow (Al Moosa University, 2026-10 - estimator: "very good way")

**A. Read the points from the Site control schematics (B-93), all of them, before generating.**
- Two formats exist:
  1. *BMS SCHEDULE tables* (DI / DO / AI / AO / ALARM / HARDWIRED INTERLOCK / COMMUNICATION): parse
     with `scripts/parse_sched.ps1` (word coordinates from MiKTeX `pdftotext -bbox`; also accepts
     "PMS SCHEDULE"). COMMUNICATION = software point.
  2. *DDC point strips* (vertical labels over DI/DO/AI/AO rows, marks are drawn dots, filled or
     hollow, "x2"/"X2" multipliers): `scripts/parse_strip.ps1` renders the strip at 200 dpi and finds
     dots/rings with a compiled C# scanner, ignores anything inside a text box, de-duplicates, and
     types any label without a detected mark from its wording - flagged "inferred".
  Use the table where a sheet has both. `scripts/combine.ps1` merges them into `points_all.csv`.
- **Read the main plant sheets by eye** (chillers, cooling towers, pump groups, refrigerant
  purge) - several equipment per sheet; the auto-parse is unreliable there.
- Present per-unit points per equipment type + a review workbook (Per Equipment / Per Sheet /
  All Points) before generating.

**B. Templates and mapping** (`scripts/io_templates.txt`, `io_map.txt`, `io_lib.ps1`): one template
per equipment kind - `@SHEET:nnn[:group]` pulls drawn points; hand-entered lines for read-by-eye or
typical points. Keep the estimator's own block values (e.g. chiller SP 15). Dry-run first
(`build_iosummary.ps1 -DryRun`) - it prints per-row IO and the expected grand total.

**C. Generate with the EquipmentList button, then complete** (`scripts/run_all.ps1`):
fully clear IOSummary (ClearIOSummary stops at row 2000) -> run `GenerateIOPointsFromEquipmentList`
(stops at row 120) -> append the remaining rows with the macro's own steps -> fill every block
bottom-up (`fill_iosummary.ps1`) -> rewrite H:L to each block's qty row (`fix_hl.ps1`) -> verify
expected == got and Workstation "Other" = 0.

**D. Layout:** points grouped by COMPONENT (Supply Fan, Exhaust Fan, Dampers, Filters, Coils &
Valves, Heat Recovery, Sensors, Pump, Chiller, Breakers & Protection, Power Metering, Software
Integration ...), DI/AI/AO/DO/SP inside each, grey bold title rows, one blank row before TOTAL.
Name every point by its drawn component (PANEL / BAG / PRE / HEPA / CARBON filter - never a bare
"FILTER STATUS").

**E. Field devices** (`fill_devices.ps1`, rules in `io_lib.ps1` Get-Device2 + `io_flags.ps1`):
only where the instrument is drawn, model = previous GTS projects' choice, column M = total
(`=<per unit>*$A<qty row>`), dampers/valves/VFDs to their selection sheets.

**F. Mark doubts cell by cell** (`apply_flags2.ps1`, rules in `io_flags2.ps1`): SP -> G, point ->
its IO cell, device -> N, device qty -> M, equipment qty -> A, schematic choice / typical -> name
cell B; reason in column Q "Check note". Assumed SP counts are always yellow.

**G. Sources workbook** (`build_io_source.ps1` first time, `refresh_sources.ps1` after edits and at
the very end): Equipment Summary, IO Points by Equipment (every
point with drawing number, how it was read, basis, device -> model, note), Source Drawings,
To Check (yellow), Field Devices. Totals must equal the IOSummary.
