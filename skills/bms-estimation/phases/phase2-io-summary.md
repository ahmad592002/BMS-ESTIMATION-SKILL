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
(modulating room thermostat).


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

- **Sort the points by type within each block**: DI, then AI, AO, DO, SP. Never leave them in the
  order they came out of the source document.
- **Put a small shaded title row above each group** - "Digital Inputs", "Analog Inputs",
  "Analog Outputs", "Digital Outputs", "Software Points" - with `Interior.ColorIndex = 15`, the same
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
