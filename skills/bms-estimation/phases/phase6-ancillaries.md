# Phase 6 - Ancillaries

Produces: `Workstation`, `DamperActuators`, `ValvesAndActuators`, `VFDs`/`MCCs`, `FCUs`/`RoomUnits`.

## Run the workbook's macros

| Sheet | Macros |
|---|---|
| Damper actuators | `DamperActuatorSelectionModule.LoadDamperActuators` then `DamperActuatorDefaultSelection` |
| Valves | `VASelectionModule.LoadValves` then `VADefaultSelection` |
| VFDs | `VFDSelectionModule.LoadVFDs` then `SelectVFDs` |
| Workstation | `WorstationModule.GenerateWorstation` (note the spelling) |

**Save before each.** `SelectVFDs` has crashed Excel; a saved file loses nothing. A `Load*` macro that
fails once via COM often succeeds on a retry with the workbook freshly opened.

## Every Load* macro loads DEFAULTS - check them before selecting

This is the biggest trap in the phase:

| Sheet | Default loaded | Fix before selecting |
|---|---|---|
| **VFDs** | **every drive at 0.75 kW / 1 HP** | Write real kW and HP into columns D and E from the equipment names or the SOW, or all drives select one frame size |
| Damper actuators | duct area = `Options` default (2.5 m²) | Enter real duct areas; torque follows area, so large AHUs need bigger actuators |
| Valves | `Options` default sizes (AHU 65, BV 200, FCU 25) | Enter flow and pressure drop so KVS sizing is real, not defaulted |

## Sheet notes

- **Workstation** - physical points (DI/AI/AO/DO) plus software points by protocol roll up to
  `Total BA`, which sizes the server/licence: `CMD.06`, `CCA-CMPXL-BA`, workstation PC, 27" monitor,
  printers, UPS. The macro sets PC / monitor / printers / UPS to 1 - set them from the client BOQ
  head-end lines (higher number), plus consoles / wall displays (Pricelist row at 0 if priced 0).
  Re-check the licence tier after any IO change. If "Other" is non-zero, go back to
  Phase 2 and tag the protocols. **Then verify the licence extensions - see below.**

### BA licence extensions - always verify, the macro over-adds

The compact licence **`CCA-CMPXL-BA` already includes 2000 BA points.** The test is on points **with
spare applied**: `Total BA * (1 + Options spare)`.

| Total BA + spare | Extension lines |
|---|---|
| **<= 2000** | **None.** The compact licence covers it - no `CCA-*-BA` extension in the BOQ at all |
| **> 2000** | Deduct the 2000 included, then cover the remainder: `CCA-1000-BA`, `CCA-500-BA`, `CCA-100-BA` (`CCA-5000-BA` only on the non-compact route) |

Compact is only chosen when BA <= 5000, FIRE <= 500, ELEC <= 500, SCADA <= 1000, METER <= 30.

**`SelectBALicenses` gets this wrong below 2000 - check its output every time:**

1. The deduction is guarded by `If (Compact And TotalBA >= 2000)`, so **under 2000 the 2000 included
   points are never deducted** and the full count falls through into the extension ladder.
2. `BA1000 = RemainingBA / 1000` assigns a Double into an Integer, so VBA **rounds instead of
   truncating** (1.656 -> 2), adding one size too many.

Worked example (SDA SCITECH Khobar): 1380 BA points, 20% spare -> 1656. Correct answer is **no
extension**. The macro produces `CCA-1000-BA` x2.

**After running `GenerateWorstation`, read the parts list and delete any `CCA-*-BA` extension line the
table above does not call for.** State in the review what the licence works out to and why.
- **DamperActuators** - `Unit Name | Actuator Title | Type | Signal | End Switch | Qty | Duct Size (m²)
  | Actuator Description | Part Number | Accessory`. Type `S.R.`/`N.S.R.`, signal `ON/OFF` or
  `Modulating` (`Options` defaults: N.S.R 2.5, S.R 2.5, F.S. 1.5 m²).
- **ValvesAndActuators** - `… | Pipe Size | Flow | P Drop | Calc KV | Valve Size | Valve KVS |
  Valve Auth | Valve Model | Valve Accessory | Actuator Model | …`. Sized by KVS (`Options` B51=2);
  threaded below the 50 mm threshold, flanged above; spring return and weather shield per `Options`;
  BV series e.g. `VKF46`; Globe preferred unless the spec says Ball/PICV. In practice GTS types the
  valve models straight into IOSummary O (two rows per valve, LEARNED C "Valves") and lists them in
  the Valves direct section - PICV coils VPF44.65F25 + SAX61P03, globe VVF42.65-50 + SKB62/F,
  butterfly VFW41.150 + SQL341E100.
- **VFDs / MCCs** - selected by HP; MCC in the BOQ only when `Options` says `Yes`.
- **FCUs / RoomUnits** - family follows the FCU protocol in `Options`: KNX -> `RDG100KN`/`RDG160KN`/
  `RDF600KN`; BIP -> `DXR2.E09/E10`; MSTP -> `RDB160BN`/`DXR2.M09/M10`; Modbus -> `RDF302`/`RDF300.02`;
  Standalone -> `RDG100`/`RDU340`.

## Exit criteria

- Licence tier matches total BA points; Workstation "Other" = 0.
- Every modulating damper and control valve in the point list has a selected actuator.
- No selection left at a loaded default that should have been real data - and any that remain are
  stated as assumptions when presenting.
- Quantities reconciled against the client BOQ (damper actuators, valves, VFDs are usually itemised
  there).
