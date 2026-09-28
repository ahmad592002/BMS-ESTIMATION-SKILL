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
  `Total BA`, which sizes the server/licence: `CMD.06`, `CCA-CMPXL-BA`, `CCA-1000-BA`, workstation PC,
  27" monitor, printers, UPS. Re-check the licence tier after any IO change. If "Other" is non-zero,
  go back to Phase 2 and tag the protocols.
- **DamperActuators** - `Unit Name | Actuator Title | Type | Signal | End Switch | Qty | Duct Size (m²)
  | Actuator Description | Part Number | Accessory`. Type `S.R.`/`N.S.R.`, signal `ON/OFF` or
  `Modulating` (`Options` defaults: N.S.R 2.5, S.R 2.5, F.S. 1.5 m²).
- **ValvesAndActuators** - `… | Pipe Size | Flow | P Drop | Calc KV | Valve Size | Valve KVS |
  Valve Auth | Valve Model | Valve Accessory | Actuator Model | …`. Sized by KVS (`Options` B51=2);
  threaded below the 50 mm threshold, flanged above; spring return and weather shield per `Options`;
  BV series e.g. `VKF46`; Globe preferred unless the spec says Ball/PICV.
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
