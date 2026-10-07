# Phase 4 - DDCSummary

Produces: per-panel IO totals and controller selection.

## Run the workbook's macros - do not hand-build

```
DDCSummaryModule.GenerateDDCSummary      then
DDCSummaryModule.SelectControllersForSheet
```

`SelectControllers` takes an argument and will fail from COM - use `SelectControllersForSheet`,
with **DDCSummary active** (it reads `ActiveSheet`): `run_macro.ps1 -Macros ... -Sheet DDCSummary`.
Then check Controllers!L > 0 - otherwise the BOQ silently has no controllers.
Save before each macro. See `references/workbook-mechanics.md`.

## Block layout

Same IO layout as IOSummary, plus a controllers region:

```
Q=Controller  R=Quantity  T=UI U=UO V=UIO W=DO X=AO Y=DI Z=RO AA=SP
```

Each block ends `TOTAL` / `SPARE` / `TOTAL WITH SPARE`, and carries `Protocol`, `Spare`, `Type` and
`Enclosure` in column Q/R.

## Controller families (`Controllers` sheet holds the capacities)

| Part | Role | Capacity |
|---|---|---|
| `PXC7.E400S/M/L` | Automation station | SP 100 / 250 / 600, max HW 100 / 200 / 400 |
| `PXC5.E24` | Compact station | UIO 16, DI 2, RO 6, SP 120, max HW 80 |
| `PXC4.E16-2` | Compact station | UIO 12, RO 4, SP 80, max HW 40 |
| `TXM1.16D` / `TXM1.8D` | Digital input | 16 / 8 DI |
| `TXM1.8U` | Universal | 8 UI |
| `TXM1.6R` | Relay output | 6 RO |
| `TXM1.8T` / `TXM1.8X` | Digital output | 8 DO |
| `TXA1.K12` / `K24` | Address key | one per 12 / 24 modules |
| `TXS1.12F10` / `TXS1.EF10` | Power supply | one per panel, more for heavy RO/DO loads |

The 250-point ceiling matches `PXC7.E400M`'s SP capacity - that is where the rule comes from.

## Exit criteria

- Every panel's module capacity >= IO demand **including spare**; station `SP` capacity >= software
  points. Show capacity against demand per panel.
- No panel over 250 points.
- If `Options` says `Modular` but the macro selected compacts (`PXC4`/`PXC5`), say so - the option may
  not be reaching the selection logic.

## Present

Per-panel bill of controllers with the capacity check alongside, and the total controller/module count
for the BOQ.
