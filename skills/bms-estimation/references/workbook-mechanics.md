# Workbook mechanics

Read before any read or write to an estimation workbook.

## Writing - Excel COM only

These workbooks are macro-heavy with many hidden lookup sheets. **A plain XLSX library strips the VBA
project and breaks every selection formula.** Always use Excel COM.

```powershell
# attach to a running instance if the file is already open
$x = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
# or start one
$x = New-Object -ComObject Excel.Application; $x.Visible = $true
$x.DisplayAlerts = $false
$wb = $x.Workbooks.Open($path)

$prev = $x.EnableEvents; $x.EnableEvents = $false   # stop Worksheet_Change firing mid-write
# ... edits ...
$x.CalculateFullRebuild()
$x.EnableEvents = $prev
$wb.Save()
```

Rules:
- **Back up first**, every time: `Copy-Item $f "$f - BACKUP <stage>.xlsm"`.
- **Save before running any macro.** Some crash Excel (`SelectVFDs` has); a saved file loses nothing.
- Set `$x.AutomationSecurity = 3` when opening only to read, so macros do not run.
- If COM returns *"Call was rejected by callee"* or *"RPC server unavailable"*, Excel is busy, showing
  a dialog, or has crashed. Check `Get-Process EXCEL`, do not retry in a loop.
- `.Select()` can fail on protected sheets - macros usually work anyway without activating.
- Merged cells: writing to the second cell of a merge does nothing and `ClearContents` throws. Values
  in the Breakdown header live in **column D**, not C.
- PowerShell format specifiers use `{0,-20}` / `{0,10}`. **`{0,>10}` is invalid** and throws.
- Cast guard: `[double]$cell.Value2` throws on header text - use `[double]::TryParse`.

## Reading without Excel

`references/dump_sheets.ps1` extracts every sheet to plain text:

```
powershell -File references/dump_sheets.ps1 -Xlsm "<file.xlsm>" -OutDir out -MaxRows 200
```

Raise `-MaxRows` for IOSummary and DDCFullSummary (they run to thousands of rows). The script can fail
against a file Excel currently has open - read through COM instead in that case.

## Formulas you must not overwrite

| Cell / column | Formula | Meaning |
|---|---|---|
| IOSummary `H:L` | `IF(C="","",IF(C*$A<blk>>0, C*$A<blk>,""))` | All-systems = per-system x QTY |
| IOSummary `O` | `IFNA(VLOOKUP($N<row>, FieldDevices!A:B, 2, FALSE), …)` | Model from the column N description |
| Workstation `B8:B12` | `SUMIF(IOSummary!P:P, <protocol>, IOSummary!L:L)` | Software points by protocol |
| Workstation `B13` | total SP minus the tagged ones | "Other" - must end at 0 |
| Breakdown `D6` | `='Cover Page'!H5` | Project name, single source |
| Breakdown `D9` | `=BOQ!J110` | Total value |
| Breakdown `D23` | `=Product_Finder_…!K14` | **Siemens cost - does NOT follow the BOQ** |
| Breakdown `I<n>` | `=F<n>*G<n>*H<n>` | qty x duration x rate |
| Cover Page `H11` | `=Breakdown!D9` | Headline price |

## Macro inventory

| Module | Key procedures |
|---|---|
| `DDCSummaryModule` | `GenerateDDCSummary`, `SelectControllersForSheet`, `GenerateFullDDCSummary`, `ClearDDCSummary` |
| `DDCListModule` | `GenerateEquipmentForDDCList`, `ClearDDCList` |
| `IOSummaryModule` | `ContorlShiftI`, `InsertNewLine`, `FillAirHandilingUnitPoints`, `AddEquipmentFromEquipmentList` |
| `DamperActuatorSelectionModule` | `LoadDamperActuators`, `DamperActuatorDefaultSelection` |
| `VASelectionModule` | `LoadValves`, `VADefaultSelection` |
| `VFDSelectionModule` | `LoadVFDs`, `SelectVFDs` |
| `RoomUnitSelectionModule` | `LoadRoomUnits`, `InsertNewFCU`, `GetFCUModel` |
| `WorstationModule` | `GenerateWorstation`, `SelectBALicenses` |
| `BOQModule` / `BOQPerPanelModule` | `StartBOQGeneration`, `StartBOQPerPanelGeneration`, `ClearBOQ` |
| `CoverPageModule` | `StartNewProject`, `Proposal_Print` |

Run them as `$x.Run("'" + $wb.Name + "'!Module.Procedure")`. `SelectControllers` needs an argument -
use `SelectControllersForSheet`.

Workbook shortcuts: **Ctrl+Shift+I** insert new line, **Ctrl+Shift+G** generate IO points from
templates. Prefer these over copy-paste - they keep formulas intact.
