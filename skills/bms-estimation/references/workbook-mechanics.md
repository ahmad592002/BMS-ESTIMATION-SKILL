# Workbook mechanics

Read before any read or write to an estimation workbook.

## Writing - Excel COM only

These workbooks are macro-heavy with many hidden lookup sheets. **A plain XLSX library strips the VBA
project and breaks every selection formula.** Always use Excel COM.

```powershell
$env:BMS_WB = "<full path>\<Project>-BMS ESTIMATION.xlsm"   # every script reads this
# Excel closed? open it the way a user would - NOT BindToMoniker (hidden instance dies with the script)
Start-Process $env:BMS_WB   # then poll GetActiveObject until the workbook is listed
$x = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $x.Workbooks | Where-Object { $_.FullName -eq $env:BMS_WB }
$x.DisplayAlerts = $false

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
- If COM returns *"Call was rejected by callee"*, *"Unable to set Calculation"* or the window shows the
  file but `Workbooks.Count` is 0: a cell is in edit mode or a dialog is open - ask the user to press
  Enter/Esc. *"RPC server unavailable"*: Excel closed or crashed. Do not retry in a loop.
- **Macros act on the ActiveWorkbook** (unqualified `Sheets()`): activate the workbook and the target
  sheet before every macro and verify `ActiveWorkbook.FullName`; keep one estimation file open.
  `scripts/run_macro.ps1` does all of this.
- Read-only check while the live window is busy: open the saved file in a separate
  `New-Object Excel.Application` (`AutomationSecurity = 3`, ReadOnly), never touching the user's window.
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
| Breakdown `D9` | `=BOQ!J<total row>` | Total value (row moves with the BOQ line count) |
| Breakdown `D23` | `=Product_Finder_…!K14` | **Siemens cost - does NOT follow the BOQ** (run pf_refill) |
| Breakdown `I23` | `=(D23*(1-E23))*(1+F23)` | Siemens discount applied ONCE here - Product Finder G11 stays 0 |
| BOQ `K<r>` | `=VLOOKUP($D<r>,Pricelist!$A$2:$E$1000,3,FALSE)` | Net cost per unit - never paste another formula here |
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
