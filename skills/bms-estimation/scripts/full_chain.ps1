# Full downstream regeneration for the workbook in $env:BMS_WB: DDC List (+format) -> DDCSummary -> controllers
# -> Full summary -> Workstation -> dampers -> valves/VFD/RoomUnits -> BOQ -> Product Finder
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = & "$sp\ddc_build.ps1" -Write; $out | Where-Object { $_ -is [string] }
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($env:BMS_WB); $x = $wb.Application
$d = $wb.Worksheets.Item("DDC List"); $lr = $d.Cells.Item($d.Rows.Count, 1).End(-4162).Row; $lc = $d.Cells.Item(2, $d.Columns.Count).End(-4159).Column; $LCOL = ($d.Cells.Item(1, $lc).Address($false, $false) -replace '\d', '')
$d.Range("A2:$LCOL$lr").Borders.LineStyle = 1; $d.Range("A2:$LCOL$lr").Borders.Weight = 2
$d.Range("A$($lr+1):$LCOL`200").Borders.LineStyle = -4142; $d.Range("A$($lr+1):A200").Interior.ColorIndex = -4142; $wb.Save()
$x.Calculation = -4135
try { for ($i = $d.Cells.FormatConditions.Count; $i -ge 1; $i--) { $fc = $d.Cells.FormatConditions.Item($i); if ($fc.Type -eq 2) { $fc.Delete() } }
  $cf = $d.Range("B2:$LCOL$lr").FormatConditions.Add(2, [Type]::Missing, '=OR(CELL("col")=COLUMN(),CELL("row")=ROW() )'); $cf.Interior.Color = 65535; $cf.StopIfTrue = $false; $cf.SetFirstPriority() } finally { $x.Calculation = -4105 }
$r1 = $d.Range($d.Cells.Item(1, 2), $d.Cells.Item(1, $lc)); $r1.Orientation = -4171; $r2 = $d.Range($d.Cells.Item(2, 2), $d.Cells.Item(2, $lc)); $r2.Interior.Color = 14806254; $r2.Font.Bold = $true; $r2.Orientation = -4171; $r2.HorizontalAlignment = -4131; $r3 = $d.Range($d.Cells.Item(3, 2), $d.Cells.Item(3, $lc)); $r3.Font.Bold = $true; $r3.Font.Color = 10498160; $d.Range($d.Cells.Item(1, 2), $d.Cells.Item(1, $lc)).EntireColumn.ColumnWidth = 2.67 # header formats
$rg4 = $d.Range("B4:$LCOL`4"); for ($i = $d.Cells.FormatConditions.Count; $i -ge 1; $i--) { $fc = $d.Cells.FormatConditions.Item($i); if ($fc.Type -eq 1) { $fc.Delete() } }; foreach ($s4 in @(@(3, 13561798, 24832), @(6, 10284031, 22428), @(5, 13551615, 393372))) { $c4 = $rg4.FormatConditions.Add(1, [int]$s4[0], '=B$3'); $c4.Interior.Color = [int]$s4[1]; $c4.Font.Color = [int]$s4[2]; $c4.StopIfTrue = $false } # row4 CF
$wb.Save(); "DDC List rows to $lr, columns to $LCOL formatted"
$s = "$sp\run_macro.ps1"
& $s -Macros 'DDCSummaryModule.GenerateDDCSummary' -Sheet 'DDCSummary'
& $s -Macros 'DDCSummaryModule.SelectControllersForSheet' -Sheet 'DDCSummary'
& $s -Macros 'DDCSummaryModule.GenerateFullDDCSummary', 'WorstationModule.GenerateWorstation' -Sheet 'DDCFullSummary'
& $s -Macros 'DamperActuatorSelectionModule.LoadDamperActuators', 'DamperActuatorSelectionModule.DamperActuatorDefaultSelection' -Sheet 'DamperActuators'
& $s -Macros 'VASelectionModule.LoadValves', 'VFDSelectionModule.LoadVFDs' -Sheet 'ValvesAndActuators'
& $s -Macros 'BOQModule.StartBOQGeneration' -Sheet 'BOQ'
& "$sp\pf_refill.ps1"


