$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$path = "C:\Users\ahmad\OneDrive\Desktop\GTS\BMS_PROJECTS\07- Al Moosa University\BMS Template 2026 V03.xlsm"
$x = New-Object -ComObject Excel.Application
$x.Visible = $true; $x.DisplayAlerts = $false
$wb = $x.Workbooks.Open($path); for ($w=0; $w -lt 120 -and -not $x.Ready; $w++) { Start-Sleep -Milliseconds 500 }
"opened " + $wb.Name
$io0 = $wb.Worksheets.Item("IOSummary"); $u = $io0.UsedRange; $lastU = $u.Row + $u.Rows.Count - 1; if ($lastU -ge 1) { $null = $io0.Range("1:$lastU").EntireRow.Delete() }; "IOSummary fully cleared (" + $lastU + " rows)"
$wb.Worksheets.Item("EquipmentList").Activate()
$null = $x.Run("'" + $wb.Name + "'!GenerateIOPointsFromEquipmentList")
"button macro done"
function Wait-Excel { for ($w=0; $w -lt 120; $w++) { try { if ($x.Ready) { $null = $x.EnableEvents; return } } catch {}; Start-Sleep -Milliseconds 500 }; throw "Excel stayed busy" }
Wait-Excel; "excel ready"
. "$sp\fill_iosummary.ps1"
. "$sp\fix_hl.ps1"