# READ-ONLY: per-equipment points, devices and notes from IOSummary + EquipmentList
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$path = $env:BMS_WB
$wbk = $null; for ($i = 0; $i -lt 30 -and -not $wbk; $i++) { try { $wbk = [Runtime.InteropServices.Marshal]::BindToMoniker($path); $null = $wbk.Name } catch { $wbk = $null; Start-Sleep 2 } }
$ios = $wbk.Worksheets.Item("IOSummary"); $lastRow = $ios.Cells.Item($ios.Rows.Count, 2).End(-4162).Row
$IOV = $ios.Range("A1:Q$lastRow").Value2
$els = $wbk.Worksheets.Item("EquipmentList"); $ELV = $els.Range("A1:F197").Value2
function N($v) { if ($v -is [double]) { $v } else { 0.0 } }
$out = @(); $sn = 0
for ($r = 1; $r -le $lastRow; $r++) {
  if ([string]$IOV[$r, 1] -ne 'QTY') { continue }
  $sn++; $q = $r + 2; $t = $q; while ([string]$IOV[$t, 2] -ne 'TOTAL') { $t++ }
  $per = @(); foreach ($k in 3..7) { $per += N $IOV[$t, $k] }
  $pts = @(); $devs = @()
  for ($k = $q + 1; $k -lt $t; $k++) { $nm = ([string]$IOV[$k, 2]).Trim(); if (-not $nm) { continue }
    $io = @(); foreach ($c in 3..7) { $io += N $IOV[$k, $c] }
    if (($io | Measure-Object -Sum).Sum -gt 0) { $pts += ("{0}[{1}]" -f $nm, (($io | ForEach-Object { [int]$_ }) -join ',')) }
    if ([string]$IOV[$k, 14]) { $devs += ("{0}->{1} x{2}" -f $IOV[$k, 14], $IOV[$k, 15], $IOV[$k, 13]) } }
  $qty = N $IOV[$q, 1]
  $out += [pscustomobject]@{ SN = $sn; Name = [string]$IOV[$q, 2]; Qty = $qty; Fn = [string]$ELV[($sn + 1), 4]; Type = [string]$ELV[($sn + 1), 5]
    DI = $per[0]; AI = $per[1]; AO = $per[2]; DO = $per[3]; SP = $per[4]; PerUnit = ($per | Measure-Object -Sum).Sum; Total = ($per | Measure-Object -Sum).Sum * $qty
    ElNote = [string]$ELV[($sn + 1), 6]; EqNote = [string]$IOV[$q, 17]; Points = ($pts -join ' ; '); Devices = ($devs -join ' ; ') }
}
$out | Export-Clixml "$env:TEMP\review_blocks.xml"
"blocks=$($out.Count) ; total points=" + ($out | Measure-Object Total -Sum).Sum

