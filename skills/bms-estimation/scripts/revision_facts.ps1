# READ-ONLY facts per phase for the revision report
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$path = $env:BMS_WB
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($path)
function N($v) { if ($v -is [double]) { $v } else { 0.0 } }
# Phase 0 Options
$o = $wb.Worksheets.Item("Options"); "== P0 Options: " + ((1..13 | ForEach-Object { [string]$o.Cells.Item($_, 1).Text + '=' + [string]$o.Cells.Item($_, 2).Text }) -join ' ; ')
# Phase 1
$el = $wb.Worksheets.Item("EquipmentList"); $last = $el.Cells.Item($el.Rows.Count, 3).End(-4162).Row; $ev = $el.Range("A2:F$last").Value2
$byB = @{}; $byT = @{}; $tot = 0; $notes = 0
for ($i = 1; $i -le $last - 1; $i++) { $q = N $ev[$i, 2]; $tot += $q; $b = ([string]$ev[$i, 3] -split ' - ')[0]; $byB[$b] = (N $byB[$b]) + $q; $t = [string]$ev[$i, 5]; $byT[$t] = (N $byT[$t]) + $q; if ([string]$ev[$i, 6]) { $notes++ } }
"== P1 EquipmentList: rows=$($last-1) units=$tot yellowNotes=$notes"
"   by building: " + (($byB.Keys | Sort-Object | ForEach-Object { "$_=$($byB[$_])" }) -join ', ')
"   by type: " + (($byT.Keys | Sort-Object | ForEach-Object { "$_=$($byT[$_])" }) -join ', ')
# Phase 2
$io = $wb.Worksheets.Item("IOSummary"); $il = $io.Cells.Item($io.Rows.Count, 2).End(-4162).Row; $iv = $io.Range("A1:Q$il").Value2
$t5 = @(0.0, 0.0, 0.0, 0.0, 0.0); $nb = 0; $yel = 0
for ($r = 1; $r -le $il; $r++) { if ([string]$iv[$r, 1] -eq 'QTY') { $nb++ }; if ([string]$iv[$r, 2] -eq 'TOTAL') { for ($k = 0; $k -lt 5; $k++) { $t5[$k] += N $iv[$r, (8 + $k)] } }; if ([string]$iv[$r, 17] -and [string]$iv[$r, 1] -ne 'QTY' -and [string]$iv[$r, 17] -ne 'Check note (yellow = verify)') { $yel++ } }
"== P2 IOSummary: blocks=$nb DI=$($t5[0]) AI=$($t5[1]) AO=$($t5[2]) DO=$($t5[3]) SP=$($t5[4]) physical=$($t5[0]+$t5[1]+$t5[2]+$t5[3]) notes=$yel"
$w = $wb.Worksheets.Item("Workstation"); "   SP by protocol: MODBUS=" + $w.Range("B8").Text + " BACNET/IP=" + $w.Range("B9").Text + " BACNET/MSTP=" + $w.Range("B10").Text + " MBUS=" + $w.Range("B11").Text + " Other=" + $w.Range("B13").Text + " TotalBA=" + $w.Range("B15").Text
$fd = $wb.Worksheets.Item("FieldDevices"); $fv = $fd.Range("A1:C130").Value2; $dev = @(); for ($r = 2; $r -le 130; $r++) { if ((N $fv[$r, 3]) -gt 0) { $dev += "$($fv[$r,2]) x$($fv[$r,3])" } }
"   field devices ($($dev.Count) types): " + ($dev -join ', ')
$va = $wb.Worksheets.Item("Valves"); $vl = @(); for ($r = 62; $r -le 80; $r++) { if ((N $va.Cells.Item($r, 3).Value2) -gt 0) { $vl += "$($va.Cells.Item($r,1).Text) x$($va.Cells.Item($r,3).Text)" } }; "   valves: " + ($vl -join ', ')
$da = $wb.Worksheets.Item("DA"); $dv = $da.UsedRange.Value2; $dl = @(); for ($r = 2; $r -le $dv.GetLength(0); $r++) { for ($c = 1; $c -le $dv.GetLength(1) - 2; $c++) { if (([string]$dv[1, $c]) -match 'Part' -and (N $dv[$r, ($c + 2)]) -gt 0) { $dl += "$($dv[$r,$c]) x$($dv[$r,($c+2)])" } } }; "   damper actuators: " + ($dl -join ', ')
# Phase 3
$d = $wb.Worksheets.Item("DDC List"); $lr = $d.Cells.Item($d.Rows.Count, 1).End(-4162).Row; $lc = $d.Cells.Item(2, $d.Columns.Count).End(-4159).Column
$names = @(5..$lr | ForEach-Object { [string]$d.Cells.Item($_, 1).Text } | Where-Object { $_ })
$bad = 0; for ($c = 2; $c -le $lc; $c++) { if ((N $d.Cells.Item(3, $c).Value2) -ne (N $d.Cells.Item(4, $c).Value2)) { $bad++ } }
"== P3 DDC List: rows=$($names.Count) equipmentColumns=$($lc-1) assignedMismatch=$bad ; DDCP=" + @($names | Where-Object { $_ -match '^DDCP' }).Count + " NET=" + @($names | Where-Object { $_ -match '^NET' }).Count + " other=" + @($names | Where-Object { $_ -notmatch '^(DDCP|NET)' }).Count + " (" + (($names | Where-Object { $_ -notmatch '^(DDCP|NET)' }) -join ', ') + ")"
# Phase 4/5
foreach ($sn in 'DDCSummary', 'DDCFullSummary') { $s = $wb.Worksheets.Item($sn); $v = $s.UsedRange.Value2; $np = 0; $tt = @(0.0, 0.0, 0.0, 0.0, 0.0); for ($r = 1; $r -le $v.GetLength(0); $r++) { if ([string]$v[$r, 1] -eq 'QTY') { $np++ }; if (([string]$v[$r, 2]).Trim() -eq 'TOTAL') { for ($k = 0; $k -lt 5; $k++) { $tt[$k] += N $v[$r, (8 + $k)] } } }; "== $sn : panels=$np points=" + ($tt -join '/') }
$ct = $wb.Worksheets.Item("Controllers"); $cl = $ct.Cells.Item($ct.Rows.Count, 1).End(-4162).Row; $cc = @(); for ($r = 2; $r -le $cl; $r++) { if ((N $ct.Cells.Item($r, 12).Value2) -gt 0) { $cc += "$($ct.Cells.Item($r,1).Text) x$($ct.Cells.Item($r,12).Text)" } }; "   controllers/modules: " + ($cc -join ', ')
$en = $wb.Worksheets.Item("Enclosures"); $ee = @(); for ($r = 2; $r -le 20; $r++) { if ((N $en.Cells.Item($r, 3).Value2) -gt 0) { $ee += "$($en.Cells.Item($r,2).Text) x$($en.Cells.Item($r,3).Text)" } }; "   enclosures: " + ($ee -join ', ')
# Phase 6/7
"   workstation: " + ((2..20 | ForEach-Object { $p = [string]$w.Cells.Item($_, 6).Text; if ($p) { "$p x$($w.Cells.Item($_,7).Text)" } }) -join ', ')
$b = $wb.Worksheets.Item("BOQ"); $cat = [ordered]@{}; $nl = 0; for ($r = 6; $r -le 104; $r++) { $c = ([string]$b.Cells.Item($r, 3).Text).Trim(); if (-not $c) { continue }; $nl++; if (-not $cat.Contains($c)) { $cat[$c] = @(0.0, 0.0) }; $cat[$c][0] += N $b.Cells.Item($r, 10).Value2; $cat[$c][1] += N $b.Cells.Item($r, 13).Value2 }
"== P7 BOQ lines=$nl markup=" + $b.Range("M2").Text + " : " + (($cat.Keys | ForEach-Object { "{0} sell {1:N0} / cost {2:N0}" -f $_, $cat[$_][0], $cat[$_][1] }) -join ' | ')
$bd = $wb.Worksheets.Item("Breakdown"); "   Breakdown: " + ((@('D6', 'G6', 'D7', 'D8', 'G9', 'D9', 'E23', 'I23', 'I24', 'I25', 'I73', 'I89', 'I91', 'I93', 'H93') | ForEach-Object { "$_=" + $bd.Range($_).Text }) -join ' ; ')
"   expenses: " + ((46..71 | ForEach-Object { $v = N $bd.Cells.Item($_, 9).Value2; if ($v) { "$($bd.Cells.Item($_,3).Text) $($bd.Cells.Item($_,6).Text)x$($bd.Cells.Item($_,7).Text)x$($bd.Cells.Item($_,8).Text)" } }) -join ', ')
"   resources: " + ((77..87 | ForEach-Object { $v = N $bd.Cells.Item($_, 9).Value2; if ($v) { "$($bd.Cells.Item($_,3).Text) $($bd.Cells.Item($_,6).Text)x$($bd.Cells.Item($_,7).Text)x$($bd.Cells.Item($_,8).Text)" } }) -join ', ')
$cp = $wb.Worksheets.Item("Cover Page"); "   Cover: To=" + $cp.Range("F3").Text + " | Company=" + $cp.Range("F4").Text + " | Ref=" + $cp.Range("C5").Text + " | Project=" + $cp.Range("H5").Text + " | Total=" + $cp.Range("H11").Text + " | Server=" + $cp.Range("H15").Text
"   Product Finder K14=" + $wb.Worksheets.Item("Product_Finder_SI_B_AUT_V27.1").Range("K14").Text


