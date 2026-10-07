# Full re-check of the estimation workbook ($env:BMS_WB) + side files. Read-only. Some checks were written for Al Moosa - adapt per project.
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$proj = Split-Path -Parent $env:BMS_WB
$pname = if ($env:BMS_PROJECT) { $env:BMS_PROJECT } else { (Split-Path -Leaf $proj) -replace '^\d+-\s*', '' }   # side files are named "<project> - ..."
$res = New-Object System.Collections.Generic.List[object]
function Chk([string]$ph, [string]$name, [string]$st, [string]$det) { $res.Add([pscustomobject]@{ Phase = $ph; Check = $name; Status = $st; Detail = $det }) }
function N($v) { if ($v -is [double]) { $v } else { 0.0 } }
$WBP = $env:BMS_WB; $wb = [Runtime.InteropServices.Marshal]::BindToMoniker($WBP); $xl = $wb.Application
$YEL = 65535

# ---------------- Phase 1: EquipmentList
$el = $wb.Worksheets.Item("EquipmentList")
$elLast = $el.Cells.Item($el.Rows.Count, 1).End(-4162).Row
$ELV = $el.Range("A1:F$elLast").Value2
$elN = $elLast - 1; $elQty = @{}; $elName = @{}; $bad = @(); $sumQ = 0
for ($i = 2; $i -le $elLast; $i++) { $s = [int](N $ELV[$i, 1]); if ($s -ne $i - 1) { $bad += "row $i SN=$s" }; $elQty[$i - 1] = N $ELV[$i, 2]; $elName[$i - 1] = ([string]$ELV[$i, 3]).Trim(); $sumQ += $elQty[$i - 1]
  if ($elQty[$i - 1] -le 0) { $bad += "SN $($i-1) qty $($elQty[$i-1])" } }
Chk 1 'EquipmentList rows / SN sequence / qty > 0' $(if ($bad) { 'FAIL' } else { 'PASS' }) "$elN rows, total qty $sumQ. $($bad -join '; ')"
$dup = $elName.Values | Group-Object | Where-Object Count -gt 1
Chk 1 'EquipmentList names unique' $(if ($dup) { 'FAIL' } else { 'PASS' }) (($dup | ForEach-Object Name) -join '; ')
$tl = @(); try { $tlv = $wb.Worksheets.Item("TemplateLists").UsedRange.Value2; foreach ($v in $tlv) { if ($v) { $tl += [string]$v } } } catch { }
$badT = @(); for ($i = 2; $i -le $elLast; $i++) { $ty = [string]$ELV[$i, 5]; if ($tl.Count -and $tl -notcontains $ty) { $badT += "SN $($i-1) '$ty'" } }
$tl = @($wb.Application.Evaluate('EquipmentTypes').Value2 | Where-Object { $_ }); $badT = @(); for ($i = 2; $i -le $elLast; $i++) { if ($tl -notcontains [string]$ELV[$i, 5]) { $badT += "SN $($i-1)" } }
Chk 1 'EquipmentList Type exists in EquipmentTypes' $(if ($badT) { 'WARN' } else { 'PASS' }) ($badT -join '; ')
$yNoNote = @(); for ($i = 2; $i -le $elLast; $i++) { $yl = ($el.Cells.Item($i, 2).Interior.Color -eq $YEL) -or ($el.Cells.Item($i, 3).Interior.Color -eq $YEL); if ($yl -and -not [string]$ELV[$i, 6]) { $yNoNote += "SN $($i-1)" } }
Chk 1 'EquipmentList yellow cells have a reason (col F)' $(if ($yNoNote) { 'FAIL' } else { 'PASS' }) ($yNoNote -join '; ')
# takeoff register vs EquipmentList (via el_map)
$asgRes = & "$sp\ddc_assign.ps1"; $unm = @($asgRes | Where-Object { $_ -is [string] }); $asg = $asgRes | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
$mis = @(); foreach ($s in 1..$elN) { $t = 0; if ($asg.ContainsKey($s)) { foreach ($k in $asg[$s].Keys) { $t += $asg[$s][$k] } }; if ($t -ne $elQty[$s]) { $mis += "SN $s EL=$($elQty[$s]) takeoff=$t" } }
Chk 1 'EquipmentList qty == takeoff register (2066 tags)' $(if ($mis -or $unm) { 'FAIL' } else { 'PASS' }) "$($unm.Count) unmatched tags. $($mis -join '; ')"

# ---------------- Phase 2: IOSummary
$io = $wb.Worksheets.Item("IOSummary"); $ioLast = $io.Cells.Item($io.Rows.Count, 2).End(-4162).Row
$IOF = $io.Range("A1:Q$ioLast").Formula; $IOV = $io.Range("A1:Q$ioLast").Value2
$blocks = @(); for ($r = 1; $r -le $ioLast; $r++) { if ([string]$IOV[$r, 1] -eq 'QTY') { $q = $r + 2; $t = $q; while ([string]$IOV[$t, 2] -ne 'TOTAL') { $t++ }; $blocks += [pscustomobject]@{ H = $r; Q = $q; T = $t; Name = ([string]$IOV[$q, 2]).Trim(); Qty = N $IOV[$q, 1] } } }
Chk 2 'IOSummary block count == EquipmentList' $(if ($blocks.Count -eq $elN) { 'PASS' } else { 'FAIL' }) "$($blocks.Count) blocks vs $elN"
$mm = @(); for ($i = 0; $i -lt [math]::Min($blocks.Count, $elN); $i++) { $b = $blocks[$i]; if ($b.Name -ne $elName[$i + 1] -or $b.Qty -ne $elQty[$i + 1]) { $mm += "#$($i+1) IO '$($b.Name)' x$($b.Qty) vs EL '$($elName[$i+1])' x$($elQty[$i+1])" } }
Chk 2 'IOSummary block name + qty == EquipmentList (same order)' $(if ($mm) { 'FAIL' } else { 'PASS' }) (($mm | Select-Object -First 8) -join '; ')
$hlBad = @(); $sumBad = @(); $tot = @(0.0, 0.0, 0.0, 0.0, 0.0)
foreach ($b in $blocks) {
  for ($r = $b.Q; $r -lt $b.T; $r++) { foreach ($k in 0..4) { $cl = [char](67 + $k); $exp = "=IF($cl$r=`"`",`"`",IF($cl$r*`$A$($b.Q)>0,$cl$r*`$A$($b.Q),`"`"))"; if ([string]$IOF[$r, (8 + $k)] -ne $exp) { $hlBad += "r$r"; break } } }
  foreach ($k in 0..4) { $cl = [char](67 + $k); $exp = "=SUM($cl$($b.Q):$cl$($b.T - 1))"; if ([string]$IOF[$b.T, (3 + $k)] -ne $exp) { $sumBad += "r$($b.T)$cl"; break } }
  foreach ($k in 0..4) { $one = N $IOV[$b.T, (3 + $k)]; $all = N $IOV[$b.T, (8 + $k)]; if ([math]::Abs($one * $b.Qty - $all) -gt 0.01) { $sumBad += "r$($b.T) all<>1xQty" }; $tot[$k] += $all } }
Chk 2 'IOSummary H:L formulas point at own qty row' $(if ($hlBad) { 'FAIL' } else { 'PASS' }) "$($hlBad.Count) rows bad $((($hlBad | Select-Object -Unique -First 10) -join ','))"
Chk 2 'IOSummary TOTAL rows sum whole block, all = 1 x qty' $(if ($sumBad) { 'FAIL' } else { 'PASS' }) (($sumBad | Select-Object -First 10) -join ',')
Chk 2 'IOSummary totals DI/AI/AO/DO/SP' 'INFO' ($tot -join ' / ')
# devices
$fd = $wb.Worksheets.Item("FieldDevices"); $FDV = $fd.UsedRange.Value2; $fdDesc = @{}; $fdModel = @{}
for ($r = 2; $r -le $FDV.GetLength(0); $r++) { $d = [string]$FDV[$r, 1]; if ($d) { $fdDesc[$d] = [string]$FDV[$r, 2]; $fdModel[[string]$FDV[$r, 2]] = $d } }
$vs = $wb.Worksheets.Item("Valves"); $direct = @{}; for ($r = 62; $r -le 100; $r++) { $a = [string]$vs.Cells.Item($r, 1).Value2; if ($a -and ([string]$vs.Cells.Item($r, 3).Formula -match 'IOSummary')) { $direct[$a] = $r } }
$devNoM = @(); $mNoDev = @(); $mWrong = @(); $unpriced = @{}; $notCounted = @{}; $naRows = @(); $yNoQ = 0; $dupNames = @()
foreach ($b in $blocks) { $seen = @{}
  for ($r = $b.Q + 1; $r -lt $b.T; $r++) {
    $nm = ([string]$IOV[$r, 2]).Trim(); $dev = [string]$IOV[$r, 14]; $mf = [string]$IOF[$r, 13]; $model = [string]$IOV[$r, 15]
    $hasIO = $false; foreach ($k in 3..7) { if ((N $IOV[$r, $k]) -ne 0) { $hasIO = $true } }
    if ($hasIO -and $nm) { if ($seen[$nm]) { $dupNames += "r$r $nm" }; $seen[$nm] = 1 }
    if ($dev -and -not $mf) { $devNoM += "r$r [$($b.Name)] $dev" }
    if ($mf -and -not $dev) { $mNoDev += "r$r" }
    if ($mf -and $mf -match '^=' -and $mf -notmatch ('\$A' + $b.Q + '$')) { $mWrong += "r$r $mf" }
    if ($mf -and $mf -notmatch '^=') { $mWrong += "r$r static $mf" }
    if ($dev) {
      if ($model -match 'Selection Sheet') { $unpriced[$dev] = (N $unpriced[$dev]) + (N $IOV[$r, 13]) }
      elseif (-not $model -or $model -match '#N/A') { $naRows += "r$r $dev" }
      elseif (-not $fdModel.ContainsKey($model) -and -not $direct.ContainsKey($model)) { $notCounted["$dev -> $model"] = (N $notCounted["$dev -> $model"]) + (N $IOV[$r, 13]) } }
  } }
Chk 2 'Device rows have a qty formula (M)' $(if ($devNoM) { 'WARN' } else { 'PASS' }) "$($devNoM.Count): $(($devNoM | Select-Object -First 6) -join '; ')"
Chk 2 'Qty (M) points at own block qty' $(if ($mWrong) { 'FAIL' } else { 'PASS' }) (($mWrong | Select-Object -First 8) -join '; ')
Chk 2 'Device model resolves (no #N/A / blank)' $(if ($naRows) { 'FAIL' } else { 'PASS' }) (($naRows | Select-Object -First 8) -join '; ')
Chk 2 'Devices with no price ("Selection Sheet")' $(if ($unpriced.Count) { 'WARN' } else { 'PASS' }) (($unpriced.Keys | ForEach-Object { "$_ x$($unpriced[$_])" }) -join '; ')
Chk 2 'Typed models counted (FieldDevices or Valves direct list)' $(if ($notCounted.Count) { 'FAIL' } else { 'PASS' }) (($notCounted.Keys | ForEach-Object { "$_ x$($notCounted[$_])" }) -join '; ')
Chk 2 'Duplicate point names inside a block' $(if ($dupNames) { 'WARN' } else { 'PASS' }) "$($dupNames.Count): $(($dupNames | Select-Object -First 8) -join '; ')"
# FieldDevices + Valves counts == IOSummary M
$sumM = @{}; foreach ($b in $blocks) { for ($r = $b.Q + 1; $r -lt $b.T; $r++) { $model = [string]$IOV[$r, 15]; if ($model -and ($IOV[$r, 13] -is [double])) { $sumM[$model] = (N $sumM[$model]) + $IOV[$r, 13] } } }
$fdBad = @(); for ($r = 2; $r -le $FDV.GetLength(0); $r++) { $m = [string]$FDV[$r, 2]; if ($m -and $m -notmatch 'Selection Sheet') { $c = N $FDV[$r, 3]; $e = N $sumM[$m]; if ([math]::Abs($c - $e) -gt 0.01) { $fdBad += "$m FD=$c IO=$e" } } }
foreach ($m in $direct.Keys) { $c = N $vs.Cells.Item($direct[$m], 3).Value2; $e = N $sumM[$m]; if ([math]::Abs($c - $e) -gt 0.01) { $fdBad += "Valves $m=$c IO=$e" } }
Chk 2 'FieldDevices / Valves qty == IOSummary device qty' $(if ($fdBad) { 'FAIL' } else { 'PASS' }) (($fdBad | Select-Object -First 8) -join '; ')
# SP protocol + Workstation
$protos = @{}; foreach ($b in $blocks) { for ($r = $b.Q + 1; $r -lt $b.T; $r++) { if ((N $IOV[$r, 7]) -gt 0) { $p = [string]$IOV[$r, 16]; $protos[$p] = (N $protos[$p]) + (N $IOV[$r, 12]) } } }
Chk 2 'SP points by protocol (col P)' $(if ($protos.ContainsKey('')) { 'WARN' } else { 'INFO' }) (($protos.Keys | ForEach-Object { "'$_'=$($protos[$_])" }) -join '; ')
$wsS = $wb.Worksheets.Item("Workstation"); $other = $wsS.Range("B13").Value2
Chk 2 'Workstation "Other" software points == 0' $(if ((N $other) -eq 0) { 'PASS' } else { 'FAIL' }) "B13 = $other ; SP total B6 = $($wsS.Range('B6').Value2)"

# ---------------- Phase 3: DDC List
$dl = $wb.Worksheets.Item("DDC List"); $dlLastC = $dl.Cells.Item(2, $dl.Columns.Count).End(-4159).Column; $dlLastR = $dl.Cells.Item($dl.Rows.Count, 1).End(-4162).Row
$DLV = $dl.Range($dl.Cells.Item(1, 1), $dl.Cells.Item($dlLastR, $dlLastC)).Value2
$nb = @(); for ($c = 2; $c -le $dlLastC; $c++) { $i = $c - 1; if ($i -le $blocks.Count) { if (([string]$DLV[2, $c]).Trim() -ne $blocks[$i - 1].Name -or (N $DLV[3, $c]) -ne $blocks[$i - 1].Qty) { $nb += "col $c" } } }
Chk 3 'DDC List columns == IOSummary blocks (name, qty)' $(if ($nb -or ($dlLastC - 1) -ne $blocks.Count) { 'FAIL' } else { 'PASS' }) "$($dlLastC - 1) columns. $($nb -join ',')"
$ua = @(); for ($c = 2; $c -le $dlLastC; $c++) { $s = 0; for ($r = 5; $r -le $dlLastR; $r++) { $s += N $DLV[$r, $c] }; if ($s -ne (N $DLV[3, $c]) -or (N $DLV[4, $c]) -ne $s) { $ua += "col $c ($s vs $($DLV[3,$c]))" } }
Chk 3 'GATE Assigned == Total, every column' $(if ($ua) { 'FAIL' } else { 'PASS' }) (($ua | Select-Object -First 8) -join '; ')
$spare = N $wb.Worksheets.Item("Options").Range("B1").Value2
$per = @{}; for ($i = 0; $i -lt $blocks.Count; $i++) { $b = $blocks[$i]; $t = 0; foreach ($k in 3..7) { $t += N $IOV[$b.T, $k] }; $per[$i + 1] = $t }
$over = @(); $maxL = 0; $names = @{}; $dupP = @(); $ptsAll = 0
for ($r = 5; $r -le $dlLastR; $r++) { $pn = [string]$DLV[$r, 1]; if (-not $pn) { continue }; if ($names[$pn]) { $dupP += $pn }; $names[$pn] = 1
  $l = 0; for ($c = 2; $c -le $dlLastC; $c++) { $l += (N $DLV[$r, $c]) * $per[$c - 1] }; $ptsAll += $l; $ws_ = [math]::Round($l * (1 + $spare)); if ($ws_ -gt $maxL) { $maxL = $ws_ }; if ($ws_ -gt 250) { $over += "$pn=$ws_" } }
Chk 3 'GATE no panel > 250 incl. spare' $(if ($over) { 'FAIL' } else { 'PASS' }) "$($names.Count) panels, max $maxL. $($over -join '; ')"
Chk 3 'Panel names unique' $(if ($dupP) { 'FAIL' } else { 'PASS' }) ($dupP -join '; ')
Chk 3 'Points across panels == IOSummary' $(if ([math]::Abs($ptsAll - ($tot | Measure-Object -Sum).Sum) -lt 0.5) { 'PASS' } else { 'FAIL' }) "panels $ptsAll vs IOSummary $(($tot | Measure-Object -Sum).Sum)"
Chk 3 'Panel rows within macro range A5:A200' $(if ($dlLastR -le 200) { 'PASS' } else { 'FAIL' }) "last row $dlLastR"
$fmtBad = @(); if ($dl.Cells.Item(2, $dlLastC).Interior.Color -ne 14806254) { $fmtBad += 'row2 fill last col' }; if ($dl.Cells.Item($dlLastR, $dlLastC).Borders.Item(9).LineStyle -ne 1) { $fmtBad += 'border last cell' }
Chk 3 'DDC List design covers all columns/rows' $(if ($fmtBad) { 'FAIL' } else { 'PASS' }) ($fmtBad -join '; ')
$mcode = $wb.VBProject.VBComponents.Item('DDCSummaryModule').CodeModule; $mtxt = $mcode.Lines(1, $mcode.CountOfLines); $lim = if ($mtxt -match 'Offset\(0, (\d+)\)\)\.Value2') { [int]$Matches[1] } else { 150 }
Chk 3 'Phase 4 macro reads all equipment columns' $(if ($dlLastC - 1 -gt $lim) { 'FAIL' } else { 'PASS' }) "$($dlLastC - 1) equipment columns; macro limit $lim"

# ---------------- Phase 4-7 (generated by estimator)
$ds = $wb.Worksheets.Item("DDCSummary"); $DSV = $ds.UsedRange.Value2; $dsPanels = @(); for ($r = 1; $r -le $DSV.GetLength(0); $r++) { if ([string]$DSV[$r, 1] -eq 'QTY' -and [string]$DSV[$r, 2]) { $dsPanels += [string]$DSV[$r, 2] } }
$miss = @($names.Keys | Where-Object { $dsPanels -notcontains $_ }); $extra = @($dsPanels | Where-Object { -not $names.ContainsKey($_) })
Chk 4 'DDCSummary panels == DDC List' $(if ($miss -or $extra) { 'FAIL' } else { 'PASS' }) "DDCSummary $($dsPanels.Count) panels vs DDC List $($names.Count); missing $($miss.Count) (e.g. $(($miss | Select-Object -First 3) -join ', ')); not in list $($extra.Count) (e.g. $(($extra | Select-Object -First 3) -join ', '))"
$need = 0; for ($r = 1; $r -le $DSV.GetLength(0); $r++) { for ($c = 1; $c -le [math]::Min(28, $DSV.GetLength(1)); $c++) { if ([string]$DSV[$r, $c] -match 'Seperate|Separate') { $need++ } } }
Chk 4 'DDCSummary "Need To Seperate"' $(if ($need) { 'FAIL' } else { 'PASS' }) "$need cells"
$bq = $wb.Worksheets.Item("BOQ"); $BQV = $bq.Range("A1:N107").Value2; $lastItem = 0; $naB = @(); for ($r = 5; $r -le 107; $r++) { if ([string]$BQV[$r, 4]) { $lastItem = $r; for ($c = 4; $c -le 13; $c++) { if ([string]$bq.Cells.Item($r, $c).Text -match '#N/A|#VALUE|#REF|Seperate') { $naB += "r$r"; break } } } }
$jf = [string]$bq.Range("J108").Formula; $mf2 = [string]$bq.Range("J110").Formula
$rngOk = ($jf -match ":J(\d+)\)" -and [int]$Matches[1] -ge $lastItem); $rngOk2 = ($mf2 -match ":M(\d+)\)" -and [int]$Matches[1] -ge $lastItem)
Chk 7 'BOQ TOTAL / TOTAL COST ranges cover all items' $(if ($rngOk -and $rngOk2) { 'PASS' } else { 'FAIL' }) "items to row $lastItem; J108 $jf ; J110 $mf2"
Chk 7 'BOQ rows with errors' $(if ($naB) { 'FAIL' } else { 'PASS' }) ($naB -join ', ')
Chk 7 'BOQ LOCAL cost not negative' $(if ((N $bq.Range("J113").Value2) -ge 0) { 'PASS' } else { 'FAIL' }) "J113 = $([math]::Round((N $bq.Range('J113').Value2),2))"
$bd = $wb.Worksheets.Item("Breakdown"); $cpg = $wb.Worksheets.Item("Cover Page")
Chk 7 'Cover Page total == Breakdown D9 == BOQ total' $(if ((N $cpg.Range("H11").Value2) -eq (N $bd.Range("D9").Value2) -and (N $bd.Range("D9").Value2) -eq (N $bq.Range("J108").Value2)) { 'PASS' } else { 'FAIL' }) "Cover $($cpg.Range('H11').Text) / Breakdown $($bd.Range('D9').Text) / BOQ $($bq.Range('J108').Text)"
$negs = @(); foreach ($r in 16..31 + 37..38 + 46..71 + 77..87) { if ((N $bd.Cells.Item($r, 9).Value2) -lt 0) { $negs += "r$r" } }
Chk 7 'Breakdown no negative cost lines' $(if ($negs) { 'FAIL' } else { 'PASS' }) ($negs -join ', ')
Chk 7 'Breakdown project details filled' $(if ([string]$bd.Range("D6").Text -and [string]$bd.Range("G6").Text -and [string]$bd.Range("D8").Text) { 'PASS' } else { 'FAIL' }) "$($bd.Range('D6').Text) | $($bd.Range('G6').Text) | $($bd.Range('D8').Text) | margin $($bd.Range('J11').Text)"
Chk 7 'Cover Page To: filled' $(if ([string]$cpg.Range("F3").Text) { 'PASS' } else { 'WARN' }) 'F3 blank'
$opt = $wb.Worksheets.Item("Options"); Chk 0 'Options FCU protocol vs drawings (IP unitary)' $(if ([string]$opt.Range("B13").Text -match 'BACnet') { 'PASS' } else { 'WARN' }) "B13 = $($opt.Range('B13').Text)"

# ---------------- side files
$root = Get-ChildItem $proj -File | Where-Object { $_.Name -notmatch '^~\$' } | ForEach-Object Name
$allowed = @((Split-Path -Leaf $env:BMS_WB), '_ESTIMATION_STATE.md', "$pname - BMS Equipment Takeoff.xlsx", "$pname - IO Summary Sources.xlsx", "$pname - DDC List Sources.xlsx") + @($env:BMS_KEEP -split ';' | Where-Object { $_ })   # BMS_KEEP: other files allowed in the root (e.g. the untouched original)
$stray = @($root | Where-Object { $allowed -notcontains $_ })
Chk 'files' 'Project root holds only final files' $(if ($stray) { 'FAIL' } else { 'PASS' }) ($stray -join '; ')
$x2 = New-Object -ComObject Excel.Application; $x2.Visible = $false; $x2.DisplayAlerts = $false
try {
  $s1 = $x2.Workbooks.Open("$proj\$pname - IO Summary Sources.xlsx", 0, $true); $u = $s1.Worksheets.Item(1).UsedRange.Value2; $lr = $u.GetLength(0)
  $w1s = $s1.Worksheets.Item(1); $lr = $w1s.Cells.Item($w1s.Rows.Count, 3).End(-4162).Row; $u = $w1s.Range("A1:S$lr").Value2
  $st = @(N $u[$lr, 15], N $u[$lr, 16], N $u[$lr, 17], N $u[$lr, 18], N $u[$lr, 19]); $s1.Close($false)
  Chk 'files' 'IO Summary Sources totals == IOSummary' $(if (($st -join '/') -eq ($tot -join '/')) { 'PASS' } else { 'FAIL' }) "Sources $($st -join '/') vs $($tot -join '/')"
  $s2 = $x2.Workbooks.Open("$proj\$pname - DDC List Sources.xlsx", 0, $true); $u2 = $s2.Worksheets.Item(1).UsedRange.Value2; $n2 = $u2.GetLength(0) - 2; $s2.Close($false)
  Chk 'files' 'DDC List Sources rows == DDC List panels' $(if ($n2 -eq $names.Count) { 'PASS' } else { 'FAIL' }) "$n2 vs $($names.Count)"
} finally { $x2.Quit() }
$stt = [IO.File]::ReadAllText("$proj\_ESTIMATION_STATE.md")
Chk 'files' 'State file totals == IOSummary' $(if ($stt -match [regex]::Escape(($tot | ForEach-Object { [string]$_ }) -join ' / ')) { 'PASS' } else { 'WARN' }) ''
$res


