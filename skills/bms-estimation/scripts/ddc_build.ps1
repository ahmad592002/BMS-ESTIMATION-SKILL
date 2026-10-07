param([switch]$Write)
# Phase 3: build the DDC List from the takeoff register (tag -> panel).
# A drawn panel whose load (physical + SP) x (1 + spare) exceeds 250 is split into n controllers in the
# SAME enclosure (enclosure count stays as drawn / BOQ). Tags are balanced greedily across controllers.
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$proj = Split-Path -Parent $env:BMS_WB
$reg = Get-Content "$sp\register.tsv" -Encoding UTF8 | ForEach-Object { , ($_ -split "`t") } | Where-Object { $_[1] -and $_[4] -and [double]$_[11] -gt 0 }
$map = Get-Content "$sp\el_map.txt" -Encoding UTF8 | Where-Object { $_ -and $_ -notmatch '^#' } | ForEach-Object { $a = $_ -split '\|'; [pscustomobject]@{ SN = [int]$a[0]; B = $a[1]; D = $a[2]; T = $a[3]; NT = $a[4] } }
$WBPATH = $env:BMS_WB
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($WBPATH); $x = $wb.Application
$spare = [double]$wb.Worksheets.Item("Options").Range("B1").Value2
$CAP = 250
$ws = $wb.Worksheets.Item("IOSummary"); $last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$LV = $ws.Range("A1:G$last").Value2
$per = @{}; $ioName = @{}; $ioQty = @{}; $sn = 0
for ($r = 1; $r -le $last; $r++) {
  if ([string]$LV[$r, 1] -eq 'QTY') { $sn++; $ioName[$sn] = [string]$LV[($r + 2), 2]; $ioQty[$sn] = [double]$LV[($r + 2), 1] }
  if ([string]$LV[$r, 2] -eq 'TOTAL') { $t = @(); for ($c = 3; $c -le 7; $c++) { $t += [double]$LV[$r, $c] }; $per[$sn] = $t } }
$nEq = $sn
# tag rows with SN + per-unit load
$items = @()
foreach ($r in $reg) { foreach ($m in $map) {
  if ($r[1] -ne $m.B -or $r[4] -notmatch ('^(' + $m.D + ')$')) { continue }
  if ($m.T -and $r[3] -notmatch $m.T) { continue }; if ($m.NT -and $r[3] -match $m.NT) { continue }
  $u = 0; foreach ($k in 0..4) { $u += $per[$m.SN][$k] }
  $items += [pscustomobject]@{ SN = $m.SN; Tag = $r[3]; Level = $r[2]; Panel = $r[6]; Qty = [double]$r[11]; Unit = $u; Load = [double]$r[11] * $u }; break } }
$hostOf = @{}; if (Test-Path "$sp\host_map.txt") { Get-Content "$sp\host_map.txt" | Where-Object { $_ -match '^\d' } | ForEach-Object { $a = $_ -split '\|'; $hostOf[[int]$a[0]] = [int]$a[1] } }
# equipment whose IOSummary qty is higher than the riser (estimator rule: take the higher of riser / client BOQ):
# the extra units have no drawn location -> add them to the panel the drawn units use, after the drawn floors
foreach ($s in 1..$nEq) { $mine = @($items | Where-Object { $_.SN -eq $s }); $have = ($mine | Measure-Object Qty -Sum).Sum
  if (-not $mine.Count -and $ioQty[$s] -gt 0 -and $hostOf.ContainsKey($s)) {
    # estimator-added equipment with no riser symbol: spread its units round-robin over the host equipment's panels
    $hp = @($items | Where-Object { $_.SN -eq $hostOf[$s] } | ForEach-Object { $_.Panel } | Select-Object -Unique | Sort-Object)
    $u = 0; foreach ($k in 0..4) { $u += $per[$s][$k] }
    for ($e = 0; $e -lt $ioQty[$s]; $e++) { $pp = $hp[$e % $hp.Count]; $lv = ($items | Where-Object { $_.Panel -eq $pp } | Select-Object -First 1).Level
      $items += [pscustomobject]@{ SN = $s; Tag = "$($ioName[$s]) #$($e+1)"; Level = $lv; Panel = $pp; Qty = 1.0; Unit = $u; Load = $u } }
    continue }
  if (-not $mine.Count -and $ioQty[$s] -gt 0) {
    # equipment drawn only as a "not counted" (qty 0) riser symbol: use that symbol's panel
    $mm = $map | Where-Object { $_.SN -eq $s } | Select-Object -First 1
    $raw = Get-Content "$sp\register.tsv" -Encoding UTF8 | ForEach-Object { , ($_ -split "`t") } | Where-Object { $mm -and $_[1] -eq $mm.B -and $_[4] -match ('^(' + $mm.D + ')$') } | Select-Object -First 1
    if ($raw) { $u = 0; foreach ($k in 0..4) { $u += $per[$s][$k] }; for ($e = 1; $e -le $ioQty[$s]; $e++) { $items += [pscustomobject]@{ SN = $s; Tag = "$($raw[3]) (counted - estimator rule)"; Level = $raw[2]; Panel = $raw[6]; Qty = 1.0; Unit = $u; Load = $u } } }
    continue }
  if ($mine.Count -and $ioQty[$s] -gt $have) { $pnl = ($mine | Group-Object Panel | Sort-Object Count -Descending | Select-Object -First 1).Name
    $u = 0; foreach ($k in 0..4) { $u += $per[$s][$k] }
    for ($e = 1; $e -le ($ioQty[$s] - $have); $e++) { $items += [pscustomobject]@{ SN = $s; Tag = "BOQ extra #$e"; Level = 'BOQ extra (location not drawn)'; Panel = $pnl; Qty = 1.0; Unit = $u; Load = $u } } } }
$pinfo = @{}; Get-Content "$sp\panels.txt" -Encoding UTF8 | ForEach-Object { $a = $_ -split '\|'; $pinfo[$a[0]] = $a }
$virtual = '^(NET-|HEAD-END)'
$rename = @{}; if (Test-Path "$sp\panel_rename.txt") { Get-Content "$sp\panel_rename.txt" | Where-Object { $_ -and $_ -notmatch '^#' } | ForEach-Object { $a = $_ -split '\|'; $rename[$a[0]] = $a[1] } }
foreach ($it in $items) { if ($rename.ContainsKey($it.Panel)) { $pinfo[$rename[$it.Panel]] = $pinfo[$it.Panel]; $it.Panel = $rename[$it.Panel] } }
$ctrl = @()   # controllers: Name, Panel, Items
foreach ($g in ($items | Group-Object Panel | Sort-Object { $pinfo[$_.Name][1] }, Name)) {
  $p = $g.Name; $load = ($g.Group | Measure-Object Load -Sum).Sum
  if ($p -match '^NET-') {
    # network units (IP unitary controllers / VRF): fill panels floor by floor, each <= 250 incl. spare
    $rank = { param($l) if ($l -match 'Basement|B0') { 0 } elseif ($l -match 'L00|Ground') { 1 } elseif ($l -match 'L0(\d)') { 1 + [int]$Matches[1] } elseif ($l -match 'Roof') { 9 } elseif ($l -match 'BOQ extra') { 10 } else { 5 } }
    $bins = @(); $cur = $null
    foreach ($it in ($g.Group | Sort-Object { & $rank $_.Level }, Tag)) {
      if (-not $cur -or ($cur.Load + $it.Load) * (1 + $spare) -gt $CAP) { $cur = [pscustomobject]@{ Load = 0.0; Items = New-Object System.Collections.Generic.List[object] }; $bins += $cur }
      $cur.Items.Add($it); $cur.Load += $it.Load }
    $n = $bins.Count
    for ($i = 0; $i -lt $n; $i++) { $lv = ($bins[$i].Items | ForEach-Object { $_.Level } | Select-Object -Unique) -join ', '
      $ctrl += [pscustomobject]@{ Name = $(if ($n -eq 1) { $p } else { "$p-$($i+1)" }); Panel = $p; Split = $n; Virtual = $true; Levels = $lv; PanelLoad = $load * (1 + $spare); Load = $bins[$i].Load; Items = $bins[$i].Items } }
    continue
  }
  $n = if ($p -match $virtual) { 1 } else { [math]::Max(1, [math]::Ceiling($load * (1 + $spare) / $CAP)) }
  while ($true) {
    $bins = @(); for ($i = 0; $i -lt $n; $i++) { $bins += [pscustomobject]@{ Load = 0.0; Items = New-Object System.Collections.Generic.List[object] } }
    # split a tag row whose load alone breaks the cap into single units
    $units = foreach ($it in $g.Group) { if ($n -gt 1 -and $it.Qty -gt 1 -and $it.Load * (1 + $spare) -gt $CAP / 2) { for ($k = 0; $k -lt $it.Qty; $k++) { [pscustomobject]@{ SN = $it.SN; Tag = "$($it.Tag) [$($k+1)/$($it.Qty)]"; Panel = $p; Qty = 1.0; Unit = $it.Unit; Load = $it.Unit } } } else { $it } }
    foreach ($it in ($units | Sort-Object Load -Descending)) { $b = $bins | Sort-Object Load | Select-Object -First 1; $b.Items.Add($it); $b.Load += $it.Load }
    $worst = ($bins | Measure-Object Load -Maximum).Maximum
    if ($p -match $virtual -or $worst * (1 + $spare) -le $CAP) { break }; $n++ }
  for ($i = 0; $i -lt $n; $i++) {
    $nm = if ($n -eq 1) { $p } else { "$p-$($i+1)" }
    $ctrl += [pscustomobject]@{ Name = $nm; Panel = $p; Split = $n; Virtual = ($p -match $virtual); PanelLoad = $load * (1 + $spare); Load = $bins[$i].Load; Items = $bins[$i].Items } }
}
# per-controller IO
foreach ($c in $ctrl) { $io = @(0.0, 0.0, 0.0, 0.0, 0.0); $q = @{}
  foreach ($it in $c.Items) { foreach ($k in 0..4) { $io[$k] += $it.Qty * $per[$it.SN][$k] }; if (-not $q.ContainsKey($it.SN)) { $q[$it.SN] = 0.0 }; $q[$it.SN] += $it.Qty }
  $c | Add-Member -Force NoteProperty IO $io; $c | Add-Member -Force NoteProperty Q $q
  $c | Add-Member -Force NoteProperty WithSpare ([math]::Round(($io[0] + $io[1] + $io[2] + $io[3] + $io[4]) * (1 + $spare), 0)) }

if ($Write) {
  $d = $wb.Worksheets.Item("DDC List")
  $wb.SaveCopyAs("$proj\Old Versions\" + [IO.Path]::GetFileNameWithoutExtension($WBPATH) + " - BACKUP before DDC List.xlsm")
  $x.ScreenUpdating = $false; $x.EnableEvents = $false
  try {
    $lastCol = [math]::Max($nEq + 1, $d.UsedRange.Column + $d.UsedRange.Columns.Count - 1)
    $lastRow = [math]::Max(200, $d.UsedRange.Row + $d.UsedRange.Rows.Count - 1)
    $d.Range($d.Cells.Item(1, 2), $d.Cells.Item(3, $lastCol)).ClearContents()
    $d.Range($d.Cells.Item(5, 1), $d.Cells.Item($lastRow, $lastCol)).ClearContents()
    $d.Range($d.Cells.Item(5, 1), $d.Cells.Item($lastRow, 1)).Interior.ColorIndex = -4142
    foreach ($cm in @($d.Comments)) { $cm.Delete() }
    $hdr = New-Object 'object[,]' 3, $nEq
    for ($s = 1; $s -le $nEq; $s++) { $hdr[0, ($s - 1)] = [double]$s; $hdr[1, ($s - 1)] = [string]$ioName[$s]; $hdr[2, ($s - 1)] = [double]$ioQty[$s] }
    $d.Range($d.Cells.Item(1, 2), $d.Cells.Item(3, $nEq + 1)).Value2 = $hdr
    $nc = $ctrl.Count; $body = New-Object 'object[,]' $nc, ($nEq + 1)
    for ($i = 0; $i -lt $nc; $i++) { $body[$i, 0] = [string]$ctrl[$i].Name
      for ($s = 1; $s -le $nEq; $s++) { if ($ctrl[$i].Q.ContainsKey($s)) { $body[$i, $s] = [double]$ctrl[$i].Q[$s] } else { $body[$i, $s] = [string]'' } } }
    $d.Range($d.Cells.Item(5, 1), $d.Cells.Item(4 + $nc, $nEq + 1)).Value2 = $body
    $lr = 4 + $nc
    $f = New-Object 'object[,]' 1, $nEq
    for ($s = 1; $s -le $nEq; $s++) { $col = $d.Cells.Item(1, $s + 1).Address($false, $false) -replace '\d', ''; $f[0, ($s - 1)] = [string]"=SUM($col`5:$col$lr)" }
    $d.Range($d.Cells.Item(4, 2), $d.Cells.Item(4, $nEq + 1)).Formula = $f
    for ($i = 0; $i -lt $nc; $i++) { $c = $ctrl[$i]; $cell = $d.Cells.Item(5 + $i, 1)
      if ($c.Split -gt 1) { $cell.Interior.Color = 65535; [void]$cell.AddComment([string]("Drawn as ONE panel " + $c.Panel + ". Load " + [math]::Round($c.PanelLoad) + " points incl. spare > 250 - split into " + $c.Split + " panels (-1 ... -" + $c.Split + ") (estimator 2026-10-02).")) }
      elseif ($c.Virtual -and $c.Panel -match '^NET-') { $cell.Interior.Color = 65535; [void]$cell.AddComment([string]("Network panel for IP unitary controllers / VRF (" + $pinfo[$c.Panel][3] + "), levels: " + $c.Levels + ". Drawn as one network " + $c.Panel + " - split into " + $c.Split + " panels, each <= 250 points incl. spare (estimator 2026-10-02).")) } }
    $x.CalculateFull()
  } finally { $x.EnableEvents = $true; $x.ScreenUpdating = $true }
  # gate 1
  $bad = 0; for ($s = 1; $s -le $nEq; $s++) { if ([double]$d.Cells.Item(3, $s + 1).Value2 -ne [double]$d.Cells.Item(4, $s + 1).Value2) { $bad++; "  column $s $($ioName[$s]): Total $($d.Cells.Item(3, $s + 1).Value2) Assigned $($d.Cells.Item(4, $s + 1).Value2)" } }
  "GATE Assigned==Total: columns failing = $bad of $nEq"
  $wb.Save(); "saved=" + $wb.Saved
}
$ctrl










