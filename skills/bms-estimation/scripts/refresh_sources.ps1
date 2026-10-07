# Rebuild "IO Summary Sources.xlsx" from the LIVE IOSummary (after estimator edits), keeping the
# drawing source of every row that still matches the previous Sources workbook.
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$proj = Split-Path -Parent $env:BMS_WB
$pname = if ($env:BMS_PROJECT) { $env:BMS_PROJECT } else { (Split-Path -Leaf $proj) -replace '^\d+-\s*', '' }   # side files are named "<project> - ..."
$srcPath = "$proj\$pname - IO Summary Sources.xlsx"
$YEL = 65535
$EDIT = 'Estimator edit (GTS standard)'

# ---------- 1. live IOSummary
$wbL = [Runtime.InteropServices.Marshal]::BindToMoniker($env:BMS_WB)
$ws = $wbL.Worksheets.Item("IOSummary")
$last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$LV = $ws.Range("A1:Q$last").Value2
$blocks = New-Object System.Collections.Generic.List[object]
for ($r = 1; $r -le $last; $r++) {
  if ([string]$LV[$r, 1] -ne 'QTY') { continue }
  $q = $r + 2; $eqN = ([string]$LV[$q, 2]).Trim(); $eqQ = [double]$LV[$q, 1]; $eqNt = [string]$LV[$q, 17]; $b = [pscustomobject]@{ Q = $q; Name = $eqN; Qty = $eqQ; EqNote = $eqNt; EqYel = $false; Rows = New-Object System.Collections.Generic.List[object] }
  if ($b.EqNote) { $b.EqYel = ($ws.Cells.Item($q, 17).Interior.Color -eq $YEL) }
  $comp = ''; $t = $q + 1
  while ([string]$LV[$t, 2] -ne 'TOTAL') {
    $nm = ([string]$LV[$t, 2]).Trim(); $io = @(); $any = $false
    for ($c = 3; $c -le 7; $c++) { $v = $LV[$t, $c]; if ($v -is [double]) { $io += $v; if ($v -ne 0) { $any = $true } } else { $io += 0 } }
    $dev = [string]$LV[$t, 14]
    if ($nm -and -not $any -and -not $dev) { $comp = $nm }
    elseif ($nm) {
      $ycells = @{}
      if ([string]$LV[$t, 17]) { foreach ($c in 2, 3, 4, 5, 6, 7, 13, 14, 15, 17) { if ($ws.Cells.Item($t, $c).Interior.Color -eq $YEL) { $ycells[$c] = $true } } }
      $mo = [string]$LV[$t, 15]; $dq0 = $LV[$t, 13]; $pr = [string]$LV[$t, 16]; $nt0 = [string]$LV[$t, 17]
      $b.Rows.Add([pscustomobject]@{ Comp = $comp; Name = $nm; IO = $io; Dev = $dev; Model = $mo; DevQ = $dq0; Proto = $pr; Note = $nt0; Yel = $ycells })
    }
    $t++
  }
  $blocks.Add($b)
}
"live blocks=$($blocks.Count)"

# ---------- 2. previous Sources workbook (sources per row)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
try {
  $old = $x.Workbooks.Open($srcPath, 0, $true)
  $S1 = $old.Worksheets.Item('Equipment Summary').UsedRange.Value2
  $S2 = $old.Worksheets.Item('IO Points by Equipment').UsedRange.Value2
  $S5 = $old.Worksheets.Item('Field Devices').UsedRange.Value2
  $S3 = $old.Worksheets.Item('Source Drawings').UsedRange.Value2; $rows3 = @(); for ($r = 2; $r -le $S3.GetLength(0); $r++) { $rw = @(); for ($c = 1; $c -le 5; $c++) { $rw += [string]$S3[$r, $c] }; $rows3 += , $rw }
  $sumOld = @{}; for ($r = 2; $r -le $S1.GetLength(0); $r++) { $n = ([string]$S1[$r, 3]).Trim(); if ($n) { $sumOld[$n] = $r } }
  $srcRow = @{}; $srcEq = @{}
  for ($r = 2; $r -le $S2.GetLength(0); $r++) { $e = ([string]$S2[$r, 2]).Trim(); $p = ([string]$S2[$r, 5]).Trim(); if (-not $p) { continue }
    $val = @([string]$S2[$r, 19], [string]$S2[$r, 20], [string]$S2[$r, 21], [string]$S2[$r, 22])
    $srcRow["$e||$($p.ToUpper())"] = $val
    if (-not $srcEq.ContainsKey($e) -and $val[0]) { $srcEq[$e] = $val } }
  $whyOld = @{}; for ($r = 2; $r -le $S5.GetLength(0); $r++) { $whyOld[[string]$S5[$r, 1]] = [string]$S5[$r, 4] }

  # ---------- 3. build rows
  $sumRows = @(); $ptRows = @(); $ptYel = @(); $chk = @(); $devAgg = [ordered]@{}; $nEdit = 0
  for ($i = 0; $i -lt $blocks.Count; $i++) {
    $b = $blocks[$i]; $sn = $i + 1; $o = $sumOld[$b.Name]
    $one = @(0, 0, 0, 0, 0)
    $ptRows += , @($sn, $b.Name, $b.Qty, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'HDR'); $ptYel += , @{}
    foreach ($p in $b.Rows) {
      for ($k = 0; $k -lt 5; $k++) { $one[$k] += $p.IO[$k] }
      $s = $srcRow["$($b.Name)||$($p.Name.ToUpper())"]
      if (-not $s) {
        $base = $srcEq[$b.Name]; $nEdit++
        $s = @($(if ($base) { $base[0] } else { '' }), $(if ($base) { $base[1] } else { '' }), $EDIT, 'Estimator edit / GTS standard')
      }
      $devTxt = ''; $dq = ''
      if ($p.Dev) { $devTxt = $p.Dev + '  ->  ' + $p.Model; if ($p.DevQ -is [double]) { $dq = [double]$p.DevQ }
        $dk = "$($p.Dev)||$($p.Model)"; if (-not $devAgg.Contains($dk)) { $devAgg[$dk] = @{ Dev = $p.Dev; Model = $p.Model; Q = 0.0; Eq = [ordered]@{} } }
        if ($p.DevQ -is [double]) { $devAgg[$dk].Q += [double]$p.DevQ }; $devAgg[$dk].Eq[$b.Name] = 1 }
      $ptRows += , @($sn, $b.Name, $b.Qty, $p.Comp, $p.Name, $p.IO[0], $p.IO[1], $p.IO[2], $p.IO[3], $p.IO[4],
        ($p.IO[0] * $b.Qty), ($p.IO[1] * $b.Qty), ($p.IO[2] * $b.Qty), ($p.IO[3] * $b.Qty), ($p.IO[4] * $b.Qty),
        $devTxt, $dq, $p.Proto, $s[0], $s[1], $s[2], $s[3], $p.Note)
      $ptYel += , $p.Yel
      if ($p.Yel.Count) { $chk += , @($sn, $b.Name, $b.Qty, $p.Name, ($p.IO[0] + $p.IO[1] + $p.IO[2] + $p.IO[3]), $p.IO[4], $s[3], $p.Note) }
    }
    if ($b.EqYel) { $chk += , @($sn, $b.Name, $b.Qty, '(equipment)', '', '', 'Equipment', $b.EqNote) }
    $c = @('', $(($b.Name -split ' - ')[0]), $b.Name, $b.Qty, '', '', '', '', '')
    if ($o) { for ($k = 4; $k -le 8; $k++) { $c[$k] = [string]$S1[$o, ($k + 1)] } }
    $sumRows += , @($sn, $c[1], $b.Name, $b.Qty, $c[4], $c[5], $c[6], $c[7], $c[8], $one[0], $one[1], $one[2], $one[3], $one[4],
      ($one[0] * $b.Qty), ($one[1] * $b.Qty), ($one[2] * $b.Qty), ($one[3] * $b.Qty), ($one[4] * $b.Qty))
  }

  # ---------- 4. write workbook
  $wb = $x.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 5) { $null = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $w1 = $wb.Worksheets.Item(1); $w1.Name = 'Equipment Summary'
  $w2 = $wb.Worksheets.Item(2); $w2.Name = 'IO Points by Equipment'
  $w3 = $wb.Worksheets.Item(3); $w3.Name = 'Source Drawings'
  $w4 = $wb.Worksheets.Item(4); $w4.Name = 'To Check (yellow)'
  $w5 = $wb.Worksheets.Item(5); $w5.Name = 'Field Devices'
  function Put($sh, [string[]]$hdr, $rows) {
    $n = $rows.Count; $c = $hdr.Count; $a = New-Object 'object[,]' ($n + 1), $c
    for ($j = 0; $j -lt $c; $j++) { $a[0, $j] = [string]$hdr[$j] }
    for ($i = 0; $i -lt $n; $i++) { for ($j = 0; $j -lt $c; $j++) { $v = $rows[$i][$j]
        if ($v -is [double] -or $v -is [int]) { if ($v -eq 0 -and $j -ge 5) { $a[($i + 1), $j] = [string]'' } else { $a[($i + 1), $j] = [double]$v } } else { $a[($i + 1), $j] = [string]$v } } }
    $sh.Range("A1").Resize($n + 1, $c).Value2 = $a
    $h = $sh.Range("A1").Resize(1, $c); $h.Font.Bold = $true; $h.Interior.Color = 0x7F3F1F; $h.Font.Color = 0xFFFFFF; $h.WrapText = $true; $h.VerticalAlignment = -4108
    $null = $sh.Range("A1").Resize($n + 1, $c).AutoFilter()
    try { $sh.Parent.Activate(); $sh.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.FreezePanes = $true } catch { }
  }
  Put $w1 @('SN', 'Building', 'Equipment (EquipmentList)', 'Qty', 'Function', 'Type', 'IO template', 'Source schematic sheet(s)', 'Basis', 'DI (1)', 'AI (1)', 'AO (1)', 'DO (1)', 'SP (1)', 'DI (all)', 'AI (all)', 'AO (all)', 'DO (all)', 'SP (all)') $sumRows
  $n1 = $sumRows.Count + 1
  $wd = 5, 8, 52, 6, 14, 18, 18, 70, 30, 6, 6, 6, 6, 6, 8, 8, 8, 8, 8; for ($j = 0; $j -lt $wd.Count; $j++) { $w1.Columns.Item($j + 1).ColumnWidth = $wd[$j] }
  $w1.Range("H2:H$n1").WrapText = $true
  $tr = $n1 + 1; $w1.Cells.Item($tr, 3).Value2 = 'TOTAL'
  foreach ($c in 10..19) { $col = $w1.Cells.Item(1, $c).Address($false, $false) -replace '\d', ''; $w1.Cells.Item($tr, $c).Formula = "=SUBTOTAL(9,$col`2:$col$n1)" }
  $w1.Range("A$tr`:S$tr").Font.Bold = $true
  for ($r = 2; $r -le $n1; $r++) { $bs = [string]$w1.Cells.Item($r, 9).Value2; $w1.Range("I$r").Interior.Color = $(if ($bs -eq 'Drawing' -or $bs -eq 'Estimator block') { 0xCCFFCC } else { 0x99FFFF }); if ($bs -match 'CHECK:') { $w1.Range("I$r").Interior.Color = $YEL; if ($bs -match 'QTY:') { $w1.Range("D$r").Interior.Color = $YEL } } }

  Put $w2 @('SN', 'Equipment', 'Qty', 'Component', 'IO point', 'DI', 'AI', 'AO', 'DO', 'SP', 'DI x Qty', 'AI x Qty', 'AO x Qty', 'DO x Qty', 'SP x Qty', 'Field device  ->  model', 'Device qty (all units)', 'SP protocol', 'Source drawing', 'Drawing title', 'How it was read', 'Basis', 'Note') $ptRows
  $n2 = $ptRows.Count + 1
  $wd = 5, 44, 6, 20, 52, 5, 5, 5, 5, 5, 7, 7, 7, 7, 7, 34, 6, 10, 34, 46, 26, 28, 40; for ($j = 0; $j -lt $wd.Count; $j++) { $w2.Columns.Item($j + 1).ColumnWidth = $wd[$j] }
  # IOSummary column -> Sources column letter
  $map = @{ 2 = 'E'; 3 = 'F'; 4 = 'G'; 5 = 'H'; 6 = 'I'; 7 = 'J'; 13 = 'Q'; 14 = 'P'; 15 = 'P'; 17 = 'W' }
  for ($i = 0; $i -lt $ptRows.Count; $i++) { $r = $i + 2
    if ($ptRows[$i][22] -eq 'HDR') { $rg = $w2.Range("A$r`:W$r"); $rg.Interior.Color = 0xF2E6D9; $rg.Font.Bold = $true; $w2.Cells.Item($r, 23).Value2 = ''; continue }
    foreach ($c in $ptYel[$i].Keys) { $w2.Range("$($map[$c])$r").Interior.Color = $YEL }
    if ($ptRows[$i][20] -eq $EDIT) { $w2.Range("U$r").Interior.Color = 0xCCFFCC } }
  $tr = $n2 + 1; $w2.Cells.Item($tr, 5).Value2 = 'TOTAL (visible rows)'
  foreach ($c in 6..15) { $col = $w2.Cells.Item(1, $c).Address($false, $false) -replace '\d', ''; $w2.Cells.Item($tr, $c).Formula = "=SUBTOTAL(9,$col`2:$col$n2)" }
  $w2.Range("A$tr`:W$tr").Font.Bold = $true

  Put $w3 @('Sheet', 'Drawing number', 'Title', 'How points were read', 'Used for equipment') $rows3; $w3.Range("E2:E$($rows3.Count + 1)").WrapText = $true
  $wd = 9, 32, 70, 30, 120; for ($j = 0; $j -lt $wd.Count; $j++) { $w3.Columns.Item($j + 1).ColumnWidth = $wd[$j] }

  Put $w4 @('SN', 'Equipment', 'Qty', 'IO point', 'Hardwired points (1)', 'SP (1)', 'Basis', 'Why it is yellow (verify)') $chk
  $wd = 5, 52, 6, 52, 10, 7, 34, 70; for ($j = 0; $j -lt $wd.Count; $j++) { $w4.Columns.Item($j + 1).ColumnWidth = $wd[$j] }

  $rows5 = @(); foreach ($k in ($devAgg.Keys | Sort-Object { - $devAgg[$_].Q })) { $d = $devAgg[$k]
    $why = $whyOld[$d.Dev]; if ($d.Dev -match '^\[Water\] \[Valve') { $why = 'GTS standard valve selection VVF42.65-50 + SKB62/F - valve on command row, actuator on feedback row; size to confirm' }
    $rows5 += , @($d.Dev, $d.Model, [double]$d.Q, [string]$why, (($d.Eq.Keys) -join '; ')) }
  Put $w5 @('Field device (FieldDevices)', 'Model', 'Total qty', 'Why / drawing symbol / previous-project choice', 'Used on equipment') $rows5
  $wd = 46, 20, 10, 70, 120; for ($j = 0; $j -lt $wd.Count; $j++) { $w5.Columns.Item($j + 1).ColumnWidth = $wd[$j] }; $w5.Range("D2:E$($rows5.Count + 1)").WrapText = $true

  foreach ($sh in @($w1, $w2, $w3, $w4, $w5)) { $sh.Cells.Font.Name = 'Calibri'; $sh.Cells.Font.Size = 10 }
  $x.CalculateFull(); $w1.Activate(); $x.ScreenUpdating = $true
  $old.Close($false)
  $bk = "$proj\Old Versions\$pname - IO Summary Sources (before regen " + (Get-Date -Format yyyy-MM-dd) + ").xlsx"
  if (-not (Test-Path $bk)) { Copy-Item $srcPath $bk }
  [IO.File]::Delete($srcPath); $wb.SaveAs($srcPath, 51)
  "point rows=$($ptRows.Count - $blocks.Count) ; rows from estimator edits=$nEdit ; to-check rows=$($chk.Count) ; devices=$($rows5.Count)"
  "totals all: DI={0} AI={1} AO={2} DO={3} SP={4}" -f $w1.Cells.Item($n1 + 1, 15).Value2, $w1.Cells.Item($n1 + 1, 16).Value2, $w1.Cells.Item($n1 + 1, 17).Value2, $w1.Cells.Item($n1 + 1, 18).Value2, $w1.Cells.Item($n1 + 1, 19).Value2
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }





