param([string]$Out)
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$sp\io_lib.ps1" | Out-Null
. "$sp\io_flags.ps1"
$YEL = 65535
. "$sp\io_flags2.ps1"
foreach ($k in @($TPL.Keys)) { foreach ($p in $TPL[$k]) {
  $ex = Get-DeviceExtra $p $k; if ($ex -and $ex[0]) { $p.Dev = $ex[0]; $p.DevQ = $ex[1]; $p.DevNote = $ex[2] }
  $ck = Get-PointChecks $p $k
  $p | Add-Member -Force -NotePropertyName Checks -NotePropertyValue $ck
  $p | Add-Member -Force -NotePropertyName Flag -NotePropertyValue (($ck | ForEach-Object { $_[1] }) -join '; ')
} }
$EqFlag = @{}; for ($ii=0; $ii -lt $eq.Count; $ii++) { $sn0 = [int](($eq[$ii] -split '\|')[0]); $kk = $map[$ii]; $nt = @()
  if ($EqQty.ContainsKey($sn0)) { $nt += ('QTY: ' + $EqQty[$sn0]) }; if ($EqName.ContainsKey($sn0)) { $nt += $EqName[$sn0] }; if ($NameNote.ContainsKey($kk) -and $NameNote[$kk]) { $nt += $NameNote[$kk] }
  if ($nt.Count) { $EqFlag[$sn0] = ($nt -join '; ') } }
# sheet titles
$title = @{}
foreach ($p in $pts) { if ($p.System -and -not $title.ContainsKey($p.Sheet)) { $title[$p.Sheet] = $p.System } }
$title['001'] = 'CHILLED WATER PLANT P&ID - CHILLERS (CONTROL SCHEMATIC SHEET-1)'
$title['046'] = 'DOMESTIC WATER SYSTEM (CONTROL SCHEMATIC SHEET-46)'
function DrawNo([string]$s) { if ($s) { "2301092-PC-AMU-DR-B-93-ZZZ-$s" } else { '' } }
function HowRead($p) { if ($p.SrcHow -like 'Template:*') { if ($p.SrcSheet) { 'Read from drawing (entered by hand)' } else { 'Typical template' } } elseif ($p.SrcHow -eq 'Schedule table') { 'BMS schedule table' } elseif ($p.SrcHow -eq 'Dot strip') { 'DDC point strip (dots)' } else { [string]$p.SrcHow } }

$sumRows = @(); $ptRows = @(); $useBy = @{}; $assm = @(); $devAgg = @{}; $ptChecks = @(); $ptIO = @()
$modelOf = @{}; foreach ($l in [IO.File]::ReadAllLines("$sp\chk\FieldDevices.txt", [Text.Encoding]::UTF8)) { $a = [regex]::Match($l, '\| A\d+=([^|]+)'); $b = [regex]::Match($l, '\| B\d+=([^|]+)'); if ($a.Success -and $b.Success) { $modelOf[$a.Groups[1].Value.Trim()] = $b.Groups[1].Value.Trim() } }
for ($i=0; $i -lt $eq.Count; $i++) {
  $e = $eq[$i] -split '\|'; $sn = [int]$e[0]; $qty = [double]$e[1]; $name = $e[2]; $fn = $e[3]; $ty = $e[4]; $key = $map[$i]
  $bld = ($name -split ' - ')[0]
  $lines = Build-Lines $TPL[$key] $key
  $one = @{ DI=0; AI=0; AO=0; DO=0; SP=0 }; $sheets = @{}; $bases = @{}
  $comp = ''
  $ptRows += ,@($sn, $name, $qty, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'HDR'); $ptChecks += ,@(); $ptIO += ,@(0,0,0,0)
  foreach ($p in $lines) {
    if ($p.PSObject.Properties['Title']) { $comp = $p.Title; continue }
    foreach ($k in 'DI','AI','AO','DO','SP') { $one[$k] += [int]$p.$k }
    if ($p.SrcSheet) { $sheets[$p.SrcSheet] = 1; if (-not $useBy.ContainsKey($p.SrcSheet)) { $useBy[$p.SrcSheet] = @{} }; $useBy[$p.SrcSheet][$name] = 1 }
    $bases[$p.Basis] = 1
    $proto = if ([int]$p.SP -gt 0) { [string]$p.C } else { '' }
    $note = @(); if ($p.SrcNote) { $note += [string]$p.SrcNote }; if ($p.C -and [int]$p.SP -eq 0) { $note += [string]$p.C }; if ($p.DevNote) { $note += [string]$p.DevNote }; if ($p.Flag) { $note += ("CHECK: " + [string]$p.Flag) }
    if ($p.Dev) { $dk = [string]$p.Dev; if (-not $devAgg.ContainsKey($dk)) { $devAgg[$dk] = @{ Q=0; Eq=@{}; Why=@{} } }; $devAgg[$dk].Q += [double]$p.DevQ * $qty; $devAgg[$dk].Eq[$name] = 1; if ($p.DevNote) { $devAgg[$dk].Why[[string]$p.DevNote] = 1 } }
    $row = @($sn, $name, $qty, $comp, [string]$p.L,
      [double]$p.DI, [double]$p.AI, [double]$p.AO, [double]$p.DO, [double]$p.SP,
      ([double]$p.DI*$qty), ([double]$p.AI*$qty), ([double]$p.AO*$qty), ([double]$p.DO*$qty), ([double]$p.SP*$qty),
      $(if ($p.Dev) { [string]$p.Dev + '  ->  ' + [string]$modelOf[[string]$p.Dev] } else { '' }), $(if ($p.Dev) { [double]$p.DevQ * $qty } else { '' }), $proto,
      (DrawNo $p.SrcSheet), $(if ($p.SrcSheet) { $title[$p.SrcSheet] } else { '' }),
      (HowRead $p), [string]$p.Basis, (($note | Select-Object -Unique) -join '; '))
    $ptRows += ,$row; $ptChecks += ,$p.Checks; $ptIO += ,@([int]$p.DI,[int]$p.AI,[int]$p.AO,[int]$p.DO)
    if ($p.Flag) { $assm += ,@($sn, $name, $qty, [string]$p.L, ([int]$p.DI + [int]$p.AI + [int]$p.AO + [int]$p.DO), [int]$p.SP, [string]$p.Basis, [string]$p.Flag) }
  }
  $basis0 = if ($bases.ContainsKey('Assumed - confirm')) { 'Includes assumed points' } elseif ($bases.ContainsKey('Typical points - no dedicated schematic')) { 'Typical points (no dedicated schematic)' } elseif ($bases.ContainsKey('Drawing - check on sheet')) { 'Drawing - some points to check' } elseif ($bases.ContainsKey('Estimator block (your entry)')) { 'Estimator block' } else { 'Drawing' }
  $basis = $basis0; if ($EqFlag.ContainsKey($sn)) { $basis = $basis0 + ' | CHECK: ' + $EqFlag[$sn] }
  $srcList = (($sheets.Keys | Sort-Object) | ForEach-Object { "B-93-$_ " + $(if ($title[$_]) { '(' + $title[$_] + ')' } else { '' }) }) -join '; '
  $sumRows += ,@($sn, $bld, $name, $qty, $fn, $ty, $key, $srcList, $basis,
    [double]$one.DI, [double]$one.AI, [double]$one.AO, [double]$one.DO, [double]$one.SP,
    ($one.DI*$qty), ($one.AI*$qty), ($one.AO*$qty), ($one.DO*$qty), ($one.SP*$qty))
}

$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$wb = $x.Workbooks.Add()
while ($wb.Worksheets.Count -lt 5) { $null = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
$w1 = $wb.Worksheets.Item(1); $w1.Name = 'Equipment Summary'
$w2 = $wb.Worksheets.Item(2); $w2.Name = 'IO Points by Equipment'
$w3 = $wb.Worksheets.Item(3); $w3.Name = 'Source Drawings'
$w4 = $wb.Worksheets.Item(4); $w4.Name = 'To Check (yellow)'
$w5 = $wb.Worksheets.Item(5); $w5.Name = 'Field Devices'
function Put($ws, [string[]]$hdr, $rows, [int[]]$txtCols) {
  $n = $rows.Count; $c = $hdr.Count
  $a = New-Object 'object[,]' ($n+1), $c
  for ($j=0; $j -lt $c; $j++) { $a[0,$j] = [string]$hdr[$j] }
  for ($i=0; $i -lt $n; $i++) { for ($j=0; $j -lt $c; $j++) { $v = $rows[$i][$j]
      if ($v -is [double] -or $v -is [int]) { if ($v -eq 0 -and $j -ge 5) { $a[($i+1),$j] = [string]'' } else { $a[($i+1),$j] = [double]$v } }
      else { $sv = [string]$v; if ($sv -match '^\d' -and ($txtCols -contains $j)) { $sv = "'" + $sv }; $a[($i+1),$j] = $sv } } }
  $ws.Range("A1").Resize($n+1, $c).Value2 = $a
  $h = $ws.Range("A1").Resize(1, $c); $h.Font.Bold = $true; $h.Interior.Color = 0x7F3F1F; $h.Font.Color = 0xFFFFFF; $h.WrapText = $true; $h.VerticalAlignment = -4108
  $null = $ws.Range("A1").Resize($n+1, $c).AutoFilter()
  $ws.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.FreezePanes = $true
}

# 1 Summary
Put $w1 @('SN','Building','Equipment (EquipmentList)','Qty','Function','Type','IO template','Source schematic sheet(s)','Basis','DI (1)','AI (1)','AO (1)','DO (1)','SP (1)','DI (all)','AI (all)','AO (all)','DO (all)','SP (all)') $sumRows @()
$n1 = $sumRows.Count + 1
$wd = 5,8,52,6,14,18,18,70,30,6,6,6,6,6,8,8,8,8,8; for ($j=0; $j -lt $wd.Count; $j++) { $w1.Columns.Item($j+1).ColumnWidth = $wd[$j] }
$w1.Range("H2:H$n1").WrapText = $true
$tr = $n1 + 1; $w1.Cells.Item($tr, 3).Value2 = 'TOTAL'
foreach ($c in 10..19) { $col = $w1.Cells.Item(1,$c).Address($false,$false) -replace '\d',''; $w1.Cells.Item($tr,$c).Formula = "=SUBTOTAL(9,$col`2:$col$n1)" }
$w1.Range("A$tr`:S$tr").Font.Bold = $true
for ($r=2; $r -le $n1; $r++) { $b = [string]$w1.Cells.Item($r,9).Value2; $w1.Range("I$r").Interior.Color = $(if ($b -eq 'Drawing' -or $b -eq 'Estimator block') { 0xCCFFCC } else { 0x99FFFF }); if ($b -match 'CHECK:') { $w1.Range("I$r").Interior.Color = $YEL; if ($b -match 'QTY:') { $w1.Range("D$r").Interior.Color = $YEL } } }

# 2 Points
$hdr2 = @('SN','Equipment','Qty','Component','IO point','DI','AI','AO','DO','SP','DI x Qty','AI x Qty','AO x Qty','DO x Qty','SP x Qty','Field device  ->  model','Device qty (all units)','SP protocol','Source drawing','Drawing title','How it was read','Basis','Note')
$rows2 = @(); foreach ($r in $ptRows) { $rows2 += ,($r[0..22]) }
Put $w2 $hdr2 $rows2 @()
$n2 = $rows2.Count + 1
$wd = 5,44,6,20,52,5,5,5,5,5,7,7,7,7,7,34,6,10,34,46,26,28,40; for ($j=0; $j -lt $wd.Count; $j++) { $w2.Columns.Item($j+1).ColumnWidth = $wd[$j] }
for ($i=0; $i -lt $ptRows.Count; $i++) {
  $r = $i + 2
  if ($ptRows[$i][22] -eq 'HDR') { $rg = $w2.Range("A$r`:W$r"); $rg.Interior.Color = 0xF2E6D9; $rg.Font.Bold = $true; $w2.Cells.Item($r,23).Value2 = '' }
  else { $ck = $ptChecks[$i]; foreach ($c in $ck) { $t = $c[0]; if ($t -eq 'IO') { $io = $ptIO[$i]; foreach ($pair in @(@('F',$io[0]),@('G',$io[1]),@('H',$io[2]),@('I',$io[3]))) { if ([int]$pair[1] -gt 0) { $w2.Range("$($pair[0])$r").Interior.Color = $YEL } } } elseif ($t -eq 'SP') { $w2.Range("J$r").Interior.Color = $YEL } elseif ($t -eq 'DEV') { $w2.Range("P$r").Interior.Color = $YEL } elseif ($t -eq 'DEVQ') { $w2.Range("Q$r").Interior.Color = $YEL } }; if ($ck.Count) { $w2.Range("W$r").Interior.Color = $YEL } }
}
$tr = $n2 + 1; $w2.Cells.Item($tr, 5).Value2 = 'TOTAL (visible rows)'
foreach ($c in 6..15) { $col = $w2.Cells.Item(1,$c).Address($false,$false) -replace '\d',''; $w2.Cells.Item($tr,$c).Formula = "=SUBTOTAL(9,$col`2:$col$n2)" }
$w2.Range("A$tr`:W$tr").Font.Bold = $true

# 3 Source drawings
$rows3 = @()
$allSheets = @(@($title.Keys) + @('002','013','023','035','047','051') | Sort-Object -Unique)
foreach ($s in $allSheets) {
  $users = if ($useBy.ContainsKey($s)) { ($useBy[$s].Keys | Sort-Object) -join '; ' } else { '(not used in IOSummary)' }
  $hw = @($pts | Where-Object { $_.Sheet -eq $s } | Select-Object -ExpandProperty Source -Unique) -join ' + '
  $rows3 += ,@("B-93-$s", (DrawNo $s), $(if ($title[$s]) { $title[$s] } else { '(no point marks - P&ID / layout / sequence only)' }), $hw, $users)
}
Put $w3 @('Sheet','Drawing number','Title','How points were read','Used for equipment') $rows3 @()
$wd = 9,32,70,30,120; for ($j=0; $j -lt $wd.Count; $j++) { $w3.Columns.Item($j+1).ColumnWidth = $wd[$j] }
$w3.Range("E2:E$($rows3.Count+1)").WrapText = $true

# 4 Assumptions
Put $w4 @('SN','Equipment','Qty','IO point','Hardwired points (1)','SP (1)','Basis','Why it is yellow (verify)') $assm @()
$wd = 5,52,6,52,10,7,34,60; for ($j=0; $j -lt $wd.Count; $j++) { $w4.Columns.Item($j+1).ColumnWidth = $wd[$j] }

$rows5 = @(); foreach ($k in ($devAgg.Keys | Sort-Object { -$devAgg[$_].Q })) { $rows5 += ,@($k, [string]$modelOf[$k], [double]$devAgg[$k].Q, (($devAgg[$k].Why.Keys | Sort-Object) -join '; '), (($devAgg[$k].Eq.Keys | Sort-Object) -join '; ')) }
Put $w5 @('Field device (FieldDevices)','Model','Total qty','Why / drawing symbol / previous-project choice','Used on equipment') $rows5 @()
$wd = 46,20,10,70,120; for ($j=0; $j -lt $wd.Count; $j++) { $w5.Columns.Item($j+1).ColumnWidth = $wd[$j] }; $w5.Range("D2:E$($rows5.Count+1)").WrapText = $true
foreach ($ws in @($w1,$w2,$w3,$w4,$w5)) { $ws.Cells.Font.Name = 'Calibri'; $ws.Cells.Font.Size = 10 }
$x.CalculateFull(); $w1.Activate(); $x.ScreenUpdating = $true
if (Test-Path $Out) { [IO.File]::Delete($Out) }
$wb.SaveAs($Out, 51)
"summary rows=" + $sumRows.Count + " ; point rows=" + ($ptRows.Count - $eq.Count) + " ; assumed rows=" + $assm.Count
"totals all: DI={0} AI={1} AO={2} DO={3} SP={4}" -f $w1.Cells.Item($n1+1,15).Value2, $w1.Cells.Item($n1+1,16).Value2, $w1.Cells.Item($n1+1,17).Value2, $w1.Cells.Item($n1+1,18).Value2, $w1.Cells.Item($n1+1,19).Value2
$wb.Close($false); $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
"saved: $Out"
