$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$sp\io_lib.ps1" | Out-Null
. "$sp\io_flags.ps1"
. "$sp\io_flags2.ps1"
$YEL = 65535
foreach ($k in @($TPL.Keys)) { foreach ($p in $TPL[$k]) {
  $ex = Get-DeviceExtra $p $k; if ($ex -and $ex[0]) { $p.Dev = $ex[0]; $p.DevQ = $ex[1]; $p.DevNote = $ex[2] }
  $p | Add-Member -Force -NotePropertyName Checks -NotePropertyValue (Get-PointChecks $p $k)
} }
$path = "C:\Users\ahmad\OneDrive\Desktop\GTS\BMS_PROJECTS\07- Al Moosa University\BMS Template 2026 V03.xlsm"
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($path); $x = $wb.Application
$wb.SaveCopyAs((Join-Path (Join-Path (Split-Path $path) "Old Versions") "BMS Template 2026 V03 - BACKUP before cell-level checks.xlsm"))
$ws = $wb.Worksheets.Item("IOSummary"); $el = $wb.Worksheets.Item("EquipmentList")
$last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$v = $ws.Range("A1:B$last").Value2
$hdr = @(); for ($r=1; $r -le $last; $r++) { if ([string]$v[$r,1] -eq 'QTY') { $hdr += $r } }
if ($hdr.Count -ne $eq.Count) { throw "blocks $($hdr.Count) != $($eq.Count)" }
$colOf = @{ SP='G'; DEV='N'; DEVQ='M' }
$x.ScreenUpdating = $false; $x.EnableEvents = $false
$nCell = 0; $nEq = 0
try {
  for ($i=0; $i -lt $eq.Count; $i++) {
    $sn = [int](($eq[$i] -split '\|')[0]); $key = $map[$i]; $h = $hdr[$i]; $q = $h + 2; $p0 = $h + 3
    $lines = Build-Lines $TPL[$key] $key; $m = $lines.Count
    # clear previous marking (keep the grey title rows)
    $ws.Range("A$q`:B$q").Interior.ColorIndex = -4142; $ws.Cells.Item($q, 17).Value2 = [string]''
    $ws.Range("C$p0").Resize($m, 15).Interior.ColorIndex = -4142
    for ($j=0; $j -lt $m; $j++) { if (-not $lines[$j].PSObject.Properties['Title']) { $ws.Range("B$($p0+$j)").Interior.ColorIndex = -4142 } }
    # equipment row
    $en = @()
    if ($EqQty.ContainsKey($sn)) { $ws.Range("A$q").Interior.Color = $YEL; $en += ('QTY: ' + $EqQty[$sn]) }
    $nm = @(); if ($EqName.ContainsKey($sn)) { $nm += $EqName[$sn] }; if ($NameNote.ContainsKey($key) -and $NameNote[$key]) { $nm += $NameNote[$key] }
    if ($nm.Count) { $ws.Range("B$q").Interior.Color = $YEL; $en += ($nm -join '; ') }
    if ($en.Count) { $ws.Cells.Item($q, 17).Value2 = [string]($en -join ' | '); $nEq++ }
    # point cells
    $mn = New-Object 'object[,]' $m, 2; $pc = New-Object 'object[,]' $m, 1; $qc = New-Object 'object[,]' $m, 1
    for ($j=0; $j -lt $m; $j++) {
      $p = $lines[$j]; $mn[$j,0] = [string]''; $mn[$j,1] = [string]''; $pc[$j,0] = [string]''; $qc[$j,0] = [string]''
      if ($p.PSObject.Properties['Title']) { continue }
      if ($p.Dev) { $mn[$j,0] = [string]("=" + [int]$p.DevQ + "*`$A$q"); $mn[$j,1] = [string]$p.Dev }
      if ([int]$p.SP -gt 0) { $pc[$j,0] = [string]$p.C } else { $pc[$j,0] = [string]$p.DevNote }
      $qc[$j,0] = [string](($p.Checks | ForEach-Object { $_[1] }) -join '; ')
    }
    $ws.Range("M$p0").Resize($m, 2).Formula = $mn; $ws.Range("P$p0").Resize($m, 1).Value2 = $pc; $ws.Range("Q$p0").Resize($m, 1).Value2 = $qc
    for ($j=0; $j -lt $m; $j++) { $p = $lines[$j]; if ($p.PSObject.Properties['Title']) { continue }; $r = $p0 + $j
      foreach ($ck in $p.Checks) { $t = $ck[0]
        if ($t -eq 'IO') { foreach ($pair in @(@('C',$p.DI),@('D',$p.AI),@('E',$p.AO),@('F',$p.DO))) { if ([int]$pair[1] -gt 0) { $ws.Range("$($pair[0])$r").Interior.Color = $YEL; $nCell++ } } }
        else { $ws.Range("$($colOf[$t])$r").Interior.Color = $YEL; $nCell++ } }
      if ($p.Checks.Count) { $ws.Range("Q$r").Interior.Color = $YEL } }
  }
  # EquipmentList: only the qty cell or the name cell
  for ($i=0; $i -lt $eq.Count; $i++) { $sn = [int](($eq[$i] -split '\|')[0]); $key = $map[$i]; $r = $i + 2
    $el.Range("B$r`:C$r").Interior.ColorIndex = -4142; $note = @()
    if ($EqQty.ContainsKey($sn)) { $el.Range("B$r").Interior.Color = $YEL; $note += ('QTY: ' + $EqQty[$sn]) }
    $nm = @(); if ($EqName.ContainsKey($sn)) { $nm += $EqName[$sn] }; if ($NameNote.ContainsKey($key) -and $NameNote[$key]) { $nm += $NameNote[$key] }
    if ($nm.Count) { $el.Range("C$r").Interior.Color = $YEL; $note += ($nm -join '; ') }
    $el.Cells.Item($r, 6).Value2 = [string]($note -join ' | ') }
  $x.CalculateFull()
} finally { $x.EnableEvents = $true; $x.ScreenUpdating = $true }
$v = $ws.Range("A1:P$last").Value2; $got = @{ DI=0; AI=0; AO=0; DO=0; SP=0 }; $unres = 0
for ($r=1; $r -le $last; $r++) { if ([string]$v[$r,2] -eq 'TOTAL') { $got.DI += [double]$v[$r,8]; $got.AI += [double]$v[$r,9]; $got.AO += [double]$v[$r,10]; $got.DO += [double]$v[$r,11]; $got.SP += [double]$v[$r,12] }
  $n = [string]$v[$r,14]; if ($n -and $n -ne 'Field Device') { $o = [string]$v[$r,15]; if (-not $o -or $o -match '#N/A') { $unres++ } } }
"IO totals: DI={0} AI={1} AO={2} DO={3} SP={4} ; unresolved devices=$unres" -f $got.DI,$got.AI,$got.AO,$got.DO,$got.SP
"yellow cells on point rows=$nCell ; equipment rows with a yellow qty/name cell=$nEq"
$wb.Save(); "saved=" + $wb.Saved