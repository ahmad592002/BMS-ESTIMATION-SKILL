param([switch]$DryRun)
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---------------- templates
$pts = @(Import-Csv "$sp\points_all.csv")
$defProto = @{ VT_SYSTEM='BACNET/IP'; LIFT='BACNET/IP'; FA='BACNET/IP'; FCU='BACNET/IP'; VAV='BACNET/IP'; CBS='MODBUS'; UPS='MODBUS'; SMART_METER='MODBUS'; MV='MODBUS' }
function Get-Device([string]$lab, [int]$di, [int]$ai, [string]$tkey) {
  $u = $lab.ToUpper()
  if ($di -gt 0 -and $u -match 'FILTER') { return @('[Filter] [D.P.S.] [100-1000 Pa] [IP54]', 1) }
  if ($di -gt 0 -and $u -match 'AIR ?FLOW STATUS|VIA DIFFERENTIAL PRESSURE|DIFFERENTIAL PRESSURE SWITCH') { return @('[Fan] [D.P.S.] [50-500 Pa] [IP54]', 1) }
  if ($di -gt 0 -and $u -match 'FLOW SWITCH' -and $tkey -match 'AHU|ERU|MAHU|ECU') { return @('[Duct] [Air Flow Switch]', 1) }
  if ($ai -gt 0) {
    if ($u -match 'DUCT|SUPPLY AIR|RETURN AIR|FRESH AIR|EXHAUST AIR|OFF-COIL|ON-COIL|COIL') {
      if ($u -match 'HUMIDITY') { return @('[Duct] [Humidity & Temp.]', [math]::Max(1, [int][math]::Ceiling($ai / 2))) }
      if ($u -match 'TEMPERATURE') { return @('[Duct] [Temperature]', $ai) }
      if ($u -match 'STATIC PRESSURE') { return @('[Duct] [Differential Pressure] [0-500 Pa] [IP54]', $ai) }
      if ($u -match 'AQ |AIR QUALITY|CO2') { return @('[Duct] [CO2]', $ai) }
      if ($u -match 'FLOW MEASURING') { return @('[Duct] [Air Flow Sensor]', $ai) }
    }
    if ($u -match 'FLOW MEASURING') { return @('[Duct] [Air Flow Sensor]', $ai) }
    if ($u -match 'WATER' -and $u -match 'TEMPERATURE|TMEPERATURE|TEMP') { return @('[Water] [Temperature] [Immersion]', $ai) }
    if ($u -match 'DIFFERENTIAL PRESSURE' -and $u -match 'WATER') { return @('[Water] [Differential pressure Transmitter]', $ai) }
    if ($u -match 'LINE PRESSURE|PRESSURE TRANSMITTER') { return @('[Water] [Pressure Sensor]', $ai) }
    if ($u -match 'ROOM TEMPERATURE|ZONE TEMPERATURE' -or ($tkey -eq 'VAV' -and $u -eq 'TEMPERATURE')) { return @('[Room Themrostat] [Modulating]', 1) }
  }
  return @('', 0)
}
$TPL = @{}; $cur = $null
foreach ($l in [IO.File]::ReadAllLines("$sp\io_templates.txt")) {
  if (-not $l.Trim() -or $l.StartsWith('#')) { continue }
  if ($l -match '^TEMPLATE (\S+)') { $cur = $Matches[1]; $TPL[$cur] = New-Object System.Collections.Generic.List[object]; continue }
  if ($l -match '^@SHEET:(\d{3})(?::(.+))?$') {
    $sh = $Matches[1]; $pre = $Matches[2]
    foreach ($p in ($pts | Where-Object { $_.Sheet -eq $sh -and (-not $pre -or $_.Group.StartsWith($pre)) })) {
      $lab = ($p.Point -replace '\s+', ' ').Trim()
      if ($lab -match '^[?&]$' -or $lab -match '^\d+\.( \d+\.)+$') { continue }
      $di=[int]$p.DI; $ai=[int]$p.AI; $ao=[int]$p.AO; $do=[int]$p.DO; $s=[int]$p.SP
      if (($di+$ai+$ao+$do+$s) -eq 0) { continue }
      if (-not $lab) {
        $gname = (($p.Group -replace '^T\d+ ?','') -replace 'S$','').Trim(); if (-not $gname) { $gname = 'Point' }
        $lab = $gname + $(if ($di) { ' status' } elseif ($ai) { ' analog input' } elseif ($ao) { ' command' } elseif ($do) { ' command' } else { ' software point' }) + ' (label not on drawing)'
      }
      if ($lab.Length -lt 6 -and $p.Group -match 'PUMPS|TANKS|FILTERS|DOSING') { $lab = (($p.Group -replace '^T\d+ ','') + ' ' + $lab).Trim() }
      $dv = Get-Device $lab $di $ai $cur
      $proto = if ($s -gt 0) { if ($p.Notes -match 'MODBUS') { 'MODBUS' } elseif ($p.Notes -match 'BACNET') { 'BACNET/IP' } elseif ($defProto.ContainsKey($cur)) { $defProto[$cur] } else { 'MODBUS' } } else { '' }
      $cm = if ($s -gt 0) { $proto } elseif ($p.Check) { "sheet $sh - type inferred, confirm" } else { '' }
      $TPL[$cur].Add([pscustomobject]@{ L=$lab; DI=$di; AI=$ai; AO=$ao; DO=$do; SP=$s; Dev=$dv[0]; DevQ=$dv[1]; C=$cm })
    }
    continue
  }
  $f = $l.Split('|')
  $di=[int]$f[1]; $ai=[int]$f[2]; $ao=[int]$f[3]; $do=[int]$f[4]; $s=[int]$f[5]
  $cm = $f[8]; if ($s -gt 0) { $cm = ($cm -split ' - ')[0].Trim() }
  $TPL[$cur].Add([pscustomobject]@{ L=$f[0]; DI=$di; AI=$ai; AO=$ao; DO=$do; SP=$s; Dev=$f[6]; DevQ=$(if ($f[7]) { [int]$f[7] } else { 0 }); C=$cm })
}
$map = @([IO.File]::ReadAllLines("$sp\io_map.txt") | Where-Object { $_.Trim() })
$eq  = @([IO.File]::ReadAllLines("$sp\eqlist_now.txt"))
if ($map.Count -ne $eq.Count) { throw "map rows $($map.Count) != equipment rows $($eq.Count)" }
foreach ($k in $map) { if (-not $TPL.ContainsKey($k)) { throw "unknown template $k" } ; if ($TPL[$k].Count -eq 0) { throw "empty template $k" } }

function Build-Lines($tpl) {
  $groups = [ordered]@{ 'Digital Inputs'=@(); 'Analog Inputs'=@(); 'Analog Outputs'=@(); 'Digital Outputs'=@(); 'Software Points'=@() }
  foreach ($p in $tpl) {
    $g = if ($p.DI -gt 0) { 'Digital Inputs' } elseif ($p.AI -gt 0) { 'Analog Inputs' } elseif ($p.AO -gt 0) { 'Analog Outputs' } elseif ($p.DO -gt 0) { 'Digital Outputs' } elseif ($p.SP -gt 0) { 'Software Points' } else { $null }
    if ($g) { $groups[$g] += $p }
  }
  $out = New-Object System.Collections.Generic.List[object]
  foreach ($g in $groups.Keys) { if ($groups[$g].Count) { $out.Add([pscustomobject]@{ Title=$g }); foreach ($p in $groups[$g]) { $out.Add($p) } } }
  if ($out.Count -eq 0) { foreach ($p in $tpl) { $out.Add($p) } }
  return ,$out
}

# expected totals
$exp = @{ DI=0; AI=0; AO=0; DO=0; SP=0 }
for ($i=0; $i -lt $eq.Count; $i++) { $q = [double](($eq[$i] -split '\|')[1]); foreach ($p in $TPL[$map[$i]]) { foreach ($k in 'DI','AI','AO','DO','SP') { $exp[$k] += $q * $p.$k } } }
"expected all-systems: DI={0} AI={1} AO={2} DO={3} SP={4}" -f $exp.DI,$exp.AI,$exp.AO,$exp.DO,$exp.SP
if ($DryRun) { for ($i=0; $i -lt $eq.Count; $i++) { $e = $eq[$i] -split '\|'; $tp = $TPL[$map[$i]]; "{0,3} {1,-70} -> {2,-18} DI{3} AI{4} AO{5} DO{6} SP{7}" -f $e[0], $e[2].Substring(0,[Math]::Min(70,$e[2].Length)), $map[$i], ($tp|Measure-Object DI -Sum).Sum, ($tp|Measure-Object AI -Sum).Sum, ($tp|Measure-Object AO -Sum).Sum, ($tp|Measure-Object DO -Sum).Sum, ($tp|Measure-Object SP -Sum).Sum }; return }

# ---------------- workbook
$x = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $x.Workbooks.Item((Split-Path -Leaf $env:BMS_WB))
$wb.SaveCopyAs((Join-Path (Split-Path $wb.FullName) "BMS Template 2026 V03 - BACKUP before IOSummary run2.xlsm"))
$ws = $wb.Worksheets.Item("IOSummary"); $el = $wb.Worksheets.Item("EquipmentList"); $iot = $wb.Worksheets.Item("IOTemplate")
$prevEv = $x.EnableEvents; $prevCalc = $x.Calculation
$x.EnableEvents = $false; $x.ScreenUpdating = $false; $x.Calculation = -4135
try {
  $last = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  if ($last -ge 1) { $null = $ws.Range("1:$last").EntireRow.Delete() }
  $macro = "'" + $wb.Name + "'!IOSummaryModule.AddEquipmentFromEquipmentList"
  for ($i = $eq.Count - 1; $i -ge 0; $i--) {
    $elRow = $i + 2
    # same steps as IOSummaryModule.AddEquipmentFromEquipmentList, done directly (no VBA dialog risk)
    $iot.Range("A4").Value2 = $el.Range("B$elRow").Value2
    $iot.Range("B2").Value2 = $el.Range("E$elRow").Value2
    $iot.Range("B4").Value2 = $el.Range("C$elRow").Value2
    $null = $iot.Range("1:30").Copy()
    $null = $ws.Range("A1").EntireRow.Insert(-4121)
    $x.CutCopyMode = $false
    $null = $iot.Range("B2").ClearContents(); $null = $iot.Range("A4:G26").ClearContents(); $null = $iot.Range("M4:N26").ClearContents()
    $lines = Build-Lines $TPL[$map[$i]]
    $n = $lines.Count + 1          # + one blank row before TOTAL
    if ($n -gt 22) { $k = $n - 22; $null = $ws.Rows.Item(6).Copy(); $null = $ws.Range("7:$(6+$k)").Insert(-4121); $x.CutCopyMode = $false }
    elseif ($n -lt 22) { $null = $ws.Range("$(5+$n):26").EntireRow.Delete() }
    $m = $lines.Count
    $bg = New-Object 'object[,]' $m, 6; $mn = New-Object 'object[,]' $m, 2; $pc = New-Object 'object[,]' $m, 1
    for ($j=0; $j -lt $m; $j++) {
      $p = $lines[$j]
      if ($p.PSObject.Properties['Title']) { $bg[$j,0] = $p.Title; for ($c=1;$c -le 5;$c++) { $bg[$j,$c] = '' }; $mn[$j,0]=''; $mn[$j,1]=''; $pc[$j,0]=''; continue }
      $bg[$j,0] = $p.L
      $bg[$j,1] = $(if ($p.DI) { [double]$p.DI } else { '' }); $bg[$j,2] = $(if ($p.AI) { [double]$p.AI } else { '' })
      $bg[$j,3] = $(if ($p.AO) { [double]$p.AO } else { '' }); $bg[$j,4] = $(if ($p.DO) { [double]$p.DO } else { '' })
      $bg[$j,5] = $(if ($p.SP) { [double]$p.SP } else { '' })
      if ($p.Dev) { $mn[$j,0] = [double]$p.DevQ; $mn[$j,1] = $p.Dev } else { $mn[$j,0] = ''; $mn[$j,1] = '' }
      $pc[$j,0] = [string]$p.C
    }
    $ws.Range("B5").Resize($m, 6).Value2 = $bg
    $ws.Range("M5").Resize($m, 2).Value2 = $mn
    $ws.Range("P5").Resize($m, 1).Value2 = $pc
    for ($j=0; $j -lt $m; $j++) { $cell = $ws.Range("B$(5+$j)"); if ($lines[$j].PSObject.Properties['Title']) { $cell.Interior.ColorIndex = 15; $cell.Font.Bold = $true } else { $cell.Interior.ColorIndex = -4142 } }
    if (($i % 25) -eq 0) { Write-Host ("  built down to EquipmentList row {0}" -f $elRow) }
  }
  $x.Calculation = -4105; $x.CalculateFullRebuild()
} finally { $x.Calculation = -4105; $x.EnableEvents = $prevEv; $x.ScreenUpdating = $true }

# ---------------- verify
$last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$v = $ws.Range("A1:P$last").Value2
$blocks = 0; $got = @{ DI=0; AI=0; AO=0; DO=0; SP=0 }; $badH = 0; $spUntagged = 0
for ($r=1; $r -le $last; $r++) {
  if ([string]$v[$r,1] -eq 'QTY') { $blocks++ }
  if ([string]$v[$r,2] -eq 'TOTAL') { $got.DI += [double]$v[$r,8]; $got.AI += [double]$v[$r,9]; $got.AO += [double]$v[$r,10]; $got.DO += [double]$v[$r,11]; $got.SP += [double]$v[$r,12] }
  if ([string]$v[$r,7] -ne '' -and [string]$v[$r,2] -ne 'TOTAL' -and [string]$v[$r,7] -ne 'SP') { if ([string]$v[$r,16] -notin 'MODBUS','BACNET/IP','BACNET/MSTP','MBUS','KNX') { $spUntagged++ } }
}
for ($r=1; $r -le $last; $r++) { if ([string]$v[$r,3] -ne '' -and [string]$v[$r,3] -match '^\d' -and [string]$v[$r,2] -ne 'TOTAL') { $f = [string]$ws.Cells.Item($r,8).Formula; if (-not $f.StartsWith('=')) { $badH++ } } }
"blocks=$blocks ; got all-systems: DI={0} AI={1} AO={2} DO={3} SP={4} ; rows with DI but no H formula={5} ; SP rows untagged={6}" -f $got.DI,$got.AI,$got.AO,$got.DO,$got.SP,$badH,$spUntagged
$wb.Save(); "saved=" + $wb.Saved
