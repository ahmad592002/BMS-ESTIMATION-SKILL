
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
  if ($l -match '^TEMPLATE (\S+)\s*\|?\s*(.*)$') { $cur = $Matches[1]; $curNote = $Matches[2].Trim(); $TPL[$cur] = New-Object System.Collections.Generic.List[object]; continue }
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
      $TPL[$cur].Add([pscustomobject]@{ L=$lab; DI=$di; AI=$ai; AO=$ao; DO=$do; SP=$s; Dev=$dv[0]; DevQ=$dv[1]; C=$cm; SrcSheet=$sh; SrcHow=[string]$p.Source; SrcGrp=[string]$p.Group; SrcRaw=[string]$p.Point; SrcNote=$(if ($p.Check) { [string]$p.Check } else { [string]$p.Notes }); Basis=$(if ($p.Check) { 'Drawing - check on sheet' } else { 'Drawing' }) })
    }
    continue
  }
  $f = $l.Split('|')
  $di=[int]$f[1]; $ai=[int]$f[2]; $ao=[int]$f[3]; $do=[int]$f[4]; $s=[int]$f[5]
  $cm = $f[8]; if ($s -gt 0) { $cm = ($cm -split ' - ')[0].Trim() }
  $txt = ($curNote + ' ' + $f[8]); $shm = [regex]::Match($txt, 'sheets? (\d{3})'); $bas = if ($txt -match 'assumed') { 'Assumed - confirm' } elseif ($txt -match 'estimator') { 'Estimator block (your entry)' } elseif ($shm.Success) { 'Drawing' } else { 'Typical points - no dedicated schematic' }
  $TPL[$cur].Add([pscustomobject]@{ L=$f[0]; DI=$di; AI=$ai; AO=$ao; DO=$do; SP=$s; Dev=$f[6]; DevQ=$(if ($f[7]) { [int]$f[7] } else { 0 }); C=$cm; SrcSheet=$(if ($shm.Success) { $shm.Groups[1].Value } else { '' }); SrcHow='Template: ' + $curNote; SrcGrp=''; SrcRaw=''; SrcNote=[string]$f[8]; Basis=$bas })
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


# ---- component grouping (estimator: group by component, not by point type) - overrides Build-Lines above
$MainComp = @{ CHILLER='Chiller'; CT='Cooling Tower Fan'; PUMP_VFD='Pump'; PUMP_CS='Pump'; PUMP_SUMP='Pump'; BOOSTER='Booster Pumps'; CIRC_PUMP='Circulating Pump';
  POOL_PUMP='Pump'; FIRE_PUMP='Fire Pumps'; IRRIGATION='Irrigation Pumps'; EF='Fan'; SMOKE_FAN='Fan'; MD='Damper'; LIFT='Lift'; VT_SYSTEM='Vertical Transportation';
  CBS='Central Battery System'; FA='Fire Alarm'; UPS='UPS'; MV='MV Switchgear'; TRANSFORMER='Transformer'; GENERATOR='Generator'; ATS='ATS'; MDB='Breakers & Protection';
  SMDB='Breakers & Protection'; DB='Breakers & Protection'; TAPOFF='Breakers & Protection'; CAPBANK='Capacitor Bank'; SMART_METER='Power Metering'; WATER_HEATER='Water Heater';
  FUEL_TANK='Fuel Tank'; FLDP='Fuel Leak Detection'; WLDP='Water Leak Detection'; LIFT_PIT='Lift Pit Leak Detection'; FM200='FM200 System'; FOAM='Foam System';
  FCS='Floor Control Station'; RLD='Refrigerant Purging Fan'; TANK='Tank Levels'; FILTER='Filter'; VALVE_2POS='Valve'; FCU='Fan Coil Unit'; VAV='VAV Box' }
function Get-Comp([string]$lab, [string]$key, [int]$sp) {
  $u = $lab.ToUpper()
  if ($key -match '^(MDB|SMDB|DB|TAPOFF|SMART_METER|MV|UPS)$') {
    if ($u -match 'OIL |OIL$|WINDING|BUCHHOLS|TAP CHANGER') { return 'Transformer' }
    if ($u -match 'BREAKER|ELR|SPD|PROTECTION|RELAY|OVERCURRENT|FAULT|DISCONNECTOR|SWITCH|VCB|EARTH|CABLE BOX') { return 'Breakers & Protection' }
    if ($u -match 'VOLTAGE|CURRENT|KW|KVA|KVAR|POWER FACTOR|FREQUENCY|WATT|PEAK|DEMAND|BATTERY AMPERAGE|CAPACITY') { return 'Power Metering' }
    return $MainComp[$key]
  }
  if ($u -match 'DAMPER') { return 'Dampers' }
  if ($u -match 'FILTER') { return 'Filters' }
  if ($u -match 'HEATER') { return 'Electric Heater' }
  if ($u -match 'BASIN SWEEPER') { return 'Basin Sweeper' }
  if ($u -match 'SUPPLY FAN') { return 'Supply Fan' }
  if ($u -match 'EXHAUST FAN|EXHAUST HOA') { return 'Exhaust Fan' }
  if ($u -match 'THERMAL WHEEL|RUN AROUND|HEAT PIPE|RECOVERY') { return 'Heat Recovery' }
  if ($u -match 'VIBRATION' -and $MainComp.ContainsKey($key)) { return $MainComp[$key] }
  if ($u -match 'VALVE|COIL|PICV|LOADING PERCENTAGE|CHILLED WATER|CHW|CONTROL BOX|^STATUS$|^START/STOP COMMAND$') { return 'Coils & Valves' }
  if ($u -match 'PUMP') { return $(if ($MainComp.ContainsKey($key) -and $MainComp[$key] -match 'Pump') { $MainComp[$key] } else { 'Pump' }) }
  if ($u -match '\bFAN\b') { return $(if ($key -match '^(CT|EF|SMOKE_FAN|RLD)$') { $MainComp[$key] } else { 'Fan' }) }
  if ($u -match 'TEMPERATURE|HUMIDITY|SENSOR|STATIC PRESSURE|FLOW MEASURING|AIR FLOW$|AQ |CO2|TRANSMITTER|LEVEL|PRESSURE|OCCUPANCY|WINDOW|DOOR CONTACT|DETECTOR|SETPOINT|TVOC|VOLATILE|PM2|OXYGEN|DIOXIDE|MONOXIDE') { return 'Sensors' }
  if ($sp -gt 0 -and $u -match 'SOFTWARE|INTEGRATION') { return 'Software Integration' }
  if ($MainComp.ContainsKey($key)) { return $MainComp[$key] }
  if ($sp -gt 0) { return 'Software Integration' }
  return 'Status & Alarms'
}
function Build-Lines($tpl, [string]$key = '') {
  $items = New-Object System.Collections.Generic.List[object]; foreach ($tp0 in $tpl) { $items.Add($tp0) }
  $cats = New-Object string[] $items.Count
  for ($i=0; $i -lt $items.Count; $i++) { $cats[$i] = Get-Comp ([string]$items[$i].L) $key ([int]$items[$i].SP) }
  # airflow-proving DPS belongs to the fan it proves: take the category of the next fan point
  for ($i=0; $i -lt $items.Count; $i++) {
    if ([string]$items[$i].L -match 'AIR ?FLOW STATUS|VIA DIFFERENTIAL PRESSURE') {
      for ($j=$i+1; $j -lt $items.Count; $j++) { if ($cats[$j] -match 'Fan') { $cats[$i] = $cats[$j]; break } }
      if ($cats[$i] -notmatch 'Fan') { for ($j=$i-1; $j -ge 0; $j--) { if ($cats[$j] -match 'Fan') { $cats[$i] = $cats[$j]; break } } }
      if ($cats[$i] -notmatch 'Fan') { $cats[$i] = $(if ($key -match 'AHU|ERU|MAHU|ECU') { 'Supply Fan' } else { 'Fan' }) }
    }
  }
  $order = New-Object System.Collections.Generic.List[string]
  foreach ($c in $cats) { if (-not $order.Contains($c)) { $order.Add($c) } }
  $rank = @{ DI=0; AI=1; AO=2; DO=3; SP=4 }
  $out = New-Object System.Collections.Generic.List[object]
  $single = ($order.Count -eq 1)
  foreach ($c in $order) {
    $grp = @(); for ($i=0; $i -lt $items.Count; $i++) { if ($cats[$i] -eq $c) { $p = $items[$i]
      $t = if ($p.DI -gt 0) { 'DI' } elseif ($p.AI -gt 0) { 'AI' } elseif ($p.AO -gt 0) { 'AO' } elseif ($p.DO -gt 0) { 'DO' } else { 'SP' }
      $grp += [pscustomobject]@{ P=$p; R=$rank[$t]; I=$i } } }
    $out.Add([pscustomobject]@{ Title=$c })
    foreach ($g in ($grp | Sort-Object R, I)) { $out.Add($g.P) }
  }
  return ,$out
}

# ---- field devices: drawn instruments, models as chosen on previous GTS projects (Ajyad, YALJ, P.Mansour, Qiddiya, RAPEH)
function Get-Device2($p, [string]$key) {
  $u = ([string]$p.L).ToUpper(); $di=[int]$p.DI; $ai=[int]$p.AI; $ao=[int]$p.AO; $do=[int]$p.DO
  $air = ($key -match 'AHU|ERU|MAHU|ECU|EF|SMOKE_FAN|FCU|VAV|RLD')
  if ($ao -gt 0) {
    if ($u -match 'DAMPER') { return @('[Damper Actuator] [Spring Return]', $ao, 'damper actuator - D.A. selection sheet') }
    if ($u -match 'PICV') { return @('[Pressure Independent Control Valve]', $ao, 'PICV - V.A. selection sheet') }
    if ($key -eq 'FCU' -and $u -match 'VALVE') { return @('[FCU Valve]', $ao, 'FCU valve - V.A. selection sheet') }
  }
  if ($do -gt 0 -and $ai -eq 0 -and $ao -eq 0) {
    if ($u -match 'DAMPER') { return @('[Damper Actuator] [Spring Return]', $do, 'damper actuator - D.A. selection sheet') }
    if ($key -match 'CHW_HEADER|VALVE_2POS' -and $u -match 'VALVE') { return @('[Butterfly Valve]', $do, 'motorized butterfly valve (client BOQ) - V.A. selection sheet') }
  }
  if ($di -gt 0) {
    if ($u -match 'FILTER' -and ($air -or $u -match 'D\.?P\.?S|DIFFERENTIAL')) { return @('[Filter] [D.P.S.] [100-1000 Pa] [IP54]', $di, 'DPS drawn - QBM81-10 as previous projects') }
    if ($u -match 'FILTER' -and $key -match 'FILTER|BOOSTER|POOL|IRRIGATION|FIRE') { return @('[Water] [Differential pressure Switch]', $di, 'filter DP switch - as previous projects') }
    if ($u -match 'AIR ?FLOW STATUS|VIA DIFFERENTIAL PRESSURE|VA DIFFERENTIAL PRESSURE|FAN DIFFERENTIAL PRESSURE SWITCH|AIRFLOW STATUS VIA DIFF') { return @('[Fan] [D.P.S.] [50-500 Pa] [IP54]', $di, 'DPS drawn - QBM81-5 as previous projects') }
    if ($u -match 'FLOW SWITCH' -and $key -match 'FCS|FIRE_PUMP') { return @('', 0, 'flow switch by fire fighting contractor') }
    if ($u -match 'FLOW SWITCH') { if ($air) { return @('[Duct] [Air Flow Switch]', $di, 'flow switch drawn') } else { return @('[Water] [Flow Switch]', $di, 'FS drawn') } }
    if ($u -match 'WATER DIFFERENTIAL PRESSURE SWITCH|PRESSURE SWITCH') { return @('[Water] [Differential pressure Switch]', $di, 'PDS drawn - PL-FD113 as previous projects') }
    if ($u -match 'HIGH LEVEL|HI LEVEL|OVERFLOW') { return @('[Water] [Level Float Switch] [Hi level alarm-76mm]', $di, 'LS drawn - AX-LS-FL-1HM as previous projects') }
    if ($u -match 'LOW LEVEL') { return @('[Water] [Level Float Switch] [Low level alarm-76mm]', $di, 'LS drawn - AX-LS-FL-1LM as previous projects') }
    if ($u -match 'LEVEL SWI') { return @('[Water] [Level Float Switch] [Hi & Low level alarm-170mm]', $di, 'LS drawn') }
  }
  if ($ai -gt 0) {
    if ($u -match 'CURRENT TRANSDUCER') { return @('[Current Transmeter] [0-10V]', $ai, 'CT - as previous projects') }
    if ($u -match 'FLOW MEASURING|AFMS|^AIR FLOW$') { return @('[Duct] [Air Flow Sensor]', $ai, 'air flow measuring station drawn') }
    if ($u -match 'STATIC PRESSURE') { return @('[Duct] [Differential Pressure] [0-500 Pa] [IP54]', $ai, 'DPT drawn - QBM3020-5 as RAPEH / SDA') }
    if ($u -match 'AQ SENSOR|AIR QUALITY') { return @('[Duct] [CO2]', $ai, 'AQ drawn - QPM2100 as Qiddiya') }
    if ($u -match 'DIOXIDE|CO2' -and $key -match 'VAV|FCU') { return @(('[Room] [CO2] [0' + [char]0x2026 + '2000 PPM]'), $ai, 'CO2 drawn - QPA2000 as previous projects') }
    if ($u -match 'HUMIDITY' -and $u -match 'DUCT|AIR|COIL|RECOVERY') { return @('[Duct] [Humidity & Temp.]', [math]::Max(1, [int][math]::Ceiling($ai / 2)), 'T/RH drawn - QFM2120 as previous projects') }
    if ($u -match 'TEMPERATURE' -and $u -match 'DUCT|AIR|COIL|DISCHARGE') { return @('[Duct] [Temperature]', $ai, 'TT drawn - QAM2112.040 as previous projects') }
    if ($u -match 'WATER|CHW|CONDENSING|HEADER' -and $u -match 'TEMP') { return @('[Water] [Temperature] [Immersion]', $ai, 'TT/TW drawn - QAE2120.010 as previous projects') }
    if ($key -eq 'HEX' -and $u -match 'TEMPERATURE') { return @('[Water] [Temperature] [Immersion]', $ai, 'TT/TW - QAE2120.010 as previous projects') }
    if ($u -match 'DIFFERENTIAL PRESSURE' -and $u -match 'WATER|HEADER') { return @('[Water] [Differential pressure Transmitter]', $ai, 'DPT drawn - QBE3000-D16 as previous projects') }
    if ($u -match 'LINE PRESSURE|PRESSURE TRANSMITTER|DISCHARGE LINE PRESSURE|^TRANSMITTERS') { return @('[Water] [Pressure Sensor]', $ai, 'PIT/PT drawn - QBE2003-P16 as previous projects') }
    if ($u -match 'TANK LEVEL|LEVEL SENSOR|LEVEL TRANSMITTER') { return @('[Water] [Ultrasonic Level Sensor] [0.25 m to 6 m tank]', $ai, 'level sensor - AX-UL-SEP380-2 as previous projects') }
    if ($u -match 'ROOM TEMPERATURE|ZONE TEMPERATURE' -or ($key -eq 'VAV' -and $u -eq 'TEMPERATURE')) { return @('[Room Themrostat] [Modulating]', 1, 'T/RH room unit - RDF440BN as previous projects') }
  }
  if ([int]$p.SP -gt 0 -and $key -eq 'CO2') { return @(('[Room] [CO2] [0' + [char]0x2026 + '2000 PPM]'), 1, 'CO2 sensor drawn - QPA2000') }
  if ([int]$p.SP -gt 0 -and $key -eq 'CONO2') { return @('[OUTDOOR] [CO] [NO2] [TEM] [0-500 PPM]', 1, 'CO/NO2 sensor drawn') }
  return @('', 0, '')
}
foreach ($k in @($TPL.Keys)) { foreach ($p in $TPL[$k]) { $d = Get-Device2 $p $k
  $p.Dev = $d[0]; $p.DevQ = $d[1]
  if ($p.PSObject.Properties['DevNote']) { $p.DevNote = $d[2] } else { $p | Add-Member -NotePropertyName DevNote -NotePropertyValue $d[2] } } }

# ---- filter names from the drawing (strip label only says FILTER STATUS) - order = left to right on the sheet
$FilterNames = @{
  '006' = @('FRESH AIR PRE-FILTER STATUS','FRESH AIR BAG FILTER STATUS','EXHAUST AIR BAG FILTER STATUS','EXHAUST AIR PANEL FILTER STATUS')
  '007' = @('FRESH AIR PRE-FILTER STATUS','FRESH AIR BAG FILTER STATUS','EXHAUST AIR BAG FILTER STATUS','EXHAUST AIR PANEL FILTER STATUS')
  '008' = @('FRESH AIR PRE-FILTER STATUS','FRESH AIR BAG FILTER STATUS','EXHAUST AIR BAG FILTER STATUS','EXHAUST AIR PANEL FILTER STATUS')
  '012' = @('PANEL FILTER STATUS','BAG FILTER STATUS'); '024' = @('PANEL FILTER STATUS','BAG FILTER STATUS'); '025' = @('PANEL FILTER STATUS','BAG FILTER STATUS')
  '026' = @('PANEL FILTER STATUS','BAG FILTER STATUS'); '049' = @('PANEL FILTER STATUS','BAG FILTER STATUS'); '053' = @('PANEL FILTER STATUS','BAG FILTER STATUS')
  '054' = @('PANEL FILTER STATUS','BAG FILTER STATUS'); '011' = @('FCU FILTER STATUS')
}
foreach ($k in @($TPL.Keys)) { $cnt = @{}
  foreach ($p in $TPL[$k]) {
    if ($p.SrcSheet -and $FilterNames.ContainsKey($p.SrcSheet) -and ([string]$p.L).Trim().ToUpper() -eq 'FILTER STATUS') {
      $n = [int]$cnt[$p.SrcSheet]; $lst = $FilterNames[$p.SrcSheet]; if ($n -lt $lst.Count) { $p.L = $lst[$n] }; $cnt[$p.SrcSheet] = $n + 1 }
    if ([string]$p.L -eq 'PRE-FILTER') { $p.L = 'PRE-FILTER DIFFERENTIAL PRESSURE SWITCH STATUS' }
  } }
