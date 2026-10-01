$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$T = Import-Csv "$sp\sched_all.tsv" -Delimiter "`t"; $S = Import-Csv "$sp\strip_all.tsv" -Delimiter "`t"
$tabSheets = @($T | Select-Object -ExpandProperty Sheet -Unique)
$all = New-Object System.Collections.Generic.List[object]
function Add-Pt($sheet,$sys,$src,$grp,$pt,$ref,$di,$do,$ai,$ao,$spt,$alm,$ilk,$inf,$notes) {
  $all.Add([pscustomobject]@{Sheet=$sheet;System=$sys;Source=$src;Group=$grp;Point=$pt;Ref=$ref;DI=[int]$di;DO=[int]$do;AI=[int]$ai;AO=[int]$ao;SP=[int]$spt;Alarm=$alm;Interlock=$ilk;Check=$inf;Notes=$notes})
}
foreach ($r in $T) { Add-Pt $r.Sheet $r.System 'Schedule table' ("$($r.Table) $($r.Section)".Trim()) $r.Description $r.Ref ([int]($r.DI -eq 'X')) ([int]($r.DO -eq 'X')) ([int]($r.AI -eq 'X')) ([int]($r.AO -eq 'X')) ([int]($r.COMM -eq 'X')) $r.ALARM $r.HW_INTERLOCK '' $r.Notes }
foreach ($r in $S) {
  if ($tabSheets -contains $r.Sheet -or $r.Sheet -eq '022' -or $r.Sheet -eq '018') { continue }
  $t = $r.Type -replace ' \(inferred\)',''; $m = [int]$r.Mult
  $chk = if ($r.Type -like '*inferred*') { 'type inferred from wording (mark not detected)' } elseif ($r.Point -in '?','&') { 'mark found, label not matched' } else { '' }
  Add-Pt $r.Sheet $r.System 'Dot strip' $r.Strip $r.Point '' ([int]($t -eq 'DI')*$m) ([int]($t -eq 'DO')*$m) ([int]($t -eq 'AI')*$m) ([int]($t -eq 'AO')*$m) 0 '' '' $chk $(if ($m -gt 1) { "x$m on drawing" } else { '' })
}
# ---- sheet 022 and 018: read manually from the drawing (X = per pump / per tower)
$c22 = 'CHILLED WATER EQUIPMENT DETAILED CONTROL SCHEMATIC'
$man = New-Object System.Collections.Generic.List[object]
function M($g,$p,$di,$do,$ai,$ao,$s) { $man.Add(@($g,$p,$di,$do,$ai,$ao,$s)) }
M 'Motorized valve - modulating' 'VALVE MODULATING COMMAND' 0 0 0 1 0
M 'Motorized valve - modulating' 'VALVE POSITION FEEDBACK' 0 0 1 0 0
M 'Motorized valve - 2 position' 'VALVE OPEN/CLOSE COMMAND' 0 1 0 0 0
M 'Motorized valve - 2 position' 'VALVE OPEN/CLOSE STATUSES (x2)' 2 0 0 0 0
M 'Cooling tower (per tower)' 'BASIN SWEEPER STATUS' 1 0 0 0 0
M 'Cooling tower (per tower)' 'BASIN SWEEPER ALARM' 1 0 0 0 0
M 'Cooling tower (per tower)' 'FAN HOA STATUS (AUTO POS.)' 1 0 0 0 0
M 'Cooling tower (per tower)' 'FAN HOA STATUS (MANUAL POS.)' 1 0 0 0 0
M 'Cooling tower (per tower)' 'FAN SPEED FEEDBACK' 0 0 1 0 0
M 'Cooling tower (per tower)' 'FAN SPEED CONTROL' 0 0 0 1 0
M 'Cooling tower (per tower)' 'FAN START/STOP COMMAND' 0 1 0 0 0
M 'Cooling tower (per tower)' 'FAN TRIP ALARM' 1 0 0 0 0
M 'Cooling tower (per tower)' 'FAN STATUS' 1 0 0 0 0
M 'Cooling tower (per tower)' 'VIBRATION DETECTOR STATUS' 1 0 0 0 0
M 'Cooling tower (per tower)' 'VFD SOFTWARE INTEGRATION' 0 0 0 0 1
M 'Cooling tower (typical, shown once)' 'CONDENSING WATER INLET / OUTLET / MAKE-UP TEMPERATURE (4 TT)' 0 0 4 0 0
M 'Cooling tower (typical, shown once)' 'MOTORIZED ISOLATION / MAKE-UP VALVES OPEN/CLOSE STATUS' 7 0 0 0 0
M 'Cooling tower (typical, shown once)' 'MOTORIZED ISOLATION / MAKE-UP VALVES OPEN/CLOSE COMMAND' 0 5 0 0 0
foreach ($g in 'Secondary CHW pump (per pump)','Primary CHW pump (per pump)','Condensing water pump (per pump)') {
  M $g 'HOA STATUS (AUTO POS.)' 1 0 0 0 0; M $g 'HOA STATUS (MANUAL POS.)' 1 0 0 0 0
  M $g 'PUMP SPEED FEEDBACK' 0 0 1 0 0; M $g 'PUMP SPEED CONTROL' 0 0 0 1 0
  M $g 'PUMP START/STOP COMMAND' 0 1 0 0 0; M $g 'PUMP TRIP ALARM' 1 0 0 0 0
  M $g 'PUMP STATUS' 1 0 0 0 0; M $g 'PUMP RUN STATUS' 1 0 0 0 0; M $g 'VFD SOFTWARE INTEGRATION' 0 0 0 0 1
}
$g = 'Cooling tower make-up pump (per pump)'
M $g 'HOA STATUS (AUTO POS.)' 1 0 0 0 0; M $g 'HOA STATUS (MANUAL POS.)' 1 0 0 0 0; M $g 'PUMP START/STOP COMMAND' 0 1 0 0 0
M $g 'PUMP TRIP ALARM' 1 0 0 0 0; M $g 'PUMP STATUS' 1 0 0 0 0; M $g 'PUMP RUN STATUS' 1 0 0 0 0
foreach ($p in 'FAN START / STOP STATUS','FAN NORMAL / FAULT STATUS','FAN H/O/A STATUS') { M 'Refrigerant purging fan' $p 1 0 0 0 0 }
foreach ($m in $man) { Add-Pt '022' $c22 'Read from drawing' $m[0] $m[1] '' $m[2] $m[3] $m[4] $m[5] $m[6] '' '' '' '' }
foreach ($p in 'FAN START / STOP STATUS','FAN NORMAL / FAULT STATUS','FAN H/O/A STATUS') { Add-Pt '018' 'REFRIGERANT LEAK DETECTION & PURGING SYSTEM CONTROLS SCHEMATIC' 'Read from drawing' 'Refrigerant purging fan' $p '' 1 0 0 0 0 '' '' '' 'Purging fans 2 nos; leak panel via BMS software interface' }
$all | Export-Csv "$sp\points_all.csv" -NoTypeInformation -Encoding UTF8
"rows: " + $all.Count
$all | Group-Object Sheet,Group | ForEach-Object { $g=$_.Group; [pscustomobject]@{Key=$_.Name;N=$g.Count;DI=($g|Measure-Object DI -Sum).Sum;DO=($g|Measure-Object DO -Sum).Sum;AI=($g|Measure-Object AI -Sum).Sum;AO=($g|Measure-Object AO -Sum).Sum;SP=($g|Measure-Object SP -Sum).Sum;Chk=@($g|Where-Object {$_.Check}).Count} } | Where-Object { $_.Key -like '022*' -or $_.Key -like '018*' } | Format-Table -AutoSize | Out-String -Width 200
