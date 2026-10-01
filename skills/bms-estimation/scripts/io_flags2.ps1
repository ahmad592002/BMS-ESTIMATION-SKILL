# ---- cell-level checks: each returns @(target, reason); target = SP | IO | DEV | DEVQ | NAME
function Get-PointChecks($p, [string]$key) {
  $c = New-Object System.Collections.Generic.List[object]; $u = ([string]$p.L).ToUpper()
  if ([int]$p.SP -gt 0 -and $AssumedSP -contains $key) { $c.Add(@('SP', "SP count $([int]$p.SP) assumed (not in drawings/spec)")) }
  if ($key -eq 'AQ' -and [int]$p.SP -gt 0) { $c.Clear(); $c.Add(@('SP', 'SP 7 = parameters listed on sheet 028 - confirm sensor')) }
  if ($p.Basis -eq 'Drawing - check on sheet') { $c.Add(@('IO', "type taken from wording - check B-93-$($p.SrcSheet)")) }
  if ($key -eq 'PUMP_SUMP' -and $u -match 'PIT HIGH LEVEL') { $c.Add(@('IO', 'pit level per pump assumed (sheet 041: per pit)')) }
  if ($key -eq 'CHW_HEADER' -and $u -match 'HEADER') { $c.Add(@('IO', 'header instrument assumed from P&ID 001')) }
  if ($p.Basis -eq 'Assumed - confirm' -and [int]$p.SP -eq 0) { $c.Add(@('IO', 'point assumed')) }
  if ($key -eq 'VAV' -and $u -match 'DIOXIDE') { $c.Add(@('DEVQ', 'CO2 per VAV per sheet 044 - BOQ: 117 T/RH/CO2 + 35 CO2')) }
  if ($u -match 'STATIC PRESSURE' -and [int]$p.AI -gt 0) { $c.Add(@('DEV', 'QBM3020-5 (RAPEH) or -25 (Ajyad/Qiddiya)?')) }
  if ($p.Dev -match 'Variable Frequency') { $c.Add(@('DEV', 'VFD not in client BMS BOQ - in BMS scope?')) }
  if ($p.Dev -match 'Ball Valve') { $c.Add(@('DEV', 'valve type/size assumed')) }
  if ($p.Dev -match 'Butterfly') { $c.Add(@('DEV', 'butterfly valve size to confirm')) }
  if ([int]$p.DO -gt 0 -and $u -match 'SOLENOID VALVE' -and -not $p.Dev) { $c.Add(@('DEV', 'solenoid valve - no FieldDevices item; by BMS or unit supplier?')) }
  if ([int]$p.DI -gt 0 -and $u -match 'DOOR CONTACT') { $c.Add(@('DEV', 'door contact - no FieldDevices item, price separately')) }
  return ,$c
}
# block-level: what the doubt is about -> QTY cell (A) or NAME cell (B)
$NameNote = @{ HEX='typical points - no HEX schedule'; TANK='typical points - no tank schedule'; STATUS_ALARM='typical status/alarm - no schematic'; STATUS_ALARM_SP10='typical points - no schematic';
  EF='EF points from sheet 045 (waste-room fan)'; SMOKE_FAN='smoke fan points from sheet 004 pattern'; NO_POINTS='no BMS points shown - confirm'; GREASE='typical - no schematic';
  FREEZER='from lab sheet 014 pattern'; GENERATOR='no generator schedule'; DOOR_CONTACT=''; LAB_INTERFACE='group interface assumed' }
$EqQty = @{ 1='qty 5 (your row)'; 14='3 detectors drawn - final count by refrigerant study'; 35='DPBS-01/02 = 2 sets?'; 42='2NOS x2 separate from WWP-07/08?'; 47='4 lift groups, lifts per group not shown';
  63='qty not shown'; 64='qty not shown'; 78='3 damper groups, qty not shown'; 79='indoor units not counted'; 93='lift count not shown'; 102='meters 2 per board assumed';
  103='54 drawn as groups, no tags - may double count'; 111='SF 8NOS paired with SEF'; 113='generator-room dampers qty not shown'; 114='drawing 139 vs BOQ 348'; 115='drawing 691 vs BOQ 449';
  117='drawing 16 vs BOQ 17'; 118='group - qty not shown'; 119='group - qty not shown'; 120='group - qty not shown'; 121='group - qty not shown'; 122='group - qty not shown';
  137='2 untagged - duplicates?'; 140='groups - qty not shown'; 164='group - qty not shown'; 165='groups - qty not shown'; 173='21 from BOQ (drawing 3 groups)';
  177='group - qty not shown'; 178='group - qty not shown'; 191='meters 2 per board assumed'; 51='number of blinds not shown'; 97='number of blinds not shown' }
$EqName = @{ 7='EAP not in legend - pump type assumed'; 15='schematic 007 (CHW run-around ERU) chosen'; 16='schematic 008 (thermal-wheel ERU) chosen'; 17='schematic 025 chosen; AHU-RF-03 drawn twice';
  21='PEF not in legend'; 68='schematic 006 chosen'; 69='schematic 026 chosen'; 70='schematic 026 chosen; auditorium CO2 to add'; 71='schematic 012 chosen (3 MAHU types drawn)';
  72='schematic 005 chosen (005/055)'; 76='DEF not in legend'; 77='PEF not in legend'; 104='schematic 008 chosen'; 105='schematic 053 chosen'; 106='schematic 025 chosen'; 176='RMU points typical'; 172='EV charging points typical' }
