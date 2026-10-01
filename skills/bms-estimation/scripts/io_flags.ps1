# ---- devices the drawings show that the first pass missed, plus yellow "check" reasons
function Get-DeviceExtra($p, [string]$key) {
  $u = ([string]$p.L).ToUpper(); $ai=[int]$p.AI; $ao=[int]$p.AO; $do=[int]$p.DO
  if ($ao -gt 0 -and $u -match 'VFD|SPEED CONTROL|^BOOSTER PUMPS CONTROL') { return @('[Variable Frequency Drive]', $ao, 'VFD drawn - VFD selection sheet', 'VFD shown on drawing but client BMS BOQ has no VFD line - confirm if VFD is in BMS scope') }
  if ($do -gt 0 -and $u -match 'RUN AROUND COILS? CONTROL VALVE|HEAT PIPE') { return @('[Ball Valve] [2 Ports]', $do, '2-position valve drawn - V.A. selection sheet', 'valve type/size assumed (2-way ball) - confirm from mechanical schedule') }
  if ($do -gt 0 -and $u -match 'SOLENOID VALVE') { return @('', 0, '', 'solenoid valve drawn - no FieldDevices item; confirm if by BMS or by unit supplier') }
  if ($ai -gt 0 -and $u -match 'TANK TEMPERATURE') { return @('[Water] [Temperature] [Immersion]', $ai, 'TT drawn - QAE2120.010 as previous projects', '') }
  if ([int]$p.DI -gt 0 -and $u -match 'DOOR CONTACT') { return @('', 0, '', 'door contact drawn / in client BOQ - no FieldDevices item, price separately') }
  return $null
}
$AssumedSP = 'STATUS_ALARM_SP10','GENERATOR','LCS','ELV','BLINDS','VRF_ODU','LAB_INTERFACE','WATER_METER','AQ'
function Get-PointFlag($p, [string]$key) {
  $r = @(); $u = ([string]$p.L).ToUpper()
  if ([int]$p.SP -gt 0 -and $AssumedSP -contains $key) { $r += "SP count $([int]$p.SP) assumed - not given in drawings/spec" }
  if ($p.Basis -eq 'Assumed - confirm' -and -not ($r -match 'SP count')) { $r += 'assumed - not from drawings' }
  if ($p.Basis -eq 'Typical points - no dedicated schematic') { $r += 'typical points - no schematic for this equipment' }
  if ($p.Basis -eq 'Drawing - check on sheet') { $r += "point type inferred from wording - check sheet B-93-$($p.SrcSheet)" }
  if ($key -eq 'PUMP_SUMP' -and $u -match 'PIT HIGH LEVEL') { $r += 'pit level split per pump assumed (sheet 041 shows levels per pit)' }
  if ($key -eq 'SMOKE_FAN') { $r += 'smoke fan points from sheet 004 fan pattern - no smoke-fan schedule' }
  if ($key -eq 'EF') { $r += 'exhaust fan points from sheet 045 (waste room fan) - no general EF schedule' }
  if ($key -eq 'AQ' -and [int]$p.SP -gt 0) { $r += 'SP = 7 parameters listed on sheet 028 - confirm sensor type' }
  if ($key -eq 'VAV' -and $u -match 'DIOXIDE') { $r += 'CO2 per VAV per sheet 044 - client BOQ lists 117 T/RH/CO2 + 35 CO2 for MAB' }
  if ($u -match 'STATIC PRESSURE' -and [int]$p.AI -gt 0) { $r += 'model QBM3020-5 (RAPEH) vs QBM3020-25 (Ajyad/Qiddiya) - confirm range' }
  if ($key -eq 'CHW_HEADER' -and $u -match 'HEADER') { $r += 'header instruments assumed from sheet 001 P&ID - confirm' }
  if ($key -eq 'HEX') { $r += 'HEX temperatures typical - no HEX schedule' }
  if ($key -eq 'TANK') { $r += 'tank points typical (client BOQ has level switches + transmitters) - no tank schedule' }
  return ($r | Select-Object -Unique) -join '; '
}
# equipment-level: quantity or schematic choice not 100% from the data (by EquipmentList SN)
$EqFlag = @{
  1='qty 5 - your row; chiller points from your block'; 7='EAP-01..03 not in legend - pump type assumed'
  12='header instrument list assumed from sheet 001'; 14='3 detectors drawn - final count by refrigerant study'
  15='schematic choice: CHW run-around ERU (007) assumed for ATC ERUs'; 16='schematic choice: thermal-wheel ERU (008) assumed for FAHU'
  17='schematic choice: CHW recirculating AHU (025) assumed; AHU-RF-03 drawn twice'; 19='smoke fan points pattern (no schedule)'; 20='smoke fan points pattern (no schedule)'
  21='PEF not in legend - assumed exhaust fan'; 25='SP count assumed'; 35='DPBS-01/02 counted as 2 sets - confirm'; 38='SP count assumed'
  42='elevator sump pumps 2NOS x2 - separate from WWP-07/08 assumed'; 47='4 lift groups drawn without count - 1 per group assumed'
  51='number of blinds not shown - integration SP assumed'; 63='number of feeder pillars not shown'; 64='number of MV panels not shown'
  67='meter qty from riser (per board NOS)'; 68='schematic choice: VRF run-around ERU (006) for EEC ERUs'; 69='schematic choice: VRF recirculating AHU (026)'
  70='schematic choice: VRF AHU (026) for auditorium AHU; CO2 sensor to add'; 71='schematic choice: kitchen MAHU type 1 (012)'; 72='schematic choice: ECU type 1 (005)'
  76='DEF not in legend'; 77='PEF not in legend'; 78='3 damper groups, qty not shown'; 79='SP per ODU assumed; indoor units not counted'
  93='lift count not shown - 1 assumed'; 97='number of blinds not shown - integration SP assumed'; 102='meter qty not shown - 2 per board assumed'
  103='54 AHUs drawn only as groups (no tags); DDCP/SCP groups may duplicate'; 104='schematic choice: thermal-wheel ERU (008) for MAB FAHU'
  105='schematic choice: kitchen MAHU type 2 (053) for KFAHU'; 106='schematic choice: CHW recirculating AHU (025)'
  111='SF 8NOS paired with SEF - type assumed'; 113='generator-room damper group qty not shown'
  114='MAB FCU 139 from drawing - client BOQ 348'; 115='MAB VAV 691 from drawing - client BOQ 449'
  116='SP count assumed'; 117='16 drawn - client BOQ 17'; 118='group - qty and SP assumed'; 119='group - qty and SP assumed'; 120='group - qty and SP assumed'
  121='group - qty not shown'; 122='group - qty not shown'; 128='SP count assumed'; 137='2 untagged heaters on sheet - confirm not duplicates'; 138='SP count assumed'
  140='inline pump groups - qty not shown'; 141='SP count assumed'; 151='typical points'; 162='SP count assumed; no generator schedule'; 164='freezer group - qty not shown'
  165='waste room equipment groups - qty not shown'; 172='EV charging - points assumed'; 173='21 from client BOQ (drawing shows 3 groups)'
  176='RMU points typical'; 177='ACB panels group - qty not shown'; 178='transformer group - qty not shown'; 191='meter qty not shown - 2 per board assumed'
  193='SP count assumed'; 194='SP count assumed'
}
