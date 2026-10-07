# Per-panel point load (per-unit IO from IOSummary TOTAL rows x assigned qty), spare from Options
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$res = & "$sp\ddc_assign.ps1"; $asg = $res | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($env:BMS_WB)
$spare = [double]$wb.Worksheets.Item("Options").Range("B1").Value2
$ws = $wb.Worksheets.Item("IOSummary"); $last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$LV = $ws.Range("A1:G$last").Value2
$per = @{}; $names = @{}; $sn = 0
for ($r = 1; $r -le $last; $r++) {
  if ([string]$LV[$r, 1] -eq 'QTY') { $sn++; $names[$sn] = [string]$LV[($r + 2), 2] }
  if ([string]$LV[$r, 2] -eq 'TOTAL') { $t = @(); for ($c = 3; $c -le 7; $c++) { $t += [double]$LV[$r, $c] }; $per[$sn] = $t }
}
$panels = [ordered]@{}
$pinfo = @{}; Get-Content "$sp\panels.txt" -Encoding UTF8 | ForEach-Object { $a = $_ -split '\|'; $pinfo[$a[0]] = $a }
foreach ($s in ($asg.Keys | Sort-Object)) { foreach ($p in $asg[$s].Keys) {
  if (-not $panels.Contains($p)) { $panels[$p] = [pscustomobject]@{ Panel = $p; IO = @(0.0, 0.0, 0.0, 0.0, 0.0); Eq = [ordered]@{} } }
  $q = $asg[$s][$p]; for ($k = 0; $k -lt 5; $k++) { $panels[$p].IO[$k] += $q * $per[$s][$k] }; $panels[$p].Eq["$s"] = $q } }
$out = foreach ($p in $panels.Values) { $phys = $p.IO[0] + $p.IO[1] + $p.IO[2] + $p.IO[3]; $tot = $phys + $p.IO[4]
  [pscustomobject]@{ Panel = $p.Panel; Bldg = $pinfo[$p.Panel][1]; Type = $pinfo[$p.Panel][3]; DI = $p.IO[0]; AI = $p.IO[1]; AO = $p.IO[2]; DO = $p.IO[3]; SP = $p.IO[4]; Phys = $phys; Total = $tot; WithSpare = [math]::Round($tot * (1 + $spare), 0); EqN = $p.Eq.Count; Eq = $p.Eq } }
$out
