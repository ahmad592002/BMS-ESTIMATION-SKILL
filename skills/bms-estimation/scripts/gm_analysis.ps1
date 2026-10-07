# READ-ONLY gross-margin analysis. No writes, no Save, no Calculate.
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$path = $env:BMS_WB
$wb = $null
for ($i = 0; $i -lt 30 -and -not $wb; $i++) { try { $wb = [Runtime.InteropServices.Marshal]::BindToMoniker($path); $null = $wb.Name } catch { $wb = $null; Start-Sleep -Seconds 2 } }
if (-not $wb) { throw "workbook not reachable" }
$b = $wb.Worksheets.Item("BOQ"); $pf = $wb.Worksheets.Item("Product_Finder_SI_B_AUT_V27.1"); $bd = $wb.Worksheets.Item("Breakdown")
$BQ = $b.Range("A1:N115").Value2; $BT = $b.Range("A1:N115")
$PF = $pf.Range("A1:L480").Value2
$BDV = $bd.Range("A1:J95").Value2
function N($v) { if ($v -is [double]) { $v } else { 0.0 } }
$list = @{}; for ($r = 17; $r -le 480; $r++) { $pn = [string]$PF[$r, 3]; if ($pn) { $list[$pn] = (N $PF[$r, 5]) * (N $PF[$r, 6]) } }
$grp = @{ Siemens = @(0.0, 0.0, 0.0); 'Al Fanar' = @(0.0, 0.0, 0.0); Local = @(0.0, 0.0, 0.0) }; $rows = @(); $na = @()
for ($r = 6; $r -le 107; $r++) { $pn = ([string]$BQ[$r, 4]).Trim(); if (-not $pn) { continue }
  $mf = ([string]$BQ[$r, 7]).Trim(); $cat = ([string]$BQ[$r, 3]).Trim(); $q = N $BQ[$r, 6]; $J = $BQ[$r, 10]; $M = $BQ[$r, 13]
  if (-not ($J -is [double])) { $na += $pn; continue }
  $k = if ($mf -eq 'Siemens') { 'Siemens' } elseif ($mf -eq 'Al Fanar') { 'Al Fanar' } else { 'Local' }
  $grp[$k][0] += $J; $grp[$k][1] += (N $M)
  if ($k -eq 'Siemens') { $L = N $list[$pn]; $grp[$k][2] += $L; $rows += [pscustomobject]@{ Pn = $pn; Cat = $cat; Q = $q; Sell = $J; Cost = (N $M); List = $L } } }
"unpriced BOQ lines: " + $(if ($na) { $na -join ', ' } else { 'none' })
"BOQ total rows: " + ((108..118 | ForEach-Object { "r$_ " + ([string]$BQ[$_, 9]).Trim() + '=' + [math]::Round((N $BQ[$_, 10]), 0) }) -join ' ; ')
"Margin (BOQ markup) = " + (N $BQ[2, 13])
foreach ($k in 'Siemens', 'Al Fanar', 'Local') { "{0,-9} selling={1,12:N0}  BOQ cost={2,12:N0}  Siemens list={3,12:N0}" -f $k, $grp[$k][0], $grp[$k][1], $grp[$k][2] }
"PF K14 (list x qty) = " + [math]::Round((N $PF[14, 11]), 0)
"--- (PF qty check skipped)"; $bqx = @{}; for ($r = 6; $r -le 107; $r++) { $pn = ([string]$BQ[$r, 4]).Trim(); if ($pn) { $bqx[$pn] = 1 } }
"--- Siemens: Pricelist cost vs Siemens list, by category"
$rows | Group-Object Cat | ForEach-Object { $c = ($_.Group | Measure-Object Cost -Sum).Sum; $l = ($_.Group | Measure-Object List -Sum).Sum; $s = ($_.Group | Measure-Object Sell -Sum).Sum; "  {0,-18} sell={1,12:N0} cost={2,12:N0} list={3,12:N0}  cost/list={4:P1}  sell/list={5:P1}" -f $_.Name, $s, $c, $l, ($c / $l), ($s / $l) }
"--- top Siemens lines by list value"
$rows | Sort-Object List -Descending | Select-Object -First 12 | ForEach-Object { "  {0,-16} qty {1,5}  sell={2,11:N0} cost={3,11:N0} list={4,11:N0} cost/list={5:P1}" -f $_.Pn, $_.Q, $_.Sell, $_.Cost, $_.List, $(if ($_.List) { $_.Cost / $_.List } else { 0 }) }
"--- Breakdown"
foreach ($r in 9, 11, 23, 24, 25, 33, 42, 73, 89, 91, 93) { "  r{0} {1,-34} D={2,14} E={3} I={4,14}" -f $r, ([string]$BDV[$r, 2]), [math]::Round((N $BDV[$r, 4]), 0), $BDV[$r, 5], [math]::Round((N $BDV[$r, 9]), 0) }
"--- expense / resource lines"
foreach ($r in (46..71) + (77..87)) { $v = N $BDV[$r, 9]; if ($v -ne 0) { "  {0,-28} {1} x {2} x {3} = {4,12:N0}" -f $BDV[$r, 3], $BDV[$r, 6], $BDV[$r, 7], $BDV[$r, 8], $v } }


