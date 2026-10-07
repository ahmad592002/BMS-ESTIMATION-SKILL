# Refill Product Finder (C17 part no, E17 qty) with every Siemens BOQ line, then report totals
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$WBPATH = $env:BMS_WB
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($WBPATH); $x = $wb.Application
$b = $wb.Worksheets.Item("BOQ"); $pf = $wb.Worksheets.Item("Product_Finder_SI_B_AUT_V27.1")
$items = @(); $sumJ = 0; $na = @(); $cat = [ordered]@{}
for ($r = 6; $r -le 107; $r++) { $pn = ([string]$b.Cells.Item($r, 4).Text).Trim(); if (-not $pn) { continue }
  $mf = ([string]$b.Cells.Item($r, 7).Text).Trim(); $q = $b.Cells.Item($r, 6).Value2; $j = $b.Cells.Item($r, 10).Value2
  if ($j -is [double]) { $sumJ += $j; $k = ([string]$b.Cells.Item($r, 3).Text).Trim(); $cat[$k] = [double]$cat[$k] + $j } else { $na += "$pn x$q" }
  if ($mf -eq 'Siemens' -and $q -is [double]) { $items += , @($pn, $q) } }
$pf.Range("C17:C480").ClearContents(); $pf.Range("E17:E480").ClearContents()
$n = $items.Count; $cc = New-Object 'object[,]' $n, 1; $ee = New-Object 'object[,]' $n, 1
for ($i = 0; $i -lt $n; $i++) { $cc[$i, 0] = [string]$items[$i][0]; $ee[$i, 0] = [double]$items[$i][1] }
$pf.Range("C17").Resize($n, 1).Value2 = $cc; $pf.Range("E17").Resize($n, 1).Value2 = $ee; $x.CalculateFull(); $wb.Save()
$bad = @(); for ($r = 17; $r -lt 17 + $n; $r++) { if ([string]$pf.Cells.Item($r, 4).Text -match '#N/A|^$') { $bad += [string]$pf.Cells.Item($r, 3).Text } }
"Product Finder: $n Siemens lines, not found: $($bad.Count) $($bad -join ', ') ; K14 = $($pf.Range('K14').Text)"
"BOQ selling (priced lines) = {0:N0} ; unpriced: {1}" -f $sumJ, ($(if ($na) { $na -join ', ' } else { 'none' }))
foreach ($k in $cat.Keys) { "  {0,-18} {1,14:N0}" -f $k, $cat[$k] }

