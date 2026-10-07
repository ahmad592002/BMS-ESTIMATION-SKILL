if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$WBPATH = $env:BMS_WB
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($WBPATH); $x = $wb.Application; $ws = $wb.Worksheets.Item("IOSummary")
$last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$v = $ws.Range("A1:B$last").Value2
$hdr = @(); for ($r=1; $r -le $last; $r++) { if ([string]$v[$r,1] -eq 'QTY') { $hdr += $r } }
$x.ScreenUpdating = $false; $x.EnableEvents = $false; try { $x.Calculation = -4135 } catch { }
$fixed = 0
try {
  foreach ($h in $hdr) {
    $q = $h + 2; $t = $q; while ([string]$v[$t,2] -ne 'TOTAL') { $t++ }
    $n = $t - $q
    $arr = New-Object 'object[,]' $n, 5
    for ($i=0; $i -lt $n; $i++) { $r = $q + $i; $k = 0
      foreach ($c in 'C','D','E','F','G') { $arr[$i,$k] = [string]('=IF(' + $c + $r + '="","",IF(' + $c + $r + '*$A' + $q + '>0,' + $c + $r + '*$A' + $q + ',""))'); $k++ } }
    $ws.Range("H$q").Resize($n, 5).Formula = $arr
    $fixed += $n
  }
} finally { try { $x.Calculation = -4105 } catch { }; $x.EnableEvents = $true; $x.ScreenUpdating = $true }
$x.CalculateFullRebuild()
$v = $ws.Range("A1:L$last").Value2
$got = @{ DI=0; AI=0; AO=0; DO=0; SP=0 }; $bad = 0
foreach ($h in $hdr) {
  $qty = [double]$v[($h+2),1]; $t = $h; while ([string]$v[$t,2] -ne 'TOTAL') { $t++ }
  $got.DI += [double]$v[$t,8]; $got.AI += [double]$v[$t,9]; $got.AO += [double]$v[$t,10]; $got.DO += [double]$v[$t,11]; $got.SP += [double]$v[$t,12]
  for ($c=3; $c -le 7; $c++) { $cc = $c + 5; $one = [double]$v[$t,$c]; $all = [double]$v[$t,$cc]; $diff = $one * $qty - $all; if ([math]::Abs($diff) -gt 0.01) { $bad++ } }
}
"rows rewritten=$fixed ; blocks=$($hdr.Count) ; block columns with mismatch=$bad"
"got: DI={0} AI={1} AO={2} DO={3} SP={4}" -f $got.DI, $got.AI, $got.AO, $got.DO, $got.SP
$wb.Save(); "saved=" + $wb.Saved


