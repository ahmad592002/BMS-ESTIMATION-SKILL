$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$sp\io_lib.ps1"     # loads $TPL, $map, $eq, Build-Lines, expected totals

if (-not $x) { $x = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
$wb = $x.Workbooks.Item("BMS Template 2026 V03.xlsm")
$ws = $wb.Worksheets.Item("IOSummary"); $el = $wb.Worksheets.Item("EquipmentList"); $iot = $wb.Worksheets.Item("IOTemplate")
$wb.SaveCopyAs((Join-Path (Split-Path $wb.FullName) "BMS Template 2026 V03 - BACKUP after button generate.xlsm"))

function Find-Blocks {
  $last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
  $v = $ws.Range("A1:B$($last+30)").Value2
  $hdr = @(); for ($r=1; $r -le $last+30; $r++) { if ([string]$v[$r,1] -eq 'QTY') { $hdr += $r } }
  return ,$hdr
}
$prevEv = $x.EnableEvents
$x.EnableEvents = $false; $x.ScreenUpdating = $false; $x.Calculation = -4135
try {
  # ---- 1. blocks for EquipmentList rows the button does not reach (row > 120)
  $hdr = Find-Blocks
  $done = $hdr.Count
  if ($done -lt $eq.Count) {
    $insertAt = $hdr[$hdr.Count-1] - 1 + 29          # same spacing as the button macro (currentRow += 29)
    for ($i = $done; $i -lt $eq.Count; $i++) {
      $elRow = $i + 2
      $iot.Range("A4").Value2 = [double]$el.Range("B$elRow").Value2
      $iot.Range("B2").Value2 = [string]$el.Range("E$elRow").Value2
      $iot.Range("B4").Value2 = [string]$el.Range("C$elRow").Value2
      $null = $iot.Range("1:30").Copy()
      $null = $ws.Range("A$insertAt").EntireRow.Insert(-4121)
      
      $null = $iot.Range("B2").ClearContents(); $null = $iot.Range("A4:G26").ClearContents(); $null = $iot.Range("M4:N26").ClearContents()
      $insertAt += 29
    }
  }
  $hdr = Find-Blocks
  if ($hdr.Count -ne $eq.Count) { throw "blocks $($hdr.Count) != equipment rows $($eq.Count)" }
  # check block order matches the EquipmentList
  for ($i=0; $i -lt $eq.Count; $i++) { $nm = [string]$ws.Cells.Item($hdr[$i]+2, 2).Value2; $want = ($eq[$i] -split '\|')[2]; if ($nm.Trim() -ne $want.Trim()) { throw "block $i name '$nm' != '$want'" } }

  # ---- 2. fill points, bottom-up so row inserts/deletes never move unfilled blocks
  for ($i = $eq.Count - 1; $i -ge 0; $i--) {
    $h = $hdr[$i]; $p0 = $h + 3; $tot = $h + 25
    if ([string]$ws.Cells.Item($tot, 2).Value2 -ne 'TOTAL') { throw "block $i : TOTAL not at row $tot" }
    $lines = Build-Lines $TPL[$map[$i]] $map[$i]
    $m = $lines.Count; $n = $m + 1
    if ($n -gt 22) { $k = $n - 22; $null = $ws.Rows.Item($p0 + 1).Copy(); $null = $ws.Range("$($p0+2):$($p0+1+$k)").Insert(-4121) }
    elseif ($n -lt 22) { $null = $ws.Range("$($p0+$n):$($p0+21)").EntireRow.Delete() }
    $bg = New-Object 'object[,]' $m, 6; $mn = New-Object 'object[,]' $m, 2; $pc = New-Object 'object[,]' $m, 1
    for ($j=0; $j -lt $m; $j++) {
      $p = $lines[$j]
      for ($c=0; $c -lt 6; $c++) { $bg[$j,$c] = [string]'' }; $mn[$j,0] = [string]''; $mn[$j,1] = [string]''; $pc[$j,0] = [string]''
      if ($p.PSObject.Properties['Title']) { $bg[$j,0] = [string]$p.Title; continue }
      $bg[$j,0] = [string]$p.L
      if ([int]$p.DI -gt 0) { $bg[$j,1] = [double]$p.DI }
      if ([int]$p.AI -gt 0) { $bg[$j,2] = [double]$p.AI }
      if ([int]$p.AO -gt 0) { $bg[$j,3] = [double]$p.AO }
      if ([int]$p.DO -gt 0) { $bg[$j,4] = [double]$p.DO }
      if ([int]$p.SP -gt 0) { $bg[$j,5] = [double]$p.SP }
      if ($p.Dev) { $mn[$j,0] = [double]$p.DevQ; $mn[$j,1] = [string]$p.Dev }
      $pc[$j,0] = [string]$p.C
    }
    $ws.Range("B$p0").Resize($m, 6).Value2 = $bg
    $ws.Range("M$p0").Resize($m, 2).Value2 = $mn
    $ws.Range("P$p0").Resize($m, 1).Value2 = $pc
    for ($j=0; $j -lt $m; $j++) { $cell = $ws.Range("B$($p0+$j)"); if ($lines[$j].PSObject.Properties['Title']) { $cell.Interior.ColorIndex = 15; $cell.Font.Bold = $true } }
  }
  $x.Calculation = -4105; $x.CalculateFullRebuild()
} catch { Write-Output ("ERR " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim() + " | i=$i"); throw } finally { $x.Calculation = -4105; $x.EnableEvents = $prevEv; $x.ScreenUpdating = $true }

# ---- 3. verify
$last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$v = $ws.Range("A1:P$last").Value2
$blocks = 0; $got = @{ DI=0; AI=0; AO=0; DO=0; SP=0 }; $spUntagged = 0; $noFormula = 0
for ($r=1; $r -le $last; $r++) {
  if ([string]$v[$r,1] -eq 'QTY') { $blocks++ }
  if ([string]$v[$r,2] -eq 'TOTAL') { $got.DI += [double]$v[$r,8]; $got.AI += [double]$v[$r,9]; $got.AO += [double]$v[$r,10]; $got.DO += [double]$v[$r,11]; $got.SP += [double]$v[$r,12]; continue }
  $g = [string]$v[$r,7]; if ($g -ne '' -and $g -ne 'SP') { if ([string]$v[$r,16] -notin 'MODBUS','BACNET/IP','BACNET/MSTP','MBUS','KNX') { $spUntagged++ } }
  if ([string]$v[$r,3] -match '^\d') { if (-not ([string]$ws.Cells.Item($r,8).Formula).StartsWith('=')) { $noFormula++ } }
}
"blocks=$blocks"
"expected : DI={0} AI={1} AO={2} DO={3} SP={4}" -f $exp.DI,$exp.AI,$exp.AO,$exp.DO,$exp.SP
"got      : DI={0} AI={1} AO={2} DO={3} SP={4}" -f $got.DI,$got.AI,$got.AO,$got.DO,$got.SP
"SP rows without protocol tag=$spUntagged ; DI rows without H formula=$noFormula"
$wb.Save(); "saved=" + $wb.Saved
