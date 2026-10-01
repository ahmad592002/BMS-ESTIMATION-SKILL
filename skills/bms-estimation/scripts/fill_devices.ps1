$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$sp\io_lib.ps1" | Out-Null
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker("C:\Users\ahmad\OneDrive\Desktop\GTS\BMS_PROJECTS\07- Al Moosa University\BMS Template 2026 V03.xlsm")
$x = $wb.Application
$wb.SaveCopyAs((Join-Path (Split-Path $wb.FullName) "BMS Template 2026 V03 - BACKUP before field devices 2.xlsm"))
$ws = $wb.Worksheets.Item("IOSummary")
$last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$v = $ws.Range("A1:B$last").Value2
$hdr = @(); for ($r=1; $r -le $last; $r++) { if ([string]$v[$r,1] -eq 'QTY') { $hdr += $r } }
if ($hdr.Count -ne $eq.Count) { throw "blocks $($hdr.Count) != $($eq.Count)" }
$x.ScreenUpdating = $false; $x.EnableEvents = $false
$written = 0; $mism = 0
try {
  for ($i=0; $i -lt $eq.Count; $i++) {
    $p0 = $hdr[$i] + 3
    $lines = Build-Lines $TPL[$map[$i]] $map[$i]
    $m = $lines.Count
    $cur = $ws.Range("B$p0").Resize($m, 1).Value2
    $mn = New-Object 'object[,]' $m, 2; $pc = New-Object 'object[,]' $m, 1
    for ($j=0; $j -lt $m; $j++) {
      $p = $lines[$j]; $want = if ($p.PSObject.Properties['Title']) { [string]$p.Title } else { [string]$p.L }
      $have = if ($m -eq 1) { [string]$cur } else { [string]$cur[($j+1),1] }
      if ($have.Trim() -ne $want.Trim()) { $mism++; throw "block $($i+1) row $($p0+$j): '$have' != '$want'" }
      $mn[$j,0] = [string]''; $mn[$j,1] = [string]''; $pc[$j,0] = [string]''
      if ($p.PSObject.Properties['Title']) { continue }
      if ($p.Dev) { $mn[$j,0] = [string]("=" + [int]$p.DevQ + "*`$A" + ($hdr[$i] + 2)); $mn[$j,1] = [string]$p.Dev; $written++ }
      if ([int]$p.SP -gt 0) { $pc[$j,0] = [string]$p.C }
      else { $notes = @(); if ($p.C) { $notes += [string]$p.C }; if ($p.DevNote) { $notes += [string]$p.DevNote }; $pc[$j,0] = [string](($notes | Select-Object -Unique) -join '; ') }
    }
    $ws.Range("M$p0").Resize($m, 2).Formula = $mn
    $ws.Range("P$p0").Resize($m, 1).Value2 = $pc
  }
  $x.CalculateFull()
} finally { $x.EnableEvents = $true; $x.ScreenUpdating = $true }
# verify: every device resolves to a model
$v = $ws.Range("A1:P$last").Value2; $unres = @(); $byModel = @{}
for ($r=1; $r -le $last; $r++) { $n = [string]$v[$r,14]; if ($n -and $n -ne 'Field Device') { $o = [string]$v[$r,15]; if (-not $o -or $o -match '#N/A|not found') { $unres += "$r : $n" } ; $byModel[$o] += [double]$v[$r,13] } }
"device rows written=$written ; unresolved=" + $unres.Count; $unres | Select-Object -First 10
$wb.Save(); "saved=" + $wb.Saved
$fd = $wb.Worksheets.Item("FieldDevices"); $fl = $fd.Cells.Item($fd.Rows.Count,1).End(-4162).Row; $fv = $fd.Range("A1:H$fl").Value2
"--- FieldDevices sheet header: " + (($fd.Range("A1:H1").Value2 | % { $_ }) -join ' | ')
