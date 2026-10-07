# "<project> - DDC List Sources.xlsx": controller loading + every tag on every controller
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$proj = Split-Path -Parent $env:BMS_WB
$pname = if ($env:BMS_PROJECT) { $env:BMS_PROJECT } else { (Split-Path -Leaf $proj) -replace '^\d+-\s*', '' }   # side files are named "<project> - ..."
$out = & "$sp\ddc_build.ps1"
$ctrl = @($out | Where-Object { $_.PSObject.Properties['IO'] -and $_.PSObject.Properties['Items'] })
$pinfo = @{}; Get-Content "$sp\panels.txt" -Encoding UTF8 | ForEach-Object { $a = $_ -split '\|'; $pinfo[$a[0]] = $a }; $renamed = @{}; Get-Content "$sp\panel_rename.txt" -Encoding UTF8 | Where-Object { $_ -and $_ -notmatch '^#' } | ForEach-Object { $a = $_ -split '\|'; if ($pinfo[$a[0]] -and -not $pinfo[$a[1]]) { $pinfo[$a[1]] = $pinfo[$a[0]]; $renamed[$a[1]] = 1 } }
$wbL = [Runtime.InteropServices.Marshal]::BindToMoniker($env:BMS_WB)
$spare = [double]$wbL.Worksheets.Item("Options").Range("B1").Value2
$d = $wbL.Worksheets.Item("DDC List"); $eqName = @{}; $nEqL = $d.Cells.Item(2, $d.Columns.Count).End(-4159).Column - 1; for ($s = 1; $s -le $nEqL; $s++) { $eqName[$s] = [string]$d.Cells.Item(2, $s + 1).Value2 }
$reg = @{}; Get-Content "$sp\register.tsv" -Encoding UTF8 | ForEach-Object { $a = $_ -split "`t"; if ($a[3]) { $reg[$a[3]] = $a } }
$rows1 = @(); $rows2 = @()
foreach ($c in $ctrl) { $pi = $pinfo[$c.Panel]; $tot = $c.IO[0] + $c.IO[1] + $c.IO[2] + $c.IO[3] + $c.IO[4]
  $note = if ($c.Panel -match '^NET-') { "Network panel (IP unitary controllers / VRF), levels $($c.Levels); drawn as one network -> $($c.Split) panels" } elseif ($c.Virtual) { 'BMS head-end integration' } elseif ($c.Split -gt 1) { "Drawn as one panel; $([math]::Round($c.PanelLoad)) pts incl. spare > 250 -> split into $($c.Split) panels -1...-$($c.Split)" } elseif ($renamed[$c.Panel]) { 'As drawn (panel named DDCP - GTS uses PXC controllers)' } else { 'As drawn' }
  $rows1 += , @($c.Name, $c.Panel, $pi[1], $pi[2], $pi[3], $pi[4], [double]$c.Split, $c.IO[0], $c.IO[1], $c.IO[2], $c.IO[3], ($c.IO[0] + $c.IO[1] + $c.IO[2] + $c.IO[3]), $c.IO[4], $tot, [double]$c.WithSpare, $(if ($c.Virtual) { '' } else { [math]::Round($c.WithSpare / 250, 2) }), $note)
  foreach ($it in ($c.Items | Sort-Object SN, Tag)) { $baseTag = $it.Tag -replace ' \[\d+/\d+\]$', ''; $r = $reg[$baseTag]
    $rows2 += , @($c.Name, $c.Panel, [double]$it.SN, $eqName[$it.SN], $it.Tag, $(if ($r) { $r[2] } else { '' }), $(if ($r) { $r[4] } else { '' }), [double]$it.Qty, $it.Unit, $(if ($r) { $r[9] } else { '' }), "B-92 BMS riser ($($pi[1]))", $(if ($r) { $r[12] } else { '' }), $(if ($r) { $r[13] } else { '' })) } }
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
try {
  $wb = $x.Workbooks.Add(); while ($wb.Worksheets.Count -lt 2) { $null = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $w1 = $wb.Worksheets.Item(1); $w1.Name = 'Controllers (loading)'; $w2 = $wb.Worksheets.Item(2); $w2.Name = 'Tags per Controller'
  function Put($sh, [string[]]$hdr, $rows) { $n = $rows.Count; $c = $hdr.Count; $a = New-Object 'object[,]' ($n + 1), $c
    for ($j = 0; $j -lt $c; $j++) { $a[0, $j] = [string]$hdr[$j] }
    for ($i = 0; $i -lt $n; $i++) { for ($j = 0; $j -lt $c; $j++) { $v = $rows[$i][$j]; if ($v -is [double] -or $v -is [int]) { $a[($i + 1), $j] = [double]$v } else { $a[($i + 1), $j] = [string]$v } } }
    $sh.Range("A1").Resize($n + 1, $c).Value2 = $a
    $h = $sh.Range("A1").Resize(1, $c); $h.Font.Bold = $true; $h.Interior.Color = 0x7F3F1F; $h.Font.Color = 0xFFFFFF; $h.WrapText = $true; $h.VerticalAlignment = -4108
    $null = $sh.Range("A1").Resize($n + 1, $c).AutoFilter()
    try { $sh.Parent.Activate(); $sh.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.FreezePanes = $true } catch { } }
  Put $w1 @('Controller (DDC List row)', 'Drawn panel', 'Building', 'Level', 'Panel type', 'Location', 'Panels after split', 'DI', 'AI', 'AO', 'DO', 'Physical', 'SP', 'Total', "Total x (1+$spare)", 'Load vs 250', 'Basis') $rows1
  $n1 = $rows1.Count + 1
  $wd = 34, 26, 7, 10, 26, 34, 9, 6, 6, 6, 6, 8, 7, 7, 9, 8, 60; for ($j = 0; $j -lt $wd.Count; $j++) { $w1.Columns.Item($j + 1).ColumnWidth = $wd[$j] }
  $w1.Range("P2:P$n1").NumberFormat = '0%'
  for ($r = 2; $r -le $n1; $r++) { $b = [string]$w1.Cells.Item($r, 17).Value2; if ($b -notlike 'As drawn*') { $w1.Range("A$r").Interior.Color = 65535; $w1.Range("Q$r").Interior.Color = 65535 } }
  $tr = $n1 + 1; $w1.Cells.Item($tr, 1).Value2 = 'TOTAL'; foreach ($col in 'H', 'I', 'J', 'K', 'L', 'M', 'N') { $w1.Range("$col$tr").Formula = "=SUBTOTAL(9,$col`2:$col$n1)" }; $w1.Range("A$tr`:Q$tr").Font.Bold = $true
  Put $w2 @('Controller', 'Drawn panel', 'SN', 'Equipment (EquipmentList / IOSummary)', 'Tag (as drawn)', 'Level (drawn)', 'Description (riser)', 'Qty', 'Points per unit (phys+SP)', 'Interface', 'Source', 'Tag source', 'Remarks') $rows2
  $wd = 34, 26, 5, 52, 30, 12, 40, 5, 9, 18, 22, 16, 50; for ($j = 0; $j -lt $wd.Count; $j++) { $w2.Columns.Item($j + 1).ColumnWidth = $wd[$j] }
  foreach ($sh in @($w1, $w2)) { $sh.Cells.Font.Name = 'Calibri'; $sh.Cells.Font.Size = 10 }
  $x.CalculateFull(); $w1.Activate(); $x.ScreenUpdating = $true
  $f = "$proj\$pname - DDC List Sources.xlsx"; if (Test-Path $f) { Copy-Item $f ("$proj\Old Versions\$pname - DDC List Sources (before regen " + (Get-Date -Format yyyy-MM-dd) + ").xlsx") -Force; [IO.File]::Delete($f) }; $wb.SaveAs($f, 51)
  "controllers=$($rows1.Count) ; tag rows=$($rows2.Count)"
  "TOTAL DI={0} AI={1} AO={2} DO={3} SP={4}" -f $w1.Range("H$tr").Value2, $w1.Range("I$tr").Value2, $w1.Range("J$tr").Value2, $w1.Range("K$tr").Value2, $w1.Range("M$tr").Value2
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }




