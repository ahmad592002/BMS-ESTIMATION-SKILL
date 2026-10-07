param([string]$Out)
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$pe = @([IO.File]::ReadAllLines("$sp\per_equipment.txt") | Where-Object { $_ -and -not $_.StartsWith('#') } | ForEach-Object { ,($_.Split('|')) })
$pts = @(Import-Csv "$sp\points_all.csv")

$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$wb = $x.Workbooks.Add()
while ($wb.Worksheets.Count -lt 3) { $null = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
$w1 = $wb.Worksheets.Item(1); $w1.Name = 'Per Equipment'
$w2 = $wb.Worksheets.Item(2); $w2.Name = 'Per Sheet'
$w3 = $wb.Worksheets.Item(3); $w3.Name = 'All Points'

function Put($ws, [string[]]$hdr, $rows, [int[]]$numCols) {
  $n = $rows.Count; $c = $hdr.Count
  $a = New-Object 'object[,]' ($n+1), $c
  for ($j=0; $j -lt $c; $j++) { $a[0,$j] = $hdr[$j] }
  for ($i=0; $i -lt $n; $i++) { for ($j=0; $j -lt $c; $j++) { $v = $rows[$i][$j]; $d = 0.0
      if (($numCols -contains $j) -and [double]::TryParse([string]$v, [ref]$d)) { $a[($i+1),$j] = $d } else { $sv = [string]$v; if ($sv -match '^\d' -and -not ($numCols -contains $j)) { $sv = "'" + $sv }; $a[($i+1),$j] = $sv } } }
  $ws.Range("A1").Resize($n+1, $c).Value2 = $a
  $h = $ws.Range("A1").Resize(1, $c); $h.Font.Bold = $true; $h.Interior.Color = 0x7F3F1F; $h.Font.Color = 0xFFFFFF; $h.WrapText = $true
  $null = $ws.Range("A1").Resize($n+1, $c).AutoFilter()
  $ws.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.FreezePanes = $true
}

# Per Equipment
$rows = @(); $i = 0
foreach ($p in $pe) { $i++; $tot = 0; foreach ($k in 4..8) { $d = 0.0; if ([double]::TryParse($p[$k], [ref]$d)) { $tot += $d } }
  $rows += ,@($i, $p[0], $p[1], $p[2], $p[4], $p[5], $p[6], $p[7], $p[8], $tot, $p[9], $p[10]) }
Put $w1 @('#','Equipment / system (per unit)','Applies to (EquipmentList)','Schematic sheet','DI','DO','AI','AO','SP','Total','Confidence','Points / notes') $rows @(0,4,5,6,7,8,9)
$wd = 5,38,34,10,6,6,6,6,6,7,26,90; for ($j=0; $j -lt $wd.Count; $j++) { $w1.Columns.Item($j+1).ColumnWidth = $wd[$j] }
for ($r=2; $r -le $rows.Count+1; $r++) {
  $cf = [string]$w1.Cells.Item($r, 11).Value2
  $col = if ($cf -like 'Read*' -or $cf -like 'Schedule*') { 0xCCFFCC } elseif ($cf -match 'check|inferred|unmatched|Partly|nothing') { 0x99FFFF } else { 0xFFFFFF }
  $w1.Range("K${r}").Interior.Color = $col
}

# Per Sheet
$grp = $pts | Group-Object Sheet | Sort-Object Name
$rows = @()
foreach ($g in $grp) { $gi = $g.Group
  $rows += ,@($g.Name, ($gi[0].System), (($gi | Select-Object -ExpandProperty Source -Unique) -join ' + '), $gi.Count,
    ($gi | Measure-Object DI -Sum).Sum, ($gi | Measure-Object DO -Sum).Sum, ($gi | Measure-Object AI -Sum).Sum, ($gi | Measure-Object AO -Sum).Sum, ($gi | Measure-Object SP -Sum).Sum,
    @($gi | Where-Object { $_.Check }).Count) }
foreach ($s in '002','013','023','035','047','051') { $rows += ,@($s, '(no point marks on this sheet - P&ID / layout / sequence only)', '-', 0, 0, 0, 0, 0, 0, 0) }
$rows = $rows | Sort-Object { $_[0] }
Put $w2 @('Sheet','System title','Source','Point rows','DI','DO','AI','AO','SP','Rows to check') $rows @(3,4,5,6,7,8,9)
$wd = 7,70,30,9,6,6,6,6,6,9; for ($j=0; $j -lt $wd.Count; $j++) { $w2.Columns.Item($j+1).ColumnWidth = $wd[$j] }

# All Points
$rows = @(); foreach ($p in ($pts | Sort-Object Sheet)) { $rows += ,@($p.Sheet, $p.System, $p.Source, $p.Group, $p.Ref, $p.Point, $p.DI, $p.DO, $p.AI, $p.AO, $p.SP, $p.Alarm, $p.Interlock, $p.Check, $p.Notes) }
Put $w3 @('Sheet','System','Source','Group / section','Ref','Point description','DI','DO','AI','AO','SP','Alarm','HW interlock','Check','Notes') $rows @(6,7,8,9,10)
$wd = 7,40,16,30,8,60,5,5,5,5,5,6,8,36,40; for ($j=0; $j -lt $wd.Count; $j++) { $w3.Columns.Item($j+1).ColumnWidth = $wd[$j] }
for ($r=2; $r -le $rows.Count+1; $r++) { if ([string]$w3.Cells.Item($r,14).Value2) { $w3.Range("N${r}").Interior.Color = 0x99FFFF } }

foreach ($ws in @($w1,$w2,$w3)) { $ws.Cells.Font.Name = 'Calibri'; $ws.Cells.Font.Size = 10 }
$w1.Activate(); $x.ScreenUpdating = $true
if (Test-Path $Out) { [IO.File]::Delete($Out) }
$wb.SaveAs($Out, 51); $wb.Close($false); $x.Quit()
"saved: $Out ; per-equipment rows=" + $pe.Count + " ; points=" + $pts.Count
