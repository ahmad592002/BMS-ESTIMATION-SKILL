param([string]$Backup, [int]$Detail = 40)
# Sheet-by-sheet value comparison: live workbook vs a backup copy (formulas compared as text)
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$proj = Split-Path -Parent $env:BMS_WB
$live = [Runtime.InteropServices.Marshal]::BindToMoniker($env:BMS_WB)
$x2 = New-Object -ComObject Excel.Application; $x2.Visible = $false; $x2.DisplayAlerts = $false; $x2.AutomationSecurity = 3
$big = 'DDCSummary|DDCFullSummary|IOTemplate|Listprice|Product_Finder_SI'
try {
  $bk = $x2.Workbooks.Open($Backup, 0, $true)
  foreach ($s in $live.Worksheets) {
    $name = $s.Name; $o = $null; try { $o = $bk.Worksheets.Item($name) } catch { "### $name : NEW sheet"; continue }
    $ur = $s.UsedRange; $or = $o.UsedRange
    $lr = [math]::Max($ur.Row + $ur.Rows.Count - 1, $or.Row + $or.Rows.Count - 1); $lc = [math]::Max($ur.Column + $ur.Columns.Count - 1, $or.Column + $or.Columns.Count - 1)
    if ($lr * $lc -gt 400000) { $lc = [math]::Min($lc, 40) }
    $a = $s.Range($s.Cells.Item(1, 1), $s.Cells.Item($lr, $lc)).Formula
    $b = $o.Range($o.Cells.Item(1, 1), $o.Cells.Item($lr, $lc)).Formula
    $diff = New-Object System.Collections.Generic.List[string]; $n = 0
    for ($r = 1; $r -le $lr; $r++) { for ($c = 1; $c -le $lc; $c++) { $va = [string]$a[$r, $c]; $vb = [string]$b[$r, $c]
        if ($va -ne $vb) { $n++; if ($diff.Count -lt $Detail) { $addr = $s.Cells.Item($r, $c).Address($false, $false); $diff.Add(("    {0}: [{1}] -> [{2}]" -f $addr, ($(if ($vb.Length -gt 60) { $vb.Substring(0, 60) + '...' } else { $vb })), ($(if ($va.Length -gt 60) { $va.Substring(0, 60) + '...' } else { $va })))) } } } }
    if ($n -gt 0) { "### $name : $n cells changed (rows 1-$lr, cols 1-$lc)"; if ($name -notmatch $big) { $diff } }
  }
  foreach ($o in $bk.Worksheets) { $f = $null; try { $f = $live.Worksheets.Item($o.Name) } catch { }; if (-not $f) { "### $($o.Name) : sheet REMOVED in new version" } }
  $bk.Close($false)
} finally { $x2.Quit() }
