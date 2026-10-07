param([string]$Src, [string]$Out, [switch]$Live)
# Dump IOSummary A:Q (formulas) to a TSV: row, A..Q, fill colour of B..Q (yellow flags)
$ErrorActionPreference = "Stop"
if ($Live) { $wb = [Runtime.InteropServices.Marshal]::BindToMoniker($Src); $x = $null }
else { $x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
       $x.AutomationSecurity = 3; $wb = $x.Workbooks.Open($Src, 0, $true) }
try {
  $ws = $wb.Worksheets.Item("IOSummary")
  $last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
  $f = $ws.Range("A1:Q$last").Formula
  $sb = New-Object System.Text.StringBuilder
  for ($r = 1; $r -le $last; $r++) {
    $cells = @($r)
    for ($c = 1; $c -le 17; $c++) { $cells += ([string]$f[$r, $c]) -replace "`t", ' ' -replace "`r?`n", ' / ' }
    [void]$sb.AppendLine(($cells -join "`t"))
  }
  [IO.File]::WriteAllText($Out, $sb.ToString(), [Text.Encoding]::UTF8)
  "dumped $last rows -> $Out"
} finally { if ($x) { $wb.Close($false); $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) } }
