param([string]$Old, [string]$New, [string]$Filter = '.')
# Block-level diff of two IOSummary dumps. Blocks keyed by equipment name (row h+2, col B).
function Read-Blocks($path) {
  $rows = Get-Content $path -Encoding UTF8 | ForEach-Object { , ($_ -split "`t") }
  $blocks = [ordered]@{}; $cur = $null
  for ($i = 0; $i -lt $rows.Count; $i++) {
    $c = $rows[$i]
    if ($c[1] -eq 'QTY') { $name = $rows[$i + 2][2]; $cur = [pscustomobject]@{ Name = $name; Hdr = [int]$c[0]; Qty = $rows[$i + 2][1]; Lines = New-Object System.Collections.Generic.List[string] }; $blocks[$name] = $cur; $i += 2; continue }
    if (-not $cur) { continue }
    if ($c[2] -eq 'TOTAL') { $cur = $null; continue }
    if (($c[2..16] -join '').Trim() -eq '' -or ($c[2].Trim() -eq '' -and ($c[15] -match '^=IFNA'))) { continue }
    # B | C..G | M | N | O | P  (skip H:L formulas)
    $cur.Lines.Add((@($c[2], $c[3], $c[4], $c[5], $c[6], $c[7], ($c[13] -replace '\$A\d+', '$Aq'), $c[14], ($c[15] -replace '^=IFNA\(VLOOKUP.*', '<lk>')) -join ' | '))
  }
  $blocks
}
$o = Read-Blocks $Old; $n = Read-Blocks $New
foreach ($k in $n.Keys) {
  if ($k -notmatch $Filter) { continue }
  if (-not $o.Contains($k)) { "=== NEW BLOCK: $k (hdr r$($n[$k].Hdr))"; $n[$k].Lines | ForEach-Object { "   + $_" }; continue }
  $a = $o[$k].Lines; $b = $n[$k].Lines
  $d = Compare-Object -ReferenceObject @($a) -DifferenceObject @($b) -SyncWindow 500
  if ($d -or $o[$k].Qty -ne $n[$k].Qty) {
    "=== $k (hdr r$($n[$k].Hdr)) qty $($o[$k].Qty) -> $($n[$k].Qty)"
    foreach ($x in $d) { if ($x.SideIndicator -eq '<=') { "   - $($x.InputObject)" } else { "   + $($x.InputObject)" } }
  }
}
foreach ($k in $o.Keys) { if ($k -match $Filter -and -not $n.Contains($k)) { "=== REMOVED BLOCK: $k" } }

