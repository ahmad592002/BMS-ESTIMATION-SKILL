# EquipmentList row -> panel quantities, from the takeoff register (B-92 risers) via el_map.txt
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$reg = Get-Content "$sp\register.tsv" -Encoding UTF8 | ForEach-Object { , ($_ -split "`t") } | Where-Object { $_[1] -and $_[4] }
$map = Get-Content "$sp\el_map.txt" -Encoding UTF8 | Where-Object { $_ -and $_ -notmatch '^#' } | ForEach-Object { $a = $_ -split '\|'; [pscustomobject]@{ SN = [int]$a[0]; B = $a[1]; D = $a[2]; T = $a[3]; NT = $a[4] } }
$used = @{}; $assign = @{}   # SN -> ordered panel -> qty
foreach ($m in $map) {
  $assign[$m.SN] = [ordered]@{}
  for ($i = 0; $i -lt $reg.Count; $i++) { $r = $reg[$i]
    if ($r[1] -ne $m.B) { continue }
    if ($r[4] -notmatch ('^(' + $m.D + ')$')) { continue }
    if ($m.T -and $r[3] -notmatch $m.T) { continue }
    if ($m.NT -and $r[3] -match $m.NT) { continue }
    $q = [double]$r[11]; if ($q -le 0) { continue }
    if ($used.ContainsKey($i)) { Write-Warning "register row $($i+2) ($($r[3]) / $($r[4])) matched by SN $($used[$i]) and SN $($m.SN)" }
    $used[$i] = $m.SN
    $p = $r[6]; if (-not $assign[$m.SN].Contains($p)) { $assign[$m.SN][$p] = 0.0 }; $assign[$m.SN][$p] += $q }
}
$un = @(); for ($i = 0; $i -lt $reg.Count; $i++) { if (-not $used.ContainsKey($i) -and [double]$reg[$i][11] -gt 0) { $un += "  unmatched register row $($i+2): $($reg[$i][1]) | $($reg[$i][3]) | $($reg[$i][4]) | q=$($reg[$i][11])" } }
$un
$assign
