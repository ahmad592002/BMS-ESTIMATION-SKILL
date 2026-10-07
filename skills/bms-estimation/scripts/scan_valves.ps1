$ErrorActionPreference = "Stop"
$root = "C:\Users\ahmad\OneDrive\Desktop\GTS\BMS_PROJECTS"
$files = @("01- Ajyad Tower - Building Management System (BMS)\TEMPLATE\AJyad Tower.xlsm", "02-YALJ\YALJ.xlsm",
  "03- Prince mansour\Estimation\Prince Mansour BMS - GTS offer.xlsm", "04- Qiddiya\Qiddiya_BMS_Estimation.xlsm",
  "05- RX premium hub\Data\Riyadh Air Premium Hub - GTS offer.xlsm", "05- RX premium hub\BMS Template 2026 V03 - RAPEH.xlsm")
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false; $x.AutomationSecurity = 3
$pat = 'Valve|VALVE|Actuator|ACTUATOR|PICV'
$mpat = '^(VV|VX|VP|VK|VF|VI|VB|VD|VM|VG|VA|SK|SA[SXLV]|SQ|SSB|SSC|SSA|SSD|SSP|GDB|GLB|GMA|GCA|ACV|STA|EVG|EVF)'
try {
  foreach ($f in $files) {
    $p = Join-Path $root $f; if (-not (Test-Path $p)) { continue }
    $wb = $x.Workbooks.Open($p, 0, $true)
    "################ $f"
    try {
      $ws = $wb.Worksheets.Item("IOSummary"); $last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
      $v = $ws.Range("A1:O$last").Value2; $blk = ''; $seen = @{}
      for ($r = 1; $r -le $last; $r++) {
        if ([string]$v[$r, 1] -eq 'QTY') { $blk = [string]$v[($r + 2), 2] }
        $n = [string]$v[$r, 14]; $o = [string]$v[$r, 15]; $b = [string]$v[$r, 2]
        if (($n -match $pat -or $o -match $mpat) -and $o -and $o -notmatch 'Selection Sheet') {
          $k = "$b|$n|$o"; if ($seen[$k]) { continue }; $seen[$k] = 1
          $io = @(); for ($c = 3; $c -le 7; $c++) { $io += [string]$v[$r, $c] }
          "  IO [{0}] {1} | io={2} | M={3} | N={4} | O={5}" -f $blk.Substring(0, [Math]::Min(35, $blk.Length)), $b, ($io -join ','), $v[$r, 13], $n, $o } }
      foreach ($sn in 'Valves') { $s = $wb.Worksheets.Item($sn); $u = $s.UsedRange.Value2
        for ($r = 2; $r -le $u.GetLength(0); $r++) { foreach ($c in 1, 5, 9, 13, 17) { $q = $u[$r, ($c + 2)]; if ($q -is [double] -and $q -gt 0) { "  VALVES-SHEET {0} | {1} | qty={2}" -f $u[$r, $c], $u[$r, ($c + 1)], $q } } } }
      $s = $wb.Worksheets.Item("ValvesAndActuators"); $u = $s.UsedRange.Value2
      for ($r = 4; $r -le $u.GetLength(0); $r++) { if ([string]$u[$r, 15]) { "  V&A {0} | {1} | qty={2} | {3} | sig={4} | DN={8} | valve={15} | act={17}" -f $u[$r, 1], $u[$r, 2], $u[$r, 3], $u[$r, 4], $u[$r, 5], $u[$r, 6], $u[$r, 7], $u[$r, 8], $u[$r, 12], $u[$r, 13], $u[$r, 14], $u[$r, 9], $u[$r, 10], $u[$r, 11], $u[$r, 16], $u[$r, 15], $u[$r, 18], $u[$r, 17] } }
    } catch { "  error: $($_.Exception.Message)" } finally { $wb.Close($false) }
  }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
