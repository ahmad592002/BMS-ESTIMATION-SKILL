param([switch]$DryRun)
# GTS standard (Al Moosa final, 2026-10-07): every valve = 2 rows with priced models
#   command row  -> [Water] [Valve]          <valve body>
#   feedback row -> [Water] [Valve Actuator] <actuator>
#   PICV (AHU/FAHU/ERU cooling coils)  VPF44.65F25 + SAX61P03   (drawings/BOQ say PICV -> never price a globe)
#   solenoid / run-around (globe)      VVF42.65-50 + SKB62/F
#   butterfly (HEX, CHW header, ICV)   VFW41.150  + SQL341E100  (set by hand on the butterfly rows)
# Sizes are almost never in the tender data -> every model cell yellow 'size not in drawings'.
# The template's "V.A. Selection Sheet" devices give no price. Models typed in O must be listed in the
# Valves sheet direct section (=SUMIFS(IOSummary!M:M, IOSummary!O:O, A..)) so the BOQ macro picks them up.
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$path = $env:BMS_WB
$wb = [Runtime.InteropServices.Marshal]::BindToMoniker($path); $x = $wb.Application
$ws = $wb.Worksheets.Item("IOSummary"); $YEL = 65535
$MODELS = @{ 'PICV' = @('VPF44.65F25', 'SAX61P03'); 'GLOBE' = @('VVF42.65-50', 'SKB62/F') }
$NOTE = 'valve size not in drawings - GTS standard model, DN65 assumed - confirm size'
$split = @{
  '[Pressure Independent Control Valve]' = @{ Cmd = 'CHILLED WATER PICV VALVE MODULATING CONTROL'; Fb = 'CHILLED WATER PICV VALVE POSITION FEEDBACK' }
  'SOLENOID'   = @{ Cmd = 'SOLENOID VALVE OPEN/CLOSE COMMAND'; Fb = 'SOLENOID VALVE OPEN/CLOSE STATUS' }
  'RUN AROUND' = @{ Cmd = 'RUN AROUND COILS CONTROL VALVE OPEN/CLOSE COMMAND'; Fb = 'RUN AROUND COILS CONTROL VALVE ACTUATOR' } }
$log = New-Object System.Collections.Generic.List[string]
$last = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
$F = $ws.Range("A1:Q$last").Formula
# block qty row for every row
$qOf = @{}; $q = 0; $bn = @{}; $name = ''
for ($r = 1; $r -le $last; $r++) { if ([string]$F[$r, 1] -eq 'QTY') { $q = $r + 2; $name = [string]$F[$q, 2] }; $qOf[$r] = $q; $bn[$r] = $name }
$x.ScreenUpdating = $false; $x.EnableEvents = $false
try {
  for ($r = $last; $r -ge 1; $r--) {
    $dev = [string]$F[$r, 14]; $nm = ([string]$F[$r, 2]).Trim()
    $kind = $null
    if ($dev -eq '[Pressure Independent Control Valve]') { $kind = $dev }
    elseif ($dev -eq '[Globe Valve] [2 Ports]' -and $nm -like 'SOLENOID VALVE*') { $kind = 'SOLENOID' }
    elseif ($dev -eq '[Globe Valve] [2 Ports]' -and $nm -like 'RUN AROUND*') { $kind = 'RUN AROUND' }
    if (-not $kind) { continue }
    $pair2 = if ($kind -eq '[Pressure Independent Control Valve]') { $MODELS['PICV'] } else { $MODELS['GLOBE'] }; $VALVE = $pair2[0]; $ACT = $pair2[1]
    $q = $qOf[$r]; $mf = [string]$F[$r, 13]; $k = 1; if ($mf -match '^=(\d+(\.\d+)?)\*') { $k = [double]$Matches[1] }
    $io = @(); for ($c = 3; $c -le 7; $c++) { $io += [string]$F[$r, $c] }
    $cmd = @('', '', $io[2], $io[3], ''); $fb = @($io[0], $io[1], '', '', '')
    $log.Add(("r{0} [{1}] {2} io={3} -> cmd(AO/DO)={4} | fb(DI/AI)={5}" -f $r, $bn[$r], $nm, ($io -join ','), ($cmd -join ','), ($fb -join ',')))
    if ($DryRun) { continue }
    $wasY = @(); for ($c = 3; $c -le 7; $c++) { $wasY += ($ws.Cells.Item($r, $c).Interior.Color -eq $YEL) }
    $oldQ = [string]$F[$r, 17]
    [void]$ws.Rows.Item($r + 1).Insert()
    foreach ($pair in @(@($r, $split[$kind].Cmd, $cmd, '[Water] [Valve]', $VALVE, @(2, 3)), @(($r + 1), $split[$kind].Fb, $fb, '[Water] [Valve Actuator]', $ACT, @(0, 1)))) {
      $rr = $pair[0]; $vals = $pair[2]
      $ws.Cells.Item($rr, 2).Value2 = [string]$pair[1]
      for ($i = 0; $i -lt 5; $i++) { $cell = $ws.Cells.Item($rr, 3 + $i)
        if ($vals[$i] -ne '') { $cell.Value2 = [double]$vals[$i] } else { $cell.Value2 = [string]'' }
        if ($vals[$i] -ne '' -and $wasY[$i]) { $cell.Interior.Color = $YEL } else { $cell.Interior.ColorIndex = -4142 } }
      $ws.Cells.Item($rr, 2).Interior.ColorIndex = -4142
      $ws.Cells.Item($rr, 13).Formula = [string]("=" + $k + "*`$A$q")
      $ws.Cells.Item($rr, 14).Value2 = [string]$pair[3]
      $ws.Cells.Item($rr, 15).Value2 = [string]$pair[4]
      $ws.Cells.Item($rr, 16).Value2 = [string]''
      $ws.Range($ws.Cells.Item($rr, 13), $ws.Cells.Item($rr, 15)).Interior.Color = [int]$YEL
      $nq = if ($rr -eq $r -and $oldQ) { "$oldQ; $NOTE" } else { $NOTE }
      $ws.Cells.Item($rr, 17).Value2 = [string]$nq; $ws.Cells.Item($rr, 17).Interior.Color = $YEL
    }
  }
  # Valves sheet: make sure the models are in the IOSummary-direct list
  $vs = $wb.Worksheets.Item("Valves")
  foreach ($pn in ($MODELS.Values | ForEach-Object { $_ })) {
    $found = $false; for ($r = 62; $r -le 100; $r++) { if ([string]$vs.Cells.Item($r, 1).Value2 -eq $pn) { $found = $true } }
    if (-not $found) { $r = 62; while ([string]$vs.Cells.Item($r, 1).Value2) { $r++ }
      if (-not $DryRun) { $vs.Cells.Item($r, 1).Value2 = [string]$pn
        $d = $wb.Worksheets.Item("Pricelist").UsedRange.Find($pn, [Type]::Missing, -4163, 1)
        if ($d) { $vs.Cells.Item($r, 2).Value2 = [string]$wb.Worksheets.Item("Pricelist").Cells.Item($d.Row, 2).Value2 }
        else { $log.Add("WARNING $pn is not in the Pricelist - add it (net price) or the BOQ shows #N/A") }
        $vs.Cells.Item($r, 3).Formula = [string]"=SUMIFS(IOSummary!M:M, IOSummary!O:O, A$r)" }
      $log.Add("Valves!A$r added $pn (direct from IOSummary)") } }
} catch { "ERROR line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim()) :: $($_.Exception.Message)"; throw } finally { $x.EnableEvents = $true; $x.ScreenUpdating = $true }
$log



