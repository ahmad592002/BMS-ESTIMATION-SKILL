param([string]$Out)
$ErrorActionPreference = "Stop"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path

$ifMap = @{ HW='Hardwired I/O'; BIP='BACnet/IP (ICT network)'; MBT='Modbus TCP/IP (ICT network)'; MBR='Modbus RTU (RS-485)'; TCP='TCP/IP data point + hardwired'; BSC='BACnet/SC + hardwired' }
$fnMap = @{ M='Monitor'; MC='Monitor&Control' }

# ---- panels
$panels = [ordered]@{}
foreach ($l in [IO.File]::ReadAllLines("$sp\panels.txt")) {
  if ($l.Trim() -eq '') { continue }
  $p = $l.Split('|'); $panels[$p[0]] = [pscustomobject]@{ Tag=$p[0]; Bldg=$p[1]; Level=$p[2]; Type=$p[3]; Loc=$p[4] }
}

function Expand-Token([string]$t) {
  $m = [regex]::Match($t, '^(.*)\{(\d+)\.\.(\d+)\}(.*)$')
  if (-not $m.Success) { return ,@($t) }
  $w = $m.Groups[2].Value.Length; $a = [int]$m.Groups[2].Value; $b = [int]$m.Groups[3].Value
  $r = @(); for ($i=$a; $i -le $b; $i++) { $r += ($m.Groups[1].Value + $i.ToString().PadLeft($w,'0') + $m.Groups[4].Value) }
  return ,$r
}

$rows = New-Object System.Collections.Generic.List[object]
foreach ($l in [IO.File]::ReadAllLines("$sp\items.txt")) {
  if ($l.Trim() -eq '' -or $l.StartsWith('#')) { continue }
  $f = $l.Split('|')
  $bldg,$lvl,$pan,$spec,$desc,$cat,$if,$fn = $f[0..7]; $rem = if ($f.Count -gt 8) { $f[8] } else { '' }
  if (-not $panels.Contains($pan)) { throw "Unknown panel '$pan' in: $l" }
  foreach ($tokRaw in $spec.Split(';')) {
    $tok = $tokRaw.Trim()
    $items = @()
    if ($tok -match '^#(\d+)\s+(.+)$') {
      $n = [int]$Matches[1]; $name = $Matches[2]
      if ($n -eq 1) { $items += ,@($name, 1, 'Untagged symbol on drawing') }
      else { for ($i=1; $i -le $n; $i++) { $items += ,@("$name #$i", 1, "Untagged - drawn as ${n}NOS") } }
    } elseif ($tok -match '^\?(\d+)\s+(.+)$') {
      $items += ,@($Matches[2], [int]$Matches[1], 'Group - quantity not on drawing (assumed)')
    } elseif ($tok.StartsWith('!')) {
      foreach ($e in (Expand-Token $tok.Substring(1))) { $items += ,@($e, 0, 'Duplicate symbol - not counted') }
    } else {
      $exp = Expand-Token $tok
      foreach ($e in $exp) { $items += ,@($e, 1, 'Drawn tag') }
    }
    foreach ($it in $items) {
      $pp = $panels[$pan]
      $rows.Add([pscustomobject]@{ Bldg=$bldg; Level=$lvl; Tag=$it[0]; Desc=$desc; Cat=$cat; Panel=$pan; PType=$pp.Type; PLoc=$pp.Loc
        Iface=$ifMap[$if]; Fn=$fnMap[$fn]; Qty=$it[1]; Src=$it[2]; Rem=$rem })
    }
  }
}
"register rows: " + $rows.Count + " ; qty: " + (($rows | Measure-Object Qty -Sum).Sum)

# ---- Excel
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$wb = $x.Workbooks.Add()
while ($wb.Worksheets.Count -lt 4) { $null = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
$wsR = $wb.Worksheets.Item(1); $wsR.Name = 'Equipment Register'
$wsP = $wb.Worksheets.Item(2); $wsP.Name = 'Per DDC Panel'
$wsS = $wb.Worksheets.Item(3); $wsS.Name = 'Summary'
$wsI = $wb.Worksheets.Item(4); $wsI.Name = 'Drawing Issues'

function Write-Block($ws, $topLeft, [object[]]$header, $data) {
  $n = $data.Count; $c = $header.Count
  $arr = New-Object 'object[,]' ($n+1), $c
  for ($j=0; $j -lt $c; $j++) { $arr[0,$j] = [string]$header[$j] }
  for ($i=0; $i -lt $n; $i++) { for ($j=0; $j -lt $c; $j++) { $arr[($i+1),$j] = $data[$i][$j] } }
  $start = $ws.Range($topLeft)
  $rng = $start.Resize($n+1, $c); $rng.Value2 = $arr
  $h = $start.Resize(1, $c); $h.Font.Bold = $true; $h.Interior.Color = 0x7F3F1F; $h.Font.Color = 0xFFFFFF; $h.WrapText = $true
  return $rng
}

# Register
$hdr = 'SN','Building','Level (drawn)','Equipment Tag','Equipment Description','Category (template type)','Connected To (DDC / RIO / PLC)','Panel Type','Panel Location','Interface','BMS Function','Qty','Tag Source','Remarks'
$data = @(); $sn = 0
foreach ($r in $rows) { $sn++; $data += ,@([double]$sn, $r.Bldg, $r.Level, $r.Tag, $r.Desc, $r.Cat, $r.Panel, $r.PType, $r.PLoc, $r.Iface, $r.Fn, [double]$r.Qty, $r.Src, $r.Rem) }
$rng = Write-Block $wsR 'A1' $hdr $data
$last = $rows.Count + 1
$wsR.Range("A1:N$last").AutoFilter() | Out-Null
$wsR.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.FreezePanes = $true
$widths = 6,9,13,30,42,20,26,22,30,26,15,6,30,55
for ($j=0; $j -lt $widths.Count; $j++) { $wsR.Columns.Item($j+1).ColumnWidth = $widths[$j] }
# grey out duplicates, yellow assumed groups
for ($i=0; $i -lt $rows.Count; $i++) {
  $rr = $i + 2
  if ($rows[$i].Qty -eq 0) { $wsR.Range("A${rr}:N${rr}").Font.Color = 0x808080; $wsR.Range("A${rr}:N${rr}").Font.Italic = $true }
  elseif ($rows[$i].Src -like 'Group*') { $wsR.Range("L${rr}:M${rr}").Interior.Color = 0x99FFFF }
}
$wsR.Cells.Item($last+2, 11).Value2 = 'TOTAL QTY'; $wsR.Cells.Item($last+2, 11).Font.Bold = $true
$wsR.Cells.Item($last+2, 12).Formula = "=SUBTOTAL(9,L2:L$last)"; $wsR.Cells.Item($last+2, 12).Font.Bold = $true

# Per DDC panel
$cats = 'AHU','Fan','Pump','CWS','FCU','VAV','Electrical Panel','Electrical Equipment','Mechanical Equipment','CO System','Tank','Elevator','Other'
$ph = @('Panel','Building','Level','Panel Type','Location','Register Rows','Total Qty') + $cats + @('Equipment connected (tags)')
$pdata = @()
foreach ($k in $panels.Keys) {
  $pp = $panels[$k]
  $mine = @($rows | Where-Object { $_.Panel -eq $k -and $_.Qty -gt 0 })
  $tags = ($mine | ForEach-Object { if ($_.Qty -gt 1) { "$($_.Tag) (x$($_.Qty))" } else { $_.Tag } }) -join ', '
  if ($tags.Length -gt 32000) { $tags = $tags.Substring(0,32000) + ' ...' }
  $line = @($k, $pp.Bldg, $pp.Level, $pp.Type, $pp.Loc, $null, $null)
  foreach ($c in $cats) { $line += $null }
  $line += $tags
  $pdata += ,$line
}
$null = Write-Block $wsP 'A1' $ph $pdata
$np = $pdata.Count + 1
$reg = "'Equipment Register'!"
$wsP.Range("F2:F$np").Formula = "=COUNTIFS(${reg}`$G:`$G,`$A2,${reg}`$L:`$L,"">0"")"
$wsP.Range("G2:G$np").Formula = "=SUMIFS(${reg}`$L:`$L,${reg}`$G:`$G,`$A2)"
for ($j=0; $j -lt $cats.Count; $j++) {
  $col = $wsP.Cells.Item(2, 8+$j).Address($false,$false) -replace '\d',''
  $wsP.Range("${col}2:${col}$np").Formula = "=SUMIFS(${reg}`$L:`$L,${reg}`$G:`$G,`$A2,${reg}`$F:`$F,${col}`$1)"
}
$wsP.Cells.Item($np+1,1).Value2 = 'TOTAL'; $wsP.Cells.Item($np+1,1).Font.Bold = $true
for ($j=6; $j -le 7+$cats.Count; $j++) { $c = $wsP.Cells.Item($np+1,$j); $col = $c.Address($false,$false) -replace '\d',''; $c.Formula = "=SUM(${col}2:${col}$np)"; $c.Font.Bold = $true }
$wsP.Range("A1:U$np").AutoFilter() | Out-Null
$wsP.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.SplitColumn = 1; $x.ActiveWindow.FreezePanes = $true
$wsP.Columns.Item(1).ColumnWidth = 26; $wsP.Columns.Item(5).ColumnWidth = 30; $wsP.Columns.Item(4).ColumnWidth = 22
for ($j=6; $j -le 20; $j++) { $wsP.Columns.Item($j).ColumnWidth = 9 }
$wsP.Columns.Item(21).ColumnWidth = 120; $wsP.Range("U2:U$np").WrapText = $true
$wsP.Range("A1:T$np").VerticalAlignment = -4160

# Summary: building x category
$bl = 'ATC','EEC','MAB'
$sh = @('Category') + $bl + @('Total')
$sdata = @(); foreach ($c in $cats) { $sdata += ,(@($c) + @($null,$null,$null,$null,$null)) }
$null = Write-Block $wsS 'A1' $sh $sdata
$ns = $cats.Count + 1
foreach ($j in 0..2) { $col = [char](66+$j); $wsS.Range("${col}2:${col}$ns").Formula = "=SUMIFS(${reg}`$L:`$L,${reg}`$B:`$B,${col}`$1,${reg}`$F:`$F,`$A2)" }
$wsS.Range("E2:E$ns").Formula = '=SUM(B2:D2)'
$wsS.Cells.Item($ns+1,1).Value2 = 'TOTAL'; foreach ($j in 2..5) { $c=$wsS.Cells.Item($ns+1,$j); $col=[char](64+$j); $c.Formula = "=SUM(${col}2:${col}$ns)" }
$wsS.Range("A$($ns+1):E$($ns+1)").Font.Bold = $true
$wsS.Columns.Item(1).ColumnWidth = 24

# Summary: equipment family x building (static from register)
$r0 = $ns + 4
$wsS.Cells.Item($r0-1,1).Value2 = 'Equipment family by building (Qty, duplicates excluded)'; $wsS.Cells.Item($r0-1,1).Font.Bold = $true
$fam = $rows | Where-Object { $_.Qty -gt 0 } | Group-Object Desc | Sort-Object Name
$fdata = @()
foreach ($g in $fam) {
  $line = @($g.Name)
  $tot = 0
  foreach ($b in $bl) { $q = ($g.Group | Where-Object { $_.Bldg -eq $b } | Measure-Object Qty -Sum).Sum; if (-not $q) { $q = 0 }; $line += [double]$q; $tot += $q }
  $line += [double]$tot; $line += (($g.Group | Select-Object -ExpandProperty Cat -Unique) -join ', ')
  $fdata += ,$line
}
$null = Write-Block $wsS "A$r0" (@('Equipment family') + $bl + @('Total','Category')) $fdata
$wsS.Columns.Item(1).ColumnWidth = 60; $wsS.Columns.Item(7).ColumnWidth = 24

# Summary: drawing vs client BOQ
$r1 = $r0 + $fdata.Count + 3
$wsS.Cells.Item($r1-1,1).Value2 = 'Drawing count vs client BOQ'; $wsS.Cells.Item($r1-1,1).Font.Bold = $true
$cmp = @(
  ,@('ATC - IP unitary controller for FCU (UCF)', "=SUMIFS(${reg}L:L,${reg}B:B,""ATC"",${reg}F:F,""FCU"")", 110.0, '')
  ,@('MAB - IP unitary controller for FCU (UCF)', "=SUMIFS(${reg}L:L,${reg}B:B,""MAB"",${reg}F:F,""FCU"")", 348.0, 'Drawing and BOQ disagree (Q2)')
  ,@('MAB - IP unitary controller for VAV (UCV)', "=SUMIFS(${reg}L:L,${reg}B:B,""MAB"",${reg}F:F,""VAV"")", 449.0, 'Drawing and BOQ disagree (Q2)')
  ,@('ATC - DDCP + DDCP/SCP enclosures', 10.0, 10.0, 'Drawing 5 DDCP + 5 DDCP/SCP; BOQ 4 + 6')
  ,@('MAB - DDCP + DDCP/SCP enclosures', 30.0, 30.0, 'Drawing 14 + 16 = BOQ 14 + 16')
  ,@('EEC - DDCP + DDCP/SCP enclosures', 6.0, 0.0, 'Not in EEC BOQ (Q16)')
  ,@('MAB - Zone air quality transmitters', 10.0, 10.0, '')
  ,@('ATC - Zone air quality transmitters', 4.0, 4.0, '')
  ,@('MAB - CO/NO2 sensors', 16.0, 17.0, 'BOQ +1')
  ,@('MAB - Electric door contacts', 21.0, 21.0, 'Drawing shows 3 groups without qty; BOQ qty used')
)
$null = Write-Block $wsS "A$r1" @('Item','Drawing','Client BOQ','Note') $cmp
for ($i=0; $i -lt $cmp.Count; $i++) { $v = $cmp[$i][1]; if ($v -is [string]) { $wsS.Cells.Item($r1+1+$i, 2).Formula = $v } }
$wsS.Columns.Item(4).ColumnWidth = 14

# Issues
$iss = @(
  ,@('1','ATC','Chillers','Riser shows 5 chillers: CH-01/02/03 on RIO-ATC-B01-05 and CH-04/05 on RIO-ATC-B01-04. Schematic B-93-023 says only 3 chiller sets run on emergency power.','Count 5')
  ,@('2','MAB','FCU / VAV','Riser: 139 UCF / 691 UCV. Client BOQ: 348 FCU / 449 VAV. This register follows the drawing.','Confirm which governs')
  ,@('3','MAB','Roof AHUs','AHUs drawn only as groups (AHU 5NOS x10 + 4NOS = 54), no tags. Each roof elec room shows 3 groups (2 DDCP + DDCP/SCP) - the DDCP/SCP group may be the same AHUs shown for MCC/power monitoring (would make 39).','Need mechanical AHU schedule')
  ,@('4','ATC','AHU-RF-03','Tag drawn twice on DDCP-ATC-R02-01 (Roof 02 and Roof 01).','Counted once')
  ,@('5','ATC','RIO-ATC-B01-04','Same tag used for the RIO serving CT-01..05 (Elec Room R1-001) and the RIO serving CH-04/05 (Chiller Plant).','Drafting error - RIO tag to be confirmed')
  ,@('6','ATC','CU-03/04, EF-01-03','CU-03/04 drawn 3 times (DDCP-ATC-L01-01, DDCP-ATC-L01-02, RIO-ATC-B01-05); EF-01-03 drawn twice.','Counted once, on DDCP-ATC-L01-01')
  ,@('7','ATC','AHU-00-04/05','Drawn on both DDCP/SCP-ATC-L00-01 and DDCP-ATC-L00-01.','Counted on DDCP/SCP-ATC-L00-01')
  ,@('8','ATC','DB-ATC-L01-L1, EDB-ATC-L01-01','Drawn on DDCP-ATC-L01-01 and DDCP-ATC-L00-01.','Counted on DDCP-ATC-L01-01')
  ,@('9','ATC','ATC-WWP-07','Drawn on DDCP/SCP-ATC-B01-01 and DDCP/SCP-ATC-B01-02 (as WWP-03/06 & 07).','Counted on DDCP/SCP-ATC-B01-01')
  ,@('10','ATC','ATC-SF-01/02','Sand filters drawn on DDCP/SCP-ATC-B01-02 and DDCP-ATC-B01-01.','Counted on DDCP/SCP-ATC-B01-02')
  ,@('11','ATC','EAP-01..03, PEF, DEF','Abbreviations not in legend B-00-ZZZ-002.','Assumed pumps / exhaust fans')
  ,@('12','ATC','Lifts','4 lift groups drawn ("LIFTS"/"LIFT") without numbers or tags.','1 interface per group assumed')
  ,@('13','ATC','PLC-ATC-B01-01','MV / transformer PLC drawn in ATC but client BOQ lists the MV/LV PLC only under MAB.','Kept in scope')
  ,@('14','EEC','L01 boards','L01 distribution boards labelled DB-EEC-L00-xx (same as ground floor).','Counted as separate L01 boards')
  ,@('15','EEC','ATS-MAB-B01-02','EEC basement panel shows ATS tagged MAB-B01-02 (same tag exists in MAB).','Not counted - confirm EEC ATS')
  ,@('16','EEC','VRF indoor units','Indoor units by VRF supplier - quantity not shown.','Qty 0 - need VRF schedule')
  ,@('17','EEC','Smart meters / dampers','Smart meter and motorized damper quantities not shown.','Assumed: 2 meters per board, 1 damper group per panel')
  ,@('18','MAB','MAB-WWP-02 / 03, AMU-GTP-01','WWP-02 drawn on DDCP-MAB-B01-02 and DDCP-MAB-B01-04; WWP-03 twice on DDCP-MAB-B01-02; AMU-GTP-01 twice on DDCP-MAB-B01-02.','Counted once each')
  ,@('19','MAB','ATS-MAB-B01-04/05','Drawn on DDCP/SCP-MAB-B01-04 and DDCP-MAB-B01-07.','Counted on DDCP/SCP-MAB-B01-04')
  ,@('20','MAB','DB-MAB-L00-S7..S11','Drawn on both DDCP/SCP-MAB-L00-01 and -02.','Counted on DDCP/SCP-MAB-L00-01')
  ,@('21','MAB','Smart meters','MDB / EMDB / MCC / EMCC meters drawn without quantity.','Assumed 2 per board = 44')
  ,@('22','MAB','Lab equipment','Fume hoods, bio safety cabinets, HEPA, blowers, compressors, freezers, inline pumps drawn as groups without quantity.','1 interface per group assumed')
  ,@('23','MAB','Central water heaters','8 tagged (MAB/ATC/EEC-CWH) on DDCP/SCP-MAB-B01-04 plus "Central water heater 2NOS" on DDCP-MAB-B01-02.','Counted 10')
  ,@('24','All','Source data','No mechanical / electrical equipment schedules or SLDs in the tender folder - quantities come from the BMS system architecture risers only.','Request schedules')
)
$null = Write-Block $wsI 'A1' @('#','Building','Item','Finding','Treatment in register') $iss
$wsI.Columns.Item(1).ColumnWidth = 5; $wsI.Columns.Item(2).ColumnWidth = 9; $wsI.Columns.Item(3).ColumnWidth = 28; $wsI.Columns.Item(4).ColumnWidth = 100; $wsI.Columns.Item(5).ColumnWidth = 40
$wsI.Range("A1:E$($iss.Count+1)").WrapText = $true; $wsI.Range("A1:E$($iss.Count+1)").VerticalAlignment = -4160

foreach ($ws in @($wsR,$wsP,$wsS,$wsI)) { $ws.Cells.Font.Name = 'Calibri'; $ws.Cells.Font.Size = 10 }
$x.CalculateFull()
$wsR.Activate()
$x.ScreenUpdating = $true
if (Test-Path $Out) { Remove-Item $Out }
$wb.SaveAs($Out, 51)
"Summary totals: " + (($bl | ForEach-Object { $col=[char](66+[array]::IndexOf($bl,$_)); "$_=" + $wsS.Range("$col$($ns+1)").Value2 }) -join ' ') + " total=" + $wsS.Range("E$($ns+1)").Value2
"Per-panel total qty=" + $wsP.Cells.Item($np+1,7).Value2 + " rows=" + $wsP.Cells.Item($np+1,6).Value2
$wb.Close($false); $x.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
"saved: $Out"
