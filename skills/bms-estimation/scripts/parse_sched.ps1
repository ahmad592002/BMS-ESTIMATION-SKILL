param([string]$Dir, [string]$Out)
$ErrorActionPreference = "Stop"
$PT = "C:\Users\ahmad\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdftotext.exe"
$tmp = Join-Path $env:TEMP "bbx.html"
$rx = [regex]'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>'
$res = New-Object System.Collections.Generic.List[string]
$res.Add("Sheet`tSystem`tTable`tSection`tRef`tDescription`tDI`tDO`tAI`tAO`tALARM`tHW_INTERLOCK`tCOMM`tNotes")
foreach ($pdf in Get-ChildItem $Dir -Filter "*B-93-ZZZ-*.pdf" | Sort-Object Name) {
  & $PT -bbox $pdf.FullName $tmp 2>$null
  $html = [IO.File]::ReadAllText($tmp)
  $W = foreach ($m in $rx.Matches($html)) { [pscustomobject]@{ x0=[double]$m.Groups[1].Value; y0=[double]$m.Groups[2].Value; x1=[double]$m.Groups[3].Value; y1=[double]$m.Groups[4].Value; t=[Net.WebUtility]::HtmlDecode($m.Groups[5].Value) } }
  $W = @($W)
  $sheet = ($pdf.BaseName -split '-')[-1]
  # system titles (may be several per sheet)
  $plain = ($W | Sort-Object y0, x0 | ForEach-Object { $_.t }) -join ' '
  $titles = @([regex]::Matches($plain, 'NOT TO SCALE\s+(.+?)(?=\s+\d+\s+SCALE:|\s+GENERAL NOTES|\s+KEY PLAN|$)') | ForEach-Object { $_.Groups[1].Value.Trim() })
  $lt = Join-Path "C:\Users\ahmad\AppData\Local\Temp\claude\C--Users-ahmad-OneDrive-Desktop-GTS-BMS-PROJECTS-07--Al-Moosa-University\ac502cda-0cd6-4028-b4ff-ddbf68ef253e\scratchpad\txt" ($pdf.BaseName + ".txt"); $sysTitle = ((Get-Content $lt | Select-String "NOT TO SCALE\s+(.+)" -AllMatches | ForEach-Object { ($_.Matches[0].Groups[1].Value -split "\s{3,}")[0].Trim() }) | Select-Object -Unique) -join " / "
  # table anchors
  $anchors = @($W | Where-Object { $_.t -eq 'SCHEDULE' } | Where-Object { $a=$_; $W | Where-Object { ($_.t -eq 'BMS' -or $_.t -eq 'PMS') -and [math]::Abs($_.y0-$a.y0) -lt 3 -and $_.x1 -lt $a.x0 -and ($a.x0-$_.x1) -lt 15 } } | Sort-Object x0, y0)
  $ti = 0
  foreach ($a in $anchors) {
    $ti++
    # header words: rotated (tall & narrow) in the 160pt below the anchor
    $hdrW = @($W | Where-Object { $_.y0 -gt $a.y0 + 5 -and $_.y0 -lt $a.y0 + 160 -and ($_.y1-$_.y0) -gt ($_.x1-$_.x0) -and $_.x0 -gt $a.x0 - 600 -and $_.x0 -lt $a.x0 + 600 })
    if ($hdrW.Count -eq 0) { continue }
    $clusters = @(); foreach ($h in ($hdrW | Sort-Object x0)) { $c = $clusters | Where-Object { [math]::Abs($_.x - $h.x0) -lt 4 } | Select-Object -First 1; if ($c) { $c.words += $h } else { $clusters += [pscustomobject]@{ x=$h.x0; words=@($h) } } }
    $cols = @()
    foreach ($c in $clusters) {
      $lab = (($c.words | Sort-Object y0 -Descending) | ForEach-Object { $_.t }) -join ' '
      $k = switch -regex ($lab) { 'DIGITAL INPUT' {'DI'} 'DIGITAL OUTPUT' {'DO'} 'ANALOG INPUT' {'AI'} 'ANALOG OUTPUT' {'AO'} 'ALARM' {'ALARM'} 'HARDWIRED|INTERLOCK' {'HW_INTERLOCK'} 'COMMUNICATION|INTERFACE' {'COMM'} default {$null} }
      if ($k) { $cols += [pscustomobject]@{ k=$k; x=$c.x } }
    }
    if (-not ($cols | Where-Object k -eq 'DI')) { continue }
    $hdrBottom = ($hdrW | Measure-Object y1 -Maximum).Maximum
    $colL = ($cols | Measure-Object x -Minimum).Minimum; $colR = ($cols | Measure-Object x -Maximum).Maximum
    $desc = $W | Where-Object { $_.t -eq 'DESCRIPTION' -and $_.y0 -gt $a.y0 -and $_.y0 -lt $hdrBottom -and $_.x0 -lt $colL } | Sort-Object x0 -Descending | Select-Object -First 1
    $sref = $W | Where-Object { $_.t -eq 'SCHEMATIC' -and $_.y0 -gt $a.y0 -and $_.y0 -lt $hdrBottom -and $_.x0 -lt $colL } | Sort-Object x0 -Descending | Select-Object -First 1
    $notesH = $W | Where-Object { $_.t -eq 'NOTES' -and $_.y0 -gt $a.y0 -and $_.y0 -lt $hdrBottom -and $_.x0 -gt $colR } | Select-Object -First 1
    $tblL = if ($sref) { $sref.x0 - 25 } else { $colL - 380 }
    $descR = $colL - 3; $dc = if ($desc) { ($desc.x0 + $desc.x1)/2 } else { ($tblL + $descR)/2 }
    $descL = [math]::Max($tblL + 40, 2*$dc - $descR)
    $notesR = if ($notesH) { 2*(($notesH.x0+$notesH.x1)/2) - ($colR + 30) } else { $colR + 400 }
    # next anchor below in same column of tables bounds this table
    $below = $anchors | Where-Object { $_.y0 -gt $a.y0 + 20 -and [math]::Abs($_.x0 - $a.x0) -lt 200 } | Sort-Object y0 | Select-Object -First 1
    $tblB = if ($below) { $below.y0 - 5 } else { 99999 }
    $body = @($W | Where-Object { $_.y0 -gt $hdrBottom + 2 -and $_.y0 -lt $tblB -and $_.x0 -ge $tblL -and $_.x1 -le $notesR + 5 -and ($_.y1-$_.y0) -le ($_.x1-$_.x0) + 15 })
    # stop at first big gap / non-table text: rows are lines of words; build lines
    $xs = @($body | Where-Object { $_.t -eq 'X' -and $_.x0 -ge $colL - 8 -and $_.x0 -le $colR + 12 })
    $lines = @(); foreach ($wd in ($body | Where-Object { $_.x1 -lt $colL - 1 } | Sort-Object y0)) { $l = $lines | Where-Object { [math]::Abs($_.y - $wd.y0) -lt 4 } | Select-Object -First 1; if ($l) { $l.w += $wd } else { $lines += [pscustomobject]@{ y=$wd.y0; w=@($wd) } } }
    $xrows = @(); foreach ($xw in ($xs | Sort-Object y0)) { $r = $xrows | Where-Object { [math]::Abs($_.y - $xw.y0) -lt 6 } | Select-Object -First 1; if ($r) { $r.x += $xw } else { $xrows += [pscustomobject]@{ y=$xw.y0; x=@($xw) } } }
    # bound the table: stop when vertical gap between consecutive rows/lines > 120pt
    $section = ''
    $events = @(); foreach ($l in $lines) { $events += [pscustomobject]@{ y=$l.y; kind='L'; o=$l } }; foreach ($r in $xrows) { $events += [pscustomobject]@{ y=$r.y; kind='X'; o=$r } }
    $events = $events | Sort-Object y
    $lastY = $hdrBottom
    $usedLines = @{}
    foreach ($e in $events) {
      if ($e.y - $lastY -gt 140) { break }
      $lastY = $e.y
      if ($e.kind -eq 'L') {
        $hasX = $xrows | Where-Object { [math]::Abs($_.y - $e.y) -lt 14 }
        if (-not $hasX) {
          $txt = (($e.o.w | Sort-Object x0) | ForEach-Object { $_.t }) -join ' '
          $inDesc = $e.o.w | Where-Object { $_.x0 -ge $descL - 2 }
          if (-not $inDesc -or ((@($e.o.w | ForEach-Object { $_.y1 - $_.y0 }) | Measure-Object -Average).Average) -lt 7) { $section = $txt }
        }
        continue
      }
      $r = $e.o
      $near = @($lines | Where-Object { [math]::Abs($_.y - $r.y) -lt 14 })
      $dw = @($near | ForEach-Object { $_.w } | Where-Object { $_.x0 -ge $descL - 2 } | Sort-Object y0, x0)
      $rw = @($near | ForEach-Object { $_.w } | Where-Object { $_.x0 -lt $descL - 2 } | Sort-Object x0)
      $nw = @($body | Where-Object { $_.x0 -gt $colR + 12 -and [math]::Abs($_.y0 - $r.y) -lt 14 } | Sort-Object y0, x0)
      $flags = @{ DI=''; DO=''; AI=''; AO=''; ALARM=''; HW_INTERLOCK=''; COMM='' }
      foreach ($xw in $r.x) { $c = $cols | Sort-Object { [math]::Abs($_.x - $xw.x0) } | Select-Object -First 1; $flags[$c.k] = 'X' }
      $line = @($sheet, $sysTitle, "T$ti", $section, (($rw | ForEach-Object t) -join ' '), (($dw | ForEach-Object t) -join ' '), $flags.DI, $flags.DO, $flags.AI, $flags.AO, $flags.ALARM, $flags.HW_INTERLOCK, $flags.COMM, (($nw | ForEach-Object t) -join ' '))
      $res.Add(($line -join "`t"))
    }
  }
  Write-Host ("{0}: tables={1} rows so far={2} | {3}" -f $sheet, $anchors.Count, ($res.Count-1), $sysTitle)
}
[IO.File]::WriteAllLines($Out, $res)
