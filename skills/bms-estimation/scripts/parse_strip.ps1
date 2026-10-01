param([string]$Dir, [string]$Out, [string]$Only = '')
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
public static class DotFinder {
  static bool D(byte[] b,int s,int w,int h,int x,int y){ if(x<0||y<0||x>=w||y>=h) return false; int i=y*s+x*3; return (b[i]+b[i+1]+b[i+2])<300; }
  public static int[] Find(byte[] b,int s,int w,int h,int yc,int band){
    var r=new System.Collections.Generic.List<int>(); int x=2;
    while(x<w-2){ int hy=-1;
      for(int dy=-band;dy<=band && hy<0;dy++){ int y=yc+dy; if(!D(b,s,w,h,x,y)) continue; int c=0; for(int a=-1;a<=1;a++) for(int q=-1;q<=1;q++) if(D(b,s,w,h,x+a,y+q)) c++; if(c>=9) hy=y; }
      for(int dy=-band;dy<=band && hy<0;dy++){ int y=yc+dy; if(D(b,s,w,h,x,y)||D(b,s,w,h,x+1,y)||D(b,s,w,h,x-1,y)) continue;
        for(int rr=3;rr<=6 && hy<0;rr++){ int ok=0; for(int k=0;k<16;k++){ double an=k*System.Math.PI/8; int px=x+(int)System.Math.Round(rr*System.Math.Cos(an)), py=y+(int)System.Math.Round(rr*System.Math.Sin(an)); if(D(b,s,w,h,px,py)||D(b,s,w,h,px+1,py)||D(b,s,w,h,px,py+1)) ok++; }
          int inner=0; for(int a=-(rr-2);a<=rr-2;a++) for(int q=-(rr-2);q<=rr-2;q++) if(D(b,s,w,h,x+a,y+q)) inner++;
          if(ok>=15 && inner<=2*(rr-1)) hy=y; } }
      if(hy>=0){ r.Add(x); r.Add(hy); x+=8; } else x++; }
    return r.ToArray(); }
}
"@
$PT  = "C:\Users\ahmad\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdftotext.exe"
$PPM = "C:\Users\ahmad\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdftoppm.exe"
$sp = Split-Path -Parent $MyInvocation.MyCommand.Path
$tmp = Join-Path $env:TEMP "bbs.html"
$DPI = 200; $K = $DPI / 72.0
$rx = [regex]'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>'
$res = New-Object System.Collections.Generic.List[string]
$res.Add("Sheet`tSystem`tStrip`tPoint`tType`tMult`tLabelX`tDotX")

foreach ($pdf in Get-ChildItem $Dir -Filter "*B-93-ZZZ-*.pdf" | Sort-Object Name) {
  $sheet = ($pdf.BaseName -split '-')[-1]
  if ($Only -and $sheet -ne $Only) { continue }
  & $PT -bbox $pdf.FullName $tmp 2>$null
  $html = [IO.File]::ReadAllText($tmp)
  $W = @(foreach ($m in $rx.Matches($html)) { [pscustomobject]@{ x0=[double]$m.Groups[1].Value; y0=[double]$m.Groups[2].Value; x1=[double]$m.Groups[3].Value; y1=[double]$m.Groups[4].Value; t=[Net.WebUtility]::HtmlDecode($m.Groups[5].Value) } })
  $lt = Join-Path "$sp\txt" ($pdf.BaseName + ".txt")
  $sysTitle = ((Get-Content $lt | Select-String "NOT TO SCALE\s+(.+)" | ForEach-Object { ($_.Matches[0].Groups[1].Value -split "\s{3,}")[0].Trim() }) | Select-Object -Unique) -join " / "
  # strips: DI with DO, AI, AO stacked under it
  $strips = @()
  foreach ($di in ($W | Where-Object { $_.t -eq 'DI' })) {
    $f = { param($t) $W | Where-Object { $_.t -eq $t -and [math]::Abs($_.x0 - $di.x0) -lt 6 -and $_.y0 -gt $di.y0 -and $_.y0 -lt $di.y0 + 110 } | Sort-Object y0 | Select-Object -First 1 }
    $do = & $f 'DO'; $ai = & $f 'AI'; $ao = & $f 'AO'
    if ($do -and $ai -and $ao) { $strips += [pscustomobject]@{ x=$di.x0; xr=$di.x1; rows=@(@('DI',$di),@('DO',$do),@('AI',$ai),@('AO',$ao)) } }
  }
  if ($strips.Count -eq 0) { Write-Host "$sheet : no strip | $sysTitle"; continue }
  $si = 0
  foreach ($s in ($strips | Sort-Object { $_.rows[0][1].y0 }, x)) {
    $si++
    $top = $s.rows[0][1].y0
    # rotated labels above the strip
    $lab = @($W | Where-Object { ($_.y1-$_.y0) -gt 1.15*($_.x1-$_.x0) -and $_.y1 -lt $top + 2 -and $_.y0 -gt $top - 700 -and $_.x0 -gt $s.x })
    # strip right end: horizontal row line ends; approximate with right-most label
    if ($lab.Count -eq 0) { continue }
    # cluster labels into columns
    $cl = @(); foreach ($wd in ($lab | Sort-Object x0)) { $c = $cl | Where-Object { [math]::Abs($_.x0 - $wd.x0) -lt 2.5 } | Select-Object -First 1; if ($c) { $c.w += $wd; if ($wd.x1 -gt $c.x1) { $c.x1 = $wd.x1 } } else { $cl += [pscustomobject]@{ x0=$wd.x0; x1=$wd.x1; w=@($wd) } } }
    # keep only columns whose lowest word sits close above the strip (within 120pt)
    $cl = @($cl | Where-Object { (($_.w | Measure-Object y1 -Maximum).Maximum) -gt $top - 140 })
    # merge wrapped label lines (columns < 10pt apart) into one label
    $labels = @(); foreach ($c in ($cl | Sort-Object x0)) { $p = $labels | Select-Object -Last 1; if ($p -and ($c.x0 - $p.x1) -lt 3.5) { $p.w += $c.w; $p.x1 = $c.x1 } else { $labels += [pscustomobject]@{ x0=$c.x0; x1=$c.x1; w=@($c.w) } } }
    $right = ($labels | Measure-Object x1 -Maximum).Maximum + 40
    $bottom = $s.rows[3][1].y1
    # render strip band
    $rxp = [int][math]::Floor(($s.xr + 4) * $K); $ryp = [int][math]::Floor(($top - 4) * $K)
    $rwp = [int][math]::Ceiling(($right - $s.xr - 4) * $K); $rhp = [int][math]::Ceiling(($bottom - $top + 8) * $K)
    $png = Join-Path $env:TEMP "strip_$sheet`_$si"
    & $PPM -png -r $DPI -x $rxp -y $ryp -W $rwp -H $rhp -singlefile $pdf.FullName $png 2>$null
    $bmp = [Drawing.Bitmap]::FromFile("$png.png")
    $rect = New-Object Drawing.Rectangle 0,0,$bmp.Width,$bmp.Height
    $bd = $bmp.LockBits($rect, [Drawing.Imaging.ImageLockMode]::ReadOnly, [Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $stride = $bd.Stride; $buf = New-Object byte[] ($stride * $bmp.Height)
    [Runtime.InteropServices.Marshal]::Copy($bd.Scan0, $buf, 0, $buf.Length); $bmp.UnlockBits($bd)
    $ImgW = $bmp.Width; $Hd = $bmp.Height; $bmp.Dispose()
    $dark = { param($px,$py) if ($px -lt 0 -or $py -lt 0 -or $px -ge $ImgW -or $py -ge $Hd) { return $false }; $i = $py*$stride + $px*3; return (($buf[$i] + $buf[$i+1] + $buf[$i+2]) -lt 300) }
    $seen = @{}; $usedLab = @{}
    foreach ($rw in $s.rows) {
      $yc = (($rw[1].y0 + $rw[1].y1) / 2 - ($top - 4)) * $K
      # the row line may be a few px off the label centre: search band +-4px for best row
      $raw = [DotFinder]::Find($buf, $stride, $ImgW, $Hd, [int]$yc, 9); $dots = @(); for ($q=0; $q -lt $raw.Length; $q+=2) { $dots += ,@($raw[$q], $raw[$q+1]) }
      foreach ($dp in $dots) { $dx = $dp[0]; $dyPt = ($top - 4) + $dp[1] / $K
        $dotPt = $s.xr + 4 + $dx / $K
        if ($W | Where-Object { $dotPt -ge $_.x0 - 1 -and $dotPt -le $_.x1 + 1 -and $dyPt -ge $_.y0 - 1 -and $dyPt -le $_.y1 + 1 }) { continue }
        $cand = $labels | Where-Object { $_.x1 -le $dotPt + 2 -and ($dotPt - $_.x1) -lt 22 } | Sort-Object { $dotPt - $_.x1 } | Select-Object -First 1
        $name = if ($cand) { (($cand.w | Sort-Object x0, @{e={$_.y0}; Descending=$true}) | ForEach-Object t) -join ' ' } else { '?' }
        $mult = $W | Where-Object { $_.t -match '^[xX]\d+$' -and $_.x0 -gt $dotPt -and $_.x0 -lt $dotPt + 22 -and [math]::Abs((($_.y0+$_.y1)/2) - (($rw[1].y0+$rw[1].y1)/2)) -lt 9 } | Select-Object -First 1
        $key = "$name|$($rw[0])"
        if ($seen.ContainsKey($key) -and (@($seen[$key] | Where-Object { [math]::Abs($_ - $dotPt) -lt 7 }).Count -gt 0)) { continue }
        $seen[$key] = @($seen[$key]) + $dotPt
        if ($cand) { $usedLab[[string]$cand.x0] = 1 }
        $m = if ($mult) { $mult.t.Substring(1) } else { '1' }
        $res.Add(("{0}`t{1}`tS{2}`t{3}`t{4}`t{5}`t{6:N0}`t{7:N0}" -f $sheet, $sysTitle, $si, $name, $rw[0], $m, $(if($cand){$cand.x1}else{0}), $dotPt))
      }
    }
    foreach ($lb in $labels) {
      if ($usedLab.ContainsKey([string]$lb.x0)) { continue }
      $name = (($lb.w | Sort-Object x0, @{e={$_.y0}; Descending=$true}) | ForEach-Object t) -join ' '
      if ($name.Length -lt 4) { continue }
      $ty = switch -regex ($name) { 'MODULATING (COMMAND|CONTROL)|SPEED CONTROL|SETPOINT' {'AO'; break} 'COMMAND' {'DO'; break} 'SENSOR|TRANSMITTER|FEEDBACK|MEASURING|TEMPERATURE|FLOW RATE|VELOCITY|METER' {'AI'; break} 'STATUS|ALARM|SWITCH|CONTACT|POSITION|DETECTOR' {'DI'; break} default {'?'} }
      $res.Add(("{0}`t{1}`tS{2}`t{3}`t{4}`t{5}`t{6:N0}`t{7}" -f $sheet, $sysTitle, $si, $name, "$ty (inferred)", '1', $lb.x1, ''))
    }
    [IO.File]::Delete("$png.png")
  }
  Write-Host ("{0}: strips={1} rows so far={2} | {3}" -f $sheet, $strips.Count, ($res.Count-1), $sysTitle)
}
[IO.File]::WriteAllLines($Out, $res)
