param([string]$Xlsm,[string]$OutDir,[int]$MaxRows=25)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z=[System.IO.Compression.ZipFile]::OpenRead($Xlsm)
function GetXml($name){ $e=$z.GetEntry($name); if(-not $e){return $null}; $s=$e.Open(); $r=New-Object System.IO.StreamReader($s); $t=$r.ReadToEnd(); $r.Close(); [xml]$t }
$wb=GetXml "xl/workbook.xml"
$rels=GetXml "xl/_rels/workbook.xml.rels"
$ss=GetXml "xl/sharedStrings.xml"
$strs=@()
if($ss){ foreach($si in $ss.sst.si){ $strs+=(($si.SelectNodes('.//*[local-name()="t"]') | ForEach-Object {$_.InnerText}) -join '') } }
New-Item -ItemType Directory -Force $OutDir | Out-Null
foreach($sh in $wb.workbook.sheets.sheet){
  $rid=$sh.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $tgt=($rels.Relationships.Relationship | Where-Object {$_.Id -eq $rid}).Target
  $tgt=$tgt -replace '^/xl/','' -replace '^\.\./',''
  $x=GetXml ("xl/"+$tgt); if(-not $x){continue}
  $safe=($sh.name -replace '[^A-Za-z0-9_\- ]','_')
  $lines=@()
  $n=0
  foreach($row in $x.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]')){
    $n++; if($n -gt $MaxRows){break}
    $cells=@()
    foreach($c in $row.SelectNodes('*[local-name()="c"]')){
      $v=$c.SelectSingleNode('*[local-name()="v"]')
      $isn=$c.SelectSingleNode('*[local-name()="is"]')
      $val=""
      if($c.t -eq "s" -and $v){ $i=[int]$v.InnerText; if($i -lt $strs.Count){$val=$strs[$i]} }
      elseif($isn){ $val=(($isn.SelectNodes('.//*[local-name()="t"]')|ForEach-Object{$_.InnerText}) -join '') }
      elseif($v){ $val=$v.InnerText }
      if($val -ne ""){ $cells+=($c.r+"="+($val -replace '\s+',' ')) }
    }
    if($cells.Count){ $lines += ("r"+$row.r+" | "+($cells -join " | ")) }
  }
  Set-Content -Encoding utf8 (Join-Path $OutDir ($safe+".txt")) ($lines -join "`n")
}
$z.Dispose()
Write-Output "done"
