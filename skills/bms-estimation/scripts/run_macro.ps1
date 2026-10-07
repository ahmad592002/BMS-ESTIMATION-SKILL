param([string[]]$Macros, [string]$Sheet)
# Run workbook macros in order, saving before and after each one. Waits for Excel to be ready.
if (-not $env:BMS_WB) { throw 'Set $env:BMS_WB to the full path of the estimation workbook (open it in Excel first with Start-Process)' }
$ErrorActionPreference = "Stop"
$WBPATH = $env:BMS_WB
function Wait-Ready($app) { for ($i = 0; $i -lt 600; $i++) { try { if ($app.Ready) { return } } catch { }; Start-Sleep -Milliseconds 500 } }
function Retry([scriptblock]$sb) { for ($i = 0; $i -lt 60; $i++) { try { return (& $sb) } catch { if ($_.Exception.HResult -eq -2146777998 -or $_.Exception.Message -match '0x800AC472|busy|rejected') { Start-Sleep -Seconds 2 } else { throw } } }; throw "Excel stayed busy" }
$wb = Retry { [Runtime.InteropServices.Marshal]::BindToMoniker($WBPATH) }
$x = $wb.Application
Wait-Ready $x
Retry { $wb.Activate() } | Out-Null; if ($Sheet) { Retry { $wb.Worksheets.Item($Sheet).Activate() } | Out-Null }; if ((Retry { $x.ActiveWorkbook.FullName }) -ne $wb.FullName) { throw "active workbook is not $($wb.Name) - stop" }; "active=" + (Retry { $x.ActiveWorkbook.Name }) + " / " + (Retry { $wb.ActiveSheet.Name })
foreach ($m in $Macros) {
  Wait-Ready $x; Retry { $wb.Save() } | Out-Null; Retry { $wb.Activate() } | Out-Null; if ((Retry { $x.ActiveWorkbook.FullName }) -ne $wb.FullName) { throw "active workbook changed - stop before $m" }
  $t0 = Get-Date
  try { Retry { [void]$x.Run("'" + $wb.Name + "'!" + $m) } | Out-Null } catch { if ($_.Exception.Message -match '0x800A9C68') { "  $m failed once (0x800A9C68) - retrying"; Wait-Ready $x; Start-Sleep -Seconds 3; Retry { [void]$x.Run("'" + $wb.Name + "'!" + $m) } | Out-Null } else { throw } }
  "{0} done in {1}s" -f $m, [int]((Get-Date) - $t0).TotalSeconds
  Wait-Ready $x; Retry { $wb.Save() } | Out-Null
}
"saved=" + (Retry { $wb.Saved })



