#Check for [Name of program here]
#get AXP Process
$axp = Get-Process [ProcessName] -ErrorAction SilentlyContinue
if ($axp) {
  # try gracefully first
  $axp.CloseMainWindow()
  # kill after five seconds
  Sleep 5
  if (!$axp.HasExited) {
    $axp | Stop-Process -Force
  }
}
Remove-Variable axp

Write-Host "[ProcessName] has been killed"

#Check for Index Image Import
# get III Process
$IndImp = Get-Process [ProcessName] -ErrorAction SilentlyContinue
if ($IndImp) {
  # try gracefully first
  $IndImp.CloseMainWindow()
  # kill after five seconds
  Sleep 5
  if (!$IndImp.HasExited) {
    $IndImp | Stop-Process -Force
  }
}
Remove-Variable IndImp
sleep 5

Write-Host "[ProcessName] has been killed"

#Restart [Name of program here]

Write-Host "Restarting [Name of program here]"

Start-Process -FilePath "E:\path\to\name\of\program\here.exe"
Write-Host "[Name of program here] has been restarted"
pause
