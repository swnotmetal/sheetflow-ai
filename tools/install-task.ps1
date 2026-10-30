# Registers the two-minute pass, so the owner never runs a command.
#
# The whole point of the local-Excel direction is that the sheet is the only interface. A shop owner
# will not open a terminal after lunch to "resume the run", so the pass has to happen on its own:
# every two minutes, invisible, for as long as the machine is on.
#
#   powershell -ExecutionPolicy Bypass -File tools\install-task.ps1
#   powershell -ExecutionPolicy Bypass -File tools\install-task.ps1 -Remove
#
# It is the *same* executable the owner double-clicks, with one override: the review timeout is 0,
# so a pass never sits waiting. Mail that is drafted and left unanswered stays suspended on disk and
# is picked up by a later pass — which is exactly what the ledger is for.

param(
  [string]$Folder = $PSScriptRoot,
  [string]$TaskName = 'sheetflow',
  [int]$EveryMinutes = 2,
  [switch]$Remove
)

$ErrorActionPreference = 'Stop'

$exe = Join-Path (Resolve-Path (Join-Path $Folder '..')).Path 'dist\sheetflow\sheetflow.exe'
if (-not (Test-Path $exe) -and -not $Remove) {
  throw "No executable at $exe. Run 'npm run build' first."
}

if ($Remove) {
  if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed the scheduled task '$TaskName'."
  } else {
    Write-Host "No scheduled task named '$TaskName'."
  }
  return
}

$action = New-ScheduledTaskAction -Execute $exe -Argument 'inbox --no-wait' -WorkingDirectory (Split-Path $exe)
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
  -RepetitionInterval (New-TimeSpan -Minutes $EveryMinutes)

$settings = New-ScheduledTaskSettingsSet `
  -MultipleInstances IgnoreNew `
  -StartWhenAvailable `
  -DontStopIfGoingOnBatteries `
  -AllowStartIfOnBatteries `
  -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

# No interactive logon: the pass must run whether or not anyone is signed in to watch it.
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
  -Settings $settings -Principal $principal `
  -Description 'sheetflow: drafts order mails into the shop workbook, every few minutes.' `
  -Force | Out-Null

Write-Host "Registered '$TaskName': $exe inbox, every $EveryMinutes minute(s), no window."
Write-Host "Remove it with:  powershell -File tools\install-task.ps1 -Remove"
