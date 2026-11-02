#requires -Version 5.1
<#
.SYNOPSIS
    Report the display method or preview/restore it on a THAWED audit-only pilot.
.NOTES
    v1.1.0. Never requests a reboot or changes Deep Freeze state.
#>
[CmdletBinding()]
param(
    [ValidateSet('Report','Apply','Restore')][string]$Mode='Report',
    [string]$PersistentStateRoot='',
    [string]$Title='IT REPAIR IN PROGRESS',
    [string]$Message='This computer is temporarily unavailable. Repairs may take several minutes. Please do not turn it off. For assistance, call ext. 2181.'
)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'RepairScreen.psm1') -Force -DisableNameChecking
if ($Mode -eq 'Report') {Get-ComptonRepairEdition | Format-List *;return}
$principal=New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {throw 'Run elevated.'}
$mutex=New-Object Threading.Mutex($false,'Global\Compton_RebootRecoveryWatchdog')
$held=$false
try {
try {$held=$mutex.WaitOne(0)} catch [Threading.AbandonedMutexException] {$held=$true}
if (-not $held) {throw 'The watchdog is currently running. Retry the preview after that check finishes.'}
$dfc=@("$env:WINDIR\System32\DFC.exe","$env:WINDIR\SysWOW64\DFC.exe","${env:ProgramFiles(x86)}\Faronics\Deep Freeze\Install C-0\DFC.exe") | Where-Object {Test-Path -LiteralPath $_ -PathType Leaf} | Select-Object -First 1
if (-not $dfc) {throw 'DFC is unavailable; cannot confirm thawed state.'}
$process=Start-Process -FilePath $dfc -ArgumentList 'get /ISFROZEN' -PassThru -WindowStyle Hidden
try {
    if (-not $process.WaitForExit(10000)) {try {$process.Kill()} catch {};throw 'DFC query timed out.'}
    $process.Refresh()
    if ($process.ExitCode -ne 0) {throw 'Preview/restore requires confirmed THAWED state.'}
} finally {$process.Dispose()}
if ($PersistentStateRoot -notmatch '^\\\\[^\\]+\\[^\\]+(?:\\.*)?$') {throw 'Supply the same real persistent UNC root used by the watchdog.'}
$computerFolder=Join-Path $PersistentStateRoot $env:COMPUTERNAME
if (-not (Test-Path -LiteralPath $computerFolder -PathType Container)) {throw 'Provision the private per-computer folder first.'}
$task=Get-ScheduledTask -TaskName 'Compton. Reboot Loop Recovery Watchdog' -TaskPath '\' -ErrorAction SilentlyContinue
if ($null -ne $task -and $Mode -eq 'Apply') {
    foreach ($action in @($task.Actions)) {
        if ($null -ne $action.PSObject.Properties['Arguments'] -and [string]$action.Arguments -match '(?i)(?:^|\s)-Recover(?:\s|$)') {throw 'Use an audit-only watchdog task for this visual preview. Do not preview on an active automatic recovery.'}
    }
}
$live=Join-Path $computerFolder 'RecoveryState.json'
if ($Mode -eq 'Apply' -and (Test-Path -LiteralPath $live)) {
    $incident=Get-Content -LiteralPath $live -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($incident.Phase -notin @('Completed')) {throw 'Resolve the existing recovery incident before previewing display changes.'}
}
$previewRoot=Join-Path $computerFolder 'RepairScreenPreview'
New-Item -Path $previewRoot -ItemType Directory -Force | Out-Null
$contextFile=Join-Path $previewRoot 'Context.json'
$context=$null
if (Test-Path -LiteralPath $contextFile) {
    $context=Get-Content -LiteralPath $contextFile -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($context.ComputerName -ne $env:COMPUTERNAME) {throw 'Preview context belongs to a different computer.'}
    [void][guid]::Parse($context.IncidentId)
}
if ($Mode -eq 'Apply') {
    if ($null -eq $context -or $context.Phase -eq 'Restored') {
        $context=[pscustomobject]@{ComputerName=$env:COMPUTERNAME;IncidentId=[guid]::NewGuid().ToString();Phase='Active'}
        $context | ConvertTo-Json | Set-Content -LiteralPath $contextFile -Encoding UTF8
    }
    $folder=Join-Path $previewRoot $context.IncidentId
    Enable-ComptonRepairScreen -IncidentFolder $folder -IncidentId $context.IncidentId -ComputerName $env:COMPUTERNAME -Title $Title -Message $Message | Format-List *
    Write-Host 'Preview configured. Check the next lock/sign-in screen; current desktop users may receive a dismissible message. Restore before enabling recovery or freezing.'
    return
}
if ($null -eq $context) {throw 'No preview backup is available.'}
$folder=Join-Path $previewRoot $context.IncidentId
$result=Restore-ComptonRepairScreen -IncidentFolder $folder -IncidentId $context.IncidentId -ComputerName $env:COMPUTERNAME
$context.Phase='Restored';$context | ConvertTo-Json | Set-Content -LiteralPath $contextFile -Encoding UTF8
$result | Format-List *

} finally {if ($held) {$mutex.ReleaseMutex()};$mutex.Dispose()}
