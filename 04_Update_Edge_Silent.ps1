#requires -Version 5.1
<#
.SYNOPSIS
    Retirement stub for the deprecated standalone Microsoft Edge MSI updater.

.DESCRIPTION
    Performs no Edge download, installation, process termination, or policy
    modification. It records that Script 04 has been intentionally retired.

.NOTES
    ScriptName:    04_Update_Edge_Silent.ps1
    ScriptVersion: 2.0.0
    Status:        Deprecated
#>

[CmdletBinding()]
param(
    [switch]$KillEdgeProcesses,
    [string]$LogDirectory = 'C:\Logs'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptName    = '04_Update_Edge_Silent.ps1'
$scriptVersion = '2.0.0'
$startTime     = Get-Date
$logPath       = Join-Path $LogDirectory '04_Update_Edge_Silent.log'
$ndjsonPath    = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$latestPath    = Join-Path $LogDirectory '04_Update_Edge_Silent.latest.json'

if (-not (Test-Path -LiteralPath $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$endTime = Get-Date
# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

$event = [ordered]@{
    '@timestamp'    = $endTime.ToUniversalTime().ToString('o')
    EventType       = 'maintenance.deprecated'
    SchemaVersion   = '1.0'
    ComputerName    = $env:COMPUTERNAME
    Domain          = $env:USERDOMAIN
    ScriptName      = $scriptName
    ScriptVersion   = $scriptVersion
    Status          = 'Deprecated'
    ExitCode        = 0
    StartTime       = $startTime.ToUniversalTime().ToString('o')
    EndTime         = $endTime.ToUniversalTime().ToString('o')
    DurationSeconds = [math]::Round(($endTime - $startTime).TotalSeconds, 3)
    WarningCount    = 0
    ErrorCount      = 0
    Replacement     = 'Microsoft Edge Update servicing; health visibility is collected by 14_Endpoint_Health_Inventory.ps1'
    Message         = 'The standalone Edge MSI updater is retired. No Edge processes were stopped and no system changes were made.'
    LegacySwitchReceived = [bool]$KillEdgeProcesses
}

$line = '{0} [INFO] Script 04 is deprecated. No Edge installation or process termination was performed.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8

$compact = $event | ConvertTo-Json -Depth 6 -Compress
$pretty  = $event | ConvertTo-Json -Depth 6
Write-MaintenanceTelemetryLine -Path $ndjsonPath -JsonLine $compact

$tempPath = '{0}.{1}.tmp' -f $latestPath, ([guid]::NewGuid().Guid)
[IO.File]::WriteAllText($tempPath, $pretty, (New-Object Text.UTF8Encoding($false)))
Move-Item -LiteralPath $tempPath -Destination $latestPath -Force

exit 0
