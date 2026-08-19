# ============================================================================
# ScriptVersion: 1.0.0
# DateUpdated: 2026-08-17
# ============================================================================

#requires -Version 5.1
#requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TargetScript,
    [string]$TargetArguments = '',
    [string]$CorrelationId,
    [switch]$BypassMaintenanceWindow
)
Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'
$root='C:\Scripts'; $logRoot='C:\Logs'
Import-Module (Join-Path $root 'Maintenance.Framework.psm1') -Force -ErrorAction Stop
$policy=Get-MaintenancePolicy -Path (Join-Path $root 'Maintenance.Policy.json')
$scriptName=[IO.Path]::GetFileName($TargetScript)
$correlation=Get-MaintenanceCorrelationId -CorrelationId $CorrelationId
$env:COMPTON_MAINTENANCE_CORRELATION_ID=$correlation
$runId=[guid]::NewGuid().Guid; $started=Get-Date; $exitCode=3; $status='CriticalFailure'; $lock=$null; $failure=$null
$logPath=Join-Path $logRoot ('Launcher-{0}.log' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
Rotate-MaintenanceLog -Path $logPath -MaxSizeMB ([int]$policy.LogMaxSizeMB) -RetentionDays ([int]$policy.LogRetentionDays)
function Write-LauncherLog([string]$Message,[string]$Level='INFO') {
    $line='[{0}] [{1}] [CorrelationId:{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$Level,$correlation,$Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}
try {
    Write-LauncherLog "Starting $scriptName. RunId=$runId"
    Write-MaintenanceEvent -EventId 1000 -Message "Starting $scriptName. RunId=$runId CorrelationId=$correlation" | Out-Null
    if(-not (Test-Path -LiteralPath $TargetScript -PathType Leaf)){ throw "Target script is missing: $TargetScript" }
    if(-not $BypassMaintenanceWindow) {
        $window=Test-MaintenanceWindow -ScriptName $scriptName -Policy $policy
        if(-not $window.Allowed){ $status='OutsideMaintenanceWindow'; $exitCode=1; Write-LauncherLog $window.Reason 'WARN'; throw [System.OperationCanceledException]::new($window.Reason) }
    }
    $dependencies=Test-MaintenanceDependencies -ScriptName $scriptName -Policy $policy
    foreach($d in $dependencies.Results){ Write-LauncherLog ("Dependency {0}:{1} Passed={2} Required={3}" -f $d.Type,$d.Value,$d.Passed,$d.Required) $(if($d.Passed){'INFO'}else{'WARN'}) }
    if(-not $dependencies.Passed){ $status='DependencyFailure'; $exitCode=2; throw "Required dependency validation failed." }
    $lock=Enter-MaintenanceExecutionLock -TimeoutSeconds ([int]$policy.ExecutionLockTimeoutSeconds)
    Write-LauncherLog 'Global execution lock acquired.'
    $ps=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $args='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $TargetScript
    if(-not [string]::IsNullOrWhiteSpace($TargetArguments)){ $args+=' '+$TargetArguments }
    $process=Start-Process -FilePath $ps -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
    $exitCode=[int]$process.ExitCode
    $status=if($exitCode -eq 0){'Success'}elseif($exitCode -eq 1){'SuccessWithWarnings'}else{'Failed'}
    Write-LauncherLog "Target completed. ExitCode=$exitCode Status=$status" $(if($exitCode -le 1){'OK'}else{'ERROR'})
} catch [System.OperationCanceledException] {
    $failure=$_.Exception.Message
    if($status -eq 'CriticalFailure'){ $status='OutsideMaintenanceWindow'; $exitCode=1 }
} catch {
    $failure=$_.Exception.Message
    if($exitCode -eq 3){ $status='Failed' }
    Write-LauncherLog $failure 'ERROR'
} finally {
    if($lock){ Exit-MaintenanceExecutionLock -Lock $lock; Write-LauncherLog 'Global execution lock released.' }
    $ended=Get-Date
    $record=[ordered]@{
        SchemaVersion=1;ComputerName=$env:COMPUTERNAME;ScriptName=$scriptName;RunId=$runId;CorrelationId=$correlation
        Status=$status;ExitCode=$exitCode;StartTime=$started.ToString('o');EndTime=$ended.ToString('o')
        DurationSeconds=[math]::Round(($ended-$started).TotalSeconds,3);FailureMessage=$failure
        FrameworkVersion=(Get-MaintenanceConfiguration).FrameworkVersion;PolicyVersion=[string]$policy.PolicyVersion
        LastCheckIn=$ended.ToUniversalTime().ToString('o')
    }
    try { Write-MaintenanceTelemetry -Record $record } catch {}
    try { Publish-MaintenanceFleetStatus -Record $record -StatusRoots @($policy.FleetStatusRoots) | Out-Null } catch {}
    $eventType=if($exitCode -eq 0){'Information'}elseif($exitCode -eq 1){'Warning'}else{'Error'}
    $eventId=if($exitCode -eq 0){1001}elseif($status -eq 'OutsideMaintenanceWindow'){1700}elseif($status -eq 'DependencyFailure'){1800}else{1900}
    Write-MaintenanceEvent -EventId $eventId -EntryType $eventType -Message "$scriptName completed. Status=$status ExitCode=$exitCode RunId=$runId CorrelationId=$correlation Failure=$failure" | Out-Null
}
exit $exitCode
﻿#requires -Version 5.1
#requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TargetScript,
    [string]$TargetArguments = '',
    [string]$CorrelationId,
    [switch]$BypassMaintenanceWindow
)
Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'
$root='C:\Scripts'; $logRoot='C:\Logs'
Import-Module (Join-Path $root 'Maintenance.Framework.psm1') -Force -ErrorAction Stop
$policy=Get-MaintenancePolicy -Path (Join-Path $root 'Maintenance.Policy.json')
$scriptName=[IO.Path]::GetFileName($TargetScript)
$correlation=Get-MaintenanceCorrelationId -CorrelationId $CorrelationId
$env:COMPTON_MAINTENANCE_CORRELATION_ID=$correlation
$runId=[guid]::NewGuid().Guid; $started=Get-Date; $exitCode=3; $status='CriticalFailure'; $lock=$null; $failure=$null
$logPath=Join-Path $logRoot ('Launcher-{0}.log' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
Rotate-MaintenanceLog -Path $logPath -MaxSizeMB ([int]$policy.LogMaxSizeMB) -RetentionDays ([int]$policy.LogRetentionDays)
function Write-LauncherLog([string]$Message,[string]$Level='INFO') {
    $line='[{0}] [{1}] [CorrelationId:{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$Level,$correlation,$Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}
try {
    Write-LauncherLog "Starting $scriptName. RunId=$runId"
    Write-MaintenanceEvent -EventId 1000 -Message "Starting $scriptName. RunId=$runId CorrelationId=$correlation" | Out-Null
    if(-not (Test-Path -LiteralPath $TargetScript -PathType Leaf)){ throw "Target script is missing: $TargetScript" }
    if(-not $BypassMaintenanceWindow) {
        $window=Test-MaintenanceWindow -ScriptName $scriptName -Policy $policy
        if(-not $window.Allowed){ $status='OutsideMaintenanceWindow'; $exitCode=1; Write-LauncherLog $window.Reason 'WARN'; throw [System.OperationCanceledException]::new($window.Reason) }
    }
    $dependencies=Test-MaintenanceDependencies -ScriptName $scriptName -Policy $policy
    foreach($d in $dependencies.Results){ Write-LauncherLog ("Dependency {0}:{1} Passed={2} Required={3}" -f $d.Type,$d.Value,$d.Passed,$d.Required) $(if($d.Passed){'INFO'}else{'WARN'}) }
    if(-not $dependencies.Passed){ $status='DependencyFailure'; $exitCode=2; throw "Required dependency validation failed." }
    $lock=Enter-MaintenanceExecutionLock -TimeoutSeconds ([int]$policy.ExecutionLockTimeoutSeconds)
    Write-LauncherLog 'Global execution lock acquired.'
    $ps=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $args='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $TargetScript
    if(-not [string]::IsNullOrWhiteSpace($TargetArguments)){ $args+=' '+$TargetArguments }
    $process=Start-Process -FilePath $ps -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
    $exitCode=[int]$process.ExitCode
    $status=if($exitCode -eq 0){'Success'}elseif($exitCode -eq 1){'SuccessWithWarnings'}else{'Failed'}
    Write-LauncherLog "Target completed. ExitCode=$exitCode Status=$status" $(if($exitCode -le 1){'OK'}else{'ERROR'})
} catch [System.OperationCanceledException] {
    $failure=$_.Exception.Message
    if($status -eq 'CriticalFailure'){ $status='OutsideMaintenanceWindow'; $exitCode=1 }
} catch {
    $failure=$_.Exception.Message
    if($exitCode -eq 3){ $status='Failed' }
    Write-LauncherLog $failure 'ERROR'
} finally {
    if($lock){ Exit-MaintenanceExecutionLock -Lock $lock; Write-LauncherLog 'Global execution lock released.' }
    $ended=Get-Date
    $record=[ordered]@{
        SchemaVersion=1;ComputerName=$env:COMPUTERNAME;ScriptName=$scriptName;RunId=$runId;CorrelationId=$correlation
        Status=$status;ExitCode=$exitCode;StartTime=$started.ToString('o');EndTime=$ended.ToString('o')
        DurationSeconds=[math]::Round(($ended-$started).TotalSeconds,3);FailureMessage=$failure
        FrameworkVersion=(Get-MaintenanceConfiguration).FrameworkVersion;PolicyVersion=[string]$policy.PolicyVersion
        LastCheckIn=$ended.ToUniversalTime().ToString('o')
    }
    try { Write-MaintenanceTelemetry -Record $record } catch {}
    try { Publish-MaintenanceFleetStatus -Record $record -StatusRoots @($policy.FleetStatusRoots) | Out-Null } catch {}
    $eventType=if($exitCode -eq 0){'Information'}elseif($exitCode -eq 1){'Warning'}else{'Error'}
    $eventId=if($exitCode -eq 0){1001}elseif($status -eq 'OutsideMaintenanceWindow'){1700}elseif($status -eq 'DependencyFailure'){1800}else{1900}
    Write-MaintenanceEvent -EventId $eventId -EntryType $eventType -Message "$scriptName completed. Status=$status ExitCode=$exitCode RunId=$runId CorrelationId=$correlation Failure=$failure" | Out-Null
}
exit $exitCode
