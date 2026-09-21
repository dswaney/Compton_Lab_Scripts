#requires -Version 5.1
#requires -RunAsAdministrator
# =====================================================================
# ScriptName:    Invoke-MaintenanceScript.ps1
# ScriptVersion: 2.0.0
# LastUpdated:   2026-09-21
# Purpose:       Executes only locally allowlisted maintenance actions with
#                fixed script mappings and fixed arguments.
# =====================================================================

<#
.SYNOPSIS
    Secure launcher for approved Compton maintenance actions.

.DESCRIPTION
    Accepts an ActionId rather than a script path. Every ActionId maps locally
    to one fixed script name and a fixed argument list. A guarded legacy
    parameter set temporarily translates only exact C:\Scripts paths and exact
    approved argument combinations; the supplied path is never executed.

    Remediation metadata is separately validated and is passed only to the
    Microsoft Edge Update remediation action. The launcher enforces the target
    script maintenance policy, dependency checks, and the global execution lock.

.NOTES
    Run Register-Tasks_SYSTEM.ps1 after deploying this version so existing
    scheduled tasks are rewritten to use ActionId. Legacy translation exists
    only to prevent a failed reconciliation from disabling an existing task.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ParameterSetName = 'Action')]
    [ValidateSet(
        'UpdateScripts',
        'EnableWindowsUpdateServices',
        'RemoveUserProfiles',
        'SundayLabApplicationMaintenance',
        'WeekendAppsUpdate',
        'VendorDriverUpdate',
        'WindowsUpdatePass',
        'ForceRebootInstallUpdates',
        'DisableWindowsUpdateServices',
        'SystemRepairWeekly',
        'EndpointHealthInventory',
        'SyncSystemTime',
        'MicrosoftEdgeUpdateRepair'
    )]
    [string]$ActionId,

    [Parameter(Mandatory, ParameterSetName = 'Legacy')]
    [string]$TargetScript,

    [Parameter(ParameterSetName = 'Legacy')]
    [string]$TargetArguments = '',

    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$')]
    [string]$CorrelationId,

    [ValidatePattern('^[A-Fa-f0-9]{64}$')]
    [string]$FindingFingerprint,

    [ValidateRange(1, 5)]
    [int]$AttemptNumber = 1
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:ScriptName = 'Invoke-MaintenanceScript.ps1'
$script:ScriptVersion = '2.0.0'
$script:Root = 'C:\Scripts'
$script:LogRoot = 'C:\Logs'
$script:RunId = [guid]::NewGuid().Guid
$script:StartTime = Get-Date
$script:ExitCode = 3
$script:Status = 'CriticalFailure'
$script:FailureMessage = $null
$script:Lock = $null
$script:LogSession = $null
$script:WorkingLogPath = $null
$script:PublishedLogPath = $null
$script:TargetScriptName = $null
$script:CorrelationId = $null
$script:LegacyInvocation = [bool]($PSCmdlet.ParameterSetName -eq 'Legacy')

$frameworkPath = Join-Path $script:Root 'Maintenance.Framework.psm1'
$policyPath = Join-Path $script:Root 'Maintenance.Policy.json'
Import-Module -Name $frameworkPath -Force -DisableNameChecking -ErrorAction Stop
$policy = Get-MaintenancePolicy -Path $policyPath
$maintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot $script:Root -LogRoot $script:LogRoot
$script:CorrelationId = Get-MaintenanceCorrelationId -CorrelationId $CorrelationId
$env:COMPTON_MAINTENANCE_CORRELATION_ID = $script:CorrelationId

# This is the complete execution allowlist. Script names are leaf names only and
# are resolved beneath C:\Scripts. No caller-controlled paths enter this table.
$actionCatalog = @{
    UpdateScripts = [pscustomobject]@{
        ScriptName = '00_Update-Scripts-FromShare.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    EnableWindowsUpdateServices = [pscustomobject]@{
        ScriptName = '01_Enable_Windows_Update_Services.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    RemoveUserProfiles = [pscustomobject]@{
        ScriptName = '02_Remove_User_Profiles.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    SundayLabApplicationMaintenance = [pscustomobject]@{
        ScriptName = '04_Sunday_Lab_Application_Maintenance.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    WeekendAppsUpdate = [pscustomobject]@{
        ScriptName = '03_Weekend_Apps_Update.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    VendorDriverUpdate = [pscustomobject]@{
        ScriptName = '05_Weekend_HP_Drivers_Update.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    WindowsUpdatePass = [pscustomobject]@{
        ScriptName = '06_Weekend_Windows_Updates.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    ForceRebootInstallUpdates = [pscustomobject]@{
        ScriptName = '07_Force_Reboot_Install_Updates.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    DisableWindowsUpdateServices = [pscustomobject]@{
        ScriptName = '09_Disable_Windows_Update_Services.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    SystemRepairWeekly = [pscustomobject]@{
        ScriptName = '08_System_Repair.ps1'
        FixedArguments = [string[]]@('-AllowCopilotRemoval')
        Kind = 'ScheduledMaintenance'
    }
    EndpointHealthInventory = [pscustomobject]@{
        ScriptName = '14_Endpoint_Health_Inventory.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    SyncSystemTime = [pscustomobject]@{
        ScriptName = '10_Sync_System_Time.ps1'
        FixedArguments = [string[]]@()
        Kind = 'ScheduledMaintenance'
    }
    MicrosoftEdgeUpdateRepair = [pscustomobject]@{
        ScriptName = 'Repair-MicrosoftEdgeUpdate.ps1'
        FixedArguments = [string[]]@()
        Kind = 'Remediation'
    }
}

function Resolve-LegacyActionId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SuppliedPath,
        [string]$SuppliedArguments = ''
    )

    if (-not [IO.Path]::IsPathRooted($SuppliedPath)) {
        throw 'Legacy target must be an absolute path beneath C:\Scripts.'
    }

    $resolvedRoot = [IO.Path]::GetFullPath($script:Root).TrimEnd('\')
    $resolvedSuppliedPath = [IO.Path]::GetFullPath($SuppliedPath)
    $suppliedParent = [IO.Path]::GetDirectoryName($resolvedSuppliedPath).TrimEnd('\')
    if ($suppliedParent -ine $resolvedRoot) {
        throw "Legacy target is outside the approved script root: $resolvedSuppliedPath"
    }

    $scriptName = [IO.Path]::GetFileName($resolvedSuppliedPath)
    $normalizedArguments = ($SuppliedArguments -replace '\s+', ' ').Trim()
    $resolvedActionId = switch ($scriptName) {
        '00_Update-Scripts-FromShare.ps1'           { 'UpdateScripts' }
        '01_Enable_Windows_Update_Services.ps1'     { 'EnableWindowsUpdateServices' }
        '02_Remove_User_Profiles.ps1'               { 'RemoveUserProfiles' }
        '03_Weekend_Apps_Update.ps1'                { 'WeekendAppsUpdate' }
        '04_Sunday_Lab_Application_Maintenance.ps1' { 'SundayLabApplicationMaintenance' }
        '05_Weekend_HP_Drivers_Update.ps1'          { 'VendorDriverUpdate' }
        '06_Weekend_Windows_Updates.ps1'            { 'WindowsUpdatePass' }
        '07_Force_Reboot_Install_Updates.ps1'       { 'ForceRebootInstallUpdates' }
        '08_System_Repair.ps1'                      { 'SystemRepairWeekly' }
        '09_Disable_Windows_Update_Services.ps1'    { 'DisableWindowsUpdateServices' }
        '10_Sync_System_Time.ps1'                   { 'SyncSystemTime' }
        '14_Endpoint_Health_Inventory.ps1'          { 'EndpointHealthInventory' }
        default { throw "Legacy script is not allowlisted: $scriptName" }
    }

    if ($scriptName -eq '08_System_Repair.ps1') {
        if ($normalizedArguments -ne '-AllowCopilotRemoval') {
            throw "Legacy SystemRepairWeekly requires the fixed argument '-AllowCopilotRemoval'."
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($normalizedArguments)) {
        throw "Legacy arguments are not approved for $scriptName."
    }

    return $resolvedActionId
}

if ($script:LegacyInvocation) {
    $ActionId = Resolve-LegacyActionId -SuppliedPath $TargetScript -SuppliedArguments $TargetArguments
}

function Write-LauncherLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $line = '[{0}] [{1}] [{2}] [CorrelationId:{3}] [ActionId:{4}] {5}' -f `
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),
        $env:COMPUTERNAME,
        $Level,
        $script:CorrelationId,
        $ActionId,
        $Message

    if ($script:WorkingLogPath) {
        Add-Content -LiteralPath $script:WorkingLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    }

    $color = switch ($Level) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host $line -ForegroundColor $color
}

function Get-AllowedAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Id)

    if (-not $actionCatalog.ContainsKey($Id)) {
        throw "ActionId is not allowlisted: $Id"
    }

    $definition = $actionCatalog[$Id]
    $leafName = [IO.Path]::GetFileName([string]$definition.ScriptName)
    if ($leafName -ne [string]$definition.ScriptName) {
        throw "Invalid internal allowlist entry for $Id. ScriptName must be a leaf name."
    }

    return $definition
}

function Get-ChildArguments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Definition,
        [Parameter(Mandatory)][string]$TargetPath
    )

    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($value in @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$TargetPath)) {
        $arguments.Add([string]$value) | Out-Null
    }
    foreach ($value in @($Definition.FixedArguments)) {
        $arguments.Add([string]$value) | Out-Null
    }

    if ([string]$Definition.Kind -eq 'Remediation') {
        if ([string]::IsNullOrWhiteSpace($FindingFingerprint)) {
            throw "FindingFingerprint is required for remediation ActionId $ActionId."
        }

        foreach ($value in @(
            '-CorrelationId', $script:CorrelationId,
            '-FindingFingerprint', $FindingFingerprint.ToLowerInvariant(),
            '-AttemptNumber', [string]$AttemptNumber
        )) {
            $arguments.Add([string]$value) | Out-Null
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($FindingFingerprint)) {
        throw 'FindingFingerprint can be supplied only for an allowlisted remediation action.'
    }

    return [string[]]$arguments.ToArray()
}

try {
    $definition = Get-AllowedAction -Id $ActionId
    $script:TargetScriptName = [string]$definition.ScriptName
    $targetPath = Join-Path $script:Root $script:TargetScriptName
    $resolvedRoot = [IO.Path]::GetFullPath($script:Root).TrimEnd('\')
    $resolvedTarget = [IO.Path]::GetFullPath($targetPath)
    $expectedParent = [IO.Path]::GetDirectoryName($resolvedTarget).TrimEnd('\')
    if ($expectedParent -ine $resolvedRoot) {
        throw "Resolved target escaped the approved script root: $resolvedTarget"
    }

    $script:LogSession = New-MaintenanceStagedLog `
        -ScriptName ("Launcher-{0}.ps1" -f $ActionId) `
        -LogRoot $script:LogRoot `
        -StagingRoot $maintenanceConfig.LogStagingRoot `
        -ComputerName $env:COMPUTERNAME `
        -Timestamp $script:StartTime
    $script:WorkingLogPath = [string]$script:LogSession.WorkingPath
    $script:PublishedLogPath = [string]$script:LogSession.PublishedPath

    Write-LauncherLog "Starting allowlisted action. RunId=$($script:RunId); Target=$($script:TargetScriptName); Kind=$($definition.Kind)." 'INFO'
    if ($script:LegacyInvocation) {
        Write-LauncherLog 'Accepted a legacy scheduled-task invocation after exact allowlist translation; run Register-Tasks_SYSTEM.ps1 to migrate this task to ActionId.' 'WARN'
    }
    Write-MaintenanceEvent -EventId 1000 -Message ("Starting ActionId={0}; Target={1}; RunId={2}; CorrelationId={3}" -f `
        $ActionId, $script:TargetScriptName, $script:RunId, $script:CorrelationId) | Out-Null

    if (-not (Test-Path -LiteralPath $resolvedTarget -PathType Leaf)) {
        throw "Allowlisted target script is missing: $resolvedTarget"
    }

    $window = Test-MaintenanceWindow -ScriptName $script:TargetScriptName -Policy $policy
    if (-not $window.Allowed) {
        $script:Status = 'OutsideMaintenanceWindow'
        $script:ExitCode = 20
        Write-LauncherLog $window.Reason 'WARN'
        throw [System.OperationCanceledException]::new($window.Reason)
    }

    $dependencies = Test-MaintenanceDependencies -ScriptName $script:TargetScriptName -Policy $policy
    foreach ($dependency in @($dependencies.Results)) {
        Write-LauncherLog ("Dependency {0}:{1}; Passed={2}; Required={3}" -f `
            $dependency.Type, $dependency.Value, $dependency.Passed, $dependency.Required) `
            $(if ($dependency.Passed) { 'INFO' } else { 'WARN' })
    }
    if (-not $dependencies.Passed) {
        $script:Status = 'DependencyFailure'
        $script:ExitCode = 2
        throw 'Required dependency validation failed.'
    }

    $childArguments = Get-ChildArguments -Definition $definition -TargetPath $resolvedTarget
    $script:Lock = Enter-MaintenanceExecutionLock -TimeoutSeconds ([int]$policy.ExecutionLockTimeoutSeconds)
    Write-LauncherLog 'Global maintenance execution lock acquired.' 'INFO'

    $powerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    & $powerShellExe @childArguments
    $script:ExitCode = [int]$LASTEXITCODE
    $script:Status = switch ($script:ExitCode) {
        0    { 'Success' }
        1    { 'SuccessWithWarnings' }
        20   { 'Blocked' }
        3010 { 'RebootRequired' }
        default { 'Failed' }
    }

    Write-LauncherLog "Target completed. ExitCode=$($script:ExitCode); Status=$($script:Status)." `
        $(if ($script:ExitCode -in @(0,1,3010)) { 'OK' } elseif ($script:ExitCode -eq 20) { 'WARN' } else { 'ERROR' })
}
catch [System.OperationCanceledException] {
    $script:FailureMessage = $_.Exception.Message
    if ($script:Status -eq 'CriticalFailure') {
        $script:Status = 'OutsideMaintenanceWindow'
        $script:ExitCode = 20
    }
}
catch {
    $script:FailureMessage = $_.Exception.Message
    if ($script:ExitCode -eq 3) {
        $script:Status = 'Failed'
    }
    try { Write-LauncherLog $script:FailureMessage 'ERROR' } catch {}
}
finally {
    if ($script:Lock) {
        Exit-MaintenanceExecutionLock -Lock $script:Lock
        try { Write-LauncherLog 'Global maintenance execution lock released.' 'INFO' } catch {}
    }

    $ended = Get-Date
    $record = [pscustomobject][ordered]@{
        '@timestamp'      = $ended.ToUniversalTime().ToString('o')
        EventType         = 'maintenance.launcher'
        SchemaVersion     = '2.0'
        ComputerName      = $env:COMPUTERNAME
        ScriptName        = $script:ScriptName
        ScriptVersion     = $script:ScriptVersion
        ActionId          = $ActionId
        ActionKind        = if ($actionCatalog.ContainsKey($ActionId)) { [string]$actionCatalog[$ActionId].Kind } else { $null }
        TargetScriptName  = $script:TargetScriptName
        FindingFingerprint = if ($FindingFingerprint) { $FindingFingerprint.ToLowerInvariant() } else { $null }
        AttemptNumber     = if ($actionCatalog.ContainsKey($ActionId) -and [string]$actionCatalog[$ActionId].Kind -eq 'Remediation') { $AttemptNumber } else { $null }
        LegacyInvocation  = $script:LegacyInvocation
        RunId             = $script:RunId
        CorrelationId     = $script:CorrelationId
        Status            = $script:Status
        ExitCode          = $script:ExitCode
        StartTime         = $script:StartTime.ToUniversalTime().ToString('o')
        EndTime           = $ended.ToUniversalTime().ToString('o')
        DurationSeconds   = [math]::Round(($ended - $script:StartTime).TotalSeconds, 3)
        FailureMessage    = $script:FailureMessage
        FrameworkVersion  = (Get-MaintenanceConfiguration).FrameworkVersion
        PolicyVersion     = [string]$policy.PolicyVersion
        TextLogPath       = $script:PublishedLogPath
        LastCheckIn       = $ended.ToUniversalTime().ToString('o')
    }

    try { Write-MaintenanceTelemetry -Record $record } catch {}
    try { Publish-MaintenanceFleetStatus -Record $record -StatusRoots @($policy.FleetStatusRoots) | Out-Null } catch {}

    $eventType = if ($script:ExitCode -eq 0) { 'Information' } elseif ($script:ExitCode -in @(1,20,3010)) { 'Warning' } else { 'Error' }
    $eventId = if ($script:ExitCode -eq 0) { 1001 } elseif ($script:Status -eq 'OutsideMaintenanceWindow') { 1700 } elseif ($script:Status -eq 'DependencyFailure') { 1800 } else { 1900 }
    try {
        Write-MaintenanceEvent -EventId $eventId -EntryType $eventType `
            -Message ("ActionId={0} completed. Target={1}; Status={2}; ExitCode={3}; RunId={4}; CorrelationId={5}; Failure={6}" -f `
                $ActionId, $script:TargetScriptName, $script:Status, $script:ExitCode,
                $script:RunId, $script:CorrelationId, $script:FailureMessage) | Out-Null
    }
    catch {}

    if ($script:LogSession) {
        try {
            Write-LauncherLog ("Finalizing launcher. Status={0}; ExitCode={1}." -f $script:Status, $script:ExitCode) `
                $(if ($script:ExitCode -in @(0,1,3010)) { 'OK' } elseif ($script:ExitCode -eq 20) { 'WARN' } else { 'ERROR' })
            $publishResult = Publish-MaintenanceLog -LogSession $script:LogSession
            if (-not $publishResult.Published) {
                Write-Warning "Completed launcher log remains in staging: $($publishResult.Path)"
            }
        }
        catch {
            Write-Warning "Failed to publish the completed launcher log: $($_.Exception.Message)"
        }
    }
}

exit $script:ExitCode
