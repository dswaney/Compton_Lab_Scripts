#requires -version 5.1
<#
.SYNOPSIS
    Disables Windows Update background activity after the weekly maintenance window.

.DESCRIPTION
    Stops and disables selected Windows Update services, applies automatic-update policy,
    disables selected Windows Update scheduled tasks, verifies the resulting state, and
    writes structured telemetry for Elastic ingestion.

.NOTES
    ScriptName:    09_Disable_Windows_Update_Services.ps1
    ScriptVersion: 2.0.6
    LastUpdated:   2026-08-17
    Changes:       v2.0.3 uses Maintenance.Framework v2.4 staged text logging.

    Script 01 is expected to restore the services and required scheduled tasks before the
    next maintenance cycle. Disabling BITS can affect non-Windows-Update software that uses
    BITS; use -KeepBITSAvailable when that compatibility is required.
#>

[CmdletBinding()]
param(
    [switch]$KeepBITSAvailable,
    [switch]$SkipScheduledTasks,
    [switch]$SkipUpdatePolicy
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptName = '09_Disable_Windows_Update_Services.ps1'
$ScriptVersion = '2.0.6'
$RunId = [guid]::NewGuid().Guid
$StartTime = Get-Date
$UtcTimestamp = $StartTime.ToUniversalTime().ToString('o')
$LogDirectory = 'C:\Logs'
$RuntimeLogPath = $null
$PublishedLogPath = $null
$LogSession = $null
$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestTelemetryPath = Join-Path $LogDirectory '09_Disable_Windows_Update_Services.latest.json'

$script:WarningCount = 0
$script:ErrorCount = 0
$script:OperationResults = New-Object System.Collections.Generic.List[object]

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 09 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '09_Disable_Windows_Update_Services.log',
        '*-09_Disable_Windows_Update_Services-*.log'
    ) | Out-Null

$LogSession = New-MaintenanceStagedLog `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $env:COMPUTERNAME `
    -Timestamp $StartTime

$RuntimeLogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

function Initialize-LogDirectory {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }
}

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
    $normalizedLevel = switch ($Level) {
        'OK'   { 'SUCCESS' }
        'WARN' { 'WARNING' }
        default { $Level }
    }
    $line = "$timestamp [$computerName] [$normalizedLevel] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow; $script:WarningCount++ }
        'ERROR' { Write-Host $line -ForegroundColor Red; $script:ErrorCount++ }
    }

    try {
        $activeLogDirectory = Split-Path -Parent $RuntimeLogPath
        if (-not (Test-Path -LiteralPath $activeLogDirectory -PathType Container)) {
            New-Item -Path $activeLogDirectory -ItemType Directory -Force | Out-Null
        }
        Add-Content -LiteralPath $RuntimeLogPath -Value $line -Encoding UTF8
    }
    catch { }
}

function Add-OperationResult {
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Status,
        [string]$Message,
        $Before,
        $After
    )

    $script:OperationResults.Add([pscustomobject]@{
        Type      = $Type
        Name      = $Name
        Status    = $Status
        Message   = $Message
        Before    = $Before
        After     = $After
        Timestamp = (Get-Date).ToUniversalTime().ToString('o')
    })
}

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-ComputerDomain {
    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.PartOfDomain -and $computerSystem.Domain) { return [string]$computerSystem.Domain }
        return [string]$env:USERDOMAIN
    }
    catch {
        return [string]$env:USERDOMAIN
    }
}

function Get-WindowsBuildInfo {
    $currentVersion = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    try {
        $cv = Get-ItemProperty -Path $currentVersion -ErrorAction Stop
        return [pscustomobject]@{
            ProductName    = [string]$cv.ProductName
            EditionId      = [string]$cv.EditionID
            DisplayVersion = [string]$cv.DisplayVersion
            CurrentBuild   = [string]$cv.CurrentBuild
            UBR            = [int]$cv.UBR
            FullBuild      = if ($null -ne $cv.UBR) { "{0}.{1}" -f $cv.CurrentBuild, $cv.UBR } else { [string]$cv.CurrentBuild }
        }
    }
    catch {
        return [pscustomobject]@{
            ProductName = $null; EditionId = $null; DisplayVersion = $null
            CurrentBuild = $null; UBR = $null; FullBuild = $null
        }
    }
}

function Get-ServiceSnapshot {
    param([Parameter(Mandatory)][string[]]$Names)

    foreach ($name in $Names) {
        try {
            $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction Stop
            [pscustomobject]@{
                Name      = [string]$service.Name
                DisplayName = [string]$service.DisplayName
                Exists    = $true
                State     = [string]$service.State
                StartMode = [string]$service.StartMode
                ProcessId = [int]$service.ProcessId
            }
        }
        catch {
            [pscustomobject]@{
                Name = $name; DisplayName = $null; Exists = $false
                State = 'NotFound'; StartMode = 'NotFound'; ProcessId = 0
            }
        }
    }
}

function Invoke-ScCommand {
    param([Parameter(Mandatory)][string[]]$Arguments)

    $output = & "$env:SystemRoot\System32\sc.exe" @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = ($output | Out-String).Trim()
        Succeeded = ($exitCode -eq 0)
    }
}



function New-StringArrayForJson {
    param([object[]]$Values)

    [string[]]$items = @(
        $Values |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    # PowerShell 5.1 can serialize some empty collections unexpectedly when they
    # are produced through conditional expressions. Return an explicit object[]
    # so ConvertTo-Json emits [] consistently.
    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Stop-ServiceProcessSafely {
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$WaitSeconds = 8
    )

    $result = [ordered]@{
        Attempted      = $false
        Succeeded      = $false
        Skipped        = $false
        Reason         = $null
        ProcessId      = $null
        HostedServices = @()
    }

    try {
        $svc = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $Name.Replace("'","''")) -ErrorAction Stop
        $pid = [int]$svc.ProcessId
        $result.ProcessId = $pid

        if ($pid -le 0) {
            $result.Skipped = $true
            $result.Reason = 'Service has no active process ID.'
            return [pscustomobject]$result
        }

        $hosted = @(
            Get-CimInstance Win32_Service -Filter ("ProcessId={0}" -f $pid) -ErrorAction Stop |
            Where-Object { [int]$_.ProcessId -eq $pid } |
            Select-Object -ExpandProperty Name
        )
        $result.HostedServices = @($hosted)

        if ($hosted.Count -gt 1) {
            $result.Skipped = $true
            $result.Reason = "PID $pid is shared by services: $($hosted -join ', '). Process kill was skipped."
            return [pscustomobject]$result
        }

        $result.Attempted = $true
        Stop-Process -Id $pid -Force -ErrorAction Stop

        $deadline = (Get-Date).AddSeconds($WaitSeconds)
        do {
            Start-Sleep -Milliseconds 250
            $snapshot = @(Get-ServiceSnapshot -Names @($Name))[0]
            if (-not $snapshot.Exists -or $snapshot.State -eq 'Stopped') {
                $result.Succeeded = $true
                $result.Reason = "Force-terminated exclusive service process PID $pid."
                return [pscustomobject]$result
            }
        } while ((Get-Date) -lt $deadline)

        $result.Reason = "PID $pid was terminated, but service state did not verify as Stopped within $WaitSeconds second(s)."
    }
    catch {
        $result.Reason = "Safe process termination failed: $($_.Exception.Message)"
    }

    return [pscustomobject]$result
}

function Set-ServiceStartupDisabledWithRetry {
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$Attempts = 3,
        [int]$DelayMilliseconds = 750
    )

    $messages = New-Object System.Collections.Generic.List[string]
    $succeeded = $false

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            Set-Service -Name $Name -StartupType Disabled -ErrorAction Stop
            $messages.Add("Set-Service startup disable succeeded on attempt $attempt.")
        }
        catch {
            $messages.Add("Set-Service attempt $attempt failed: $($_.Exception.Message)")
        }

        Start-Sleep -Milliseconds $DelayMilliseconds
        $snapshot = @(Get-ServiceSnapshot -Names @($Name))[0]
        if ($snapshot.Exists -and $snapshot.StartMode -eq 'Disabled') {
            $succeeded = $true
            break
        }

        $scConfig = Invoke-ScCommand -Arguments @('config', $Name, 'start=', 'disabled')
        if ($scConfig.Succeeded) {
            $messages.Add("sc.exe config succeeded on attempt $attempt.")
        }
        else {
            $messages.Add("sc.exe config attempt $attempt failed. ExitCode=$($scConfig.ExitCode). $($scConfig.Output)")
        }

        Start-Sleep -Milliseconds $DelayMilliseconds
        $snapshot = @(Get-ServiceSnapshot -Names @($Name))[0]
        if ($snapshot.Exists -and $snapshot.StartMode -eq 'Disabled') {
            $succeeded = $true
            break
        }
    }

    return [pscustomobject]@{
        Succeeded = $succeeded
        Messages  = @($messages | ForEach-Object { $_ })
    }
}

function Stop-AndDisableService {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$AllowMissing,
        [switch]$AllowProtectedFailure
    )

    $before = @(Get-ServiceSnapshot -Names @($Name))[0]
    if (-not $before.Exists) {
        $level = if ($AllowMissing) { 'INFO' } else { 'WARN' }
        Write-Status "Service is not present: $Name" $level
        Add-OperationResult -Type 'Service' -Name $Name -Status 'NotFound' -Message 'Service was not installed.' -Before $before -After $before
        return
    }

    $messages = New-Object System.Collections.Generic.List[string]
    $operationFailed = $false

    if ($before.State -ne 'Stopped') {
        $normalStopVerified = $false

        try {
            Stop-Service -Name $Name -Force -ErrorAction Stop
            $serviceObject = Get-Service -Name $Name -ErrorAction Stop
            $serviceObject.WaitForStatus('Stopped', [timespan]::FromSeconds(12))
            $normalStopVerified = $true
            $messages.Add('Stopped with Stop-Service within 12 seconds.')
        }
        catch {
            $messages.Add("Stop-Service did not verify within 12 seconds: $($_.Exception.Message)")

            $scStop = Invoke-ScCommand -Arguments @('stop', $Name)
            if ($scStop.Succeeded -or $scStop.Output -match 'service has not been started|service is not started') {
                $messages.Add('Fallback stop requested with sc.exe.')

                try {
                    $serviceObject = Get-Service -Name $Name -ErrorAction Stop
                    $serviceObject.WaitForStatus('Stopped', [timespan]::FromSeconds(6))
                    $normalStopVerified = $true
                    $messages.Add('Service stopped after sc.exe fallback.')
                }
                catch {
                    $messages.Add('Service still did not verify as stopped after sc.exe fallback.')
                }
            }
            else {
                $messages.Add("sc.exe stop failed. ExitCode=$($scStop.ExitCode). $($scStop.Output)")
            }
        }

        if (-not $normalStopVerified) {
            $killResult = Stop-ServiceProcessSafely -Name $Name -WaitSeconds 6

            if ($killResult.Succeeded) {
                $messages.Add($killResult.Reason)
            }
            elseif ($killResult.Skipped) {
                $operationFailed = $true
                $messages.Add("Process kill skipped for safety. $($killResult.Reason)")
            }
            else {
                $operationFailed = $true
                $messages.Add($killResult.Reason)
            }
        }
    }
    else {
        $messages.Add('Service was already stopped.')
    }

    $startupResult = Set-ServiceStartupDisabledWithRetry -Name $Name -Attempts 3 -DelayMilliseconds 750
    foreach ($msg in $startupResult.Messages) {
        $messages.Add($msg)
    }
    if (-not $startupResult.Succeeded) {
        $operationFailed = $true
        $messages.Add('Startup type did not verify as Disabled after 3 attempts.')
    }

    Start-Sleep -Milliseconds 1000
    $after = @(Get-ServiceSnapshot -Names @($Name))[0]
    $verified = $after.Exists -and $after.StartMode -eq 'Disabled' -and $after.State -eq 'Stopped'

    if ($verified) {
        Write-Status "Service state change: $Name | State: $($before.State) -> $($after.State) | StartMode: $($before.StartMode) -> $($after.StartMode)" 'OK'
        Add-OperationResult -Type 'Service' -Name $Name -Status 'Success' -Message ($messages -join ' ') -Before $before -After $after
    }
    else {
        $message = ($messages -join ' ') + " Final state: State=$($after.State), StartMode=$($after.StartMode)."
        if ($operationFailed -and $AllowProtectedFailure) {
            Write-Status "Protected service could not be fully disabled: $Name. State=$($after.State); StartMode=$($after.StartMode)" 'WARN'
            Add-OperationResult -Type 'Service' -Name $Name -Status 'ProtectedService' -Message $message -Before $before -After $after
        }
        elseif ($operationFailed) {
            Write-Status "Failed to fully disable service $Name. State=$($after.State); StartMode=$($after.StartMode)" 'ERROR'
            Add-OperationResult -Type 'Service' -Name $Name -Status 'Failed' -Message $message -Before $before -After $after
        }
        else {
            Write-Status "Service $Name did not remain fully disabled. State=$($after.State); StartMode=$($after.StartMode)" 'WARN'
            Add-OperationResult -Type 'Service' -Name $Name -Status 'VerificationFailed' -Message $message -Before $before -After $after
        }
    }
}

function Get-RegistryValueSnapshot {
    param([string]$Path, [string]$Name)

    try {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return [pscustomobject]@{ Path = $Path; Name = $Name; Exists = $true; Value = $item.$Name }
    }
    catch {
        return [pscustomobject]@{ Path = $Path; Name = $Name; Exists = $false; Value = $null }
    }
}

function Set-RegistryDwordVerified {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Value
    )

    $before = Get-RegistryValueSnapshot -Path $Path -Name $Name
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force -ErrorAction Stop | Out-Null
        $after = Get-RegistryValueSnapshot -Path $Path -Name $Name
        if ($after.Exists -and [int]$after.Value -eq $Value) {
            Write-Status "Verified policy value: $Path\$Name = $Value" 'OK'
            Add-OperationResult -Type 'RegistryPolicy' -Name "$Path\$Name" -Status 'Success' -Message 'Registry policy value applied and verified.' -Before $before -After $after
        }
        else {
            Write-Status "Policy value failed verification: $Path\$Name" 'ERROR'
            Add-OperationResult -Type 'RegistryPolicy' -Name "$Path\$Name" -Status 'VerificationFailed' -Message 'Registry write completed but the expected value was not read back.' -Before $before -After $after
        }
    }
    catch {
        $after = Get-RegistryValueSnapshot -Path $Path -Name $Name
        Write-Status "Failed to set $Path\${Name}: $($_.Exception.Message)" 'ERROR'
        Add-OperationResult -Type 'RegistryPolicy' -Name "$Path\$Name" -Status 'Failed' -Message $_.Exception.Message -Before $before -After $after
    }
}

function Get-ScheduledTaskSnapshot {
    param([string]$TaskPath, [string]$TaskName)

    try {
        $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
        return [pscustomobject]@{
            FullName = "$TaskPath$TaskName"
            Exists = $true
            Enabled = [bool]$task.Settings.Enabled
            State = [string]$task.State
        }
    }
    catch {
        return [pscustomobject]@{
            FullName = "$TaskPath$TaskName"
            Exists = $false
            Enabled = $null
            State = 'NotFound'
        }
    }
}

function Disable-ScheduledTaskVerified {
    param(
        [Parameter(Mandatory)][string]$TaskPath,
        [Parameter(Mandatory)][string]$TaskName
    )

    $before = Get-ScheduledTaskSnapshot -TaskPath $TaskPath -TaskName $TaskName
    if (-not $before.Exists) {
        Write-Status "Scheduled task is not present: $($before.FullName)" 'INFO'
        Add-OperationResult -Type 'ScheduledTask' -Name $before.FullName -Status 'NotFound' -Message 'Task is not present on this Windows build.' -Before $before -After $before
        return
    }

    try {
        if ($before.Enabled) {
            Disable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
        }
        $after = Get-ScheduledTaskSnapshot -TaskPath $TaskPath -TaskName $TaskName
        if ($after.Exists -and -not $after.Enabled) {
            Write-Status "Verified scheduled task disabled: $($after.FullName)" 'OK'
            Add-OperationResult -Type 'ScheduledTask' -Name $after.FullName -Status 'Success' -Message 'Task disabled and verified.' -Before $before -After $after
        }
        else {
            Write-Status "Scheduled task did not remain disabled: $($before.FullName)" 'WARN'
            Add-OperationResult -Type 'ScheduledTask' -Name $before.FullName -Status 'VerificationFailed' -Message 'Task remained enabled after the disable request.' -Before $before -After $after
        }
    }
    catch {
        $after = Get-ScheduledTaskSnapshot -TaskPath $TaskPath -TaskName $TaskName
        Write-Status "Could not disable scheduled task $($before.FullName): $($_.Exception.Message)" 'WARN'
        Add-OperationResult -Type 'ScheduledTask' -Name $before.FullName -Status 'Failed' -Message $_.Exception.Message -Before $before -After $after
    }
}

function Get-PendingRebootState {
    $reasons = New-Object System.Collections.Generic.List[string]
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons.Add('ComponentBasedServicing') }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons.Add('WindowsUpdate') }

    try {
        $sessionManager = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop
        if ($sessionManager.PendingFileRenameOperations) { $reasons.Add('PendingFileRenameOperations') }
        if ($sessionManager.PendingFileRenameOperations2) { $reasons.Add('PendingFileRenameOperations2') }
    }
    catch { }

    return [pscustomobject]@{
        Pending = ($reasons.Count -gt 0)
        Reasons = @($reasons)
    }
}

function Write-Telemetry {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [string]$FailureMessage,
        $ServicesBefore,
        $ServicesAfter,
        $TasksAfter,
        $PolicyAfter,
        $PendingReboot
    )

    $endTime = Get-Date
    $duration = [math]::Round(($endTime - $StartTime).TotalSeconds, 3)
    # PowerShell 5.1 can throw 'Argument types do not match' when a generic List[object]
    # is wrapped directly in @(...). Convert it explicitly to a normal object array.
    $operations = if ($null -ne $script:OperationResults) {
        @($script:OperationResults | ForEach-Object { $_ })
    }
    else {
        @()
    }

    $serviceOperations = @($operations | Where-Object { $_.Type -eq 'Service' })
    $taskOperations = @($operations | Where-Object { $_.Type -eq 'ScheduledTask' })
    $policyOperations = @($operations | Where-Object { $_.Type -eq 'RegistryPolicy' })

    $event = [ordered]@{
        '@timestamp' = $endTime.ToUniversalTime().ToString('o')
        EventType = 'maintenance.execution'
        EventCategory = 'windows_update_control'
        ComputerName = $env:COMPUTERNAME
        Domain = Get-ComputerDomain
        ScriptName = $ScriptName
        ScriptVersion = $ScriptVersion
        RunId = $RunId
        Status = $Status
        ExitCode = $ExitCode
        StartTime = $StartTime.ToString('o')
        EndTime = $endTime.ToString('o')
        DurationSeconds = $duration
        ErrorCount = $script:ErrorCount
        WarningCount = $script:WarningCount
        FailureMessage = $FailureMessage
        TextLogPath = $PublishedLogPath

        windows_update_control = [ordered]@{
            KeepBITSAvailable = [bool]$KeepBITSAvailable
            ScheduledTaskChangesSkipped = [bool]$SkipScheduledTasks
            UpdatePolicyChangesSkipped = [bool]$SkipUpdatePolicy

            Reboot = [ordered]@{
                Recommended = $true
                Pending = if ($PendingReboot) { [bool]$PendingReboot.Pending } else { $null }
                PendingRebootReasonNames = New-StringArrayForJson -Values $(if ($PendingReboot) { @($PendingReboot.Reasons) } else { @() })
            }

            Services = [ordered]@{
                Targeted = $serviceOperations.Count
                DisabledVerified = @($serviceOperations | Where-Object { $_.Status -eq 'Success' }).Count
                FailedNames = New-StringArrayForJson -Values @(
                    $serviceOperations |
                    Where-Object { $_.Status -in @('Failed','VerificationFailed') } |
                    ForEach-Object { $_.Name }
                )
                Before = @($ServicesBefore)
                After = @($ServicesAfter)
            }

            ScheduledTasks = [ordered]@{
                Targeted = $taskOperations.Count
                DisabledVerified = @($taskOperations | Where-Object { $_.Status -eq 'Success' }).Count
                FailedNames = New-StringArrayForJson -Values @(
                    $taskOperations |
                    Where-Object { $_.Status -in @('Failed','VerificationFailed') } |
                    ForEach-Object { $_.Name }
                )
                After = @($TasksAfter)
            }

            Policies = [ordered]@{
                Targeted = $policyOperations.Count
                Verified = @($policyOperations | Where-Object { $_.Status -eq 'Success' }).Count
                FailureNames = New-StringArrayForJson -Values @(
                    $policyOperations |
                    Where-Object { $_.Status -in @('Failed','VerificationFailed') } |
                    ForEach-Object { $_.Name }
                )
                After = @($PolicyAfter)
            }

            OperationResults = $operations
            Windows = Get-WindowsBuildInfo
        }
    }

    try {
        $wuc = $event.windows_update_control
        Write-Status ("Elastic Script 09 summary: ServicesFailed={0}; TasksFailed={1}; PolicyFailures={2}; PendingReboot={3}; RebootReasons={4}" -f `
            @($wuc.Services.FailedNames).Count,
            @($wuc.ScheduledTasks.FailedNames).Count,
            @($wuc.Policies.FailureNames).Count,
            $wuc.Reboot.Pending,
            (@($wuc.Reboot.PendingRebootReasonNames) -join ', ')) 'INFO'
    }
    catch { }

    $jsonCompact = $event | ConvertTo-Json -Depth 12 -Compress
    $jsonFormatted = $event | ConvertTo-Json -Depth 12

    try { Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $jsonCompact }
    catch { Write-Status "Could not append telemetry to ${TelemetryPath}: $($_.Exception.Message)" 'ERROR' }

    try { [System.IO.File]::WriteAllText($LatestTelemetryPath, $jsonFormatted, (New-Object System.Text.UTF8Encoding($false))) }
    catch { Write-Status "Could not write latest telemetry to ${LatestTelemetryPath}: $($_.Exception.Message)" 'ERROR' }
}

Initialize-LogDirectory
Write-Status "Starting $ScriptName version $ScriptVersion. Run ID: $RunId" 'INFO'
Write-Status "Active staged text log: $RuntimeLogPath" 'INFO'
Write-Status "Completed text log publish path: $PublishedLogPath" 'INFO'

$finalStatus = 'Failed'
$finalExitCode = 1
$failureMessage = $null
$servicesBefore = @()
$servicesAfter = @()
$tasksAfter = @()
$policyAfter = @()
$pendingReboot = $null

try {
    if (-not (Test-IsAdministrator)) {
        throw 'This script must be run as Administrator or SYSTEM.'
    }

    $servicesToDisable = New-Object System.Collections.Generic.List[string]
    $servicesToDisable.Add('wuauserv')
    if (-not $KeepBITSAvailable) { $servicesToDisable.Add('bits') }
    $servicesToDisable.Add('dosvc')
    $servicesToDisable.Add('UsoSvc')
    $servicesToDisable.Add('WaaSMedicSvc')

    $servicesBefore = @(Get-ServiceSnapshot -Names @($servicesToDisable))

    foreach ($serviceName in $servicesToDisable) {
        Stop-AndDisableService -Name $serviceName -AllowMissing -AllowProtectedFailure:($serviceName -eq 'WaaSMedicSvc')
    }

    if ($KeepBITSAvailable) {
        Write-Status 'BITS was intentionally left available because -KeepBITSAvailable was specified.' 'INFO'
    }

    if (-not $SkipUpdatePolicy) {
        $auPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
        Set-RegistryDwordVerified -Path $auPolicyPath -Name 'NoAutoUpdate' -Value 1
        Set-RegistryDwordVerified -Path $auPolicyPath -Name 'AUOptions' -Value 1
    }
    else {
        Write-Status 'Windows Update policy changes were skipped by request.' 'INFO'
    }

    $tasks = @(
        @{ Path = '\Microsoft\Windows\WindowsUpdate\';      Name = 'Scheduled Start' },
        @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan' },
        @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan Static Task' },
        @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'USO_UxBroker' },
        @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Reboot' },
        @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Maintenance Install' },
        @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Refresh Settings' },
        @{ Path = '\Microsoft\Windows\WaaSMedic\';          Name = 'PerformRemediation' }
    )

    if (-not $SkipScheduledTasks) {
        foreach ($task in $tasks) {
            Disable-ScheduledTaskVerified -TaskPath $task.Path -TaskName $task.Name
        }
    }
    else {
        Write-Status 'Windows Update scheduled-task changes were skipped by request.' 'INFO'
    }

    $servicesAfter = @(Get-ServiceSnapshot -Names @($servicesToDisable))
    $tasksAfter = @($tasks | ForEach-Object { Get-ScheduledTaskSnapshot -TaskPath $_.Path -TaskName $_.Name })
    $policyAfter = @(
        Get-RegistryValueSnapshot -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name 'NoAutoUpdate'
        Get-RegistryValueSnapshot -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name 'AUOptions'
    )
    $pendingReboot = Get-PendingRebootState

    $failedOperations = @($script:OperationResults | Where-Object { $_.Status -in @('Failed','VerificationFailed') })
    if ($script:ErrorCount -gt 0 -or @($failedOperations | Where-Object { $_.Type -in @('Service','RegistryPolicy') }).Count -gt 0) {
        $finalStatus = 'CompletedWithFailures'
        $finalExitCode = 2
        Write-Status 'Windows Update background controls completed with one or more failures.' 'ERROR'
    }
    elseif ($script:WarningCount -gt 0 -or $failedOperations.Count -gt 0) {
        $finalStatus = 'SuccessWithWarnings'
        $finalExitCode = 0
        Write-Status 'Windows Update background controls completed with warnings.' 'WARN'
    }
    else {
        $finalStatus = 'SuccessRebootRecommended'
        $finalExitCode = 0
        Write-Status 'Windows Update background controls were disabled and verified.' 'OK'
    }

    Write-Status 'A reboot is recommended so service and task state is fully settled.' 'INFO'
}
catch {
    $failureMessage = $_.Exception.Message
    $finalStatus = 'Failed'
    $finalExitCode = 1
    Write-Status "Script failed: $failureMessage" 'ERROR'
}
finally {
    if (-not $pendingReboot) { $pendingReboot = Get-PendingRebootState }
    if (-not $servicesAfter -or $servicesAfter.Count -eq 0) {
        $targetNames = @('wuauserv','bits','dosvc','UsoSvc','WaaSMedicSvc')
        $servicesAfter = @(Get-ServiceSnapshot -Names $targetNames)
    }

    try {
        Write-Telemetry -Status $finalStatus -ExitCode $finalExitCode -FailureMessage $failureMessage `
            -ServicesBefore $servicesBefore -ServicesAfter $servicesAfter -TasksAfter $tasksAfter `
            -PolicyAfter $policyAfter -PendingReboot $pendingReboot
    }
    catch {
        # Telemetry must never replace the maintenance result with an unhandled exception.
        Write-Status "Telemetry generation failed: $($_.Exception.Message)" 'ERROR'
        if ($finalExitCode -eq 0) {
            $finalStatus = 'TelemetryFailure'
            $finalExitCode = 4
        }
    }

    # This is the final append to the active text log. Nothing should write to
    # the file after it is published into C:\Logs for Elastic.
    Write-Status ("Finalizing {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
        $ScriptName,
        $finalStatus,
        $finalExitCode,
        $script:WarningCount,
        $script:ErrorCount) $(if ($finalExitCode -eq 0) { 'OK' } else { 'ERROR' })

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 09 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 09 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $finalExitCode
