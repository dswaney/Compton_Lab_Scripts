# =====================================================================
# ScriptName: 01_Enable_Windows_Update_Services.ps1
# ScriptVersion: 2.8.5
# LastUpdated: 2026-09-03
# Purpose: Restore Windows Update services, tasks, policy settings,
#          Windows 11 UI preferences, and classic right-click context menu behavior for all users;
#          verify required services are running, retry startup failures
#          up to 4 total attempts, and force a reboot if critical
#          services still refuse to start.
# Changes:  v2.8.5 aligns updater-task validation and repair with the 01:00 Sunday master schedule.
#           v2.8.4 guarantees structured failure telemetry for bootstrap/framework/privilege/runtime failures.
#           Adds FailureStage and FailureMessage fields and ensures service-recovery reboot failures pass through finally.
#           v2.8.3 uses Maintenance.Framework v2.4 staged text logging so Elastic only sees the completed immutable log.
# Fix:      Removed StartWhenAvailable from task creation and added direct
# =====================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$script:ScriptName       = '01_Enable_Windows_Update_Services.ps1'
$script:ScriptVersion    = '2.8.5'
$script:ExecutionStart   = Get-Date
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:ExitCode         = 0
$script:FinalStatus      = 'Running'
$script:RebootInitiated  = $false
$script:RebootReason     = $null
$script:TaskReconcileStatus = 'NotRun'
$script:ServiceNames     = @('wuauserv','bits','dosvc','UsoSvc','WaaSMedicSvc','CryptSvc','msiserver')
$script:ServiceStateBefore = @()
$script:ServiceStateAfter  = @()
$script:PolicyConflictsBefore = @()
$script:TelemetryDirectory = 'C:\Logs'
$script:NdjsonPath = Join-Path $script:TelemetryDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestJsonPath = Join-Path $script:TelemetryDirectory '01_Enable_Windows_Update_Services.latest.json'
$script:LogPath = $null
$script:PublishedLogPath = $null
$script:LogSession = $null
$script:FailureStage = $null
$script:FailureMessage = $null
$script:BootstrapExitCode = $null


# Load the shared framework from the same directory as this script.
# ============================================================================
# BOOTSTRAP THE LATEST 00 UPDATER BEFORE IMPORTING MAINTENANCE.FRAMEWORK
# ============================================================================
# Older lab computers may have an incomplete updater or may not yet have
# Maintenance.Framework.psm1. Refresh and run 00 first so C:\Scripts is
# synchronized before script 01 depends on any shared maintenance components.


function Stop-ServiceWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$TimeoutSeconds = $ServiceStopTimeoutSeconds,
        [int]$KillWaitSeconds = $ServiceKillWaitSeconds
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if (-not $service) {
        Write-Status -Level INFO -Message "Service not found: $Name"
        return $true
    }

    if ($service.Status -eq 'Stopped') {
        Write-Status -Level INFO -Message "Service is already stopped: $Name"
        return $true
    }

    Write-Status -Level INFO -Message "Stopping service $Name with a timeout of $TimeoutSeconds second(s)..."

    try {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Status -Level WARN -Message "Initial Stop-Service request failed for $Name : $($_.Exception.Message)"
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        Start-Sleep -Seconds 1
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if (-not $service -or $service.Status -eq 'Stopped') {
            Write-Status -Level OK -Message "Service stopped successfully: $Name"
            return $true
        }
    }
    while ((Get-Date) -lt $deadline)

    Write-Status -Level WARN -Message "Service $Name did not stop within $TimeoutSeconds second(s). Attempting forced process termination."

    try {
        $serviceInfo = Get-CimInstance Win32_Service -Filter "Name='$Name'" -ErrorAction Stop
        $servicePid = [int]$serviceInfo.ProcessId

        if ($servicePid -gt 0) {
            $process = Get-Process -Id $servicePid -ErrorAction SilentlyContinue

            if ($process) {
                Write-Status -Level WARN -Message "Terminating PID $servicePid hosting service $Name."
                Stop-Process -Id $servicePid -Force -ErrorAction Stop
            }
            else {
                Write-Status -Level INFO -Message "Service $Name reported PID $servicePid, but the process is already gone."
            }
        }
        else {
            Write-Status -Level WARN -Message "Service $Name has no active process ID to terminate."
        }
    }
    catch {
        Write-Status -Level WARN -Message "Forced process termination failed for service $Name : $($_.Exception.Message)"
    }

    $killDeadline = (Get-Date).AddSeconds($KillWaitSeconds)

    do {
        Start-Sleep -Seconds 1
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if (-not $service -or $service.Status -eq 'Stopped') {
            Write-Status -Level OK -Message "Service stopped after forced termination: $Name"
            return $true
        }
    }
    while ((Get-Date) -lt $killDeadline)

    Write-Status -Level ERROR -Message "Service $Name is still not stopped after timeout and forced termination."
    return $false
}

function Write-EarlyFailureTelemetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FailureStage,
        [Parameter(Mandatory)][string]$FailureMessage,
        [int]$ExitCode = 1
    )

    try {
        if (-not (Test-Path -LiteralPath $script:TelemetryDirectory -PathType Container)) {
            New-Item -Path $script:TelemetryDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $end = Get-Date
        $duration = [math]::Round(($end - $script:ExecutionStart).TotalSeconds, 3)
        $domain = $env:USERDNSDOMAIN
        if ([string]::IsNullOrWhiteSpace($domain)) {
            try { $domain = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Domain } catch { $domain = $null }
        }

        # Keep the same top-level schema already used by Script 01 so the existing
        # compton-maintenance-json pipeline and mappings can ingest the failure.
        $event = [ordered]@{
            EventType                = 'maintenance.execution'
            ComputerName             = $env:COMPUTERNAME
            Domain                   = $domain
            ScriptName               = $script:ScriptName
            ScriptVersion            = $script:ScriptVersion
            Status                   = 'Failed'
            ExitCode                 = $ExitCode
            StartTime                = $script:ExecutionStart.ToString('o')
            EndTime                  = $end.ToString('o')
            DurationSeconds          = $duration
            ErrorCount               = 1
            WarningCount             = $script:WarningCount
            Timestamp                = (Get-Date).ToUniversalTime().ToString('o')
            FailureStage             = $FailureStage
            FailureMessage           = $FailureMessage
            RebootInitiated          = $false
            RebootReason             = $null
            TaskReconciliationStatus = 'NotRun'
            PolicyConflictDetected   = $false
            ServicesChecked          = 0
            ServicesRepairedCount    = 0
            ApprovedServiceState     = $false
            TextLogPath              = $null
        }

        $jsonCompact = $event | ConvertTo-Json -Depth 8 -Compress
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $line = $jsonCompact + [Environment]::NewLine

        # Retry because Elastic/file readers can briefly hold the telemetry file.
        $written = $false
        $lastError = $null
        for ($attempt = 1; $attempt -le 10; $attempt++) {
            try {
                [System.IO.File]::AppendAllText($script:NdjsonPath, $line, $utf8NoBom)
                $written = $true
                break
            }
            catch {
                $lastError = $_.Exception.Message
                Start-Sleep -Milliseconds 250
            }
        }

        if (-not $written) {
            throw "Unable to append early failure telemetry after 10 attempts: $lastError"
        }

        try {
            [System.IO.File]::WriteAllText(
                $script:LatestJsonPath,
                ($event | ConvertTo-Json -Depth 8),
                $utf8NoBom
            )
        }
        catch {
            Write-Host "Early failure latest.json write failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        Write-Host ("Early failure telemetry written. Stage={0}; ExitCode={1}" -f $FailureStage, $ExitCode) -ForegroundColor Yellow
    }
    catch {
        Write-Host "CRITICAL: Early failure telemetry could not be written: $($_.Exception.Message)" -ForegroundColor Red
    }
}


function Invoke-LatestUpdaterBootstrap {
    [CmdletBinding()]
    param()

    $localScriptsRoot = 'C:\Scripts'
    $localUpdaterPath = Join-Path $localScriptsRoot '00_Update-Scripts-FromShare.ps1'

    $sourceRoots = @(
        '\\filesvr\Labscripts',
        '\\10.2.3.30\Labscripts'
    )

    $sourceRoot = $null

    foreach ($candidate in $sourceRoots) {
        try {
            if (Test-Path -LiteralPath $candidate -PathType Container -ErrorAction Stop) {
                $sourceRoot = $candidate
                break
            }
        }
        catch {
            # Try the next source.
        }
    }

    if ([string]::IsNullOrWhiteSpace($sourceRoot)) {
        throw 'Updater bootstrap failed because neither \\filesvr\Labscripts nor \\10.2.3.30\Labscripts is available.'
    }

    $sourceUpdaterPath = Join-Path $sourceRoot '00_Update-Scripts-FromShare.ps1'

    if (-not (Test-Path -LiteralPath $sourceUpdaterPath -PathType Leaf)) {
        throw "Updater bootstrap failed because the source updater does not exist: $sourceUpdaterPath"
    }

    if (-not (Test-Path -LiteralPath $localScriptsRoot -PathType Container)) {
        New-Item -Path $localScriptsRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    # Refuse to replace a working local updater with a syntactically broken source copy.
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $sourceUpdaterPath,
        [ref]$tokens,
        [ref]$parseErrors
    )

    if ($parseErrors -and $parseErrors.Count -gt 0) {
        $summary = @(
            $parseErrors |
            ForEach-Object { "Line $($_.Extent.StartLineNumber): $($_.Message)" }
        ) -join ' | '

        throw "Updater bootstrap source failed PowerShell parser validation. $summary"
    }

    $sourceHash = (Get-FileHash -LiteralPath $sourceUpdaterPath -Algorithm SHA256 -ErrorAction Stop).Hash
    $localHash = $null

    if (Test-Path -LiteralPath $localUpdaterPath -PathType Leaf) {
        try {
            $localHash = (Get-FileHash -LiteralPath $localUpdaterPath -Algorithm SHA256 -ErrorAction Stop).Hash
        }
        catch {
            $localHash = $null
        }
    }

    if ($localHash -ne $sourceHash) {
        Write-Host "[BOOTSTRAP] Updating local 00 updater from $sourceRoot" -ForegroundColor Yellow

        Copy-Item -LiteralPath $sourceUpdaterPath -Destination $localUpdaterPath -Force -ErrorAction Stop
        Unblock-File -LiteralPath $localUpdaterPath -ErrorAction SilentlyContinue

        $copiedHash = (Get-FileHash -LiteralPath $localUpdaterPath -Algorithm SHA256 -ErrorAction Stop).Hash

        if ($copiedHash -ne $sourceHash) {
            throw 'Updater bootstrap copy verification failed: the local 00 updater does not match the source copy.'
        }

        Write-Host '[BOOTSTRAP] Latest 00 updater copied successfully.' -ForegroundColor Green
    }
    else {
        Write-Host '[BOOTSTRAP] Local 00 updater already matches the deployment share.' -ForegroundColor Green
    }

    Write-Host '[BOOTSTRAP] Running 00 updater before continuing script 01...' -ForegroundColor Cyan

    $arguments = @(
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $localUpdaterPath)
    )

    $bootstrapProcess = Start-Process `
        -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList $arguments `
        -WorkingDirectory $localScriptsRoot `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    if ($bootstrapProcess.ExitCode -ne 0) {
        $script:BootstrapExitCode = [int]$bootstrapProcess.ExitCode
        throw "Latest 00 updater returned exit code $($bootstrapProcess.ExitCode). Script 01 will stop because C:\Scripts may be incomplete."
    }

    $frameworkPath = Join-Path $localScriptsRoot 'Maintenance.Framework.psm1'
    if (-not (Test-Path -LiteralPath $frameworkPath -PathType Leaf)) {
        throw "00 updater completed but Maintenance.Framework.psm1 is still missing: $frameworkPath"
    }

    Write-Host '[BOOTSTRAP] 00 updater completed and Maintenance.Framework.psm1 is present. Continuing script 01.' -ForegroundColor Green
}

try {
    Invoke-LatestUpdaterBootstrap
}
catch {
    $script:FinalStatus = 'Failed'
    $script:FailureStage = 'Bootstrap'
    $script:FailureMessage = $_.Exception.Message
    $script:ExitCode = if ($null -ne $script:BootstrapExitCode -and $script:BootstrapExitCode -ne 0) {
        [int]$script:BootstrapExitCode
    }
    else {
        1
    }

    Write-Host ("Script 01 bootstrap failed: {0}" -f $script:FailureMessage) -ForegroundColor Red
    Write-EarlyFailureTelemetry `
        -FailureStage $script:FailureStage `
        -FailureMessage $script:FailureMessage `
        -ExitCode $script:ExitCode

    exit $script:ExitCode
}

try {
    $MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
    [int]$ServiceStopTimeoutSeconds = 30
    [int]$ServiceKillWaitSeconds = 10
    Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
    $MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

    $requiredFrameworkVersion = [version]'2.4.0'
    $currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

    if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
        throw "Script 01 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
    }

    Archive-MaintenanceLogs `
        -ScriptName $script:ScriptName `
        -LogRoot $script:TelemetryDirectory `
        -AdditionalPatterns @(
            '01_Enable_Windows_Update_Services.log',
            '*-01_Enable_Windows_Update_Services-*.log'
        ) | Out-Null

    $script:LogSession = New-MaintenanceStagedLog `
        -ScriptName $script:ScriptName `
        -LogRoot $script:TelemetryDirectory `
        -StagingRoot $MaintenanceConfig.LogStagingRoot `
        -ComputerName $env:COMPUTERNAME `
        -Timestamp $script:ExecutionStart

    $script:LogPath = [string]$script:LogSession.WorkingPath
    $script:PublishedLogPath = [string]$script:LogSession.PublishedPath
}
catch {
    $script:FinalStatus = 'Failed'
    $script:ExitCode = 1
    $script:FailureStage = 'FrameworkInitialization'
    $script:FailureMessage = $_.Exception.Message

    Write-Host ("Script 01 framework initialization failed: {0}" -f $script:FailureMessage) -ForegroundColor Red
    Write-EarlyFailureTelemetry `
        -FailureStage $script:FailureStage `
        -FailureMessage $script:FailureMessage `
        -ExitCode $script:ExitCode

    exit $script:ExitCode
}


function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $elasticLevel = switch ($Level) {
        'OK'    { 'SUCCESS' }
        'WARN'  { 'WARNING' }
        default { $Level }
    }

    $line = '{0} [{1}] [{2}] {3}' -f $timestamp, $env:COMPUTERNAME, $elasticLevel, $Message
    $color = switch ($Level) {
        'INFO'  { 'Cyan' }
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
    }

    Write-Host $line -ForegroundColor $color

    try {
        $activeLogDirectory = Split-Path -Parent $script:LogPath
        if (-not (Test-Path -LiteralPath $activeLogDirectory -PathType Container)) {
            New-Item -Path $activeLogDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host "Unable to write maintenance log: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}


function Get-FileSha256Safe {
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
        }
    }
    catch {
    }

    return $null
}

function Get-ServiceTelemetrySnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Names)

    $results = foreach ($name in $Names) {
        try {
            $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction Stop
            [pscustomobject]@{
                Name        = $name
                Exists      = $true
                State       = [string]$service.State
                StartMode   = [string]$service.StartMode
                ProcessId   = [int]$service.ProcessId
            }
        }
        catch {
            [pscustomobject]@{
                Name        = $name
                Exists      = $false
                State       = $null
                StartMode   = $null
                ProcessId   = 0
            }
        }
    }

    return @($results)
}

function Get-WindowsUpdatePolicyConflicts {
    [CmdletBinding()]
    param()

    $conflicts = New-Object System.Collections.Generic.List[string]
    $wuPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'

    foreach ($item in @(
        @{ Path = $wuPath; Name = 'WUServer' },
        @{ Path = $wuPath; Name = 'WUStatusServer' },
        @{ Path = $wuPath; Name = 'UpdateServiceUrlAlternate' },
        @{ Path = $wuPath; Name = 'DisableWindowsUpdateAccess' },
        @{ Path = $auPath; Name = 'UseWUServer' }
    )) {
        try {
            $value = (Get-ItemProperty -Path $item.Path -Name $item.Name -ErrorAction Stop).$($item.Name)
            if ($null -ne $value -and $value -notin @(0, '')) {
                $conflicts.Add("$($item.Path)\\$($item.Name)=$value")
            }
        }
        catch { }
    }

    try {
        $noAutoUpdate = (Get-ItemProperty -Path $auPath -Name 'NoAutoUpdate' -ErrorAction Stop).NoAutoUpdate
        if ([int]$noAutoUpdate -ne 0) { $conflicts.Add("$auPath\\NoAutoUpdate=$noAutoUpdate") }
    }
    catch { }

    return @($conflicts)
}

function Write-ExecutionTelemetry {
    [CmdletBinding()]
    param()

    try {
        if (-not (Test-Path -LiteralPath $script:TelemetryDirectory -PathType Container)) {
            New-Item -Path $script:TelemetryDirectory -ItemType Directory -Force | Out-Null
        }

        $end = Get-Date
        $duration = [math]::Round(($end - $script:ExecutionStart).TotalSeconds, 3)
        $afterByName = @{}
        foreach ($service in $script:ServiceStateAfter) { $afterByName[$service.Name] = $service }

        $serviceChanges = foreach ($before in $script:ServiceStateBefore) {
            $after = $afterByName[$before.Name]
            if ($null -ne $after -and ($before.State -ne $after.State -or $before.StartMode -ne $after.StartMode)) {
                [pscustomobject]@{
                    Name            = $before.Name
                    StateBefore     = $before.State
                    StateAfter      = $after.State
                    StartModeBefore = $before.StartMode
                    StartModeAfter  = $after.StartMode
                }
            }
        }

        $required = @('wuauserv','bits','dosvc','UsoSvc')
        $servicesFailed = @($script:ServiceStateAfter | Where-Object { $_.Name -in $required -and $_.State -ne 'Running' } | ForEach-Object Name)
        $servicesRunning = @($script:ServiceStateAfter | Where-Object State -eq 'Running' | ForEach-Object Name)
        $policyConflictsAfter = @(Get-WindowsUpdatePolicyConflicts)

        $domain = $env:USERDNSDOMAIN
        if ([string]::IsNullOrWhiteSpace($domain)) {
            try { $domain = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Domain } catch { $domain = $null }
        }

        $event = [ordered]@{
            EventType                 = 'maintenance.execution'
            ComputerName              = $env:COMPUTERNAME
            Domain                    = $domain
            ScriptName                = $script:ScriptName
            ScriptVersion             = $script:ScriptVersion
            Status                    = $script:FinalStatus
            ExitCode                  = $script:ExitCode
            StartTime                 = $script:ExecutionStart.ToString('o')
            EndTime                   = $end.ToString('o')
            DurationSeconds           = $duration
            ErrorCount                = $script:ErrorCount
            WarningCount              = $script:WarningCount
            Timestamp                 = (Get-Date).ToUniversalTime().ToString('o')
            FailureStage              = $script:FailureStage
            FailureMessage            = $script:FailureMessage
            RebootInitiated           = $script:RebootInitiated
            RebootReason              = $script:RebootReason
            TaskReconciliationStatus  = $script:TaskReconcileStatus
            PolicyConflictDetected    = ($script:PolicyConflictsBefore.Count -gt 0)
            PolicyConflictsBefore     = @($script:PolicyConflictsBefore)
            PolicyConflictsAfter      = $policyConflictsAfter
            ServicesChecked           = $script:ServiceNames.Count
            ServicesRunningAfter      = $servicesRunning
            ServicesFailed            = $servicesFailed
            ServicesRepairedCount     = @($serviceChanges).Count
            ServiceChanges            = @($serviceChanges)
            ServiceStateBefore        = @($script:ServiceStateBefore)
            ServiceStateAfter         = @($script:ServiceStateAfter)
            ApprovedServiceState      = ($servicesFailed.Count -eq 0)
            TextLogPath                = $script:PublishedLogPath
        }

        $jsonCompact = $event | ConvertTo-Json -Depth 8 -Compress
        Write-MaintenanceTelemetryLine -Path $script:NdjsonPath -JsonLine $jsonCompact
        $event | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:LatestJsonPath -Encoding UTF8
    }
    catch {
        Write-Host "Telemetry write failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-ServiceStartRegistry {
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName,

        [Parameter(Mandatory)]
        [int]$StartValue
    )

    $paths = @(
        "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName",
        "HKLM:\SYSTEM\ControlSet001\Services\$ServiceName"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            try {
                Set-ItemProperty -Path $path -Name Start -Value $StartValue -Type DWord -ErrorAction Stop
                Write-Status "Set registry Start=$StartValue for $ServiceName at $path" 'OK'
            }
            catch {
                Write-Status "Failed setting Start for $ServiceName at $path : $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Status "Registry path not found for $ServiceName at $path" 'WARN'
        }
    }
}

function Set-ServiceStartupAndStart {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Automatic','Manual')]
        [string]$StartupType
    )

    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop

        try {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction Stop
            Write-Status "Set startup type for $Name to $StartupType" 'OK'
        }
        catch {
            Write-Status "Set-Service failed for $Name. Trying sc.exe config..." 'WARN'
            $startValue = if ($StartupType -eq 'Automatic') { 'auto' } else { 'demand' }
            & sc.exe config $Name start= $startValue | Out-Null
            Write-Status "Configured startup type for $Name via sc.exe" 'OK'
        }

        try {
            Start-Service -Name $Name -ErrorAction Stop
            Write-Status "Started service: $Name" 'OK'
        }
        catch {
            Write-Status "Could not start service $Name immediately: $($_.Exception.Message)" 'WARN'
        }
    }
    catch {
        Write-Status "Service not found or inaccessible: $Name" 'WARN'
    }
}

function Get-ServiceStateSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        return (Get-Service -Name $Name -ErrorAction Stop).Status
    }
    catch {
        return $null
    }
}

function Wait-ForServiceRunning {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [int]$TimeoutSeconds = 15
    )

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($stopWatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $status = Get-ServiceStateSafe -Name $Name
        if ($status -eq 'Running') {
            return $true
        }

        Start-Sleep -Seconds 2
    }

    return $false
}

function Ensure-ServiceRunningWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Automatic','Manual')]
        [string]$StartupType,

        [int]$MaxAttempts = 4,

        [int]$WaitPerAttemptSeconds = 15
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $currentState = Get-ServiceStateSafe -Name $Name

        if ($currentState -eq 'Running') {
            Write-Status "Service $Name is already running." 'OK'
            return $true
        }

        Write-Status "Attempt $attempt of $MaxAttempts to start service $Name..." 'INFO'

        try {
            Set-ServiceStartupAndStart -Name $Name -StartupType $StartupType
        }
        catch {
            Write-Status "Unexpected error while attempting to start $Name : $($_.Exception.Message)" 'WARN'
        }

        if (Wait-ForServiceRunning -Name $Name -TimeoutSeconds $WaitPerAttemptSeconds) {
            Write-Status "Verified service is running: $Name" 'OK'
            return $true
        }

        $stateAfterWait = Get-ServiceStateSafe -Name $Name
        Write-Status "Service $Name did not reach Running state after attempt $attempt. Current state: $stateAfterWait" 'WARN'

        if ($attempt -lt $MaxAttempts) {
            Start-Sleep -Seconds 5
        }
    }

    Write-Status "Service $Name failed to reach Running state after $MaxAttempts attempts." 'ERROR'
    return $false
}

function Force-RebootNow {
    param(
        [string]$Reason = 'Required Windows Update services failed to start after multiple attempts.'
    )

    $script:RebootInitiated = $true
    $script:RebootReason = $Reason
    $script:FinalStatus = 'FailedRebootScheduled'
    $script:ExitCode = 1
    $script:FailureStage = 'ServiceRecovery'
    $script:FailureMessage = $Reason
    Write-Status "FORCING REBOOT: $Reason" 'ERROR'

    try {
        shutdown.exe /r /f /t 30 /c "$Reason" | Out-Null
        Write-Status "Forced reboot command issued successfully. System will restart in 30 seconds." 'ERROR'
    }
    catch {
        Write-Status "Failed to issue shutdown.exe reboot command: $($_.Exception.Message)" 'ERROR'
    }

    # Throw instead of exiting so the script-level catch/finally can publish
    # structured failure telemetry before the scheduled reboot occurs.
    throw $Reason
}

function Enable-ScheduledTaskSafe {
    param(
        [Parameter(Mandatory)]
        [string]$TaskPath,

        [Parameter(Mandatory)]
        [string]$TaskName
    )

    try {
        $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
        if ($task.State -eq 'Disabled') {
            Enable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
            Write-Status "Enabled scheduled task: $TaskPath$TaskName" 'OK'
        }
        else {
            Write-Status "Scheduled task already enabled or available: $TaskPath$TaskName" 'INFO'
        }
    }
    catch {
        Write-Status "Scheduled task not found or could not be enabled: $TaskPath$TaskName" 'WARN'
    }
}

function Remove-RegistryValueSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        if (Test-Path $Path) {
            $prop = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $prop) {
                Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                Write-Status "Removed $Path\$Name" 'OK'
            }
            else {
                Write-Status "Registry value not present: $Path\$Name" 'INFO'
            }
        }
        else {
            Write-Status "Registry path not present: $Path" 'INFO'
        }
    }
    catch {
        Write-Status "Failed to remove $Path\$Name : $($_.Exception.Message)" 'WARN'
    }
}

function Set-RegistryDwordSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [int]$Value
    )

    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }

        New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force | Out-Null
        Write-Status "Set $Path\$Name = $Value" 'OK'
    }
    catch {
        Write-Status "Failed to set $Path\$Name : $($_.Exception.Message)" 'ERROR'
    }
}

function Get-ScriptHeaderValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$HeaderName
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $pattern = '^\s*#\s*' + [regex]::Escape($HeaderName) + '\s*:\s*(.+?)\s*$'
        $match = Select-String -LiteralPath $Path -Pattern $pattern -CaseSensitive:$false -ErrorAction Stop | Select-Object -First 1
        if ($match -and $match.Matches.Count -gt 0) {
            return $match.Matches[0].Groups[1].Value.Trim()
        }
    }
    catch {
        Write-Status "Unable to read $HeaderName from $Path : $($_.Exception.Message)" 'WARN'
    }

    return $null
}

function Convert-ToVersionObjectSafe {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$VersionText
    )

    if ([string]::IsNullOrWhiteSpace($VersionText)) {
        return [version]'0.0'
    }

    try {
        return [version]$VersionText.Trim()
    }
    catch {
        $cleanVersion = ($VersionText -replace '[^0-9\.]', '').Trim('.')
        if ([string]::IsNullOrWhiteSpace($cleanVersion)) {
            return [version]'0.0'
        }

        try {
            return [version]$cleanVersion
        }
        catch {
            Write-Status "Unable to parse version [$VersionText]. Treating as 0.0." 'WARN'
            return [version]'0.0'
        }
    }
}

function Get-WeeklySundayTriggerSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Triggers
    )

    foreach ($trigger in $Triggers) {
        try {
            $days = [string]$trigger.DaysOfWeek
            $startBoundary = [string]$trigger.StartBoundary
            $startTime = $null

            if (-not [string]::IsNullOrWhiteSpace($startBoundary)) {
                $startTime = ([datetime]$startBoundary).ToString('HH:mm')
            }

            if ($days -match 'Sunday' -and $startTime -eq '01:00') {
                return $true
            }
        }
        catch {
        }
    }

    return $false
}

function Test-CheckForUpdatedScriptsTaskCompliant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Task,

        [Parameter(Mandatory)]
        [string]$ExpectedScriptPath,

        [Parameter(Mandatory)]
        [string]$ExpectedPowerShellExe
    )

    $issues = New-Object System.Collections.Generic.List[string]

    $action = $Task.Actions | Select-Object -First 1
    if ($null -eq $action) {
        $issues.Add('Task has no action.') | Out-Null
    }
    else {
        $actualExecute = [string]$action.Execute
        $actualArguments = [string]$action.Arguments

        if (($actualExecute -ine $ExpectedPowerShellExe) -and ((Split-Path -Leaf $actualExecute) -ine 'powershell.exe')) {
            $issues.Add("Action executable is not Windows PowerShell. Current: $actualExecute") | Out-Null
        }

        if ($actualArguments -notmatch [regex]::Escape($ExpectedScriptPath)) {
            $issues.Add("Action does not point to expected updater script. Current arguments: $actualArguments") | Out-Null
        }

        if ($actualArguments -notmatch '-NoProfile') {
            $issues.Add('Action is missing -NoProfile.') | Out-Null
        }

        if ($actualArguments -notmatch '-ExecutionPolicy\s+Bypass') {
            $issues.Add('Action is missing -ExecutionPolicy Bypass.') | Out-Null
        }
    }

    if ($Task.Principal.UserId -notmatch 'SYSTEM') {
        $issues.Add("Task principal is not SYSTEM. Current: $($Task.Principal.UserId)") | Out-Null
    }

    if ($Task.Principal.RunLevel -ne 'Highest') {
        $issues.Add("Task run level is not Highest. Current: $($Task.Principal.RunLevel)") | Out-Null
    }

    if (-not (Get-WeeklySundayTriggerSummary -Triggers $Task.Triggers)) {
        $issues.Add('Task trigger is not Sunday at 01:00.') | Out-Null
    }

    return $issues
}


function Ensure-CheckForUpdatedScriptsTask {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,

        [string]$TaskName = '01. Check for Updated Scripts'
    )

    $windowsPowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    if (-not (Test-Path -LiteralPath $windowsPowerShellExe)) {
        Write-Status "Windows PowerShell executable not found: $windowsPowerShellExe" 'ERROR'
        return $false
    }

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Status "Cannot register task because updater script is missing: $ScriptPath" 'ERROR'
        return $false
    }

    try {
        Import-Module ScheduledTasks -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Status "ScheduledTasks module is unavailable: $($_.Exception.Message)" 'ERROR'
        return $false
    }

    $argumentString = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
    $needsRebuild = $true

    try {
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        $issues = Test-CheckForUpdatedScriptsTaskCompliant -Task $existingTask -ExpectedScriptPath $ScriptPath -ExpectedPowerShellExe $windowsPowerShellExe

        if ($issues.Count -eq 0) {
            Write-Status "Scheduled task '$TaskName' is already configured correctly." 'OK'
            $needsRebuild = $false
        }
        else {
            Write-Status "Scheduled task '$TaskName' needs repair:" 'WARN'
            foreach ($issue in $issues) {
                Write-Status " - $issue" 'WARN'
            }
        }
    }
    catch {
        Write-Status "Scheduled task '$TaskName' does not exist and will be created." 'WARN'
    }

    if (-not $needsRebuild) {
        return $true
    }

    try {
        $action = New-ScheduledTaskAction -Execute $windowsPowerShellExe -Argument $argumentString
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -WeeksInterval 1 -At ([datetime]::Today.Add([timespan]::Parse('01:00')))
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Hours 12)

        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Write-Status "Scheduled task '$TaskName' now points to: $ScriptPath" 'OK'
        Write-Status "Task action: $windowsPowerShellExe $argumentString" 'INFO'
        Write-Status "Task trigger: Sunday at 01:00 as SYSTEM with highest privileges." 'OK'
        return $true
    }
    catch {
        Write-Status "Failed to create/repair scheduled task '$TaskName': $($_.Exception.Message)" 'ERROR'
        return $false
    }
}


try {
    if (-not (Test-IsAdmin)) {
        Write-Status 'This script must be run as Administrator.' 'ERROR'
        $script:FinalStatus = 'Failed'
        $script:ExitCode = 1
        $script:FailureStage = 'PrivilegeCheck'
        $script:FailureMessage = 'Administrative privileges are required.'
        throw $script:FailureMessage
    }

    $script:ServiceStateBefore = @(Get-ServiceTelemetrySnapshot -Names $script:ServiceNames)
    $script:PolicyConflictsBefore = @(Get-WindowsUpdatePolicyConflicts)
Write-Status "Initializing script..." 'INFO'
# Restore registry startup values first
# Common defaults used for Windows Update-related services:
# wuauserv = Manual (3)
# bits = Manual (3)
# dosvc = Automatic family (2)
# UsoSvc = Automatic (2)
# WaaSMedicSvc = Manual/triggered on many systems (3)

Set-ServiceStartRegistry -ServiceName 'wuauserv'     -StartValue 3
Set-ServiceStartRegistry -ServiceName 'bits'         -StartValue 3
Set-ServiceStartRegistry -ServiceName 'dosvc'        -StartValue 2
Set-ServiceStartRegistry -ServiceName 'UsoSvc'       -StartValue 2
Set-ServiceStartRegistry -ServiceName 'WaaSMedicSvc' -StartValue 3

# Initial restore and startup
Set-ServiceStartupAndStart -Name 'wuauserv' -StartupType Manual
Set-ServiceStartupAndStart -Name 'bits'     -StartupType Manual
Set-ServiceStartupAndStart -Name 'dosvc'    -StartupType Automatic
Set-ServiceStartupAndStart -Name 'UsoSvc'   -StartupType Automatic

# WaaSMedicSvc can be protected; set registry above, then try starting via sc.exe
try {
    & sc.exe config WaaSMedicSvc start= demand | Out-Null
    Write-Status "Configured WaaSMedicSvc startup via sc.exe" 'OK'
}
catch {
    Write-Status "Could not configure WaaSMedicSvc via sc.exe" 'WARN'
}

try {
    & sc.exe start WaaSMedicSvc | Out-Null
    Write-Status "Attempted to start WaaSMedicSvc" 'INFO'
}
catch {
    Write-Status "Could not start WaaSMedicSvc directly" 'WARN'
}

# Restore Automatic Updates policy
$wuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$auPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'

Set-RegistryDwordSafe -Path $auPolicyPath -Name 'NoAutoUpdate' -Value 0
Set-RegistryDwordSafe -Path $auPolicyPath -Name 'AUOptions'    -Value 3

# Remove common WSUS redirection values if they were previously set
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUServer'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUStatusServer'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'UpdateServiceUrlAlternate'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'SetProxyBehaviorForUpdateDetection'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'DisableWindowsUpdateAccess'
Remove-RegistryValueSafe -Path $auPolicyPath -Name 'UseWUServer'

# Re-enable common update scheduled tasks
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

foreach ($task in $tasks) {
    Enable-ScheduledTaskSafe -TaskPath $task.Path -TaskName $task.Name
}

# Restart key services in a sensible order
$restartOrder = @('bits', 'dosvc', 'wuauserv', 'UsoSvc')
foreach ($svc in $restartOrder) {
    try {
        Restart-Service -Name $svc -Force -ErrorAction Stop
        Write-Status "Restarted service: $svc" 'OK'
    }
    catch {
        Write-Status "Could not restart $svc : $($_.Exception.Message)" 'WARN'
    }
}

# Verify and retry critical services
$requiredServices = @(
    @{ Name = 'wuauserv'; StartupType = 'Manual' },
    @{ Name = 'bits';     StartupType = 'Manual' },
    @{ Name = 'dosvc';    StartupType = 'Automatic' },
    @{ Name = 'UsoSvc';   StartupType = 'Automatic' }
)

$failedServices = @()

foreach ($requiredService in $requiredServices) {
    $serviceStarted = Ensure-ServiceRunningWithRetry -Name $requiredService.Name -StartupType $requiredService.StartupType -MaxAttempts 4 -WaitPerAttemptSeconds 15
    if (-not $serviceStarted) {
        $failedServices += $requiredService.Name
    }
}

if ($failedServices.Count -gt 0) {
    $failedList = $failedServices -join ', '
    Write-Status "One or more critical Windows Update services failed to start: $failedList" 'ERROR'
    Force-RebootNow -Reason "Windows Update service recovery failed. Services not running: $failedList"
}


# ------------------------------------------------------------
# Scheduled-task reconciliation
# The register script is idempotent: it creates missing tasks and changes only
# managed tasks whose action, trigger, principal, or settings differ.
# ------------------------------------------------------------
$RegisterTasksScript = 'C:\Scripts\Register-Tasks_SYSTEM.ps1'

if (Test-Path -LiteralPath $RegisterTasksScript -PathType Leaf) {
    Write-Status 'Reconciling managed scheduled tasks...' 'INFO'

    try {
        $taskReconcileProcess = Start-Process `
            -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$RegisterTasksScript`"" `
            -Wait `
            -PassThru `
            -WindowStyle Hidden

        if ($taskReconcileProcess.ExitCode -eq 0) {
            $script:TaskReconcileStatus = 'Success'
            Write-Status 'Managed scheduled tasks are current.' 'OK'
        }
        else {
            $script:TaskReconcileStatus = "ExitCode:$($taskReconcileProcess.ExitCode)"
            Write-Status "Task reconciliation exited with code $($taskReconcileProcess.ExitCode)." 'WARN'
        }
    }
    catch {
        $script:TaskReconcileStatus = 'Failed'
        Write-Status "Task reconciliation failed: $($_.Exception.Message)" 'WARN'
    }
}
else {
    $script:TaskReconcileStatus = 'ScriptMissing'
    Write-Status "Task reconciliation script is missing: $RegisterTasksScript" 'WARN'
}

# Final enforcement pass. This runs after any task rebuild so older register scripts
# cannot leave "Run task as soon as possible after a scheduled start is missed" enabled.
$script:FinalStatus = if ($script:WarningCount -gt 0) { 'SuccessWithWarnings' } else { 'Success' }
$script:ExitCode = 0
Write-Status "Windows Update settings have been restored and critical services are running." 'OK'
Write-Status "No reboot required. Continuing normally." 'INFO'
}
catch {
    if ($script:ExitCode -eq 0) { $script:ExitCode = 1 }
    if ($script:FinalStatus -eq 'Running') { $script:FinalStatus = 'Failed' }
    if ([string]::IsNullOrWhiteSpace($script:FailureStage)) { $script:FailureStage = 'Execution' }
    if ([string]::IsNullOrWhiteSpace($script:FailureMessage)) { $script:FailureMessage = $_.Exception.Message }
    Write-Status "Unhandled failure: $($_.Exception.Message)" 'ERROR'
}
finally {
    $script:ServiceStateAfter = @(Get-ServiceTelemetrySnapshot -Names $script:ServiceNames)
    Write-ExecutionTelemetry

    # This is the final text-log write. Once published into C:\Logs, Elastic may
    # open the file immediately, so nothing should append to it afterward.
    Write-Status ("Completed {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
        $script:ScriptName,
        $script:FinalStatus,
        $script:ExitCode,
        $script:WarningCount,
        $script:ErrorCount) $(if ($script:ExitCode -eq 0) { 'OK' } else { 'ERROR' })

    if ($null -ne $script:LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $script:LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 01 text log for Elastic: {0}" -f $script:PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 01 completed log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $script:ExitCode
