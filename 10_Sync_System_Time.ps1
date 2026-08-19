#requires -Version 5.1
<#
.SYNOPSIS
    Configures and verifies Windows Time synchronization.

.DESCRIPTION
    Ensures the Windows Time service is configured and running, applies the
    appropriate synchronization source, performs a forced rediscovery/resync,
    verifies the final status, and writes structured maintenance telemetry.

    By default, domain-joined computers use the Active Directory domain
    hierarchy. Workgroup computers use the configured manual NTP peers.

.NOTES
    ScriptName: 10_Sync_System_Time.ps1
    Version:    2.0.10
    LastUpdated: 2026-08-17
    Changes: v2.0.10 fixes telemetry finalization parameter binding introduced in v2.0.9; New-TelemetryEvent now receives all parameters before the summary log is written.
             v2.0.9 makes Script 10 telemetry mapping-safe by normalizing Changes.Before/After into scalar or JSON-string fields
             and forces ManualNtpServers to serialize as a true JSON array when empty.
             v2.0.8 moves Script 10-specific telemetry under a dedicated time_synchronization namespace,
             adds shared-process protection before force-terminating w32time, and moves active staged logging
             to C:\ProgramData\Compton\Maintenance-Logs\Staging.
             v2.0.7 uses Maintenance.Framework v2.4 staged text logging.
#>

[CmdletBinding()]
param(
    [ValidateSet('Auto', 'DomainHierarchy', 'Manual')]
    [string]$SyncMode = 'Auto',

    [ValidateNotNullOrEmpty()]
    [string[]]$NtpServers = @(
        'time.windows.com,0x9'
    ),

    [ValidateRange(10, 300)]
    [int]$ServiceStopTimeoutSeconds = 30,

    [ValidateRange(1, 10)]
    [int]$ResyncAttempts = 3,

    [ValidateRange(2, 120)]
    [int]$ResyncRetryDelaySeconds = 10,

    [ValidateRange(1, 168)]
    [int]$MaximumSyncAgeHours = 24,

    [ValidateRange(1, 86400)]
    [int]$SpecialPollIntervalSeconds = 3600,

    [ValidateRange(0.001, 3600000)]
    [double]$ClockOffsetAlertThresholdMilliseconds = 5000,

    [string]$LogDirectory = 'C:\Logs'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ScriptName = '10_Sync_System_Time.ps1'
$script:ScriptVersion = '2.0.10'
$script:RunId = [guid]::NewGuid().Guid
$script:StartTime = Get-Date
$script:Warnings = @()
$script:Errors = @()
$script:CommandResults = @()
$script:Changes = @()
$script:FinalExitCode = 0
$script:FinalStatus = 'Failed'
$script:OverallResult = 'Failed'
$script:FailureMessage = $null

$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestJsonPath = Join-Path $LogDirectory '10_Sync_System_Time.latest.json'
$RuntimeLogPath = $null
$PublishedLogPath = $null
$LogSession = $null
$DedicatedStagingRoot = 'C:\ProgramData\Compton\Maintenance-Logs\Staging'

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 10 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName $script:ScriptName `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '10_Sync_System_Time.log',
        '*-10_Sync_System_Time-*.log'
    ) | Out-Null

if (-not (Test-Path -LiteralPath $DedicatedStagingRoot -PathType Container)) {
    New-Item -Path $DedicatedStagingRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
}

$LogSession = New-MaintenanceStagedLog `
    -ScriptName $script:ScriptName `
    -LogRoot $LogDirectory `
    -StagingRoot $DedicatedStagingRoot `
    -ComputerName $env:COMPUTERNAME `
    -Timestamp $script:StartTime

$RuntimeLogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
    $normalizedLevel = if ($Level -eq 'WARN') { 'WARNING' } else { $Level }
    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $computerName, $normalizedLevel, $Message
    try {
        $activeLogDirectory = Split-Path -Parent $RuntimeLogPath
        Ensure-Directory -Path $activeLogDirectory
        Add-Content -LiteralPath $RuntimeLogPath -Value $line -Encoding UTF8
    }
    catch {
        # Logging must not hide the primary operation result.
    }

    Write-Host $line
}

function Add-WarningMessage {
    param([Parameter(Mandatory)][string]$Message)

    $script:Warnings += [string]$Message
    Write-Log -Message $Message -Level WARN
}

function Add-ErrorMessage {
    param([Parameter(Mandatory)][string]$Message)

    $script:Errors += [string]$Message
    Write-Log -Message $Message -Level ERROR
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-ExternalCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$FilePath,
        [string[]]$Arguments = @(),
        [int[]]$SuccessExitCodes = @(0)
    )

    $started = Get-Date
    $output = @()
    $exitCode = $null
    $succeeded = $false
    $exceptionMessage = $null

    try {
        $output = @(& $FilePath @Arguments 2>&1 | ForEach-Object { [string]$_ })
        $exitCode = $LASTEXITCODE
        $succeeded = $SuccessExitCodes -contains $exitCode
    }
    catch {
        $exceptionMessage = $_.Exception.Message
    }

    $finished = Get-Date
    $result = [pscustomobject]@{
        Name            = $Name
        FilePath        = $FilePath
        Arguments       = @($Arguments)
        ExitCode        = $exitCode
        Success         = $succeeded
        StartTime       = $started.ToUniversalTime().ToString('o')
        EndTime         = $finished.ToUniversalTime().ToString('o')
        DurationSeconds = [math]::Round(($finished - $started).TotalSeconds, 3)
        Output          = @($output)
        Exception       = $exceptionMessage
    }

    $script:CommandResults += $result
    return $result
}

function Get-WindowsInformation {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $cv = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

    [pscustomobject]@{
        ProductName    = [string]$cv.ProductName
        EditionId      = [string]$cv.EditionID
        DisplayVersion = [string]$cv.DisplayVersion
        ReleaseId      = [string]$cv.ReleaseId
        BuildNumber    = [string]$os.BuildNumber
        UBR            = [int]$cv.UBR
        FullBuild      = '{0}.{1}' -f $os.BuildNumber, $cv.UBR
    }
}

function Get-DomainInformation {
    $computer = Get-CimInstance -ClassName Win32_ComputerSystem

    [pscustomobject]@{
        PartOfDomain = [bool]$computer.PartOfDomain
        Domain       = [string]$computer.Domain
        Workgroup    = if ($computer.PartOfDomain) { $null } else { [string]$computer.Workgroup }
    }
}

function Get-ServiceSnapshot {
    $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='w32time'" -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        return [pscustomobject]@{
            Exists      = $false
            State       = $null
            StartMode   = $null
            ProcessId   = $null
            ExitCode    = $null
        }
    }

    [pscustomobject]@{
        Exists      = $true
        State       = [string]$service.State
        StartMode   = [string]$service.StartMode
        ProcessId   = [int]$service.ProcessId
        ExitCode    = [int]$service.ExitCode
    }
}

function Convert-W32TimeValue {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $trimmed = $Value.Trim()
    if ($trimmed -match '^(?i)unspecified$' -or $trimmed -match '^(?i)n/a$') {
        return $null
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($trimmed, [ref]$parsed)) {
        return $parsed
    }

    return $trimmed
}

function Convert-PhaseOffsetToMeasurement {
    param([AllowNull()][string]$PhaseOffset)

    $measurement = [ordered]@{
        Raw                  = $PhaseOffset
        Milliseconds         = $null
        AbsoluteMilliseconds = $null
        Direction            = 'Unknown'
    }

    if ([string]::IsNullOrWhiteSpace($PhaseOffset)) {
        return [pscustomobject]$measurement
    }

    $match = [regex]::Match($PhaseOffset.Trim(), '([+-]?\d+(?:[\.,]\d+)?)\s*(ms|us|µs|s)\b', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
        return [pscustomobject]$measurement
    }

    $numericText = $match.Groups[1].Value.Replace(',', '.')
    $value = 0.0
    if (-not [double]::TryParse($numericText, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
        return [pscustomobject]$measurement
    }

    $milliseconds = switch ($match.Groups[2].Value.ToLowerInvariant()) {
        's'  { $value * 1000.0 }
        'ms' { $value }
        'us' { $value / 1000.0 }
        'µs' { $value / 1000.0 }
    }

    $milliseconds = [math]::Round($milliseconds, 6)
    $measurement.Milliseconds = $milliseconds
    $measurement.AbsoluteMilliseconds = [math]::Round([math]::Abs($milliseconds), 6)
    $measurement.Direction = if ($milliseconds -gt 0) { 'Ahead' } elseif ($milliseconds -lt 0) { 'Behind' } else { 'Synchronized' }
    return [pscustomobject]$measurement
}

function Convert-W32tmOutputToObject {
    param([string[]]$Lines)

    $properties = [ordered]@{}
    foreach ($line in $Lines) {
        if ($line -match '^\s*([^:]+):\s*(.*)$') {
            $name = $matches[1].Trim()
            $value = $matches[2].Trim()
            if (-not $properties.Contains($name)) {
                $properties[$name] = $value
            }
        }
    }

    return [pscustomobject]$properties
}

function Get-TimeConfigurationSnapshot {
    $parametersPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
    $clientPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient'
    $configPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config'

    $parameters = Get-ItemProperty -LiteralPath $parametersPath -ErrorAction SilentlyContinue
    $client = Get-ItemProperty -LiteralPath $clientPath -ErrorAction SilentlyContinue
    $config = Get-ItemProperty -LiteralPath $configPath -ErrorAction SilentlyContinue

    [pscustomobject]@{
        Type                       = [string]$parameters.Type
        NtpServer                  = [string]$parameters.NtpServer
        NtpClientEnabled           = [int]$client.Enabled
        SpecialPollIntervalSeconds = [int]$client.SpecialPollInterval
        AnnounceFlags              = [int]$config.AnnounceFlags
        MaxNegPhaseCorrection      = [uint32]$config.MaxNegPhaseCorrection
        MaxPosPhaseCorrection      = [uint32]$config.MaxPosPhaseCorrection
    }
}

function Get-TimeStatusSnapshot {
    $w32tm = Join-Path $env:SystemRoot 'System32\w32tm.exe'
    $statusCommand = Invoke-ExternalCommand -Name 'QueryStatus' -FilePath $w32tm -Arguments @('/query', '/status', '/verbose')
    $sourceCommand = Invoke-ExternalCommand -Name 'QuerySource' -FilePath $w32tm -Arguments @('/query', '/source')
    $configCommand = Invoke-ExternalCommand -Name 'QueryConfiguration' -FilePath $w32tm -Arguments @('/query', '/configuration')

    $parsedStatus = Convert-W32tmOutputToObject -Lines $statusCommand.Output

    # /query /status already contains the authoritative Source field and may
    # succeed on systems where /query /source returns Access Denied (0x80070005).
    $source = if ($parsedStatus.PSObject.Properties['Source'] -and -not [string]::IsNullOrWhiteSpace([string]$parsedStatus.Source)) {
        [string]$parsedStatus.Source
    }
    elseif ($sourceCommand.Success -and $sourceCommand.Output.Count -gt 0) {
        ($sourceCommand.Output -join ' ').Trim()
    }
    else {
        $null
    }

    $lastSyncRaw = $null
    if ($null -ne $parsedStatus.PSObject.Properties['Last Successful Sync Time']) {
        $lastSyncRaw = [string]$parsedStatus.'Last Successful Sync Time'
    }

    $lastSyncValue = Convert-W32TimeValue -Value $lastSyncRaw
    $lastSyncDate = if ($lastSyncValue -is [datetime]) { $lastSyncValue } else { $null }
    $syncAgeHours = if ($null -ne $lastSyncDate) {
        [math]::Round(((Get-Date) - $lastSyncDate).TotalHours, 3)
    }
    else {
        $null
    }

    $phaseOffsetRaw = if ($parsedStatus.PSObject.Properties['Phase Offset']) { [string]$parsedStatus.'Phase Offset' } else { $null }
    $offsetMeasurement = Convert-PhaseOffsetToMeasurement -PhaseOffset $phaseOffsetRaw

    [pscustomobject]@{
        QuerySucceeded         = [bool]$statusCommand.Success
        Source                 = $source
        LastSuccessfulSyncTime = if ($null -ne $lastSyncDate) { $lastSyncDate.ToUniversalTime().ToString('o') } else { $lastSyncRaw }
        SyncAgeHours           = $syncAgeHours
        LeapIndicator          = if ($parsedStatus.PSObject.Properties['Leap Indicator']) { [string]$parsedStatus.'Leap Indicator' } else { $null }
        Stratum                = if ($parsedStatus.PSObject.Properties['Stratum']) { [string]$parsedStatus.Stratum } else { $null }
        Precision              = if ($parsedStatus.PSObject.Properties['Precision']) { [string]$parsedStatus.Precision } else { $null }
        RootDelay              = if ($parsedStatus.PSObject.Properties['Root Delay']) { [string]$parsedStatus.'Root Delay' } else { $null }
        RootDispersion         = if ($parsedStatus.PSObject.Properties['Root Dispersion']) { [string]$parsedStatus.'Root Dispersion' } else { $null }
        ReferenceId            = if ($parsedStatus.PSObject.Properties['ReferenceId']) { [string]$parsedStatus.ReferenceId } else { $null }
        PhaseOffset            = $offsetMeasurement.Raw
        ClockOffsetMilliseconds = $offsetMeasurement.Milliseconds
        AbsoluteClockOffsetMilliseconds = $offsetMeasurement.AbsoluteMilliseconds
        ClockOffsetDirection   = $offsetMeasurement.Direction
        PollInterval           = if ($parsedStatus.PSObject.Properties['Poll Interval']) { [string]$parsedStatus.'Poll Interval' } else { $null }
        StatusOutput           = @($statusCommand.Output)
        ConfigurationOutput    = @($configCommand.Output)
    }
}

function Write-TimeStatusSummary {
    param(
        [Parameter(Mandatory)][string]$Label,
        $Status
    )

    if ($null -eq $Status) {
        Write-Log -Message ("Time status {0}: unavailable." -f $Label) -Level WARN
        return
    }

    $lastSync = if ($Status.LastSuccessfulSyncTime) { $Status.LastSuccessfulSyncTime } else { 'Unknown' }
    $syncAge = if ($null -ne $Status.SyncAgeHours) { '{0} hours' -f $Status.SyncAgeHours } else { 'Unknown' }
    $source = if ($Status.Source) { $Status.Source } else { 'Unknown' }
    $stratum = if ($Status.Stratum) { $Status.Stratum } else { 'Unknown' }
    $phaseOffset = if ($Status.PhaseOffset) { $Status.PhaseOffset } else { 'Unknown' }
    $clockOffset = if ($null -ne $Status.ClockOffsetMilliseconds) { '{0} ms' -f $Status.ClockOffsetMilliseconds } else { 'Unknown' }
    $absoluteOffset = if ($null -ne $Status.AbsoluteClockOffsetMilliseconds) { '{0} ms' -f $Status.AbsoluteClockOffsetMilliseconds } else { 'Unknown' }
    $direction = if ($Status.ClockOffsetDirection) { $Status.ClockOffsetDirection } else { 'Unknown' }

    Write-Log -Message ("Time status {0}: Source={1}; LastSuccessfulSync={2}; SyncAge={3}; Stratum={4}; PhaseOffset={5}; ClockOffset={6}; AbsoluteClockOffset={7}; Direction={8}." -f `
        $Label, $source, $lastSync, $syncAge, $stratum, $phaseOffset, $clockOffset, $absoluteOffset, $direction)
}

function Get-RecentTimeServiceEvents {
    param([int]$Hours = 24)

    $start = (Get-Date).AddHours(-$Hours)
    $events = @()

    try {
        $events = @(Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = 'Microsoft-Windows-Time-Service'
            StartTime    = $start
        } -ErrorAction Stop | Select-Object -First 50)
    }
    catch {
        Add-WarningMessage -Message ('Unable to read recent Windows Time events: {0}' -f $_.Exception.Message)
    }

    @($events | ForEach-Object {
        [pscustomobject]@{
            TimeCreated = if ($_.TimeCreated) { $_.TimeCreated.ToUniversalTime().ToString('o') } else { $null }
            Id          = $_.Id
            Level       = $_.LevelDisplayName
            Message     = [string]$_.Message
        }
    })
}

function Stop-W32TimeSafely {
    $service = Get-Service -Name 'w32time' -ErrorAction Stop
    if ($service.Status -eq 'Stopped') {
        return [pscustomobject]@{
            WasRunning             = $false
            Stopped                = $true
            ForcedKill             = $false
            ForcedKillSkipped      = $false
            ProcessId              = 0
            HostedServices         = @()
            ForcedKillReason       = $null
        }
    }

    Stop-Service -Name 'w32time' -Force -ErrorAction SilentlyContinue
    $deadline = (Get-Date).AddSeconds($ServiceStopTimeoutSeconds)

    do {
        Start-Sleep -Milliseconds 500
        $service.Refresh()
    } while ($service.Status -ne 'Stopped' -and (Get-Date) -lt $deadline)

    $forcedKill = $false
    $forcedKillSkipped = $false
    $forcedKillReason = $null
    $processId = 0
    $hostedServices = @()

    if ($service.Status -ne 'Stopped') {
        $cim = Get-CimInstance -ClassName Win32_Service -Filter "Name='w32time'" -ErrorAction Stop
        $processId = [int]$cim.ProcessId

        if ($processId -gt 0) {
            $hostedServices = @(
                Get-CimInstance -ClassName Win32_Service -Filter ("ProcessId={0}" -f $processId) -ErrorAction Stop |
                Where-Object { [int]$_.ProcessId -eq $processId } |
                Select-Object -ExpandProperty Name
            )

            if ($hostedServices.Count -eq 1 -and $hostedServices[0] -ieq 'w32time') {
                Stop-Process -Id $processId -Force -ErrorAction Stop
                $forcedKill = $true
                $forcedKillReason = "Force-terminated exclusive w32time process PID $processId."
                Start-Sleep -Seconds 1
            }
            else {
                $forcedKillSkipped = $true
                $forcedKillReason = "Process PID $processId is shared by services: $($hostedServices -join ', '). Force termination was skipped for safety."
                Add-WarningMessage -Message $forcedKillReason
            }
        }
        else {
            $forcedKillSkipped = $true
            $forcedKillReason = 'w32time did not stop, but no active service process ID was available for safe termination.'
            Add-WarningMessage -Message $forcedKillReason
        }

        $service.Refresh()
    }

    [pscustomobject]@{
        WasRunning        = $true
        Stopped           = ($service.Status -eq 'Stopped')
        ForcedKill        = $forcedKill
        ForcedKillSkipped = $forcedKillSkipped
        ProcessId         = $processId
        HostedServices    = @($hostedServices)
        ForcedKillReason  = $forcedKillReason
    }
}

function Set-TimeServiceStartup {
    $before = Get-ServiceSnapshot
    $changed = $false

    if (-not $before.Exists) {
        throw 'The Windows Time service (w32time) is not installed.'
    }

    if ($before.StartMode -ne 'Auto') {
        Set-Service -Name 'w32time' -StartupType Automatic
        $changed = $true
    }

    $after = Get-ServiceSnapshot
    $script:Changes += [pscustomobject]@{
        Component = 'ServiceStartup'
        Before    = $before.StartMode
        After     = $after.StartMode
        Changed   = $changed
        Verified  = ($after.StartMode -eq 'Auto')
    }

    if ($after.StartMode -ne 'Auto') {
        throw 'Failed to configure the Windows Time service for automatic startup.'
    }
}

function Resolve-SyncMode {
    param(
        [Parameter(Mandatory)][string]$RequestedMode,
        [Parameter(Mandatory)]$DomainInfo
    )

    if ($RequestedMode -eq 'Auto') {
        if ($DomainInfo.PartOfDomain) {
            return 'DomainHierarchy'
        }

        return 'Manual'
    }

    return $RequestedMode
}

function Set-TimeConfiguration {
    param(
        [Parameter(Mandatory)][ValidateSet('DomainHierarchy', 'Manual')]
        [string]$ResolvedMode
    )

    $w32tm = Join-Path $env:SystemRoot 'System32\w32tm.exe'
    $before = Get-TimeConfigurationSnapshot
    $arguments = @()

    if ($ResolvedMode -eq 'DomainHierarchy') {
        $arguments = @('/config', '/syncfromflags:domhier', '/update')
    }
    else {
        $peerList = ($NtpServers -join ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($peerList)) {
            throw 'At least one NTP server is required for Manual synchronization mode.'
        }

        $arguments = @('/config', ('/manualpeerlist:{0}' -f $peerList), '/syncfromflags:manual', '/reliable:no', '/update')

        $clientPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient'
        Set-ItemProperty -LiteralPath $clientPath -Name 'SpecialPollInterval' -Type DWord -Value $SpecialPollIntervalSeconds
    }

    $result = Invoke-ExternalCommand -Name 'ConfigureWindowsTime' -FilePath $w32tm -Arguments $arguments
    if (-not $result.Success) {
        throw ('w32tm configuration failed with exit code {0}. Output: {1}' -f $result.ExitCode, ($result.Output -join ' | '))
    }

    $after = Get-TimeConfigurationSnapshot
    $expectedType = if ($ResolvedMode -eq 'DomainHierarchy') { 'NT5DS' } else { 'NTP' }
    $verified = ($after.Type -eq $expectedType)

    if ($ResolvedMode -eq 'Manual') {
        $requestedPeers = ($NtpServers -join ' ').Trim()
        $verified = $verified -and ($after.NtpServer -eq $requestedPeers)
    }

    $script:Changes += [pscustomobject]@{
        Component = 'TimeConfiguration'
        Before    = $before
        After     = $after
        Changed   = (($before.Type -ne $after.Type) -or ($before.NtpServer -ne $after.NtpServer) -or ($before.SpecialPollIntervalSeconds -ne $after.SpecialPollIntervalSeconds))
        Verified  = $verified
    }

    if (-not $verified) {
        throw ('Windows Time configuration verification failed. Expected synchronization type {0}; found {1}.' -f $expectedType, $after.Type)
    }

    return [pscustomobject]@{
        Before   = $before
        After    = $after
        Verified = $verified
    }
}

function Start-TimeServiceVerified {
    Start-Service -Name 'w32time' -ErrorAction Stop
    $service = Get-Service -Name 'w32time' -ErrorAction Stop
    $service.WaitForStatus('Running', [timespan]::FromSeconds(30))

    $snapshot = Get-ServiceSnapshot
    if ($snapshot.State -ne 'Running') {
        throw 'The Windows Time service did not reach the Running state.'
    }

    return $snapshot
}


function Test-TimeCommandAccessDenied {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$CommandResult
    )

    $outputText = @($CommandResult.Output) -join ' '

    return (
        $CommandResult.ExitCode -eq 5 -or
        $CommandResult.ExitCode -eq -2147024891 -or
        $outputText -match '0x80070005' -or
        $outputText -match '(?i)Access is denied'
    )
}

function Invoke-TimeResynchronization {
    $w32tm = Join-Path $env:SystemRoot 'System32\w32tm.exe'
    $attemptResults = @()

    $rediscover = Invoke-ExternalCommand -Name 'RediscoverTimeSource' -FilePath $w32tm -Arguments @('/resync', '/rediscover')
    $attemptResults += [pscustomobject]@{
        Attempt  = 0
        Type     = 'Rediscover'
        Success  = $rediscover.Success
        ExitCode = $rediscover.ExitCode
        Output   = @($rediscover.Output)
    }

    for ($attempt = 1; $attempt -le $ResyncAttempts; $attempt++) {
        $result = Invoke-ExternalCommand -Name ('ResyncAttempt{0}' -f $attempt) -FilePath $w32tm -Arguments @('/resync', '/force')
        $attemptResults += [pscustomobject]@{
            Attempt  = $attempt
            Type     = 'ForcedResync'
            Success  = $result.Success
            ExitCode = $result.ExitCode
            Output   = @($result.Output)
        }

        if ($result.Success) {
            return [pscustomobject]@{
                Success      = $true
                AccessDenied = $false
                Attempts     = $attempt
                Results      = $attemptResults
            }
        }

        if ($attempt -lt $ResyncAttempts) {
            Start-Sleep -Seconds $ResyncRetryDelaySeconds
        }
    }

    $accessDenied = $false
    foreach ($attemptResult in @($attemptResults)) {
        if (Test-TimeCommandAccessDenied -CommandResult $attemptResult) {
            $accessDenied = $true
            break
        }
    }

    return [pscustomobject]@{
        Success      = $false
        AccessDenied = $accessDenied
        Attempts     = $ResyncAttempts
        Results      = $attemptResults
    }
}

function Test-TimeHealth {
    param(
        [Parameter(Mandatory)]$Status,
        [Parameter(Mandatory)][string]$ResolvedMode
    )

    $reasons = @()

    if (-not $Status.QuerySucceeded) {
        $reasons += 'StatusQueryFailed'
    }

    if ([string]::IsNullOrWhiteSpace($Status.Source)) {
        $reasons += 'TimeSourceMissing'
    }
    elseif ($Status.Source -match '^(?i)(Local CMOS Clock|Free-running System Clock)$') {
        $reasons += 'UnsynchronizedLocalClock'
    }

    if ([string]$Status.LeapIndicator -match '^(?i)3\b|not synchronized') {
        $reasons += 'LeapIndicatorUnsynchronized'
    }

    if ($null -eq $Status.SyncAgeHours) {
        $reasons += 'LastSuccessfulSyncUnknown'
    }
    elseif ($Status.SyncAgeHours -gt $MaximumSyncAgeHours) {
        $reasons += 'LastSuccessfulSyncTooOld'
    }

    if ($ResolvedMode -eq 'DomainHierarchy' -and $Status.Source -match '^(?i)time\.windows\.com') {
        $reasons += 'UnexpectedPublicSourceForDomainHierarchy'
    }

    [pscustomobject]@{
        Healthy = ($reasons.Count -eq 0)
        Reasons = $reasons
    }
}

function Write-Telemetry {
    param(
        [Parameter(Mandatory)]$Event
    )

    Ensure-Directory -Path $LogDirectory

    $compact = $Event | ConvertTo-Json -Depth 12 -Compress
    Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compact

    $temporaryPath = '{0}.{1}.tmp' -f $LatestJsonPath, $script:RunId
    $Event | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $temporaryPath -Encoding UTF8
    Move-Item -LiteralPath $temporaryPath -Destination $LatestJsonPath -Force
}


function New-StringArrayForJson {
    param([object[]]$Values)

    [string[]]$items = @(
        $Values |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Convert-ChangeForTelemetry {
    param(
        [Parameter(Mandatory)]
        [object]$Change
    )

    $beforeValue = $null
    $afterValue  = $null
    $beforeJson  = $null
    $afterJson   = $null

    $before = $Change.Before
    $after  = $Change.After

    $beforeIsComplex = $false
    $afterIsComplex  = $false

    if ($null -ne $before) {
        $beforeIsComplex = (
            $before -is [System.Collections.IDictionary] -or
            ($before -is [System.Collections.IEnumerable] -and -not ($before -is [string])) -or
            ($before -is [pscustomobject])
        )
    }

    if ($null -ne $after) {
        $afterIsComplex = (
            $after -is [System.Collections.IDictionary] -or
            ($after -is [System.Collections.IEnumerable] -and -not ($after -is [string])) -or
            ($after -is [pscustomobject])
        )
    }

    if ($beforeIsComplex) {
        try {
            $beforeJson = $before | ConvertTo-Json -Depth 8 -Compress
        }
        catch {
            $beforeJson = [string]$before
        }
    }
    elseif ($null -ne $before) {
        $beforeValue = [string]$before
    }

    if ($afterIsComplex) {
        try {
            $afterJson = $after | ConvertTo-Json -Depth 8 -Compress
        }
        catch {
            $afterJson = [string]$after
        }
    }
    elseif ($null -ne $after) {
        $afterValue = [string]$after
    }

    [pscustomobject]@{
        Component   = [string]$Change.Component
        Changed     = [bool]$Change.Changed
        Verified    = [bool]$Change.Verified
        BeforeValue = $beforeValue
        AfterValue  = $afterValue
        BeforeJson  = $beforeJson
        AfterJson   = $afterJson
    }
}

function New-TelemetryEvent {
    param(
        $DomainInformation,
        $WindowsInformation,
        $ServiceBefore,
        $ServiceAfter,
        $ConfigurationResult,
        $StatusBefore,
        $StatusAfter,
        $HealthResult,
        $ResyncResult,
        $RecentEvents,
        [string]$ResolvedMode
    )

    $endTime = Get-Date
    [ordered]@{
        '@timestamp'       = $endTime.ToUniversalTime().ToString('o')
        EventType          = 'maintenance.execution'
        MaintenanceArea    = 'time_synchronization'
        ComputerName       = $env:COMPUTERNAME
        Domain             = if ($DomainInformation) { $DomainInformation.Domain } else { $env:USERDOMAIN }
        ScriptName         = $script:ScriptName
        ScriptVersion      = $script:ScriptVersion
        RunId              = $script:RunId
        Status             = $script:FinalStatus
        OverallResult      = $script:OverallResult
        ExitCode           = $script:FinalExitCode
        StartTime          = $script:StartTime.ToUniversalTime().ToString('o')
        EndTime            = $endTime.ToUniversalTime().ToString('o')
        DurationSeconds    = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 3)
        ErrorCount         = $script:Errors.Count
        WarningCount       = $script:Warnings.Count
        FailureMessage     = $script:FailureMessage
        TextLogPath        = $PublishedLogPath

        time_synchronization = [ordered]@{
            Execution = [ordered]@{
                RunningAccount  = [Security.Principal.WindowsIdentity]::GetCurrent().Name
                RunningAsSystem = ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18')
                IsAdministrator = Test-IsAdministrator
                Errors          = @($script:Errors | ForEach-Object { [string]$_ })
                Warnings        = @($script:Warnings | ForEach-Object { [string]$_ })
            }

            Configuration = [ordered]@{
                RequestedSyncMode  = $SyncMode
                ResolvedSyncMode   = $ResolvedMode
                ManualNtpServers   = New-StringArrayForJson -Values $(if ($ResolvedMode -eq 'Manual') { @($NtpServers) } else { @() })
                MaximumSyncAgeHours = $MaximumSyncAgeHours
                ClockOffsetAlertThresholdMilliseconds = $ClockOffsetAlertThresholdMilliseconds
                ResyncAttemptsConfigured = $ResyncAttempts
                DomainInformation = $DomainInformation
                TimeConfiguration = $ConfigurationResult
            }

            Health = [ordered]@{
                ClockOffsetBeforeMilliseconds = if ($StatusBefore) { $StatusBefore.ClockOffsetMilliseconds } else { $null }
                AbsoluteClockOffsetBeforeMilliseconds = if ($StatusBefore) { $StatusBefore.AbsoluteClockOffsetMilliseconds } else { $null }
                ClockOffsetAfterMilliseconds = if ($StatusAfter) { $StatusAfter.ClockOffsetMilliseconds } else { $null }
                AbsoluteClockOffsetAfterMilliseconds = if ($StatusAfter) { $StatusAfter.AbsoluteClockOffsetMilliseconds } else { $null }
                ClockOffsetExceedsThreshold = if ($StatusAfter -and $null -ne $StatusAfter.AbsoluteClockOffsetMilliseconds) {
                    [bool]($StatusAfter.AbsoluteClockOffsetMilliseconds -gt $ClockOffsetAlertThresholdMilliseconds)
                }
                else {
                    $null
                }
                TimeHealth = $HealthResult
            }

            Windows          = $WindowsInformation
            ServiceBefore    = $ServiceBefore
            ServiceAfter     = $ServiceAfter
            TimeStatusBefore = $StatusBefore
            TimeStatusAfter  = $StatusAfter
            Resync           = $ResyncResult
            Changes          = @(
                $script:Changes |
                ForEach-Object { Convert-ChangeForTelemetry -Change $_ }
            )
            CommandResults   = @($script:CommandResults)
            RecentTimeEvents = @($RecentEvents)
        }
    }
}

Ensure-Directory -Path $LogDirectory
Write-Log -Message ('Starting {0} version {1}. Run ID: {2}' -f $script:ScriptName, $script:ScriptVersion, $script:RunId)
Write-Log -Message ('Active staged text log: {0}' -f $RuntimeLogPath)
Write-Log -Message ('Dedicated staging root: {0}' -f $DedicatedStagingRoot)
Write-Log -Message ('Completed text log publish path: {0}' -f $PublishedLogPath)

$domainInfo = $null
$windowsInfo = $null
$serviceBefore = $null
$serviceAfter = $null
$configurationResult = $null
$statusBefore = $null
$statusAfter = $null
$healthResult = $null
$resyncResult = $null
$recentEvents = @()
$resolvedMode = $null

try {
    if (-not (Test-IsAdministrator)) {
        throw 'Administrator privileges are required to configure and synchronize the Windows Time service.'
    }

    $windowsInfo = Get-WindowsInformation
    $domainInfo = Get-DomainInformation
    $resolvedMode = Resolve-SyncMode -RequestedMode $SyncMode -DomainInfo $domainInfo

    if ($resolvedMode -eq 'DomainHierarchy' -and -not $domainInfo.PartOfDomain) {
        throw 'DomainHierarchy mode was requested, but this computer is not joined to a domain.'
    }

    Write-Log -Message ('Resolved synchronization mode: {0}.' -f $resolvedMode)

    $serviceBefore = Get-ServiceSnapshot
    $statusBefore = Get-TimeStatusSnapshot
    Write-Log -Message ('Windows Time service before: State={0}; StartMode={1}.' -f $serviceBefore.State, $serviceBefore.StartMode)
    Write-TimeStatusSummary -Label 'before' -Status $statusBefore

    Set-TimeServiceStartup

    # w32tm /config /update requires the Windows Time service to be running.
    # Ensure it is started and verified before applying configuration.
    $serviceAfterStartup = Start-TimeServiceVerified
    Write-Log -Message 'Windows Time service is running before configuration.' -Level SUCCESS

    $configurationResult = Set-TimeConfiguration -ResolvedMode $resolvedMode
    Write-Log -Message ('Time configuration: Type={0} -> {1}; NtpServer={2} -> {3}; PollIntervalSeconds={4} -> {5}.' -f `
        $configurationResult.Before.Type, $configurationResult.After.Type, `
        $configurationResult.Before.NtpServer, $configurationResult.After.NtpServer, `
        $configurationResult.Before.SpecialPollIntervalSeconds, $configurationResult.After.SpecialPollIntervalSeconds) -Level SUCCESS

    $resyncResult = Invoke-TimeResynchronization
    if (-not $resyncResult.Success) {
        if ($resyncResult.AccessDenied) {
            Add-WarningMessage -Message ('Forced Windows Time resynchronization could not be initiated because Windows returned Access Denied (0x80070005). Existing synchronization health will be evaluated from w32tm /query /status.')
        }
        else {
            $lastResyncAttempt = @($resyncResult.Results) | Select-Object -Last 1
            $lastOutput = if ($lastResyncAttempt) { @($lastResyncAttempt.Output) -join ' | ' } else { 'No output captured.' }
            Add-ErrorMessage -Message ('Windows Time resynchronization failed after {0} attempts. LastExitCode={1}; LastOutput={2}' -f `
                $resyncResult.Attempts, $lastResyncAttempt.ExitCode, $lastOutput)
        }
    }
    else {
        Write-Log -Message ('Windows Time resynchronization succeeded on attempt {0}.' -f $resyncResult.Attempts) -Level SUCCESS
    }

    Start-Sleep -Seconds 2
    $serviceAfter = Get-ServiceSnapshot
    $statusAfter = Get-TimeStatusSnapshot
    Write-Log -Message ('Windows Time service after: State={0}; StartMode={1}.' -f $serviceAfter.State, $serviceAfter.StartMode) -Level SUCCESS
    Write-TimeStatusSummary -Label 'after' -Status $statusAfter
    if ($null -ne $statusAfter.AbsoluteClockOffsetMilliseconds) {
        $thresholdExceeded = $statusAfter.AbsoluteClockOffsetMilliseconds -gt $ClockOffsetAlertThresholdMilliseconds
        $thresholdMessage = 'Clock offset threshold check: AbsoluteOffset={0} ms; Threshold={1} ms; Exceeded={2}.' -f `
            $statusAfter.AbsoluteClockOffsetMilliseconds, $ClockOffsetAlertThresholdMilliseconds, $thresholdExceeded
        if ($thresholdExceeded) {
            Add-WarningMessage -Message $thresholdMessage
        }
        else {
            Write-Log -Message $thresholdMessage -Level SUCCESS
        }
    }
    $healthResult = Test-TimeHealth -Status $statusAfter -ResolvedMode $resolvedMode
    $recentEvents = Get-RecentTimeServiceEvents -Hours 24

    if ($serviceAfter.State -ne 'Running') {
        Add-ErrorMessage -Message 'The Windows Time service is not running after maintenance.'
    }

    if (-not $healthResult.Healthy) {
        foreach ($reason in $healthResult.Reasons) {
            Add-ErrorMessage -Message ('Time synchronization health check failed: {0}.' -f $reason)
        }
    }
    elseif ($resyncResult -and $resyncResult.AccessDenied) {
        Write-Log -Message ('Existing Windows Time synchronization is healthy despite Access Denied on forced resync. Source={0}; SyncAgeHours={1}; AbsoluteClockOffsetMilliseconds={2}.' -f `
            $statusAfter.Source, $statusAfter.SyncAgeHours, $statusAfter.AbsoluteClockOffsetMilliseconds) -Level SUCCESS
    }

    if ($script:Errors.Count -gt 0) {
        $script:FinalStatus = 'Failed'
        $script:OverallResult = 'SynchronizationFailed'
        $script:FinalExitCode = 2
    }
    elseif ($script:Warnings.Count -gt 0) {
        $script:FinalStatus = 'SuccessWithWarnings'
        $script:OverallResult = if ($script:Changes.Count -gt 0) { 'ConfiguredAndSynchronized' } else { 'AlreadyCompliant' }
        $script:FinalExitCode = 0
    }
    else {
        $script:FinalStatus = 'Success'
        $script:OverallResult = if (@($script:Changes | Where-Object { $_.Changed }).Count -gt 0) {
            'ConfiguredAndSynchronized'
        }
        else {
            'AlreadyCompliant'
        }
        $script:FinalExitCode = 0
    }
}
catch {
    $script:FailureMessage = $_.Exception.Message
    Add-ErrorMessage -Message $script:FailureMessage
    $script:FinalStatus = 'Failed'
    $script:OverallResult = 'Failed'
    $script:FinalExitCode = 1

    if ($null -eq $serviceAfter) {
        $serviceAfter = Get-ServiceSnapshot
    }

    if ($null -eq $statusAfter) {
        try {
            $statusAfter = Get-TimeStatusSnapshot
        }
        catch {
            # Preserve the primary failure.
        }
    }
}
finally {
    try {
        if ($null -eq $recentEvents -or $recentEvents.Count -eq 0) {
            $recentEvents = Get-RecentTimeServiceEvents -Hours 24
        }

        $event = New-TelemetryEvent `
            -DomainInformation $domainInfo `
            -WindowsInformation $windowsInfo `
            -ServiceBefore $serviceBefore `
            -ServiceAfter $serviceAfter `
            -ConfigurationResult $configurationResult `
            -StatusBefore $statusBefore `
            -StatusAfter $statusAfter `
            -HealthResult $healthResult `
            -ResyncResult $resyncResult `
            -RecentEvents $recentEvents `
            -ResolvedMode $resolvedMode

        try {
            $changeCount = @($event.time_synchronization.Changes).Count
            Write-Log -Message ("Elastic time-sync telemetry summary: Changes={0}; ManualNtpServers={1}; Status={2}; OverallResult={3}" -f `
                $changeCount,
                @($event.time_synchronization.Configuration.ManualNtpServers).Count,
                $event.Status,
                $event.OverallResult)
        }
        catch { }

        Write-Telemetry -Event $event
    }
    catch {
        $telemetryFailure = ('Telemetry finalization failed: {0}' -f $_.Exception.Message)
        Write-Log -Message $telemetryFailure -Level ERROR
        if ($script:FinalExitCode -eq 0) {
            $script:FinalExitCode = 1
            $script:FinalStatus = 'TelemetryFailure'
            $script:OverallResult = 'TelemetryFailure'
        }
    }

    # Final append before the completed immutable file enters C:\Logs.
    Write-Log -Message ('Completed with status {0}, result {1}, and exit code {2}.' -f `
        $script:FinalStatus,
        $script:OverallResult,
        $script:FinalExitCode) `
        -Level $(if ($script:FinalExitCode -eq 0) { 'SUCCESS' } else { 'ERROR' })

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 10 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 10 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $script:FinalExitCode
