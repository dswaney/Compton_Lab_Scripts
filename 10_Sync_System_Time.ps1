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
    Version:    2.0.1
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

    [string]$LogDirectory = 'C:\Logs'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ScriptName = '10_Sync_System_Time.ps1'
$script:ScriptVersion = '2.0.1'
$script:RunId = [guid]::NewGuid().Guid
$script:StartTime = Get-Date
$script:Warnings = New-Object System.Collections.Generic.List[string]
$script:Errors = New-Object System.Collections.Generic.List[string]
$script:CommandResults = New-Object System.Collections.Generic.List[object]
$script:Changes = New-Object System.Collections.Generic.List[object]
$script:FinalExitCode = 0
$script:FinalStatus = 'Failed'
$script:OverallResult = 'Failed'
$script:FailureMessage = $null

$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestJsonPath = Join-Path $LogDirectory '10_Sync_System_Time.latest.json'
$RuntimeLogPath = Join-Path $LogDirectory '10_Sync_System_Time.log'

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

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

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    try {
        Add-Content -LiteralPath $RuntimeLogPath -Value $line -Encoding UTF8
    }
    catch {
        # Logging must not hide the primary operation result.
    }

    Write-Host $line
}

function Add-WarningMessage {
    param([Parameter(Mandatory)][string]$Message)

    $script:Warnings.Add($Message)
    Write-Log -Message $Message -Level WARN
}

function Add-ErrorMessage {
    param([Parameter(Mandatory)][string]$Message)

    $script:Errors.Add($Message)
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

    $script:CommandResults.Add($result)
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
    $source = if ($sourceCommand.Success -and $sourceCommand.Output.Count -gt 0) {
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
        PhaseOffset            = if ($parsedStatus.PSObject.Properties['Phase Offset']) { [string]$parsedStatus.'Phase Offset' } else { $null }
        PollInterval           = if ($parsedStatus.PSObject.Properties['Poll Interval']) { [string]$parsedStatus.'Poll Interval' } else { $null }
        StatusOutput           = @($statusCommand.Output)
        ConfigurationOutput    = @($configCommand.Output)
    }
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
        if ($_.Exception.Message -match 'No events were found') {
            Write-Log -Message 'No recent Windows Time service events were found in the selected period.' -Level INFO
        }
        else {
            Add-WarningMessage -Message ('Unable to read recent Windows Time events: {0}' -f $_.Exception.Message)
        }
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
            WasRunning  = $false
            Stopped     = $true
            ForcedKill  = $false
        }
    }

    Stop-Service -Name 'w32time' -Force -ErrorAction SilentlyContinue
    $deadline = (Get-Date).AddSeconds($ServiceStopTimeoutSeconds)

    do {
        Start-Sleep -Milliseconds 500
        $service.Refresh()
    } while ($service.Status -ne 'Stopped' -and (Get-Date) -lt $deadline)

    $forcedKill = $false
    if ($service.Status -ne 'Stopped') {
        $cim = Get-CimInstance -ClassName Win32_Service -Filter "Name='w32time'"
        if ($cim.ProcessId -gt 0) {
            Stop-Process -Id $cim.ProcessId -Force -ErrorAction Stop
            $forcedKill = $true
            Start-Sleep -Seconds 1
        }

        $service.Refresh()
    }

    [pscustomobject]@{
        WasRunning = $true
        Stopped    = ($service.Status -eq 'Stopped')
        ForcedKill = $forcedKill
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
    $script:Changes.Add([pscustomobject]@{
        Component = 'ServiceStartup'
        Before    = $before.StartMode
        After     = $after.StartMode
        Changed   = $changed
        Verified  = ($after.StartMode -eq 'Auto')
    })

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

    $script:Changes.Add([pscustomobject]@{
        Component = 'TimeConfiguration'
        Before    = $before
        After     = $after
        Changed   = (($before.Type -ne $after.Type) -or ($before.NtpServer -ne $after.NtpServer) -or ($before.SpecialPollIntervalSeconds -ne $after.SpecialPollIntervalSeconds))
        Verified  = $verified
    })

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

function Invoke-TimeResynchronization {
    $w32tm = Join-Path $env:SystemRoot 'System32\w32tm.exe'
    $attemptResults = New-Object System.Collections.Generic.List[object]

    $rediscover = Invoke-ExternalCommand -Name 'RediscoverTimeSource' -FilePath $w32tm -Arguments @('/resync', '/rediscover')
    $attemptResults.Add([pscustomobject]@{
        Attempt = 0
        Type    = 'Rediscover'
        Success = $rediscover.Success
        ExitCode = $rediscover.ExitCode
        Output  = @($rediscover.Output)
    })

    for ($attempt = 1; $attempt -le $ResyncAttempts; $attempt++) {
        $result = Invoke-ExternalCommand -Name ('ResyncAttempt{0}' -f $attempt) -FilePath $w32tm -Arguments @('/resync', '/force')
        $attemptResults.Add([pscustomobject]@{
            Attempt  = $attempt
            Type     = 'ForcedResync'
            Success  = $result.Success
            ExitCode = $result.ExitCode
            Output   = @($result.Output)
        })

        if ($result.Success) {
            return [pscustomobject]@{
                Success  = $true
                Attempts = $attempt
                Results  = $attemptResults.ToArray()
            }
        }

        if ($attempt -lt $ResyncAttempts) {
            Start-Sleep -Seconds $ResyncRetryDelaySeconds
        }
    }

    return [pscustomobject]@{
        Success  = $false
        Attempts = $ResyncAttempts
        Results  = $attemptResults.ToArray()
    }
}

function Test-TimeHealth {
    param(
        [Parameter(Mandatory)]$Status,
        [Parameter(Mandatory)][string]$ResolvedMode
    )

    $reasons = New-Object System.Collections.Generic.List[string]

    if (-not $Status.QuerySucceeded) {
        $reasons.Add('StatusQueryFailed')
    }

    if ([string]::IsNullOrWhiteSpace($Status.Source)) {
        $reasons.Add('TimeSourceMissing')
    }
    elseif ($Status.Source -match '^(?i)(Local CMOS Clock|Free-running System Clock)$') {
        $reasons.Add('UnsynchronizedLocalClock')
    }

    if ([string]$Status.LeapIndicator -match '^(?i)3\b|not synchronized') {
        $reasons.Add('LeapIndicatorUnsynchronized')
    }

    if ($null -eq $Status.SyncAgeHours) {
        $reasons.Add('LastSuccessfulSyncUnknown')
    }
    elseif ($Status.SyncAgeHours -gt $MaximumSyncAgeHours) {
        $reasons.Add('LastSuccessfulSyncTooOld')
    }

    if ($ResolvedMode -eq 'DomainHierarchy' -and $Status.Source -match '^(?i)time\.windows\.com') {
        $reasons.Add('UnexpectedPublicSourceForDomainHierarchy')
    }

    [pscustomobject]@{
        Healthy = ($reasons.Count -eq 0)
        Reasons = $reasons.ToArray()
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
        Errors             = $script:Errors.ToArray()
        Warnings           = $script:Warnings.ToArray()
        FailureMessage     = $script:FailureMessage
        RunningAccount     = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        RunningAsSystem    = ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18')
        IsAdministrator    = Test-IsAdministrator
        RequestedSyncMode  = $SyncMode
        ResolvedSyncMode   = $ResolvedMode
        ManualNtpServers   = if ($ResolvedMode -eq 'Manual') { @($NtpServers) } else { @() }
        MaximumSyncAgeHours = $MaximumSyncAgeHours
        ResyncAttemptsConfigured = $ResyncAttempts
        Windows            = $WindowsInformation
        DomainInformation  = $DomainInformation
        ServiceBefore      = $ServiceBefore
        ServiceAfter       = $ServiceAfter
        Configuration      = $ConfigurationResult
        TimeStatusBefore   = $StatusBefore
        TimeStatusAfter    = $StatusAfter
        TimeHealth         = $HealthResult
        Resync             = $ResyncResult
        Changes            = $script:Changes.ToArray()
        CommandResults     = $script:CommandResults.ToArray()
        RecentTimeEvents   = @($RecentEvents)
    }
}

Ensure-Directory -Path $LogDirectory
Write-Log -Message ('Starting {0} version {1}. Run ID: {2}' -f $script:ScriptName, $script:ScriptVersion, $script:RunId)

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

    Set-TimeServiceStartup

    # w32tm /config /update requires the Windows Time service to be running.
    # Start and verify the service before applying the synchronization settings.
    $serviceBeforeConfiguration = Start-TimeServiceVerified
    Write-Log -Message 'Windows Time service is running before configuration.' -Level SUCCESS

    $configurationResult = Set-TimeConfiguration -ResolvedMode $resolvedMode
    $serviceAfterConfiguration = Get-ServiceSnapshot

    $resyncResult = Invoke-TimeResynchronization
    if (-not $resyncResult.Success) {
        Add-ErrorMessage -Message ('Windows Time resynchronization failed after {0} attempts.' -f $resyncResult.Attempts)
    }
    else {
        Write-Log -Message ('Windows Time resynchronization succeeded on attempt {0}.' -f $resyncResult.Attempts) -Level SUCCESS
    }

    Start-Sleep -Seconds 2
    $serviceAfter = Get-ServiceSnapshot
    $statusAfter = Get-TimeStatusSnapshot
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

        Write-Telemetry -Event $event
        Write-Log -Message ('Completed with status {0}, result {1}, and exit code {2}.' -f $script:FinalStatus, $script:OverallResult, $script:FinalExitCode) -Level $(if ($script:FinalExitCode -eq 0) { 'SUCCESS' } else { 'ERROR' })
    }
    catch {
        Write-Host ('Telemetry finalization failed: {0}' -f $_.Exception.Message)
        if ($script:FinalExitCode -eq 0) {
            $script:FinalExitCode = 1
        }
    }
}

exit $script:FinalExitCode
