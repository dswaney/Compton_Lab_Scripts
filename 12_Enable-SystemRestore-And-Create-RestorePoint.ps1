<#
.SYNOPSIS
    Enables System Restore on the Windows client operating-system drive, creates a verified
    restore point, and retains a configurable number of restore points created by this script.

.NOTES
    ScriptName: 12_Enable-SystemRestore-And-Create-RestorePoint.ps1
    ScriptVersion: 2.0.8
    LastUpdated: 2026-08-17
    Changes: v2.0.8 normalizes restore-point enumeration to arrays in retention and latest-point logic so a single restore point does not trigger a StrictMode Count-property failure.
              v2.0.7 allows an empty restore-point inventory during pre/post checks, which is valid on systems with no existing restore points.
              v2.0.6 adds Elastic mapping-safe array serialization for services, restore-point inventories, and retention results; normalizes Windows 11 product naming; and adds a concise Elastic restore-point summary.
              v2.0.5 uses Maintenance.Framework v2.4 staged text logging.
    Designed for Windows PowerShell 5.1 on Windows 10/11.
#>

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$RestorePointDescription = 'Compton IT - Weekly Restore Point',

    [ValidateRange(1, 30)]
    [int]$VerificationTimeoutMinutes = 5,

    [ValidateRange(1, 20)]
    [int]$MaxManagedRestorePoints = 2,

    [ValidateRange(30, 3600)]
    [int]$ServiceStartTimeoutSeconds = 60,

    [switch]$SkipRetentionCleanup
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptName = '12_Enable-SystemRestore-And-Create-RestorePoint.ps1'
$ScriptVersion = '2.0.8'
$RunId = [guid]::NewGuid().Guid
$RunStart = Get-Date
$LogDirectory = 'C:\Logs'
$LogPath = $null
$PublishedLogPath = $null
$LogSession = $null
$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestTelemetryPath = Join-Path $LogDirectory '12_Enable-SystemRestore-And-Create-RestorePoint.latest.json'

$script:WarningCount = 0
$script:ErrorCount = 0
$script:FailureMessage = $null
$script:FinalStatus = 'Running'
$script:FinalExitCode = 1
$script:RestoreProtectionEnableAttempted = $false
$script:RestoreProtectionEnabled = $false
$script:RestorePointCreationAttempted = $false
$script:RestorePointVerified = $false
$script:CreatedRestorePoint = $null
$script:RetentionResults = New-Object System.Collections.Generic.List[object]
$script:ServiceResults = New-Object System.Collections.Generic.List[object]
$script:RestorePointsBefore = @()
$script:RestorePointsAfter = @()
$script:FrequencyBefore = $null
$script:FrequencyAfter = $null
$script:FrequencyChanged = $false

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 12 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '12_Enable-SystemRestore-And-Create-RestorePoint.log',
        '*-12_Enable-SystemRestore-And-Create-RestorePoint-*.log'
    ) | Out-Null

$LogSession = New-MaintenanceStagedLog `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $env:COMPUTERNAME `
    -Timestamp $RunStart

$LogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
    $normalizedLevel = switch ($Level) {
        'OK'   { 'SUCCESS' }
        'WARN' { 'WARNING' }
        default { $Level }
    }
    $line = '{0} [{1}] [{2}] {3}' -f $timestamp, $computerName, $normalizedLevel, $Message

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }

    try {
        $activeLogDirectory = Split-Path -Parent $LogPath
        Ensure-Directory -Path $activeLogDirectory
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning ('Unable to write to the log file: {0}' -f $_.Exception.Message)
    }
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

function Invoke-SelfElevation {
    if (Test-IsAdministrator) {
        return
    }

    if (-not $PSCommandPath) {
        throw 'Administrative privileges are required. Open Windows PowerShell as Administrator and run the script again.'
    }

    Write-Host 'Administrative privileges are required. Requesting elevation...' -ForegroundColor Yellow

    $escapedDescription = $RestorePointDescription.Replace('"', '\"')
    $argumentParts = @(
        '-NoProfile'
        '-ExecutionPolicy Bypass'
        ('-File "{0}"' -f $PSCommandPath)
        ('-RestorePointDescription "{0}"' -f $escapedDescription)
        ('-VerificationTimeoutMinutes {0}' -f $VerificationTimeoutMinutes)
        ('-MaxManagedRestorePoints {0}' -f $MaxManagedRestorePoints)
        ('-ServiceStartTimeoutSeconds {0}' -f $ServiceStartTimeoutSeconds)
    )

    if ($SkipRetentionCleanup) {
        $argumentParts += '-SkipRetentionCleanup'
    }

    Start-Process `
        -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList ($argumentParts -join ' ') `
        -Verb RunAs | Out-Null

    exit 0
}

function Convert-ToIsoUtc {
    param([AllowNull()]$Date)

    if ($null -eq $Date) { return $null }
    return $Date.ToUniversalTime().ToString('o')
}

function Get-WindowsIdentity {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $cvPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $cv = Get-ItemProperty -LiteralPath $cvPath -ErrorAction SilentlyContinue

    $productName = if ($cv) { [string]$cv.ProductName } else { [string]$os.Caption }
    $buildText = [string]$os.BuildNumber
    $buildNumber = 0
    [void][int]::TryParse($buildText, [ref]$buildNumber)

    if ($buildNumber -ge 22000 -and $productName -match '^Windows 10') {
        $productName = $productName -replace '^Windows 10', 'Windows 11'
    }

    [pscustomobject]@{
        ProductName    = $productName
        EditionId      = if ($cv) { [string]$cv.EditionID } else { $null }
        DisplayVersion = if ($cv) { [string]$cv.DisplayVersion } else { $null }
        ReleaseId      = if ($cv) { [string]$cv.ReleaseId } else { $null }
        Build          = $buildText
        UBR            = if ($cv -and $null -ne $cv.UBR) { [int]$cv.UBR } else { $null }
        FullBuild      = if ($cv -and $null -ne $cv.UBR) { '{0}.{1}' -f $buildText, $cv.UBR } else { $buildText }
        ProductType    = [int]$os.ProductType
        SystemDrive    = [string]$os.SystemDrive
    }
}

function Get-OperatingSystemDrive {
    param([Parameter(Mandatory)]$WindowsIdentity)

    if ([string]::IsNullOrWhiteSpace([string]$WindowsIdentity.SystemDrive)) {
        throw 'Unable to determine the Windows operating-system drive.'
    }

    return ('{0}\' -f ([string]$WindowsIdentity.SystemDrive).TrimEnd('\'))
}

function Get-RestorePointCreationFrequency {
    $registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
    $valueName = 'SystemRestorePointCreationFrequency'

    try {
        $value = Get-ItemPropertyValue -LiteralPath $registryPath -Name $valueName -ErrorAction Stop
        return [int]$value
    }
    catch {
        return $null
    }
}

function Set-RestorePointCreationFrequency {
    $registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
    $valueName = 'SystemRestorePointCreationFrequency'

    $script:FrequencyBefore = Get-RestorePointCreationFrequency

    if (-not (Test-Path -LiteralPath $registryPath)) {
        New-Item -Path $registryPath -Force -ErrorAction Stop | Out-Null
    }

    if ($script:FrequencyBefore -ne 0) {
        New-ItemProperty `
            -LiteralPath $registryPath `
            -Name $valueName `
            -PropertyType DWord `
            -Value 0 `
            -Force `
            -ErrorAction Stop | Out-Null

        $script:FrequencyChanged = $true
        Write-Log 'Configured Windows to permit restore-point creation on each scheduled run.' 'OK'
    }
    else {
        Write-Log 'Restore-point creation frequency is already configured for each run.'
    }

    $script:FrequencyAfter = Get-RestorePointCreationFrequency
    if ($script:FrequencyAfter -ne 0) {
        throw 'SystemRestorePointCreationFrequency could not be verified as 0.'
    }
}


function Test-SystemRestorePolicy {
    [CmdletBinding()]
    param()

    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore'
    $disableSr = $null
    $disableConfig = $null

    if (Test-Path -LiteralPath $policyPath) {
        try {
            $policy = Get-ItemProperty -LiteralPath $policyPath -ErrorAction Stop
            if ($null -ne $policy.DisableSR) {
                $disableSr = [int]$policy.DisableSR
            }
            if ($null -ne $policy.DisableConfig) {
                $disableConfig = [int]$policy.DisableConfig
            }
        }
        catch {
            Write-Log ('Unable to read System Restore policy: {0}' -f $_.Exception.Message) 'WARN'
        }
    }

    if ($disableSr -eq 1) {
        throw 'System Restore is disabled by policy (DisableSR=1). Remove or change the policy before restore points can be created.'
    }

    if ($disableConfig -eq 1) {
        Write-Log 'System Restore configuration is restricted by policy (DisableConfig=1), but restore-point creation will still be attempted.' 'WARN'
    }

    Write-Log 'System Restore policy does not explicitly disable restore-point creation.' 'OK'
}

function Enable-SystemRestoreProtection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Drive,
        [ValidateRange(5,120)]
        [int]$InitializationWaitSeconds = 20
    )

    $script:RestoreProtectionEnableAttempted = $true
    Write-Log ('Ensuring System Restore protection is enabled for {0}' -f $Drive) 'INFO'

    # Keep the supported PowerShell cmdlet call.
    Enable-ComputerRestore -Drive $Drive -ErrorAction Stop
    $script:RestoreProtectionEnabled = $true
    Write-Log ('System Restore protection enable command completed for {0}' -f $Drive) 'OK'

    # Call the underlying SystemRestore.Enable() method directly as well so
    # we can capture its return code. Microsoft documents that this method
    # returns before monitoring/filter-driver initialization is complete.
    try {
        $restoreClass = Get-WmiObject `
            -Namespace 'root\default' `
            -List `
            -Class 'SystemRestore' `
            -ErrorAction Stop

        if ($restoreClass) {
            $enableResult = $restoreClass.Enable($Drive)
            $enableCode = [int]$enableResult.ReturnValue

            if ($enableCode -eq 0) {
                Write-Log ('Native SystemRestore.Enable returned success for {0}.' -f $Drive) 'OK'
            }
            else {
                Write-Log ('Native SystemRestore.Enable returned code {0} for {1}.' -f $enableCode, $Drive) 'WARN'
            }
        }
    }
    catch {
        Write-Log ('Unable to obtain native SystemRestore.Enable return code: {0}' -f $_.Exception.Message) 'WARN'
    }

    Write-Log ('Waiting {0} second(s) for System Restore monitoring to initialize...' -f $InitializationWaitSeconds) 'INFO'
    Start-Sleep -Seconds $InitializationWaitSeconds
}

function Get-ServiceSnapshot {
    param([Parameter(Mandatory)][string]$Name)

    $service = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $Name.Replace("'", "''")) -ErrorAction SilentlyContinue
    if (-not $service) {
        return [pscustomobject]@{
            Name      = $Name
            Exists    = $false
            State     = $null
            StartMode = $null
            ProcessId = $null
        }
    }

    [pscustomobject]@{
        Name      = [string]$service.Name
        Exists    = $true
        State     = [string]$service.State
        StartMode = [string]$service.StartMode
        ProcessId = [int]$service.ProcessId
    }
}

function Initialize-ShadowCopyServices {
    # VSS and swprv are always expected. srservice is checked when present.
    # On some Windows builds the System Restore service is not exposed as a
    # separate service, so absence of srservice is informational rather than fatal.
    $serviceDefinitions = @(
        [pscustomobject]@{ Name = 'VSS';       DesiredStartup = 'Manual'; Required = $true  },
        [pscustomobject]@{ Name = 'swprv';     DesiredStartup = 'Manual'; Required = $true  },
        [pscustomobject]@{ Name = 'srservice'; DesiredStartup = 'Manual'; Required = $false }
    )

    foreach ($definition in $serviceDefinitions) {
        $serviceName = $definition.Name
        $before = Get-ServiceSnapshot -Name $serviceName
        $attempted = $false
        $operationStatus = 'NoChange'
        $message = $null

        if (-not $before.Exists) {
            if ($definition.Required) {
                $operationStatus = 'NotFound'
                $message = 'Required service was not found.'
                Write-Log ('Required service {0} was not found.' -f $serviceName) 'WARN'
            }
            else {
                $operationStatus = 'NotPresentOnBuild'
                $message = 'Optional System Restore service is not present on this Windows build.'
                Write-Log ('Optional service {0} is not present on this Windows build.' -f $serviceName) 'INFO'
            }
        }
        else {
            # A disabled dependency causes Checkpoint-Computer to fail with
            # "the service cannot be started because it is disabled".
            if ($before.StartMode -eq 'Disabled') {
                $attempted = $true
                Write-Log ('Service {0} is disabled. Changing startup type to {1}...' -f $serviceName, $definition.DesiredStartup) 'INFO'

                try {
                    Set-Service -Name $serviceName -StartupType $definition.DesiredStartup -ErrorAction Stop
                    Write-Log ('Startup type for {0} changed to {1}.' -f $serviceName, $definition.DesiredStartup) 'OK'
                }
                catch {
                    # Some protected services reject Set-Service; try sc.exe.
                    $scStart = if ($definition.DesiredStartup -eq 'Manual') { 'demand' } else { 'auto' }
                    $scOutput = & "$env:SystemRoot\System32\sc.exe" config $serviceName start= $scStart 2>&1

                    if ($LASTEXITCODE -ne 0) {
                        $operationStatus = 'Failed'
                        $message = ($scOutput -join ' ')
                        Write-Log ('Unable to change startup type for service {0}: {1}' -f $serviceName, $message) 'WARN'
                    }
                    else {
                        Write-Log ('Startup type for {0} changed through sc.exe.' -f $serviceName) 'OK'
                    }
                }
            }

            $current = Get-ServiceSnapshot -Name $serviceName

            if ($current.State -ne 'Running') {
                $attempted = $true
                Write-Log ('Starting restore-point dependency service {0}...' -f $serviceName) 'INFO'

                try {
                    Start-Service -Name $serviceName -ErrorAction Stop
                    $serviceController = Get-Service -Name $serviceName -ErrorAction Stop
                    $serviceController.WaitForStatus(
                        [System.ServiceProcess.ServiceControllerStatus]::Running,
                        [TimeSpan]::FromSeconds($ServiceStartTimeoutSeconds)
                    )
                    $operationStatus = 'Started'
                    Write-Log ('Service {0} is running.' -f $serviceName) 'OK'
                }
                catch {
                    $operationStatus = 'Failed'
                    $message = $_.Exception.Message
                    Write-Log ('Service {0} could not be started: {1}' -f $serviceName, $message) 'WARN'
                }
            }
            else {
                Write-Log ('Required service {0} is already running.' -f $serviceName) 'INFO'
            }
        }

        $after = Get-ServiceSnapshot -Name $serviceName
        $verified = if (-not $after.Exists -and -not $definition.Required) {
            $true
        }
        else {
            ($after.Exists -and $after.State -eq 'Running')
        }

        [void]$script:ServiceResults.Add([pscustomobject]@{
            Name            = $serviceName
            Required        = [bool]$definition.Required
            Before          = $before
            StartAttempted  = $attempted
            OperationStatus = $operationStatus
            Message         = $message
            After           = $after
            VerifiedRunning = $verified
        })

        $serviceLogLevel = if ($verified) { 'OK' } else { 'WARN' }
        Write-Log ('Restore service state: Name={0}; Required={1}; State={2} -> {3}; StartMode={4} -> {5}; VerifiedRunning={6}; Operation={7}.' -f `
            $serviceName, $definition.Required, $before.State, $after.State, $before.StartMode, $after.StartMode, $verified, $operationStatus) $serviceLogLevel

        if (-not $verified -and $definition.Required) {
            throw ('Required restore-point service {0} could not be verified as running.' -f $serviceName)
        }
    }
}

function Convert-RestorePointTime {
    param([AllowNull()][string]$CreationTime)

    if ([string]::IsNullOrWhiteSpace($CreationTime)) {
        return $null
    }

    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($CreationTime)
    }
    catch {
        return $null
    }
}

function Convert-RestorePointRecord {
    param([Parameter(Mandatory)]$RestorePoint)

    $created = Convert-RestorePointTime -CreationTime ([string]$RestorePoint.CreationTime)

    [pscustomobject]@{
        SequenceNumber = [uint32]$RestorePoint.SequenceNumber
        Description    = [string]$RestorePoint.Description
        CreationTime   = Convert-ToIsoUtc -Date $created
        RestorePointType = if ($null -ne $RestorePoint.RestorePointType) { [int]$RestorePoint.RestorePointType } else { $null }
        EventType      = if ($null -ne $RestorePoint.EventType) { [int]$RestorePoint.EventType } else { $null }
        Managed        = ([string]$RestorePoint.Description -eq $RestorePointDescription)
    }
}

function Get-AllRestorePoints {
    try {
        [object[]]$points = @(
            Get-ComputerRestorePoint -ErrorAction Stop |
            Sort-Object -Property SequenceNumber -Descending
        )

        return $points
    }
    catch {
        Write-Log ('Unable to enumerate restore points: {0}' -f $_.Exception.Message) 'WARN'
        return @()
    }
}

function Write-RestorePointInventorySummary {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$RestorePoints
    )

    $managed = @($RestorePoints | Where-Object { $_.Managed })
    $nonManaged = @($RestorePoints | Where-Object { -not $_.Managed })
    $latest = $RestorePoints | Sort-Object -Property SequenceNumber -Descending | Select-Object -First 1

    Write-Log ('Restore-point inventory {0}: Total={1}; Managed={2}; NonManaged={3}.' -f `
        $Label, $RestorePoints.Count, $managed.Count, $nonManaged.Count)

    if ($latest) {
        Write-Log ('Latest restore point {0}: SequenceNumber={1}; Created={2}; Managed={3}; Description={4}.' -f `
            $Label, $latest.SequenceNumber, $latest.CreationTime, $latest.Managed, $latest.Description)
    }
    else {
        Write-Log ('Latest restore point {0}: none found.' -f $Label) 'INFO'
    }
}

function Get-LatestRestorePoint {
    $points = @(Get-AllRestorePoints)
    if ($points.Count -eq 0) { return $null }
    return $points[0]
}

function Initialize-SystemRestoreNativeApi {
    if ('SystemRestore.NativeMethods' -as [type]) {
        return
    }

    $source = @'
using System;
using System.Runtime.InteropServices;

namespace SystemRestore
{
    public static class NativeMethods
    {
        [DllImport("SrClient.dll", SetLastError = true)]
        public static extern uint SRRemoveRestorePoint(uint restorePointSequenceNumber);
    }
}
'@

    Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
}

function Remove-RestorePointBySequenceNumber {
    param([Parameter(Mandatory)][uint32]$SequenceNumber)

    Initialize-SystemRestoreNativeApi
    $result = [SystemRestore.NativeMethods]::SRRemoveRestorePoint($SequenceNumber)

    if ($result -eq 0) {
        return
    }

    $message = (New-Object ComponentModel.Win32Exception([int]$result)).Message
    throw ('SRRemoveRestorePoint failed for sequence {0}. Win32Result={1}; Message={2}' -f $SequenceNumber, $result, $message)
}

function Remove-ObsoleteManagedRestorePoints {
    param(
        [Parameter(Mandatory)][string]$ManagedDescription,
        [Parameter(Mandatory)][int]$KeepCount
    )

    $all = @(Get-AllRestorePoints)
    $managed = @($all |
        Where-Object { [string]$_.Description -eq $ManagedDescription } |
        Sort-Object -Property SequenceNumber -Descending)

    $toKeep = @($managed | Select-Object -First $KeepCount)
    $toRemove = @($managed | Select-Object -Skip $KeepCount)

    Write-Log (
        'Managed restore-point retention: ManagedTotal={0}; Keeping={1}; Removing={2}; NonManagedPreserved={3}' -f
        $managed.Count,
        $toKeep.Count,
        $toRemove.Count,
        ($all.Count - $managed.Count)
    )

    foreach ($restorePoint in $toRemove) {
        $created = Convert-RestorePointTime -CreationTime ([string]$restorePoint.CreationTime)
        $createdText = if ($created) { $created.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Unknown' }
        $resultStatus = 'Failed'
        $failure = $null

        Write-Log (
            'Deleting obsolete managed restore point: SequenceNumber={0}; Created={1}; Description={2}' -f
            $restorePoint.SequenceNumber,
            $createdText,
            $restorePoint.Description
        )

        try {
            Remove-RestorePointBySequenceNumber -SequenceNumber ([uint32]$restorePoint.SequenceNumber)
            $resultStatus = 'Deleted'
            Write-Log ('Deleted managed restore point sequence {0}.' -f $restorePoint.SequenceNumber) 'OK'
        }
        catch {
            $failure = $_.Exception.Message
            Write-Log ('Unable to delete managed restore point sequence {0}: {1}' -f $restorePoint.SequenceNumber, $failure) 'WARN'
        }

        [void]$script:RetentionResults.Add([pscustomobject]@{
            SequenceNumber = [uint32]$restorePoint.SequenceNumber
            Description    = [string]$restorePoint.Description
            CreationTime   = Convert-ToIsoUtc -Date $created
            Status         = $resultStatus
            Error          = $failure
        })
    }
}


function New-NativeSystemRestorePoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Description
    )

    Write-Log ('Attempting native SystemRestore.CreateRestorePoint fallback: {0}' -f $Description) 'INFO'

    $restoreClass = Get-WmiObject `
        -Namespace 'root\default' `
        -List `
        -Class 'SystemRestore' `
        -ErrorAction Stop

    if (-not $restoreClass) {
        throw 'The root\default:SystemRestore WMI class is not available.'
    }

    # EventType 100 = BEGIN_SYSTEM_CHANGE
    # RestorePointType 12 = MODIFY_SETTINGS
    try {
        $result = $restoreClass.CreateRestorePoint(
            $Description,
            12,
            100
        )

        $returnCode = [int]$result.ReturnValue
        Write-Log ('Native SystemRestore.CreateRestorePoint returned code {0}.' -f $returnCode) `
            $(if ($returnCode -eq 0) { 'OK' } else { 'WARN' })

        if ($returnCode -ne 0) {
            throw ('Native SystemRestore.CreateRestorePoint failed with return code {0}.' -f $returnCode)
        }

        return $true
    }
    catch {
        throw
    }
}

function New-VerifiedRestorePoint {
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][int]$TimeoutMinutes
    )

    $startedAt = Get-Date
    $safeDescription = $Description.Trim()

    if ($safeDescription.Length -gt 256) {
        $safeDescription = $safeDescription.Substring(0, 256)
        Write-Log 'Restore-point description exceeded 256 characters and was truncated.' 'WARN'
    }

    $sequenceNumbersBefore = @(
        Get-AllRestorePoints | ForEach-Object { [uint32]$_.SequenceNumber }
    )

    $script:RestorePointCreationAttempted = $true
    Write-Log ('Creating restore point: {0}' -f $safeDescription)

    $creationSucceeded = $false
    $maxCreationAttempts = 6
    $retryDelaySeconds = 10
    $lastCreationError = $null

    for ($attempt = 1; $attempt -le $maxCreationAttempts; $attempt++) {
        try {
            Write-Log ('Restore-point creation attempt {0} of {1} using Checkpoint-Computer.' -f $attempt, $maxCreationAttempts) 'INFO'

            Checkpoint-Computer `
                -Description $safeDescription `
                -RestorePointType MODIFY_SETTINGS `
                -ErrorAction Stop

            Write-Log 'Checkpoint-Computer completed without error.' 'OK'
            $creationSucceeded = $true
            break
        }
        catch {
            $lastCreationError = $_.Exception.Message
            Write-Log ('Checkpoint-Computer attempt {0} failed: {1}' -f $attempt, $lastCreationError) 'WARN'

            $serviceDisabledCondition = (
                $lastCreationError -match 'service cannot be started because it is disabled' -or
                $lastCreationError -match 'does not have enabled devices associated with it' -or
                $lastCreationError -match 'ServiceDisabled' -or
                $lastCreationError -match '0x80070422'
            )

            if (-not $serviceDisabledCondition) {
                throw
            }

            # Try the lower-level WMI interface too. If System Restore is still
            # initializing, it can fail with the same underlying condition.
            try {
                $null = New-NativeSystemRestorePoint -Description $safeDescription
                Write-Log 'Native SystemRestore.CreateRestorePoint completed successfully.' 'OK'
                $creationSucceeded = $true
                break
            }
            catch {
                $lastCreationError = $_.Exception.Message
                Write-Log ('Native fallback attempt {0} also failed: {1}' -f $attempt, $lastCreationError) 'WARN'
            }

            if ($attempt -lt $maxCreationAttempts) {
                Write-Log ('System Restore may still be initializing. Waiting {0} seconds before retrying...' -f $retryDelaySeconds) 'INFO'
                Start-Sleep -Seconds $retryDelaySeconds
            }
        }
    }

    if (-not $creationSucceeded) {
        throw ('Restore-point creation failed after {0} attempts. Last error: {1}' -f $maxCreationAttempts, $lastCreationError)
    }

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    do {
        Start-Sleep -Seconds 5
        $points = @(Get-AllRestorePoints)
        $candidate = $points |
            Where-Object {
                $sequenceNumbersBefore -notcontains [uint32]$_.SequenceNumber -and
                [string]$_.Description -eq $safeDescription
            } |
            Sort-Object -Property SequenceNumber -Descending |
            Select-Object -First 1

        if ($candidate) {
            $creationDate = Convert-RestorePointTime -CreationTime ([string]$candidate.CreationTime)

            if ($creationDate -and $creationDate -ge $startedAt.AddMinutes(-1)) {
                $script:RestorePointVerified = $true
                $script:CreatedRestorePoint = Convert-RestorePointRecord -RestorePoint $candidate

                Write-Log (
                    'Restore point verified: SequenceNumber={0}; Created={1}; Description={2}' -f
                    $candidate.SequenceNumber,
                    $creationDate.ToString('yyyy-MM-dd HH:mm:ss'),
                    $candidate.Description
                ) 'OK'

                return $candidate
            }
        }
    }
    while ((Get-Date) -lt $deadline)

    throw ('The restore point could not be verified within {0} minute(s).' -f $TimeoutMinutes)
}

function Write-JsonAtomically {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Json
    )

    $directory = Split-Path -Parent $Path
    Ensure-Directory -Path $directory

    $temporaryPath = '{0}.{1}.tmp' -f $Path, $RunId
    [System.IO.File]::WriteAllText($temporaryPath, $Json, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
}


function New-ObjectArrayForJson {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    [object[]]$items = @(
        $InputObject |
        ForEach-Object { $_ }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Write-Telemetry {
    param([AllowNull()]$WindowsIdentity)

    try {
        Ensure-Directory -Path $LogDirectory

        $runEnd = Get-Date
        $duration = [math]::Round(($runEnd - $RunStart).TotalSeconds, 3)

        $managedBefore = @($script:RestorePointsBefore | Where-Object { $_.Managed })
        $managedAfter = @($script:RestorePointsAfter | Where-Object { $_.Managed })
        $nonManagedBefore = @($script:RestorePointsBefore | Where-Object { -not $_.Managed })
        $nonManagedAfter = @($script:RestorePointsAfter | Where-Object { -not $_.Managed })

        $event = [ordered]@{
            '@timestamp' = (Get-Date).ToUniversalTime().ToString('o')
            EventType = 'maintenance.system_restore'
            ComputerName = $env:COMPUTERNAME
            Domain = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { $env:USERDOMAIN }
            ScriptName = $ScriptName
            ScriptVersion = $ScriptVersion
            RunId = $RunId
            Status = $script:FinalStatus
            ExitCode = $script:FinalExitCode
            StartTime = Convert-ToIsoUtc -Date $RunStart
            EndTime = Convert-ToIsoUtc -Date $runEnd
            DurationSeconds = $duration
            WarningCount = $script:WarningCount
            ErrorCount = $script:ErrorCount
            FailureMessage = $script:FailureMessage
            TextLogPath = $PublishedLogPath
            RunningAccount = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            IsSystem = ([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem)
            Windows = $WindowsIdentity
            RestorePointDescription = $RestorePointDescription
            VerificationTimeoutMinutes = $VerificationTimeoutMinutes
            MaxManagedRestorePoints = $MaxManagedRestorePoints
            SkipRetentionCleanup = [bool]$SkipRetentionCleanup
            RestoreProtection = [ordered]@{
                Drive = if ($WindowsIdentity) { '{0}\' -f ([string]$WindowsIdentity.SystemDrive).TrimEnd('\') } else { $null }
                EnableAttempted = $script:RestoreProtectionEnableAttempted
                EnableCommandSucceeded = $script:RestoreProtectionEnabled
            }
            CreationFrequency = [ordered]@{
                Before = $script:FrequencyBefore
                After = $script:FrequencyAfter
                Changed = $script:FrequencyChanged
                Verified = ($script:FrequencyAfter -eq 0)
            }
            ShadowCopyServices = New-ObjectArrayForJson -InputObject $script:ServiceResults
            RestorePointCreation = [ordered]@{
                Attempted = $script:RestorePointCreationAttempted
                Verified = $script:RestorePointVerified
                CreatedRestorePoint = $script:CreatedRestorePoint
            }
            RestorePointInventory = [ordered]@{
                TotalBefore = $script:RestorePointsBefore.Count
                TotalAfter = $script:RestorePointsAfter.Count
                ManagedBefore = $managedBefore.Count
                ManagedAfter = $managedAfter.Count
                NonManagedBefore = $nonManagedBefore.Count
                NonManagedAfter = $nonManagedAfter.Count
                RestorePointsBefore = New-ObjectArrayForJson -InputObject $script:RestorePointsBefore
                RestorePointsAfter = New-ObjectArrayForJson -InputObject $script:RestorePointsAfter
            }
            Retention = [ordered]@{
                CleanupSkipped = [bool]$SkipRetentionCleanup
                DeletedCount = @($script:RetentionResults | ForEach-Object { $_ } | Where-Object { $_.Status -eq 'Deleted' }).Count
                FailedCount = @($script:RetentionResults | ForEach-Object { $_ } | Where-Object { $_.Status -eq 'Failed' }).Count
                Results = New-ObjectArrayForJson -InputObject $script:RetentionResults
                NonManagedRestorePointsPreserved = $true
            }
        }

        try {
            Write-Log ('Elastic system-restore summary: Status={0}; ProtectionEnabled={1}; RestorePointVerified={2}; TotalBefore={3}; TotalAfter={4}; ManagedBefore={5}; ManagedAfter={6}; RetentionDeleted={7}; RetentionFailed={8}.' -f `
                $event.Status,
                $event.RestoreProtection.EnableCommandSucceeded,
                $event.RestorePointCreation.Verified,
                $event.RestorePointInventory.TotalBefore,
                $event.RestorePointInventory.TotalAfter,
                $event.RestorePointInventory.ManagedBefore,
                $event.RestorePointInventory.ManagedAfter,
                $event.Retention.DeletedCount,
                $event.Retention.FailedCount) 'INFO'
        }
        catch { }

        $prettyJson = $event | ConvertTo-Json -Depth 12
        $compactJson = $event | ConvertTo-Json -Depth 12 -Compress

        Write-JsonAtomically -Path $LatestTelemetryPath -Json $prettyJson
        Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compactJson
    }
    catch {
        Write-Warning ('Unable to write telemetry: {0}' -f $_.Exception.Message)
    }
}

$windowsIdentity = $null

try {
    Invoke-SelfElevation
    Ensure-Directory -Path $LogDirectory
    Write-Log ('===== System Restore script v{0} started; RunId={1} =====' -f $ScriptVersion, $RunId)
    Write-Log ('Active staged text log: {0}' -f $LogPath)
    Write-Log ('Completed text log publish path: {0}' -f $PublishedLogPath)

    if (-not (Test-IsAdministrator)) {
        throw 'Administrative privileges are required.'
    }

    $windowsIdentity = Get-WindowsIdentity

    if ($windowsIdentity.ProductType -ne 1) {
        throw 'System Restore checkpoints are supported on Windows client operating systems, not Windows Server.'
    }

    $osDrive = Get-OperatingSystemDrive -WindowsIdentity $windowsIdentity
    Write-Log ('Detected operating-system drive: {0}' -f $osDrive)

    $script:RestorePointsBefore = @(
        Get-AllRestorePoints | ForEach-Object { Convert-RestorePointRecord -RestorePoint $_ }
    )
    Write-RestorePointInventorySummary -Label 'before' -RestorePoints $script:RestorePointsBefore

    Test-SystemRestorePolicy
    Enable-SystemRestoreProtection -Drive $osDrive -InitializationWaitSeconds 20
    Set-RestorePointCreationFrequency
    Initialize-ShadowCopyServices

    try {
        $systemRestoreClass = Get-WmiObject -Namespace 'root\default' -List -Class 'SystemRestore' -ErrorAction Stop
        if ($systemRestoreClass) {
            Write-Log 'SystemRestore WMI provider is available.' 'OK'
        }
    }
    catch {
        Write-Log ('SystemRestore WMI provider pre-check failed: {0}' -f $_.Exception.Message) 'WARN'
    }


    [void](New-VerifiedRestorePoint `
        -Description $RestorePointDescription `
        -TimeoutMinutes $VerificationTimeoutMinutes)

    if ($SkipRetentionCleanup) {
        Write-Log 'Managed restore-point retention cleanup was skipped by parameter.' 'WARN'
    }
    else {
        Remove-ObsoleteManagedRestorePoints `
            -ManagedDescription $RestorePointDescription `
            -KeepCount $MaxManagedRestorePoints
    }

    $script:RestorePointsAfter = @(
        Get-AllRestorePoints | ForEach-Object { Convert-RestorePointRecord -RestorePoint $_ }
    )
    Write-RestorePointInventorySummary -Label 'after' -RestorePoints $script:RestorePointsAfter

    $retentionFailures = @($script:RetentionResults | ForEach-Object { $_ } | Where-Object { $_.Status -eq 'Failed' }).Count
    if ($retentionFailures -gt 0 -or $script:WarningCount -gt 0) {
        $script:FinalStatus = 'SuccessWithWarnings'
    }
    else {
        $script:FinalStatus = 'Success'
    }

    $script:FinalExitCode = 0
    Write-Log 'System Restore is enabled and the new restore point was created and verified.' 'OK'
    Write-Log ('===== System Restore script completed with status {0} =====' -f $script:FinalStatus) 'OK'
}
catch {
    $script:FailureMessage = $_.Exception.Message
    $script:FinalStatus = 'Failed'
    $script:FinalExitCode = 1
    Write-Log ('System Restore operation failed: {0}' -f $script:FailureMessage) 'ERROR'
    Write-Log '===== System Restore script completed with errors =====' 'ERROR'
}
finally {
    if ($script:RestorePointsAfter.Count -eq 0) {
        $script:RestorePointsAfter = @(
            Get-AllRestorePoints | ForEach-Object { Convert-RestorePointRecord -RestorePoint $_ }
        )
    }

    if (-not $windowsIdentity) {
        try { $windowsIdentity = Get-WindowsIdentity } catch { $windowsIdentity = $null }
    }

    try {
        Write-Telemetry -WindowsIdentity $windowsIdentity
    }
    catch {
        Write-Log ('Telemetry write failed: {0}' -f $_.Exception.Message) 'ERROR'
        if ($script:FinalExitCode -eq 0) {
            $script:FinalStatus = 'TelemetryFailure'
            $script:FinalExitCode = 4
        }
    }

    # Final append before the completed immutable text log enters C:\Logs.
    Write-Log ('Finalizing {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}' -f `
        $ScriptName,
        $script:FinalStatus,
        $script:FinalExitCode,
        $script:WarningCount,
        $script:ErrorCount) $(if ($script:FinalExitCode -eq 0) { 'OK' } else { 'ERROR' })

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 12 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 12 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $script:FinalExitCode
