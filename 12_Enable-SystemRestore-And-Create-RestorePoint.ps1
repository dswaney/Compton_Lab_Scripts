<#
.SYNOPSIS
    Enables System Restore on the Windows client operating-system drive, creates a verified
    restore point, and retains a configurable number of restore points created by this script.

.NOTES
    ScriptName: 12_Enable-SystemRestore-And-Create-RestorePoint.ps1
    ScriptVersion: 2.0.0
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
$ScriptVersion = '2.0.0'
$RunId = [guid]::NewGuid().Guid
$RunStart = Get-Date
$LogDirectory = 'C:\Logs'
$LogPath = Join-Path $LogDirectory '12_Enable-SystemRestore-And-Create-RestorePoint.log'
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
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

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
    $line = '[{0}] [{1,-5}] {2}' -f $timestamp, $Level, $Message

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }

    try {
        Ensure-Directory -Path $LogDirectory
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

    [pscustomobject]@{
        ProductName    = if ($cv) { [string]$cv.ProductName } else { [string]$os.Caption }
        EditionId      = if ($cv) { [string]$cv.EditionID } else { $null }
        DisplayVersion = if ($cv) { [string]$cv.DisplayVersion } else { $null }
        ReleaseId      = if ($cv) { [string]$cv.ReleaseId } else { $null }
        Build          = [string]$os.BuildNumber
        UBR            = if ($cv -and $null -ne $cv.UBR) { [int]$cv.UBR } else { $null }
        FullBuild      = if ($cv -and $null -ne $cv.UBR) { '{0}.{1}' -f $os.BuildNumber, $cv.UBR } else { [string]$os.BuildNumber }
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

function Enable-SystemRestoreProtection {
    param([Parameter(Mandatory)][string]$Drive)

    $script:RestoreProtectionEnableAttempted = $true
    Write-Log ('Ensuring System Restore protection is enabled for {0}' -f $Drive)

    Enable-ComputerRestore -Drive $Drive -ErrorAction Stop
    $script:RestoreProtectionEnabled = $true
    Write-Log ('System Restore protection enable command completed for {0}' -f $Drive) 'OK'
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
    foreach ($serviceName in @('VSS', 'swprv')) {
        $before = Get-ServiceSnapshot -Name $serviceName
        $attempted = $false
        $operationStatus = 'NoChange'
        $message = $null

        if (-not $before.Exists) {
            $operationStatus = 'NotFound'
            $message = 'Required service was not found.'
            Write-Log ('Service {0} was not found.' -f $serviceName) 'WARN'
        }
        elseif ($before.State -ne 'Running') {
            $attempted = $true
            Write-Log ('Starting required service {0}...' -f $serviceName)

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
            Write-Log ('Required service {0} is already running.' -f $serviceName)
        }

        $after = Get-ServiceSnapshot -Name $serviceName
        $verified = ($after.Exists -and $after.State -eq 'Running')

        [void]$script:ServiceResults.Add([pscustomobject]@{
            Name            = $serviceName
            Before          = $before
            StartAttempted  = $attempted
            OperationStatus = $operationStatus
            Message         = $message
            After           = $after
            VerifiedRunning = $verified
        })

        if (-not $verified) {
            Write-Log ('Service {0} is not verified as running; restore-point creation will still be attempted.' -f $serviceName) 'WARN'
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
        return @(Get-ComputerRestorePoint -ErrorAction Stop | Sort-Object -Property SequenceNumber -Descending)
    }
    catch {
        Write-Log ('Unable to enumerate restore points: {0}' -f $_.Exception.Message) 'WARN'
        return @()
    }
}

function Get-LatestRestorePoint {
    $points = Get-AllRestorePoints
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

    $all = Get-AllRestorePoints
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

    Checkpoint-Computer `
        -Description $safeDescription `
        -RestorePointType MODIFY_SETTINGS `
        -ErrorAction Stop

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    do {
        Start-Sleep -Seconds 5
        $points = Get-AllRestorePoints
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
            ShadowCopyServices = @($script:ServiceResults)
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
                RestorePointsBefore = $script:RestorePointsBefore
                RestorePointsAfter = $script:RestorePointsAfter
            }
            Retention = [ordered]@{
                CleanupSkipped = [bool]$SkipRetentionCleanup
                DeletedCount = @($script:RetentionResults | Where-Object { $_.Status -eq 'Deleted' }).Count
                FailedCount = @($script:RetentionResults | Where-Object { $_.Status -eq 'Failed' }).Count
                Results = @($script:RetentionResults)
                NonManagedRestorePointsPreserved = $true
            }
        }

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

    Enable-SystemRestoreProtection -Drive $osDrive
    Set-RestorePointCreationFrequency
    Initialize-ShadowCopyServices

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

    $retentionFailures = @($script:RetentionResults | Where-Object { $_.Status -eq 'Failed' }).Count
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

    Write-Telemetry -WindowsIdentity $windowsIdentity
}

exit $script:FinalExitCode
