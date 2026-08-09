#requires -Version 5.1
#requires -RunAsAdministrator
<#
.SYNOPSIS
    Creates one authoritative weekly endpoint health snapshot for Elastic.

.DESCRIPTION
    Collects hardware, operating system, performance, storage, security,
    networking, update, event-log, service, and management-agent health.

    The script is fault-tolerant: individual collectors can fail without
    preventing the remaining endpoint snapshot from being written.

.OUTPUTS
    C:\Logs\Maintenance-Telemetry.ndjson
    C:\Logs\14_Endpoint_Health_Inventory.latest.json
    C:\Logs\14_Endpoint_Health_Inventory.log

.NOTES
    ScriptName:    14_Endpoint_Health_Inventory.ps1
    ScriptVersion: 1.1.0
    Designed for:  Windows PowerShell 5.1
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 120)]
    [int]$CpuSampleSeconds = 5,

    [ValidateRange(1, 168)]
    [int]$EventLookbackHours = 168,

    [ValidateRange(1, 365)]
    [int]$UpdateLookbackDays = 60,

    [ValidateRange(1, 100)]
    [int]$CpuWarningPercent = 85,

    [ValidateRange(1, 100)]
    [int]$MemoryWarningPercent = 85,

    [ValidateRange(1, 100)]
    [int]$DiskWarningPercent = 85,

    [ValidateRange(1, 100)]
    [int]$DiskCriticalPercent = 95,

    [ValidateRange(1, 10000)]
    [int]$CriticalEventWarningCount = 5,

    [string[]]$CriticalServices = @(
        'WinDefend',
        'mpssvc',
        'W32Time',
        'EventLog',
        'Schedule',
        'RpcSs',
        'Dnscache',
        'LanmanWorkstation'
    ),

    [string[]]$ManagementAgentPatterns = @(
        'Action1',
        'Wazuh',
        'Elastic Agent',
        'Microsoft Intune',
        'Tanium',
        'CrowdStrike',
        'SentinelOne',
        'Qualys',
        'Ninja',
        'ConnectWise',
        'Kaseya'
    ),

    [string]$LogDirectory = 'C:\Logs'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ScriptName       = '14_Endpoint_Health_Inventory.ps1'
$script:ScriptVersion    = '1.1.0'
$script:RunId            = [guid]::NewGuid().Guid
$script:StartTime        = Get-Date
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:Findings         = New-Object System.Collections.Generic.List[object]
$script:CollectorResults = New-Object System.Collections.Generic.List[object]

$script:NdjsonPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestPath = Join-Path $LogDirectory '14_Endpoint_Health_Inventory.latest.json'
$script:LogPath    = Join-Path $LogDirectory '14_Endpoint_Health_Inventory.log'

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $line = '{0} [{1}] {2}' -f $timestamp, $Level, $Message

    try {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    } catch {
        Write-Verbose "Unable to write runtime log: $($_.Exception.Message)"
    }

    switch ($Level) {
        'WARN'  { $script:WarningCount++ }
        'ERROR' { $script:ErrorCount++ }
    }

    Write-Verbose $line
}

function Add-Finding {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Healthy','Info','Warning','Critical')]
        [string]$Severity,

        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Check,
        [Parameter(Mandatory)][string]$Message,
        $Value = $null
    )

    $script:Findings.Add([pscustomobject]@{
        Severity = $Severity
        Category = $Category
        Check    = $Check
        Message  = $Message
        Value    = $Value
    })
}

function Invoke-Collector {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    $started = Get-Date
    try {
        $data = & $ScriptBlock
        $result = [pscustomobject]@{
            Name            = $Name
            Status          = 'Success'
            StartTime       = $started.ToUniversalTime().ToString('o')
            EndTime         = (Get-Date).ToUniversalTime().ToString('o')
            DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 3)
            Error           = $null
        }
        $script:CollectorResults.Add($result)
        return $data
    } catch {
        $message = $_.Exception.Message
        Write-Log -Level 'ERROR' -Message ("Collector '{0}' failed: {1}" -f $Name, $message)
        Add-Finding -Severity 'Warning' -Category 'Collection' -Check $Name -Message $message
        $script:CollectorResults.Add([pscustomobject]@{
            Name            = $Name
            Status          = 'Failed'
            StartTime       = $started.ToUniversalTime().ToString('o')
            EndTime         = (Get-Date).ToUniversalTime().ToString('o')
            DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 3)
            Error           = $message
        })
        return $null
    }
}

function Convert-CimDate {
    param($Value)
    if (-not $Value) { return $null }

    try {
        if ($Value -is [datetime]) { return $Value }
        return [Management.ManagementDateTimeConverter]::ToDateTime([string]$Value)
    } catch {
        try { return [datetime]$Value } catch { return $null }
    }
}

function Get-RegistryValueSafe {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
        return $item.$Name
    } catch {
        return $null
    }
}

function Test-PendingReboot {
    $reasons = New-Object System.Collections.Generic.List[string]

    $paths = @{
        'ComponentBasedServicing' = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        'WindowsUpdate'           = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
        'PackagesPending'         = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    }

    foreach ($entry in $paths.GetEnumerator()) {
        if (Test-Path -LiteralPath $entry.Value) {
            $reasons.Add($entry.Key)
        }
    }

    $sessionPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    $pendingRename = Get-RegistryValueSafe -Path $sessionPath -Name 'PendingFileRenameOperations'
    if ($pendingRename) { $reasons.Add('PendingFileRenameOperations') }

    $pendingRename2 = Get-RegistryValueSafe -Path $sessionPath -Name 'PendingFileRenameOperations2'
    if ($pendingRename2) { $reasons.Add('PendingFileRenameOperations2') }

    $updateExe = Get-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Microsoft\Updates' -Name 'UpdateExeVolatile'
    if ($null -ne $updateExe -and [int]$updateExe -ne 0) {
        $reasons.Add('UpdateExeVolatile')
    }

    try {
        $sccm = Invoke-CimMethod -Namespace 'root\ccm\ClientSDK' `
            -ClassName 'CCM_ClientUtilities' -MethodName 'DetermineIfRebootPending' `
            -ErrorAction Stop

        if ($sccm.RebootPending -or $sccm.IsHardRebootPending) {
            $reasons.Add('ConfigurationManager')
        }
    } catch {
        # SCCM client is optional.
    }

    [pscustomobject]@{
        Pending = ($reasons.Count -gt 0)
        Reasons = @($reasons | Select-Object -Unique)
    }
}

function Get-ComputerIdentity {
    $computer = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_BIOS
    $baseboard = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
    $enclosure = Get-CimInstance Win32_SystemEnclosure -ErrorAction SilentlyContinue

    [pscustomobject]@{
        ComputerName  = $env:COMPUTERNAME
        Domain        = $computer.Domain
        DomainJoined  = [bool]$computer.PartOfDomain
        Manufacturer  = $computer.Manufacturer
        Model         = $computer.Model
        SystemType    = $computer.SystemType
        SystemFamily  = $computer.SystemFamily
        SystemSKU     = $computer.SystemSKUNumber
        SerialNumber  = $bios.SerialNumber
        BIOSVersion   = (($bios.SMBIOSBIOSVersion, $bios.Version | Where-Object { $_ }) -join ' / ')
        BIOSDate      = if ($bios.ReleaseDate) { (Convert-CimDate $bios.ReleaseDate).ToUniversalTime().ToString('o') } else { $null }
        Baseboard     = if ($baseboard) { $baseboard.Product } else { $null }
        AssetTag      = if ($enclosure) { $enclosure.SMBIOSAssetTag } else { $null }
        ChassisTypes  = if ($enclosure) { @($enclosure.ChassisTypes) } else { @() }
        TotalMemoryGB = [math]::Round($computer.TotalPhysicalMemory / 1GB, 2)
    }
}

function Get-OperatingSystemHealth {
    $os = Get-CimInstance Win32_OperatingSystem
    $lastBoot = Convert-CimDate $os.LastBootUpTime
    $now = Get-Date
    $ubr = Get-RegistryValueSafe `
        -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
        -Name 'UBR'
    $displayVersion = Get-RegistryValueSafe `
        -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
        -Name 'DisplayVersion'
    $editionId = Get-RegistryValueSafe `
        -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
        -Name 'EditionID'

    $pending = Test-PendingReboot
    if ($pending.Pending) {
        Add-Finding -Severity 'Warning' -Category 'OperatingSystem' `
            -Check 'PendingReboot' -Message ('Pending reboot: {0}' -f ($pending.Reasons -join ', ')) `
            -Value $pending.Reasons
    } else {
        Add-Finding -Severity 'Healthy' -Category 'OperatingSystem' `
            -Check 'PendingReboot' -Message 'No pending reboot indicators were found.'
    }

    [pscustomobject]@{
        Caption              = $os.Caption
        Edition              = $editionId
        DisplayVersion       = $displayVersion
        Version              = $os.Version
        BuildNumber          = $os.BuildNumber
        UBR                  = $ubr
        FullBuild            = if ($null -ne $ubr) { '{0}.{1}' -f $os.BuildNumber, $ubr } else { [string]$os.BuildNumber }
        Architecture         = $os.OSArchitecture
        InstallDate          = if ($os.InstallDate) { (Convert-CimDate $os.InstallDate).ToUniversalTime().ToString('o') } else { $null }
        LastBootTime         = if ($lastBoot) { $lastBoot.ToUniversalTime().ToString('o') } else { $null }
        UptimeSeconds        = if ($lastBoot) { [math]::Round(($now - $lastBoot).TotalSeconds, 0) } else { $null }
        UptimeDays           = if ($lastBoot) { [math]::Round(($now - $lastBoot).TotalDays, 2) } else { $null }
        PendingReboot        = $pending.Pending
        PendingRebootReasons = $pending.Reasons
    }
}

function Get-PerformanceHealth {
    $os = Get-CimInstance Win32_OperatingSystem
    $cpuSamples = @()

    try {
        $counter = Get-Counter '\Processor(_Total)\% Processor Time' `
            -SampleInterval 1 -MaxSamples $CpuSampleSeconds -ErrorAction Stop
        $cpuSamples = @($counter.CounterSamples | ForEach-Object { [double]$_.CookedValue })
    } catch {
        $processor = Get-CimInstance Win32_Processor
        $cpuSamples = @($processor | ForEach-Object { [double]$_.LoadPercentage })
    }

    $cpuAverage = if ($cpuSamples.Count -gt 0) {
        [math]::Round(($cpuSamples | Measure-Object -Average).Average, 2)
    } else { $null }

    $cpuMaximum = if ($cpuSamples.Count -gt 0) {
        [math]::Round(($cpuSamples | Measure-Object -Maximum).Maximum, 2)
    } else { $null }

    $totalMemoryGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $freeMemoryGB  = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $usedMemoryGB  = [math]::Round($totalMemoryGB - $freeMemoryGB, 2)
    $memoryPercent = if ($totalMemoryGB -gt 0) {
        [math]::Round(($usedMemoryGB / $totalMemoryGB) * 100, 2)
    } else { $null }

    if ($null -ne $cpuAverage -and $cpuAverage -ge $CpuWarningPercent) {
        Add-Finding -Severity 'Warning' -Category 'Performance' -Check 'CPUUtilization' `
            -Message ("Average CPU utilization is {0}%." -f $cpuAverage) -Value $cpuAverage
    } else {
        Add-Finding -Severity 'Healthy' -Category 'Performance' -Check 'CPUUtilization' `
            -Message ("Average CPU utilization is {0}%." -f $cpuAverage) -Value $cpuAverage
    }

    if ($null -ne $memoryPercent -and $memoryPercent -ge $MemoryWarningPercent) {
        Add-Finding -Severity 'Warning' -Category 'Performance' -Check 'MemoryUtilization' `
            -Message ("Memory utilization is {0}%." -f $memoryPercent) -Value $memoryPercent
    } else {
        Add-Finding -Severity 'Healthy' -Category 'Performance' -Check 'MemoryUtilization' `
            -Message ("Memory utilization is {0}%." -f $memoryPercent) -Value $memoryPercent
    }

    [pscustomobject]@{
        CpuSampleSeconds    = $CpuSampleSeconds
        CpuAveragePercent   = $cpuAverage
        CpuMaximumPercent   = $cpuMaximum
        CpuSamples          = $cpuSamples
        MemoryTotalGB       = $totalMemoryGB
        MemoryUsedGB        = $usedMemoryGB
        MemoryFreeGB        = $freeMemoryGB
        MemoryUsedPercent   = $memoryPercent
    }
}

function Get-DiskHealth {
    $volumes = @()
    $logicalDisks = Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3'

    foreach ($disk in $logicalDisks) {
        $used = [double]$disk.Size - [double]$disk.FreeSpace
        $usedPercent = if ($disk.Size -gt 0) {
            [math]::Round(($used / [double]$disk.Size) * 100, 2)
        } else { $null }

        $severity = 'Healthy'
        if ($usedPercent -ge $DiskCriticalPercent) {
            $severity = 'Critical'
        } elseif ($usedPercent -ge $DiskWarningPercent) {
            $severity = 'Warning'
        }

        Add-Finding -Severity $severity -Category 'Storage' `
            -Check ('DiskUtilization:{0}' -f $disk.DeviceID) `
            -Message ("Drive {0} is {1}% used." -f $disk.DeviceID, $usedPercent) `
            -Value $usedPercent

        $volumes += [pscustomobject]@{
            DriveLetter      = $disk.DeviceID
            VolumeName       = $disk.VolumeName
            FileSystem       = $disk.FileSystem
            SizeGB           = [math]::Round($disk.Size / 1GB, 2)
            UsedGB           = [math]::Round($used / 1GB, 2)
            FreeGB           = [math]::Round($disk.FreeSpace / 1GB, 2)
            UsedPercent      = $usedPercent
            Compressed       = [bool]$disk.Compressed
        }
    }

    $physical = @()
    if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {
        foreach ($pd in Get-PhysicalDisk -ErrorAction SilentlyContinue) {
            $reliability = $null
            try {
                $reliability = Get-StorageReliabilityCounter -PhysicalDisk $pd -ErrorAction Stop
            } catch {}

            $healthSeverity = if ($pd.HealthStatus -eq 'Healthy') { 'Healthy' } else { 'Critical' }
            Add-Finding -Severity $healthSeverity -Category 'Storage' `
                -Check ('PhysicalDisk:{0}' -f $pd.FriendlyName) `
                -Message ("Physical disk {0} health is {1}." -f $pd.FriendlyName, $pd.HealthStatus) `
                -Value $pd.HealthStatus

            $physical += [pscustomobject]@{
                FriendlyName           = $pd.FriendlyName
                SerialNumber           = $pd.SerialNumber
                MediaType              = [string]$pd.MediaType
                BusType                = [string]$pd.BusType
                SizeGB                 = [math]::Round($pd.Size / 1GB, 2)
                HealthStatus           = [string]$pd.HealthStatus
                OperationalStatus      = @($pd.OperationalStatus | ForEach-Object { [string]$_ })
                FirmwareVersion        = $pd.FirmwareVersion
                Wear                   = if ($reliability) { $reliability.Wear } else { $null }
                TemperatureC           = if ($reliability) { $reliability.Temperature } else { $null }
                ReadErrorsTotal        = if ($reliability) { $reliability.ReadErrorsTotal } else { $null }
                WriteErrorsTotal       = if ($reliability) { $reliability.WriteErrorsTotal } else { $null }
                PowerOnHours           = if ($reliability) { $reliability.PowerOnHours } else { $null }
            }
        }
    } else {
        foreach ($drive in Get-CimInstance Win32_DiskDrive) {
            $physical += [pscustomobject]@{
                FriendlyName      = $drive.Model
                SerialNumber      = $drive.SerialNumber
                MediaType         = $drive.MediaType
                BusType           = $drive.InterfaceType
                SizeGB            = [math]::Round($drive.Size / 1GB, 2)
                HealthStatus      = if ($drive.Status -eq 'OK') { 'Healthy' } else { $drive.Status }
                OperationalStatus = @($drive.Status)
                FirmwareVersion   = $drive.FirmwareRevision
                Wear              = $null
                TemperatureC      = $null
                ReadErrorsTotal   = $null
                WriteErrorsTotal  = $null
                PowerOnHours      = $null
            }
        }
    }

    [pscustomobject]@{
        Volumes       = $volumes
        PhysicalDisks = $physical
    }
}

function Get-DeviceManagerHealth {
    $devices = @()

    if (Get-Command Get-PnpDevice -ErrorAction SilentlyContinue) {
        $problems = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
            Where-Object { $_.Status -ne 'OK' -or $_.Problem -ne 0 }

        foreach ($device in $problems) {
            $devices += [pscustomobject]@{
                Class        = $device.Class
                FriendlyName = $device.FriendlyName
                InstanceId   = $device.InstanceId
                Status       = [string]$device.Status
                ProblemCode  = $device.Problem
            }
        }
    } else {
        $problems = Get-CimInstance Win32_PnPEntity |
            Where-Object { $null -ne $_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode -ne 0 }

        foreach ($device in $problems) {
            $devices += [pscustomobject]@{
                Class        = $device.PNPClass
                FriendlyName = $device.Name
                InstanceId   = $device.PNPDeviceID
                Status       = $device.Status
                ProblemCode  = $device.ConfigManagerErrorCode
            }
        }
    }

    if ($devices.Count -gt 0) {
        Add-Finding -Severity 'Warning' -Category 'Hardware' -Check 'DeviceManagerProblems' `
            -Message ("{0} Device Manager problem(s) detected." -f $devices.Count) `
            -Value $devices.Count
    } else {
        Add-Finding -Severity 'Healthy' -Category 'Hardware' -Check 'DeviceManagerProblems' `
            -Message 'No Device Manager problems were detected.' -Value 0
    }

    [pscustomobject]@{
        ProblemCount = $devices.Count
        Problems     = $devices
    }
}

function Get-DefenderHealth {
    if (-not (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)) {
        Add-Finding -Severity 'Info' -Category 'Security' -Check 'Defender' `
            -Message 'Microsoft Defender cmdlets are unavailable.'
        return [pscustomobject]@{
            Available = $false
            Status    = 'Unavailable'
        }
    }

    $status = Get-MpComputerStatus
    $threats = @()
    if (Get-Command Get-MpThreatDetection -ErrorAction SilentlyContinue) {
        $threats = @(Get-MpThreatDetection -ErrorAction SilentlyContinue |
            Select-Object -First 20 ThreatID, InitialDetectionTime, LastThreatStatusChangeTime, ActionSuccess)
    }

    $healthy = [bool]$status.AntivirusEnabled -and
               [bool]$status.RealTimeProtectionEnabled -and
               -not [bool]$status.ComputerState

    $severity = if ($healthy) { 'Healthy' } else { 'Critical' }
    Add-Finding -Severity $severity -Category 'Security' -Check 'Defender' `
        -Message ("Defender antivirus enabled={0}; real-time protection enabled={1}." -f
            $status.AntivirusEnabled, $status.RealTimeProtectionEnabled)

    [pscustomobject]@{
        Available                         = $true
        AntivirusEnabled                  = [bool]$status.AntivirusEnabled
        AntispywareEnabled                = [bool]$status.AntispywareEnabled
        RealTimeProtectionEnabled         = [bool]$status.RealTimeProtectionEnabled
        BehaviorMonitorEnabled            = [bool]$status.BehaviorMonitorEnabled
        IoavProtectionEnabled             = [bool]$status.IoavProtectionEnabled
        NISEnabled                        = [bool]$status.NISEnabled
        AntivirusSignatureVersion         = $status.AntivirusSignatureVersion
        AntivirusSignatureLastUpdated     = if ($status.AntivirusSignatureLastUpdated) {
            ([datetime]$status.AntivirusSignatureLastUpdated).ToUniversalTime().ToString('o')
        } else { $null }
        QuickScanAge                      = $status.QuickScanAge
        FullScanAge                       = $status.FullScanAge
        RebootRequired                    = [bool]$status.RebootRequired
        IsTamperProtected                 = [bool]$status.IsTamperProtected
        RecentThreatDetections            = $threats
        RecentThreatDetectionCount        = $threats.Count
    }
}

function Get-FirewallHealth {
    $profiles = @()
    if (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
        foreach ($profile in Get-NetFirewallProfile) {
            $profiles += [pscustomobject]@{
                Name                   = [string]$profile.Name
                Enabled                = [bool]$profile.Enabled
                DefaultInboundAction   = [string]$profile.DefaultInboundAction
                DefaultOutboundAction  = [string]$profile.DefaultOutboundAction
                NotifyOnListen         = [bool]$profile.NotifyOnListen
                LogAllowed             = [bool]$profile.LogAllowed
                LogBlocked             = [bool]$profile.LogBlocked
                LogFileName            = $profile.LogFileName
            }
        }
    }

    $disabled = @($profiles | Where-Object { -not $_.Enabled })
    if ($profiles.Count -eq 0) {
        Add-Finding -Severity 'Warning' -Category 'Security' -Check 'Firewall' `
            -Message 'Windows Firewall profile information could not be collected.'
    } elseif ($disabled.Count -gt 0) {
        Add-Finding -Severity 'Critical' -Category 'Security' -Check 'Firewall' `
            -Message ("Disabled firewall profile(s): {0}" -f (($disabled.Name) -join ', ')) `
            -Value $disabled.Name
    } else {
        Add-Finding -Severity 'Healthy' -Category 'Security' -Check 'Firewall' `
            -Message 'All Windows Firewall profiles are enabled.'
    }

    [pscustomobject]@{
        ProfileCount         = $profiles.Count
        DisabledProfileCount = $disabled.Count
        Profiles             = $profiles
    }
}

function Get-BitLockerHealth {
    if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
        Add-Finding -Severity 'Info' -Category 'Security' -Check 'BitLocker' `
            -Message 'BitLocker cmdlets are unavailable.'
        return [pscustomobject]@{
            Available = $false
            Volumes   = @()
        }
    }

    $volumes = @()
    foreach ($volume in Get-BitLockerVolume -ErrorAction SilentlyContinue) {
        $volumes += [pscustomobject]@{
            MountPoint          = $volume.MountPoint
            VolumeType          = [string]$volume.VolumeType
            VolumeStatus        = [string]$volume.VolumeStatus
            ProtectionStatus    = [string]$volume.ProtectionStatus
            EncryptionMethod    = [string]$volume.EncryptionMethod
            EncryptionPercentage= $volume.EncryptionPercentage
            AutoUnlockEnabled   = $volume.AutoUnlockEnabled
            LockStatus          = [string]$volume.LockStatus
            KeyProtectorTypes   = @($volume.KeyProtector | ForEach-Object { [string]$_.KeyProtectorType })
        }
    }

    $osVolume = $volumes | Where-Object { $_.MountPoint -eq $env:SystemDrive } | Select-Object -First 1
    if ($osVolume -and $osVolume.ProtectionStatus -eq 'On') {
        Add-Finding -Severity 'Healthy' -Category 'Security' -Check 'BitLocker' `
            -Message 'BitLocker protection is enabled on the operating-system volume.'
    } else {
        Add-Finding -Severity 'Warning' -Category 'Security' -Check 'BitLocker' `
            -Message 'BitLocker protection is not enabled on the operating-system volume.'
    }

    [pscustomobject]@{
        Available = $true
        Volumes   = $volumes
    }
}

function Get-TpmAndSecureBootHealth {
    $tpm = $null
    if (Get-Command Get-Tpm -ErrorAction SilentlyContinue) {
        $tpmData = Get-Tpm -ErrorAction SilentlyContinue
        if ($tpmData) {
            $tpm = [pscustomobject]@{
                Present               = [bool]$tpmData.TpmPresent
                Ready                 = [bool]$tpmData.TpmReady
                Enabled               = [bool]$tpmData.TpmEnabled
                Activated             = [bool]$tpmData.TpmActivated
                Owned                 = [bool]$tpmData.TpmOwned
                RestartPending        = [bool]$tpmData.RestartPending
                ManufacturerIdTxt     = $tpmData.ManufacturerIdTxt
                ManufacturerVersion   = $tpmData.ManufacturerVersion
                ManagedAuthLevel      = [string]$tpmData.ManagedAuthLevel
            }

            $severity = if ($tpm.Present -and $tpm.Ready) { 'Healthy' } else { 'Warning' }
            Add-Finding -Severity $severity -Category 'Security' -Check 'TPM' `
                -Message ("TPM present={0}; ready={1}." -f $tpm.Present, $tpm.Ready)
        }
    }

    $secureBoot = [pscustomobject]@{
        Supported = $null
        Enabled   = $null
        Error     = $null
    }

    if (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue) {
        try {
            $secureBoot.Enabled = [bool](Confirm-SecureBootUEFI -ErrorAction Stop)
            $secureBoot.Supported = $true
        } catch {
            if ($_.Exception.Message -match 'not supported') {
                $secureBoot.Supported = $false
            } else {
                $secureBoot.Error = $_.Exception.Message
            }
        }
    }

    if ($secureBoot.Supported -eq $true) {
        $severity = if ($secureBoot.Enabled) { 'Healthy' } else { 'Warning' }
        Add-Finding -Severity $severity -Category 'Security' -Check 'SecureBoot' `
            -Message ("Secure Boot enabled={0}." -f $secureBoot.Enabled)
    } elseif ($secureBoot.Supported -eq $false) {
        Add-Finding -Severity 'Info' -Category 'Security' -Check 'SecureBoot' `
            -Message 'Secure Boot is not supported on this system.'
    } else {
        Add-Finding -Severity 'Warning' -Category 'Security' -Check 'SecureBoot' `
            -Message 'Secure Boot state could not be determined.'
    }

    [pscustomobject]@{
        TPM        = $tpm
        SecureBoot = $secureBoot
    }
}

function Get-NetworkHealth {
    $adapters = @()
    $ipConfigs = @()

    if (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue) {
        foreach ($adapter in Get-NetAdapter -ErrorAction SilentlyContinue) {
            $adapters += [pscustomobject]@{
                Name              = $adapter.Name
                InterfaceAlias    = $adapter.InterfaceAlias
                InterfaceIndex    = $adapter.ifIndex
                Description       = $adapter.InterfaceDescription
                Status            = [string]$adapter.Status
                LinkSpeed         = [string]$adapter.LinkSpeed
                MacAddress        = $adapter.MacAddress
                MediaType         = [string]$adapter.MediaType
                PhysicalMediaType = [string]$adapter.PhysicalMediaType
                DriverVersion     = $adapter.DriverVersion
                DriverDate        = if ($adapter.DriverDate) { ([datetime]$adapter.DriverDate).ToString('yyyy-MM-dd') } else { $null }
                HardwareInterface = [bool]$adapter.HardwareInterface
            }
        }
    }

    if (Get-Command Get-NetIPConfiguration -ErrorAction SilentlyContinue) {
        foreach ($config in Get-NetIPConfiguration -Detailed -ErrorAction SilentlyContinue) {
            $dnsServers = @()
            try {
                $dnsServers = @((Get-DnsClientServerAddress -InterfaceIndex $config.InterfaceIndex -ErrorAction Stop).ServerAddresses)
            } catch {}

            $ipConfigs += [pscustomobject]@{
                InterfaceAlias = $config.InterfaceAlias
                InterfaceIndex = $config.InterfaceIndex
                NetProfileName = if ($config.NetProfile) { $config.NetProfile.Name } else { $null }
                NetworkCategory= if ($config.NetProfile) { [string]$config.NetProfile.NetworkCategory } else { $null }
                IPv4Addresses  = @($config.IPv4Address | ForEach-Object { $_.IPAddress })
                IPv6Addresses  = @($config.IPv6Address | ForEach-Object { $_.IPAddress })
                IPv4Gateway    = @($config.IPv4DefaultGateway | ForEach-Object { $_.NextHop })
                IPv6Gateway    = @($config.IPv6DefaultGateway | ForEach-Object { $_.NextHop })
                DNSServers     = $dnsServers
                DhcpEnabled    = try {
                    (Get-NetIPInterface -InterfaceIndex $config.InterfaceIndex -AddressFamily IPv4 -ErrorAction Stop).Dhcp -eq 'Enabled'
                } catch { $null }
            }
        }
    }

    $upPhysical = @($adapters | Where-Object { $_.HardwareInterface -and $_.Status -eq 'Up' })
    if ($upPhysical.Count -gt 0) {
        Add-Finding -Severity 'Healthy' -Category 'Network' -Check 'ActiveAdapter' `
            -Message ("{0} active physical network adapter(s)." -f $upPhysical.Count)
    } else {
        Add-Finding -Severity 'Critical' -Category 'Network' -Check 'ActiveAdapter' `
            -Message 'No active physical network adapter was detected.'
    }

    [pscustomobject]@{
        Adapters           = $adapters
        IPConfigurations   = $ipConfigs
        ActiveAdapterCount = $upPhysical.Count
    }
}

function Get-WindowsUpdateHealth {
    $hotfixes = @()
    try {
        $hotfixes = @(Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending)
    } catch {
        $hotfixes = @()
    }

    $latest = $hotfixes | Select-Object -First 1
    $recent = @($hotfixes | Where-Object {
        $_.InstalledOn -and $_.InstalledOn -ge (Get-Date).AddDays(-$UpdateLookbackDays)
    } | Select-Object HotFixID, Description, InstalledBy, InstalledOn)

    $history = @()
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $count = $searcher.GetTotalHistoryCount()
        if ($count -gt 0) {
            $take = [math]::Min($count, 50)
            $history = @($searcher.QueryHistory(0, $take) | ForEach-Object {
                [pscustomobject]@{
                    Date       = $_.Date.ToUniversalTime().ToString('o')
                    Title      = $_.Title
                    Operation  = [string]$_.Operation
                    ResultCode = [string]$_.ResultCode
                    HResult    = ('0x{0:X8}' -f ($_.HResult -band 0xffffffffL))
                }
            })
        }
    } catch {
        Write-Log -Level 'WARN' -Message ("Windows Update history could not be collected: {0}" -f $_.Exception.Message)
    }

    $latestDate = if ($latest -and $latest.InstalledOn) { [datetime]$latest.InstalledOn } else { $null }
    $ageDays = if ($latestDate) { [math]::Round(((Get-Date) - $latestDate).TotalDays, 1) } else { $null }

    if ($null -eq $latestDate) {
        Add-Finding -Severity 'Warning' -Category 'Updates' -Check 'LastWindowsUpdate' `
            -Message 'No installed hotfix date could be determined.'
    } elseif ($ageDays -gt $UpdateLookbackDays) {
        Add-Finding -Severity 'Warning' -Category 'Updates' -Check 'LastWindowsUpdate' `
            -Message ("Latest installed hotfix is {0} days old." -f $ageDays) -Value $ageDays
    } else {
        Add-Finding -Severity 'Healthy' -Category 'Updates' -Check 'LastWindowsUpdate' `
            -Message ("Latest installed hotfix is {0} days old." -f $ageDays) -Value $ageDays
    }

    [pscustomobject]@{
        LatestHotfix = if ($latest) {
            [pscustomobject]@{
                HotFixID    = $latest.HotFixID
                Description = $latest.Description
                InstalledBy = $latest.InstalledBy
                InstalledOn = if ($latest.InstalledOn) { ([datetime]$latest.InstalledOn).ToUniversalTime().ToString('o') } else { $null }
                AgeDays     = $ageDays
            }
        } else { $null }
        RecentHotfixCount = $recent.Count
        RecentHotfixes    = $recent
        UpdateHistory     = $history
    }
}

function Get-EventLogHealth {
    $start = (Get-Date).AddHours(-$EventLookbackHours)

    $applicationCrashes = @()
    try {
        $applicationCrashes = @(Get-WinEvent -FilterHashtable @{
            LogName   = 'Application'
            StartTime = $start
            Id        = 1000,1001,1002
        } -ErrorAction Stop | Select-Object -First 100 | ForEach-Object {
            [pscustomobject]@{
                TimeCreated = $_.TimeCreated.ToUniversalTime().ToString('o')
                Id          = $_.Id
                Provider    = $_.ProviderName
                Level       = $_.LevelDisplayName
                Message     = ($_.Message -replace '\s+', ' ').Trim()
            }
        })
    } catch {
        if ($_.Exception.Message -notmatch 'No events were found') {
            Write-Log -Level 'WARN' -Message ("Application crash event query failed: {0}" -f $_.Exception.Message)
        }
    }

    $criticalEvents = @()
    foreach ($log in 'System','Application') {
        try {
            $criticalEvents += @(Get-WinEvent -FilterHashtable @{
                LogName   = $log
                StartTime = $start
                Level     = 1
            } -ErrorAction Stop | Select-Object -First 100 | ForEach-Object {
                [pscustomobject]@{
                    LogName     = $log
                    TimeCreated = $_.TimeCreated.ToUniversalTime().ToString('o')
                    Id          = $_.Id
                    Provider    = $_.ProviderName
                    Message     = ($_.Message -replace '\s+', ' ').Trim()
                }
            })
        } catch {
            if ($_.Exception.Message -notmatch 'No events were found') {
                Write-Log -Level 'WARN' -Message ("Critical event query failed for {0}: {1}" -f $log, $_.Exception.Message)
            }
        }
    }

    $severity = if ($criticalEvents.Count -ge $CriticalEventWarningCount) { 'Warning' } else { 'Healthy' }
    Add-Finding -Severity $severity -Category 'EventLogs' -Check 'CriticalEvents' `
        -Message ("{0} critical event(s) in the last {1} hours." -f $criticalEvents.Count, $EventLookbackHours) `
        -Value $criticalEvents.Count

    if ($applicationCrashes.Count -gt 0) {
        Add-Finding -Severity 'Warning' -Category 'EventLogs' -Check 'ApplicationCrashes' `
            -Message ("{0} application crash/hang event(s) in the last {1} hours." -f
                $applicationCrashes.Count, $EventLookbackHours) `
            -Value $applicationCrashes.Count
    } else {
        Add-Finding -Severity 'Healthy' -Category 'EventLogs' -Check 'ApplicationCrashes' `
            -Message ("No application crash/hang events in the last {0} hours." -f $EventLookbackHours)
    }

    [pscustomobject]@{
        LookbackHours          = $EventLookbackHours
        ApplicationCrashCount  = $applicationCrashes.Count
        ApplicationCrashes     = $applicationCrashes
        CriticalEventCount     = $criticalEvents.Count
        CriticalEvents         = $criticalEvents
    }
}

function Get-ServiceHealth {
    $results = @()

    foreach ($serviceName in $CriticalServices) {
        $service = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $serviceName.Replace("'","''")) `
            -ErrorAction SilentlyContinue

        if (-not $service) {
            $results += [pscustomobject]@{
                Name        = $serviceName
                DisplayName = $null
                Exists      = $false
                State       = 'NotFound'
                StartMode   = $null
                ProcessId   = $null
                Healthy     = $null
            }
            Add-Finding -Severity 'Info' -Category 'Services' -Check $serviceName `
                -Message ("Service {0} is not installed." -f $serviceName)
            continue
        }

        $shouldRun = $service.StartMode -notin @('Disabled','Manual')
        $healthy = if ($shouldRun) { $service.State -eq 'Running' } else { $true }

        $results += [pscustomobject]@{
            Name        = $service.Name
            DisplayName = $service.DisplayName
            Exists      = $true
            State       = $service.State
            StartMode   = $service.StartMode
            ProcessId   = $service.ProcessId
            Healthy     = $healthy
        }

        $severity = if ($healthy) { 'Healthy' } else { 'Critical' }
        Add-Finding -Severity $severity -Category 'Services' -Check $service.Name `
            -Message ("Service {0}: state={1}, startup={2}." -f
                $service.Name, $service.State, $service.StartMode)
    }

    [pscustomobject]@{
        CheckedCount  = $results.Count
        UnhealthyCount= @($results | Where-Object { $_.Healthy -eq $false }).Count
        Services      = $results
    }
}

function Get-ManagementAgentHealth {
    $serviceInventory = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue)
    $processInventory = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)
    $uninstallEntries = @()

    foreach ($path in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )) {
        $uninstallEntries += @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName })
    }

    $agents = @()

    foreach ($pattern in $ManagementAgentPatterns) {
        $escaped = [regex]::Escape($pattern)

        $matchingServices = @($serviceInventory | Where-Object {
            $_.Name -match $escaped -or $_.DisplayName -match $escaped -or $_.PathName -match $escaped
        })

        $matchingProcesses = @($processInventory | Where-Object {
            $_.Name -match $escaped -or $_.ExecutablePath -match $escaped -or $_.CommandLine -match $escaped
        })

        $matchingProducts = @($uninstallEntries | Where-Object {
            $_.DisplayName -match $escaped -or $_.Publisher -match $escaped
        })

        if ($matchingServices.Count -gt 0 -or $matchingProcesses.Count -gt 0 -or $matchingProducts.Count -gt 0) {
            $services = @($matchingServices | ForEach-Object {
                [pscustomobject]@{
                    Name        = $_.Name
                    DisplayName = $_.DisplayName
                    State       = $_.State
                    StartMode   = $_.StartMode
                    ProcessId   = $_.ProcessId
                }
            })

            $products = @($matchingProducts | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate)
            $processes = @($matchingProcesses | Select-Object Name, ProcessId, ExecutablePath)

            $serviceHealthy = if ($services.Count -gt 0) {
                @($services | Where-Object { $_.State -eq 'Running' }).Count -gt 0
            } else {
                $processes.Count -gt 0
            }

            $agents += [pscustomobject]@{
                Pattern        = $pattern
                Detected       = $true
                Healthy        = $serviceHealthy
                Services       = $services
                Processes      = $processes
                InstalledItems = $products
            }

            $severity = if ($serviceHealthy) { 'Healthy' } else { 'Warning' }
            Add-Finding -Severity $severity -Category 'ManagementAgents' -Check $pattern `
                -Message ("Management agent '{0}' detected; healthy={1}." -f $pattern, $serviceHealthy)
        }
    }

    if ($agents.Count -eq 0) {
        Add-Finding -Severity 'Warning' -Category 'ManagementAgents' -Check 'AgentDetection' `
            -Message 'No recognized endpoint management or security agent was detected.'
    }

    [pscustomobject]@{
        PatternsChecked = $ManagementAgentPatterns
        DetectedCount    = $agents.Count
        Agents           = $agents
    }
}


function Get-EdgeHealth {
    $edgePaths = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) }

    $edgePath = $edgePaths | Select-Object -First 1
    $version = $null
    $fileDate = $null
    if ($edgePath) {
        $edgeFile = Get-Item -LiteralPath $edgePath
        $version = $edgeFile.VersionInfo.ProductVersion
        $fileDate = $edgeFile.LastWriteTimeUtc.ToString('o')
    }

    $serviceNames = @('edgeupdate','edgeupdatem','MicrosoftEdgeElevationService')
    $services = @()
    foreach ($name in $serviceNames) {
        $svc = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $name) -ErrorAction SilentlyContinue
        if ($svc) {
            $services += [pscustomobject]@{
                Name      = $svc.Name
                DisplayName = $svc.DisplayName
                State     = $svc.State
                StartMode = $svc.StartMode
                PathName  = $svc.PathName
            }
        }
    }

    $tasks = @()
    if (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue) {
        $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskName -match 'MicrosoftEdgeUpdate|Microsoft Edge Update' } |
            ForEach-Object {
                $info = Get-ScheduledTaskInfo -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue
                [pscustomobject]@{
                    TaskName       = $_.TaskName
                    TaskPath       = $_.TaskPath
                    State          = [string]$_.State
                    LastRunTime    = if ($info -and $info.LastRunTime) { $info.LastRunTime.ToUniversalTime().ToString('o') } else { $null }
                    LastTaskResult = if ($info) { $info.LastTaskResult } else { $null }
                    NextRunTime    = if ($info -and $info.NextRunTime) { $info.NextRunTime.ToUniversalTime().ToString('o') } else { $null }
                }
            })
    }

    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate'
    $policy = [ordered]@{
        Present                      = Test-Path -LiteralPath $policyPath
        UpdateDefault                = Get-RegistryValueSafe -Path $policyPath -Name 'UpdateDefault'
        StableChannelUpdate          = Get-RegistryValueSafe -Path $policyPath -Name 'Update{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}'
        AutoUpdateCheckPeriodMinutes = Get-RegistryValueSafe -Path $policyPath -Name 'AutoUpdateCheckPeriodMinutes'
        TargetVersionPrefix          = Get-RegistryValueSafe -Path $policyPath -Name 'TargetVersionPrefix{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}'
        RollbackToTargetVersion      = Get-RegistryValueSafe -Path $policyPath -Name 'RollbackToTargetVersion{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}'
    }

    $updatesDisabled = ($policy.UpdateDefault -eq 0 -or $policy.StableChannelUpdate -eq 0)
    $updaterPresent = ($services.Count -gt 0 -or $tasks.Count -gt 0)

    if (-not $edgePath) {
        Add-Finding -Severity 'Warning' -Category 'Applications' -Check 'MicrosoftEdge' `
            -Message 'Microsoft Edge executable was not found.'
        $health = 'Warning'
    } elseif ($updatesDisabled) {
        Add-Finding -Severity 'Critical' -Category 'Applications' -Check 'MicrosoftEdgeUpdate' `
            -Message 'Microsoft Edge updates are disabled by policy.'
        $health = 'Critical'
    } elseif (-not $updaterPresent) {
        Add-Finding -Severity 'Warning' -Category 'Applications' -Check 'MicrosoftEdgeUpdate' `
            -Message 'Edge is installed, but no Edge Update service or scheduled task was detected.'
        $health = 'Warning'
    } else {
        Add-Finding -Severity 'Healthy' -Category 'Applications' -Check 'MicrosoftEdgeUpdate' `
            -Message ("Edge {0} is installed and Edge Update servicing is present." -f $version)
        $health = 'Healthy'
    }

    $edgeCrashes = @()
    try {
        $edgeCrashes = @(Get-WinEvent -FilterHashtable @{
            LogName   = 'Application'
            StartTime = (Get-Date).AddHours(-$EventLookbackHours)
            Id        = 1000,1001,1002
        } -ErrorAction Stop | Where-Object {
            $_.Message -match '(?i)msedge\.exe|Microsoft Edge'
        } | Select-Object -First 50 | ForEach-Object {
            [pscustomobject]@{
                TimeCreated = $_.TimeCreated.ToUniversalTime().ToString('o')
                Id          = $_.Id
                Provider    = $_.ProviderName
                Message     = ($_.Message -replace '\s+', ' ').Trim()
            }
        })
    } catch {
        if ($_.Exception.Message -notmatch 'No events were found') {
            Write-Log -Level 'WARN' -Message ("Edge crash event query failed: {0}" -f $_.Exception.Message)
        }
    }

    [pscustomobject]@{
        Installed          = [bool]$edgePath
        Version            = $version
        ExecutablePath     = $edgePath
        ExecutableModified = $fileDate
        Health             = $health
        UpdatesDisabled    = $updatesDisabled
        UpdateServices     = $services
        UpdateTasks        = $tasks
        UpdatePolicy       = [pscustomobject]$policy
        CrashCount         = $edgeCrashes.Count
        Crashes            = $edgeCrashes
    }
}

function Get-HealthSummary {
    $healthy  = @($script:Findings | Where-Object Severity -eq 'Healthy').Count
    $info     = @($script:Findings | Where-Object Severity -eq 'Info').Count
    $warnings = @($script:Findings | Where-Object Severity -eq 'Warning').Count
    $critical = @($script:Findings | Where-Object Severity -eq 'Critical').Count

    $status = if ($critical -gt 0) {
        'Critical'
    } elseif ($warnings -gt 0) {
        'Warning'
    } else {
        'Healthy'
    }

    [pscustomobject]@{
        Status        = $status
        HealthyCount  = $healthy
        InfoCount     = $info
        WarningCount  = $warnings
        CriticalCount = $critical
        Findings      = @($script:Findings)
    }
}

function Write-JsonAtomically {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Json
    )

    $temporary = '{0}.{1}.tmp' -f $Path, ([guid]::NewGuid().Guid)
    [IO.File]::WriteAllText($temporary, $Json, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Write-Telemetry {
    param([Parameter(Mandatory)]$Event)

    $jsonCompact = $Event | ConvertTo-Json -Depth 12 -Compress
    $jsonPretty  = $Event | ConvertTo-Json -Depth 12

    Write-MaintenanceTelemetryLine -Path $script:NdjsonPath -JsonLine $jsonCompact
    Write-JsonAtomically -Path $script:LatestPath -Json $jsonPretty
}

Ensure-Directory -Path $LogDirectory
Write-Log -Message ("Starting {0} version {1}. RunId={2}" -f
    $script:ScriptName, $script:ScriptVersion, $script:RunId)

$identity        = Invoke-Collector -Name 'ComputerIdentity'       -ScriptBlock { Get-ComputerIdentity }
$operatingSystem = Invoke-Collector -Name 'OperatingSystem'       -ScriptBlock { Get-OperatingSystemHealth }
$performance     = Invoke-Collector -Name 'Performance'           -ScriptBlock { Get-PerformanceHealth }
$storage         = Invoke-Collector -Name 'Storage'               -ScriptBlock { Get-DiskHealth }
$deviceManager   = Invoke-Collector -Name 'DeviceManager'         -ScriptBlock { Get-DeviceManagerHealth }
$defender        = Invoke-Collector -Name 'Defender'              -ScriptBlock { Get-DefenderHealth }
$firewall        = Invoke-Collector -Name 'Firewall'              -ScriptBlock { Get-FirewallHealth }
$bitLocker       = Invoke-Collector -Name 'BitLocker'             -ScriptBlock { Get-BitLockerHealth }
$tpmSecureBoot   = Invoke-Collector -Name 'TpmAndSecureBoot'      -ScriptBlock { Get-TpmAndSecureBootHealth }
$network         = Invoke-Collector -Name 'Network'               -ScriptBlock { Get-NetworkHealth }
$updates         = Invoke-Collector -Name 'WindowsUpdate'         -ScriptBlock { Get-WindowsUpdateHealth }
$eventLogs       = Invoke-Collector -Name 'EventLogs'             -ScriptBlock { Get-EventLogHealth }
$services        = Invoke-Collector -Name 'Services'              -ScriptBlock { Get-ServiceHealth }
$agents          = Invoke-Collector -Name 'ManagementAgents'      -ScriptBlock { Get-ManagementAgentHealth }
$edge            = Invoke-Collector -Name 'MicrosoftEdge'         -ScriptBlock { Get-EdgeHealth }

$summary = Get-HealthSummary
$endTime = Get-Date
$failedCollectors = @($script:CollectorResults | Where-Object Status -eq 'Failed').Count

$executionStatus = if ($failedCollectors -gt 0) {
    'SuccessWithWarnings'
} elseif ($summary.Status -eq 'Critical') {
    'SuccessWithCriticalFindings'
} elseif ($summary.Status -eq 'Warning') {
    'SuccessWithWarnings'
} else {
    'Success'
}

$exitCode = if ($failedCollectors -gt 0) {
    2
} elseif ($summary.Status -eq 'Critical') {
    3
} elseif ($summary.Status -eq 'Warning') {
    1
} else {
    0
}

$event = [ordered]@{
    '@timestamp' = $endTime.ToUniversalTime().ToString('o')
    EventType    = 'endpoint.health'
    SchemaVersion= '1.0'
    RunId        = $script:RunId

    ComputerName = $env:COMPUTERNAME
    Domain       = if ($identity) { $identity.Domain } else { $env:USERDOMAIN }

    ScriptName    = $script:ScriptName
    ScriptVersion = $script:ScriptVersion
    Status        = $executionStatus
    ExitCode      = $exitCode
    StartTime     = $script:StartTime.ToUniversalTime().ToString('o')
    EndTime       = $endTime.ToUniversalTime().ToString('o')
    DurationSeconds = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 3)
    WarningCount  = $script:WarningCount
    ErrorCount    = $script:ErrorCount

    HealthSummary = $summary
    Identity      = $identity
    OperatingSystem = $operatingSystem
    Performance   = $performance
    Storage       = $storage
    DeviceManager = $deviceManager
    Security      = [ordered]@{
        Defender      = $defender
        Firewall      = $firewall
        BitLocker     = $bitLocker
        TPM           = if ($tpmSecureBoot) { $tpmSecureBoot.TPM } else { $null }
        SecureBoot    = if ($tpmSecureBoot) { $tpmSecureBoot.SecureBoot } else { $null }
    }
    Network       = $network
    Updates       = $updates
    EventLogs     = $eventLogs
    Services      = $services
    ManagementAgents = $agents
    Applications   = [ordered]@{
        MicrosoftEdge = $edge
    }
    Collectors    = @($script:CollectorResults)
}

try {
    Write-Telemetry -Event $event
    Write-Log -Message ("Endpoint snapshot completed. Health={0}; ExitCode={1}; CollectorsFailed={2}" -f
        $summary.Status, $exitCode, $failedCollectors)
} catch {
    Write-Log -Level 'ERROR' -Message ("Unable to write endpoint telemetry: {0}" -f $_.Exception.Message)
    exit 4
}

exit $exitCode
