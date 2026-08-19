#requires -Version 5.1
#requires -RunAsAdministrator
<#
.SYNOPSIS
    Creates one authoritative weekly endpoint health snapshot for Elastic.

.DESCRIPTION
    Collects hardware, operating system, performance, storage, security,
    networking, update, event-log, service, and management-agent health.

    Writes one endpoint.health summary event and one endpoint.health.finding
    event for every Warning or Critical finding. The individual finding events
    support clean Kibana tables, filtering, drilldowns, and alerting.

    The script is fault-tolerant: individual collectors can fail without
    preventing the remaining endpoint snapshot from being written.

.OUTPUTS
    C:\Logs\Maintenance-Telemetry.ndjson
    C:\Logs\14_Endpoint_Health_Inventory.latest.json
    C:\Logs\14_Endpoint_Health_Inventory.log

.NOTES
    ScriptName:    14_Endpoint_Health_Inventory.ps1
    ScriptVersion: 1.2.8
    LastUpdated:   2026-08-19
    Changes:       v1.2.8 makes Secure Boot detection resilient by using Confirm-SecureBootUEFI first,
                   then falling back to the Windows SecureBoot registry state when the cmdlet returns Access Denied.
                   Telemetry now records the detection method and diagnostic error without creating a false warning
                   when the registry can determine the actual Secure Boot state.
                   v1.2.7 fixes storage manufacturer normalization so generic Windows values such as
                   '(Standard disk drives)' are treated as unknown and vendor is inferred from model/PNP data.
                   v1.2.6 enriches physical storage inventory with explicit manufacturer/brand, model,
                   serial number, firmware, interface/bus, device identifiers, partition style, boot/system flags,
                   and Windows PNP metadata to improve lab inventory and lifecycle reporting.
                   v1.2.5 adds detailed hardware inventory for CPU, GPU, memory modules, and system storage,
                   derives Building/Lab/DeviceIdentifier from the computer name for dashboard controls,
                   and places those location fields on both endpoint.health and endpoint.health.finding events.
                   v1.2.4 refines Device Manager health so benign ProblemCode=0 degraded devices are inventoried but not counted as actionable problems,
                   and returns exit code 0 for SuccessWithWarnings/SuccessWithCriticalFindings so scheduler execution is not falsely marked failed.
                   v1.2.3 stabilizes Elastic telemetry mappings for findings and collections,
                   renames pending-reboot reason telemetry, adds explicit critical-service
                   expectations, and adds a concise Elastic endpoint-health summary.
                   v1.2.2 uses Maintenance.Framework v2.4 staged text logging.
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

    [ValidateRange(0.1, 1000)]
    [double]$CapabilityAccessManagerWarningGB = 1,

    [ValidateRange(0.1, 1000)]
    [double]$CapabilityAccessManagerCriticalGB = 10,

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
$script:ScriptVersion    = '1.2.8'
$script:RunId            = [guid]::NewGuid().Guid
$script:StartTime        = Get-Date
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:Findings         = New-Object System.Collections.Generic.List[object]
$script:CollectorResults = New-Object System.Collections.Generic.List[object]

# Explicit health expectations for the default critical services.
# Any custom service supplied through -CriticalServices defaults to RequiredRunning.
$script:CriticalServiceExpectations = @{
    'WinDefend'          = 'RequiredRunning'
    'mpssvc'             = 'RequiredRunning'
    'W32Time'            = 'RequiredRunning'
    'EventLog'           = 'RequiredRunning'
    'Schedule'           = 'RequiredRunning'
    'RpcSs'              = 'RequiredRunning'
    'Dnscache'           = 'RequiredRunning'
    'LanmanWorkstation'  = 'RequiredRunning'
}

$script:NdjsonPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestPath = Join-Path $LogDirectory '14_Endpoint_Health_Inventory.latest.json'
$script:LogPath          = $null
$script:PublishedLogPath = $null
$script:LogSession       = $null

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 14 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName $script:ScriptName `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '14_Endpoint_Health_Inventory.log',
        '*-14_Endpoint_Health_Inventory-*.log'
    ) | Out-Null

$script:LogSession = New-MaintenanceStagedLog `
    -ScriptName $script:ScriptName `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $env:COMPUTERNAME `
    -Timestamp $script:StartTime

$script:LogPath = [string]$script:LogSession.WorkingPath
$script:PublishedLogPath = [string]$script:LogSession.PublishedPath

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
    $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
    $normalizedLevel = if ($Level -eq 'WARN') { 'WARNING' } else { $Level }
    $line = '{0} [{1}] [{2}] {3}' -f $timestamp, $computerName, $normalizedLevel, $Message

    try {
        $activeLogDirectory = Split-Path -Parent $script:LogPath
        Ensure-Directory -Path $activeLogDirectory
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


function New-StringArrayForJson {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    [string[]]$items = @(
        $InputObject |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function New-ObjectArrayForJson {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    [object[]]$items = @(
        $InputObject |
        Where-Object { $null -ne $_ } |
        ForEach-Object { $_ }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Convert-FindingValueForTelemetry {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $result = [ordered]@{
        ValueType    = 'Null'
        ValueNumber  = $null
        ValueBoolean = $null
        ValueText    = $null
        ValueTexts   = [object[]]@()
        ValueJson    = $null
    }

    if ($null -eq $Value) {
        return [pscustomobject]$result
    }

    if ($Value -is [bool]) {
        $result.ValueType = 'Boolean'
        $result.ValueBoolean = [bool]$Value
        return [pscustomobject]$result
    }

    if (
        $Value -is [byte] -or $Value -is [sbyte] -or
        $Value -is [int16] -or $Value -is [uint16] -or
        $Value -is [int32] -or $Value -is [uint32] -or
        $Value -is [int64] -or $Value -is [uint64] -or
        $Value -is [single] -or $Value -is [double] -or
        $Value -is [decimal]
    ) {
        $result.ValueType = 'Number'
        $result.ValueNumber = [double]$Value
        return [pscustomobject]$result
    }

    if ($Value -is [string] -or $Value -is [char]) {
        $result.ValueType = 'Text'
        $result.ValueText = [string]$Value
        return [pscustomobject]$result
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @($Value)
        $allScalarText = $true
        foreach ($item in $items) {
            if ($null -eq $item) { continue }
            if (
                $item -is [System.Collections.IDictionary] -or
                $item -is [pscustomobject] -or
                ($item -is [System.Collections.IEnumerable] -and -not ($item -is [string]))
            ) {
                $allScalarText = $false
                break
            }
        }

        if ($allScalarText) {
            $result.ValueType = 'TextArray'
            $result.ValueTexts = New-StringArrayForJson -InputObject $items
            return [pscustomobject]$result
        }
    }

    $result.ValueType = 'Json'
    try {
        $result.ValueJson = $Value | ConvertTo-Json -Depth 8 -Compress
    }
    catch {
        $result.ValueJson = [string]$Value
    }

    return [pscustomobject]$result
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

    $typedValue = Convert-FindingValueForTelemetry -Value $Value

    $script:Findings.Add([pscustomobject]@{
        Severity     = $Severity
        Category     = $Category
        Check        = $Check
        Message      = $Message
        ValueType    = $typedValue.ValueType
        ValueNumber  = $typedValue.ValueNumber
        ValueBoolean = $typedValue.ValueBoolean
        ValueText    = $typedValue.ValueText
        ValueTexts   = $typedValue.ValueTexts
        ValueJson    = $typedValue.ValueJson
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

function Get-ObjectPropertyValueSafe {
    param(
        $InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }

    return $property.Value
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
        ReasonNames = New-StringArrayForJson -InputObject @($reasons | Select-Object -Unique)
    }
}


function Get-ComputerLocationIdentity {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    $building = $null
    $lab = $null
    $deviceIdentifier = $null
    $derived = $false

    # Standard lab names such as:
    #   IB1-103-06
    #   IB1-103-AV
    #   SSB-122-01
    #
    # The first two hyphen-delimited components define the lab.
    if ($ComputerName -match '^(?<Building>[^-]+)-(?<Room>[^-]+)(?:-(?<Device>.+))?$') {
        $building = [string]$Matches.Building
        $lab = '{0}-{1}' -f $Matches.Building, $Matches.Room
        if ($Matches.ContainsKey('Device') -and -not [string]::IsNullOrWhiteSpace($Matches.Device)) {
            $deviceIdentifier = [string]$Matches.Device
        }
        $derived = $true
    }

    [pscustomobject]@{
        Building         = $building
        Lab              = $lab
        DeviceIdentifier = $deviceIdentifier
        DerivedFromName  = $derived
    }
}

function Convert-ProcessorArchitecture {
    [CmdletBinding()]
    param([AllowNull()]$Architecture)

    switch ([int]$Architecture) {
        0  { 'x86' }
        1  { 'MIPS' }
        2  { 'Alpha' }
        3  { 'PowerPC' }
        5  { 'ARM' }
        6  { 'Itanium' }
        9  { 'x64' }
        12 { 'ARM64' }
        default { if ($null -eq $Architecture) { $null } else { [string]$Architecture } }
    }
}

function Convert-SmbiosMemoryType {
    [CmdletBinding()]
    param([AllowNull()]$MemoryType)

    if ($null -eq $MemoryType) { return $null }

    switch ([int]$MemoryType) {
        18 { 'DDR' }
        19 { 'DDR2' }
        20 { 'DDR2 FB-DIMM' }
        24 { 'DDR3' }
        26 { 'DDR4' }
        27 { 'LPDDR' }
        28 { 'LPDDR2' }
        29 { 'LPDDR3' }
        30 { 'LPDDR4' }
        34 { 'DDR5' }
        35 { 'LPDDR5' }
        default { 'Type{0}' -f [int]$MemoryType }
    }
}

function Get-HardwareInventory {
    [CmdletBinding()]
    param()

    # ----- CPU -----
    $processorRows = @()
    $processors = @(Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue)

    foreach ($processor in $processors) {
        $processorRows += [pscustomobject]@{
            Manufacturer                  = ([string]$processor.Manufacturer).Trim()
            Model                         = ([string]$processor.Name).Trim()
            SocketDesignation             = $processor.SocketDesignation
            ProcessorId                   = $processor.ProcessorId
            Architecture                  = Convert-ProcessorArchitecture -Architecture $processor.Architecture
            AddressWidthBits              = $processor.AddressWidth
            Cores                         = $processor.NumberOfCores
            LogicalProcessors             = $processor.NumberOfLogicalProcessors
            MaxClockMHz                   = $processor.MaxClockSpeed
            CurrentClockMHz               = $processor.CurrentClockSpeed
            L2CacheKB                     = $processor.L2CacheSize
            L3CacheKB                     = $processor.L3CacheSize
            VirtualizationFirmwareEnabled = [bool]$processor.VirtualizationFirmwareEnabled
            VmMonitorModeExtensions       = [bool]$processor.VMMonitorModeExtensions
            SecondLevelAddressTranslation = [bool]$processor.SecondLevelAddressTranslationExtensions
        }
    }

    $totalCores = if ($processorRows.Count -gt 0) {
        [int](($processorRows | Measure-Object -Property Cores -Sum).Sum)
    } else { 0 }

    $totalLogical = if ($processorRows.Count -gt 0) {
        [int](($processorRows | Measure-Object -Property LogicalProcessors -Sum).Sum)
    } else { 0 }

    $cpu = [pscustomobject]@{
        SocketCount        = $processorRows.Count
        TotalCores         = $totalCores
        TotalLogicalProcessors = $totalLogical
        PrimaryManufacturer = if ($processorRows.Count -gt 0) { $processorRows[0].Manufacturer } else { $null }
        PrimaryModel        = if ($processorRows.Count -gt 0) { $processorRows[0].Model } else { $null }
        Processors          = New-ObjectArrayForJson -InputObject $processorRows
    }

    # ----- Physical memory -----
    $memoryRows = @()
    $physicalMemory = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue)
    foreach ($module in $physicalMemory) {
        $smbiosType = Get-ObjectPropertyValueSafe -InputObject $module -Name 'SMBIOSMemoryType'

        $memoryRows += [pscustomobject]@{
            DeviceLocator       = $module.DeviceLocator
            BankLabel           = $module.BankLabel
            CapacityGB          = if ($module.Capacity) { [math]::Round([double]$module.Capacity / 1GB, 2) } else { $null }
            Manufacturer        = ([string]$module.Manufacturer).Trim()
            PartNumber          = ([string]$module.PartNumber).Trim()
            SerialNumber        = ([string]$module.SerialNumber).Trim()
            SpeedMHz            = $module.Speed
            ConfiguredSpeedMHz  = Get-ObjectPropertyValueSafe -InputObject $module -Name 'ConfiguredClockSpeed'
            MemoryType          = Convert-SmbiosMemoryType -MemoryType $smbiosType
            SmbiosMemoryType    = $smbiosType
            FormFactor          = $module.FormFactor
            DataWidthBits       = $module.DataWidth
            TotalWidthBits      = $module.TotalWidth
            ConfiguredVoltageMv = Get-ObjectPropertyValueSafe -InputObject $module -Name 'ConfiguredVoltage'
        }
    }

    $memoryArrays = @(Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue)
    $slotCount = 0
    foreach ($array in $memoryArrays) {
        if ($null -ne $array.MemoryDevices) { $slotCount += [int]$array.MemoryDevices }
    }

    $totalInstalledMemoryGB = if ($memoryRows.Count -gt 0) {
        [math]::Round((($memoryRows | Measure-Object -Property CapacityGB -Sum).Sum), 2)
    } else { $null }

    $memory = [pscustomobject]@{
        TotalInstalledGB = $totalInstalledMemoryGB
        SlotCount        = $slotCount
        PopulatedSlots   = $memoryRows.Count
        EmptySlots       = if ($slotCount -ge $memoryRows.Count) { $slotCount - $memoryRows.Count } else { 0 }
        Modules          = New-ObjectArrayForJson -InputObject $memoryRows
    }

    # ----- GPU / display adapters -----
    $gpuRows = @()
    $videoControllers = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue)
    foreach ($gpu in $videoControllers) {
        $driverDate = $null
        if ($gpu.DriverDate) {
            try { $driverDate = (Convert-CimDate $gpu.DriverDate).ToUniversalTime().ToString('o') } catch { }
        }

        $adapterRamBytes = $null
        $adapterRamGB = $null
        if ($null -ne $gpu.AdapterRAM) {
            try {
                $adapterRamBytes = [uint64]$gpu.AdapterRAM
                if ($adapterRamBytes -gt 0) {
                    $adapterRamGB = [math]::Round([double]$adapterRamBytes / 1GB, 2)
                }
            } catch { }
        }

        $gpuRows += [pscustomobject]@{
            Manufacturer             = $gpu.AdapterCompatibility
            Model                    = $gpu.Name
            VideoProcessor           = $gpu.VideoProcessor
            PnpDeviceId              = $gpu.PNPDeviceID
            Status                   = $gpu.Status
            DriverVersion            = $gpu.DriverVersion
            DriverDate               = $driverDate
            AdapterRamBytes          = $adapterRamBytes
            AdapterRamGB             = $adapterRamGB
            CurrentHorizontalPixels  = $gpu.CurrentHorizontalResolution
            CurrentVerticalPixels    = $gpu.CurrentVerticalResolution
            CurrentRefreshRateHz     = $gpu.CurrentRefreshRate
            CurrentBitsPerPixel      = $gpu.CurrentBitsPerPixel
            VideoModeDescription     = $gpu.VideoModeDescription
            InstalledDisplayDrivers  = $gpu.InstalledDisplayDrivers
        }
    }

    $gpu = [pscustomobject]@{
        AdapterCount = $gpuRows.Count
        PrimaryManufacturer = if ($gpuRows.Count -gt 0) { $gpuRows[0].Manufacturer } else { $null }
        PrimaryModel        = if ($gpuRows.Count -gt 0) { $gpuRows[0].Model } else { $null }
        Adapters            = New-ObjectArrayForJson -InputObject $gpuRows
    }

    [pscustomobject]@{
        CPU    = $cpu
        Memory = $memory
        GPU    = $gpu
    }
}

function Get-ComputerIdentity {
    $computer = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_BIOS
    $baseboard = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
    $enclosure = Get-CimInstance Win32_SystemEnclosure -ErrorAction SilentlyContinue
    $location = Get-ComputerLocationIdentity -ComputerName $env:COMPUTERNAME

    [pscustomobject]@{
        ComputerName  = $env:COMPUTERNAME
        Building      = $location.Building
        Lab           = $location.Lab
        DeviceIdentifier = $location.DeviceIdentifier
        LocationDerivedFromName = $location.DerivedFromName
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
        ChassisTypes  = New-ObjectArrayForJson -InputObject $(if ($enclosure) { @($enclosure.ChassisTypes) } else { @() })
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
            -Check 'PendingReboot' -Message ('Pending reboot: {0}' -f ($pending.ReasonNames -join ', ')) `
            -Value $pending.ReasonNames
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
        PendingRebootReasonNames = $pending.ReasonNames
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
        CpuSamples          = New-ObjectArrayForJson -InputObject $cpuSamples
        MemoryTotalGB       = $totalMemoryGB
        MemoryUsedGB        = $usedMemoryGB
        MemoryFreeGB        = $freeMemoryGB
        MemoryUsedPercent   = $memoryPercent
    }
}


function Get-StorageManufacturerFromModel {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Manufacturer,
        [AllowNull()][string]$Model,
        [AllowNull()][string]$FriendlyName,
        [AllowNull()][string]$PnpDeviceId
    )

    if (-not [string]::IsNullOrWhiteSpace($Manufacturer)) {
        $clean = $Manufacturer.Trim()
        $normalized = ($clean -replace '[()]','').Trim()

        if ($normalized -notmatch '^(Standard disk drives|Standard disk drive|Unknown|Generic)$') {
            return $clean
        }
    }

    $source = ('{0} {1} {2}' -f $Model, $FriendlyName, $PnpDeviceId).Trim()

    $patterns = [ordered]@{
        'Samsung'         = 'SAMSUNG|MZ[VNLQ]|PM9|SM9'
        'SK hynix'        = 'SK HYNIX|SKHYNIX|HFS|HFM'
        'Western Digital' = 'WESTERN DIGITAL|WDC|WD_BLACK|WD BLUE|WD GREEN|SANDISK'
        'Micron'          = 'MICRON|MTFD'
        'KIOXIA'          = 'KIOXIA|TOSHIBA'
        'Intel'           = 'INTEL'
        'Kingston'        = 'KINGSTON'
        'Crucial'         = 'CRUCIAL'
        'Seagate'         = 'SEAGATE|ST[0-9]'
        'Solidigm'        = 'SOLIDIGM'
    }

    foreach ($entry in $patterns.GetEnumerator()) {
        if ($source -match $entry.Value) {
            return $entry.Key
        }
    }

    return $Manufacturer
}

function Get-DiskDriveMetadata {
    [CmdletBinding()]
    param()

    $rows = @()
    foreach ($drive in @(Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue)) {
        $rows += [pscustomobject]@{
            Index            = $drive.Index
            DeviceId         = $drive.DeviceID
            PnpDeviceId      = $drive.PNPDeviceID
            Manufacturer     = $drive.Manufacturer
            Model            = $drive.Model
            SerialNumber     = ([string]$drive.SerialNumber).Trim()
            FirmwareRevision = $drive.FirmwareRevision
            InterfaceType    = $drive.InterfaceType
            MediaType        = $drive.MediaType
            SizeBytes        = $drive.Size
            SizeGB           = if ($drive.Size) { [math]::Round([double]$drive.Size / 1GB, 2) } else { $null }
            Partitions       = $drive.Partitions
            Status           = $drive.Status
        }
    }

    return ,([object[]]$rows)
}

function Get-DiskHealth {
    $volumes = @()
    $diskDriveMetadata = @(Get-DiskDriveMetadata)
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

            $matchingWmi = $null
            if ($pd.SerialNumber) {
                $matchingWmi = $diskDriveMetadata | Where-Object {
                    $_.SerialNumber -and
                    ([string]$_.SerialNumber).Trim() -eq ([string]$pd.SerialNumber).Trim()
                } | Select-Object -First 1
            }

            if (-not $matchingWmi) {
                $matchingWmi = $diskDriveMetadata | Where-Object {
                    $_.Model -and $pd.FriendlyName -and
                    (
                        ([string]$_.Model).Trim() -eq ([string]$pd.FriendlyName).Trim() -or
                        ([string]$pd.FriendlyName).Trim() -like ('*{0}*' -f ([string]$_.Model).Trim())
                    )
                } | Select-Object -First 1
            }

            $pdManufacturer = Get-ObjectPropertyValueSafe -InputObject $pd -Name 'Manufacturer'
            $pdModel = Get-ObjectPropertyValueSafe -InputObject $pd -Name 'Model'
            if ([string]::IsNullOrWhiteSpace([string]$pdModel)) { $pdModel = $pd.FriendlyName }

            $manufacturer = Get-StorageManufacturerFromModel `
                -Manufacturer $(if ($pdManufacturer) { [string]$pdManufacturer } elseif ($matchingWmi) { [string]$matchingWmi.Manufacturer } else { $null }) `
                -Model ([string]$pdModel) `
                -FriendlyName ([string]$pd.FriendlyName) `
                -PnpDeviceId $(if ($matchingWmi) { [string]$matchingWmi.PnpDeviceId } else { $null })

            $diskNumber = Get-ObjectPropertyValueSafe -InputObject $pd -Name 'DeviceId'
            $diskObject = $null
            if ($null -ne $diskNumber -and (Get-Command Get-Disk -ErrorAction SilentlyContinue)) {
                try { $diskObject = Get-Disk -Number ([int]$diskNumber) -ErrorAction Stop } catch { }
            }

            $physical += [pscustomobject]@{
                DiskNumber             = $diskNumber
                DeviceId               = if ($matchingWmi) { $matchingWmi.DeviceId } else { $null }
                PnpDeviceId            = if ($matchingWmi) { $matchingWmi.PnpDeviceId } else { $null }
                FriendlyName           = $pd.FriendlyName
                Manufacturer           = $manufacturer
                RawManufacturer        = if ($pdManufacturer) { [string]$pdManufacturer } elseif ($matchingWmi) { [string]$matchingWmi.Manufacturer } else { $null }
                Model                  = [string]$pdModel
                SerialNumber           = ([string]$pd.SerialNumber).Trim()
                FirmwareVersion        = $pd.FirmwareVersion
                FirmwareRevisionWmi    = if ($matchingWmi) { $matchingWmi.FirmwareRevision } else { $null }
                MediaType              = [string]$pd.MediaType
                BusType                = [string]$pd.BusType
                InterfaceType          = if ($matchingWmi) { $matchingWmi.InterfaceType } else { $null }
                SizeGB                 = [math]::Round([double]$pd.Size / 1GB, 2)
                HealthStatus           = [string]$pd.HealthStatus
                OperationalStatus      = New-StringArrayForJson -InputObject @($pd.OperationalStatus | ForEach-Object { [string]$_ })
                PartitionStyle         = if ($diskObject) { [string]$diskObject.PartitionStyle } else { $null }
                IsBoot                 = if ($diskObject) { [bool]$diskObject.IsBoot } else { $null }
                IsSystem               = if ($diskObject) { [bool]$diskObject.IsSystem } else { $null }
                IsOffline              = if ($diskObject) { [bool]$diskObject.IsOffline } else { $null }
                IsReadOnly             = if ($diskObject) { [bool]$diskObject.IsReadOnly } else { $null }
                NumberOfPartitions     = if ($matchingWmi) { $matchingWmi.Partitions } elseif ($diskObject) { $diskObject.NumberOfPartitions } else { $null }
                Wear                   = if ($reliability) { $reliability.Wear } else { $null }
                TemperatureC           = if ($reliability) { $reliability.Temperature } else { $null }
                ReadErrorsTotal        = if ($reliability) { $reliability.ReadErrorsTotal } else { $null }
                WriteErrorsTotal       = if ($reliability) { $reliability.WriteErrorsTotal } else { $null }
                PowerOnHours           = if ($reliability) { $reliability.PowerOnHours } else { $null }
            }
        }
    } else {
        foreach ($drive in Get-CimInstance Win32_DiskDrive) {
            $manufacturer = Get-StorageManufacturerFromModel `
                -Manufacturer ([string]$drive.Manufacturer) `
                -Model ([string]$drive.Model) `
                -FriendlyName ([string]$drive.Model) `
                -PnpDeviceId ([string]$drive.PNPDeviceID)

            $diskObject = $null
            if ((Get-Command Get-Disk -ErrorAction SilentlyContinue) -and $null -ne $drive.Index) {
                try { $diskObject = Get-Disk -Number ([int]$drive.Index) -ErrorAction Stop } catch { }
            }

            $physical += [pscustomobject]@{
                DiskNumber          = $drive.Index
                DeviceId            = $drive.DeviceID
                PnpDeviceId         = $drive.PNPDeviceID
                FriendlyName        = $drive.Model
                Manufacturer        = $manufacturer
                RawManufacturer     = $drive.Manufacturer
                Model               = $drive.Model
                SerialNumber        = ([string]$drive.SerialNumber).Trim()
                FirmwareVersion     = $drive.FirmwareRevision
                FirmwareRevisionWmi = $drive.FirmwareRevision
                MediaType           = $drive.MediaType
                BusType             = if ($diskObject) { [string]$diskObject.BusType } else { $drive.InterfaceType }
                InterfaceType       = $drive.InterfaceType
                SizeGB              = [math]::Round([double]$drive.Size / 1GB, 2)
                HealthStatus        = if ($drive.Status -eq 'OK') { 'Healthy' } else { $drive.Status }
                OperationalStatus   = New-StringArrayForJson -InputObject @($drive.Status)
                PartitionStyle      = if ($diskObject) { [string]$diskObject.PartitionStyle } else { $null }
                IsBoot              = if ($diskObject) { [bool]$diskObject.IsBoot } else { $null }
                IsSystem            = if ($diskObject) { [bool]$diskObject.IsSystem } else { $null }
                IsOffline           = if ($diskObject) { [bool]$diskObject.IsOffline } else { $null }
                IsReadOnly          = if ($diskObject) { [bool]$diskObject.IsReadOnly } else { $null }
                NumberOfPartitions  = $drive.Partitions
                Wear                = $null
                TemperatureC        = $null
                ReadErrorsTotal     = $null
                WriteErrorsTotal    = $null
                PowerOnHours        = $null
            }
        }
    }

    $camWal = Get-CapabilityAccessManagerHealth

    [pscustomobject]@{
        Volumes                 = New-ObjectArrayForJson -InputObject $volumes
        PhysicalDisks           = New-ObjectArrayForJson -InputObject $physical
        PhysicalDiskCount       = $physical.Count
        PrimaryDiskManufacturer = if ($physical.Count -gt 0) { $physical[0].Manufacturer } else { $null }
        PrimaryDiskModel        = if ($physical.Count -gt 0) { $physical[0].Model } else { $null }
        PrimaryDiskBusType      = if ($physical.Count -gt 0) { $physical[0].BusType } else { $null }
        PrimaryDiskSizeGB       = if ($physical.Count -gt 0) { $physical[0].SizeGB } else { $null }
        CapabilityAccessManager = $camWal
    }
}


function Get-CapabilityAccessManagerHealth {
    $camDirectory = Join-Path $env:ProgramData 'Microsoft\Windows\CapabilityAccessManager'
    $camWalPath = Join-Path $camDirectory 'CapabilityAccessManager.db-wal'

    $result = [ordered]@{
        Path                 = $camWalPath
        Exists               = $false
        SizeBytes            = 0
        SizeMB               = 0
        SizeGB               = 0
        LastWriteTime        = $null
        SystemDrive          = $env:SystemDrive
        SystemDriveSizeGB    = $null
        PercentOfSystemDrive = 0
        WarningThresholdGB   = $CapabilityAccessManagerWarningGB
        CriticalThresholdGB  = $CapabilityAccessManagerCriticalGB
        Status               = 'Healthy'
        Error                = $null
    }

    try {
        $systemDisk = Get-CimInstance Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $env:SystemDrive) -ErrorAction SilentlyContinue
        if ($systemDisk -and $systemDisk.Size) {
            $result.SystemDriveSizeGB = [math]::Round([double]$systemDisk.Size / 1GB, 2)
        }

        if (Test-Path -LiteralPath $camWalPath -PathType Leaf) {
            $file = Get-Item -LiteralPath $camWalPath -Force -ErrorAction Stop
            $sizeBytes = [double]$file.Length
            $sizeGB = $sizeBytes / 1GB

            $result.Exists        = $true
            $result.SizeBytes     = [int64]$file.Length
            $result.SizeMB        = [math]::Round($sizeBytes / 1MB, 2)
            $result.SizeGB        = [math]::Round($sizeGB, 3)
            $result.LastWriteTime = $file.LastWriteTimeUtc.ToString('o')

            if ($systemDisk -and $systemDisk.Size -gt 0) {
                $result.PercentOfSystemDrive = [math]::Round(($sizeBytes / [double]$systemDisk.Size) * 100, 3)
            }

            if ($sizeGB -ge $CapabilityAccessManagerCriticalGB) {
                $result.Status = 'Critical'
            } elseif ($sizeGB -ge $CapabilityAccessManagerWarningGB) {
                $result.Status = 'Warning'
            }
        }

        $message = if (-not $result.Exists) {
            'CapabilityAccessManager.db-wal was not present.'
        } else {
            'CapabilityAccessManager.db-wal size is {0} GB ({1}% of {2}).' -f `
                $result.SizeGB, $result.PercentOfSystemDrive, $result.SystemDrive
        }

        Add-Finding -Severity $result.Status -Category 'Storage' `
            -Check 'CapabilityAccessManagerDbWal' -Message $message -Value $result.SizeGB

        Write-Log -Level 'INFO' -Message ("CAM WAL check: status={0}; exists={1}; sizeGB={2}; percentOfSystemDrive={3}; path={4}" -f `
            $result.Status, $result.Exists, $result.SizeGB, $result.PercentOfSystemDrive, $camWalPath)
    } catch {
        $result.Status = 'Warning'
        $result.Error = $_.Exception.Message
        Add-Finding -Severity 'Warning' -Category 'Storage' `
            -Check 'CapabilityAccessManagerDbWal' `
            -Message ("CapabilityAccessManager.db-wal could not be inspected: {0}" -f $_.Exception.Message)
        Write-Log -Level 'WARN' -Message ("CAM WAL check failed: {0}" -f $_.Exception.Message)
    }

    [pscustomobject]$result
}

function Get-DeviceManagerHealth {
    $devices = @()
    $informationalDegraded = @()

    if (Get-Command Get-PnpDevice -ErrorAction SilentlyContinue) {
        $presentDevices = @(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue)

        foreach ($device in $presentDevices) {
            $problemCode = $null
            try { $problemCode = [int]$device.Problem } catch { $problemCode = $null }

            $entry = [pscustomobject]@{
                Class        = $device.Class
                FriendlyName = $device.FriendlyName
                InstanceId   = $device.InstanceId
                Status       = [string]$device.Status
                ProblemCode  = $problemCode
            }

            if ($null -ne $problemCode -and $problemCode -ne 0) {
                $devices += $entry
            }
            elseif ($device.Status -ne 'OK') {
                $informationalDegraded += $entry
            }
        }
    }
    else {
        $presentDevices = @(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue)

        foreach ($device in $presentDevices) {
            $problemCode = $device.ConfigManagerErrorCode
            $entry = [pscustomobject]@{
                Class        = $device.PNPClass
                FriendlyName = $device.Name
                InstanceId   = $device.PNPDeviceID
                Status       = $device.Status
                ProblemCode  = $problemCode
            }

            if ($null -ne $problemCode -and [int]$problemCode -ne 0) {
                $devices += $entry
            }
            elseif ($device.Status -and $device.Status -ne 'OK') {
                $informationalDegraded += $entry
            }
        }
    }

    if ($devices.Count -gt 0) {
        Add-Finding -Severity 'Warning' -Category 'Hardware' -Check 'DeviceManagerProblems' `
            -Message ("{0} actionable Device Manager problem(s) detected." -f $devices.Count) `
            -Value $devices.Count
    }
    else {
        Add-Finding -Severity 'Healthy' -Category 'Hardware' -Check 'DeviceManagerProblems' `
            -Message 'No actionable Device Manager problems were detected.' -Value 0
    }

    if ($informationalDegraded.Count -gt 0) {
        Add-Finding -Severity 'Info' -Category 'Hardware' -Check 'DeviceManagerDegradedNonActionable' `
            -Message ("{0} degraded device(s) reported ProblemCode=0 and were not counted as actionable problems." -f $informationalDegraded.Count) `
            -Value $informationalDegraded.Count
    }

    [pscustomobject]@{
        ProblemCount                 = $devices.Count
        Problems                     = New-ObjectArrayForJson -InputObject $devices
        DegradedNonActionableCount   = $informationalDegraded.Count
        DegradedNonActionableDevices = New-ObjectArrayForJson -InputObject $informationalDegraded
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
        RecentThreatDetections            = New-ObjectArrayForJson -InputObject $threats
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
        Profiles             = New-ObjectArrayForJson -InputObject $profiles
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
            KeyProtectorTypes   = New-StringArrayForJson -InputObject @($volume.KeyProtector | ForEach-Object { [string]$_.KeyProtectorType })
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
        Volumes   = New-ObjectArrayForJson -InputObject $volumes
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
        Supported          = $null
        Enabled            = $null
        DetectionMethod    = $null
        RegistryValue      = $null
        ConfirmCmdletError = $null
        Error              = $null
    }

    # Preferred method: Microsoft's SecureBoot cmdlet. On some Windows 11 systems
    # this can return "Unable to set proper privileges. Access was denied." even
    # when the caller is elevated. If that occurs, fall back to the Windows
    # SecureBoot state maintained under HKLM.
    if (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue) {
        try {
            $secureBoot.Enabled = [bool](Confirm-SecureBootUEFI -ErrorAction Stop)
            $secureBoot.Supported = $true
            $secureBoot.DetectionMethod = 'Confirm-SecureBootUEFI'
        }
        catch {
            $secureBoot.ConfirmCmdletError = $_.Exception.Message

            if ($_.Exception.Message -match 'not supported|platform does not support|unsupported') {
                $secureBoot.Supported = $false
                $secureBoot.DetectionMethod = 'Confirm-SecureBootUEFI'
            }
        }
    }

    # Registry fallback. UEFISecureBootEnabled is maintained by Windows and is
    # readable without needing the firmware-variable privilege required by the
    # SecureBoot cmdlet.
    if ($null -eq $secureBoot.Enabled -and $secureBoot.Supported -ne $false) {
        $secureBootStatePath = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State'

        try {
            if (Test-Path -LiteralPath $secureBootStatePath) {
                $state = Get-ItemProperty -LiteralPath $secureBootStatePath `
                    -Name 'UEFISecureBootEnabled' -ErrorAction Stop

                $rawState = $state.UEFISecureBootEnabled
                if ($null -ne $rawState) {
                    $secureBoot.RegistryValue = [int]$rawState
                    $secureBoot.Enabled = ([int]$rawState -eq 1)
                    $secureBoot.Supported = $true
                    $secureBoot.DetectionMethod = 'Registry'
                }
            }
        }
        catch {
            $secureBoot.Error = $_.Exception.Message
        }
    }

    if ($secureBoot.Supported -eq $true -and $null -ne $secureBoot.Enabled) {
        $severity = if ($secureBoot.Enabled) { 'Healthy' } else { 'Warning' }

        $methodText = if ($secureBoot.DetectionMethod) {
            '; method={0}' -f $secureBoot.DetectionMethod
        } else {
            ''
        }

        Add-Finding -Severity $severity -Category 'Security' -Check 'SecureBoot' `
            -Message ("Secure Boot enabled={0}{1}." -f $secureBoot.Enabled, $methodText)
    }
    elseif ($secureBoot.Supported -eq $false) {
        Add-Finding -Severity 'Info' -Category 'Security' -Check 'SecureBoot' `
            -Message 'Secure Boot is not supported on this system.'
    }
    else {
        $diagnostic = @(
            $secureBoot.ConfirmCmdletError
            $secureBoot.Error
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        $detail = if ($diagnostic.Count -gt 0) {
            ' Diagnostic: {0}' -f (($diagnostic | Select-Object -Unique) -join ' | ')
        } else {
            ''
        }

        Add-Finding -Severity 'Warning' -Category 'Security' -Check 'SecureBoot' `
            -Message ("Secure Boot state could not be determined.{0}" -f $detail)
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
                IPv4Addresses  = New-StringArrayForJson -InputObject @($config.IPv4Address | ForEach-Object { $_.IPAddress })
                IPv6Addresses  = New-StringArrayForJson -InputObject @($config.IPv6Address | ForEach-Object { $_.IPAddress })
                IPv4Gateway    = New-StringArrayForJson -InputObject @($config.IPv4DefaultGateway | ForEach-Object {
                    Get-ObjectPropertyValueSafe -InputObject $_ -Name 'NextHop'
                } | Where-Object { $null -ne $_ })
                IPv6Gateway    = New-StringArrayForJson -InputObject @($config.IPv6DefaultGateway | ForEach-Object {
                    Get-ObjectPropertyValueSafe -InputObject $_ -Name 'NextHop'
                } | Where-Object { $null -ne $_ })
                DNSServers     = New-StringArrayForJson -InputObject $dnsServers
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
        Adapters           = New-ObjectArrayForJson -InputObject $adapters
        IPConfigurations   = New-ObjectArrayForJson -InputObject $ipConfigs
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
        RecentHotfixes    = New-ObjectArrayForJson -InputObject $recent
        UpdateHistory     = New-ObjectArrayForJson -InputObject $history
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
        ApplicationCrashes     = New-ObjectArrayForJson -InputObject $applicationCrashes
        CriticalEventCount     = $criticalEvents.Count
        CriticalEvents         = New-ObjectArrayForJson -InputObject $criticalEvents
    }
}

function Get-ServiceHealth {
    $results = @()

    foreach ($serviceName in $CriticalServices) {
        $service = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $serviceName.Replace("'","''")) `
            -ErrorAction SilentlyContinue

        $expectation = if ($script:CriticalServiceExpectations.ContainsKey($serviceName)) {
            [string]$script:CriticalServiceExpectations[$serviceName]
        }
        else {
            'RequiredRunning'
        }

        if (-not $service) {
            $healthy = ($expectation -eq 'Optional')
            $results += [pscustomobject]@{
                Name          = $serviceName
                DisplayName   = $null
                Exists        = $false
                State         = 'NotFound'
                StartMode     = $null
                ProcessId     = $null
                Expectation   = $expectation
                Healthy       = $healthy
            }

            $severity = if ($healthy) { 'Info' } else { 'Critical' }
            Add-Finding -Severity $severity -Category 'Services' -Check $serviceName `
                -Message ("Service {0} is not installed; expectation={1}." -f $serviceName, $expectation)
            continue
        }

        $healthy = switch ($expectation) {
            'RequiredRunning' {
                ($service.State -eq 'Running' -and $service.StartMode -ne 'Disabled')
            }
            'AllowedStoppedManual' {
                (
                    $service.State -eq 'Running' -or
                    ($service.State -eq 'Stopped' -and $service.StartMode -eq 'Manual')
                )
            }
            'Optional' {
                $true
            }
            default {
                ($service.State -eq 'Running')
            }
        }

        $results += [pscustomobject]@{
            Name          = $service.Name
            DisplayName   = $service.DisplayName
            Exists        = $true
            State         = $service.State
            StartMode     = $service.StartMode
            ProcessId     = $service.ProcessId
            Expectation   = $expectation
            Healthy       = $healthy
        }

        $severity = if ($healthy) { 'Healthy' } else { 'Critical' }
        Add-Finding -Severity $severity -Category 'Services' -Check $service.Name `
            -Message ("Service {0}: state={1}, startup={2}, expectation={3}." -f
                $service.Name, $service.State, $service.StartMode, $expectation)
    }

    [pscustomobject]@{
        CheckedCount  = $results.Count
        UnhealthyCount= @($results | Where-Object { $_.Healthy -eq $false }).Count
        Services      = New-ObjectArrayForJson -InputObject $results
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
            Where-Object { Get-ObjectPropertyValueSafe -InputObject $_ -Name 'DisplayName' })
    }

    $agents = @()

    foreach ($pattern in $ManagementAgentPatterns) {
        $escaped = [regex]::Escape($pattern)

        $matchingServices = @($serviceInventory | Where-Object {
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'Name') -match $escaped -or
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'DisplayName') -match $escaped -or
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'PathName') -match $escaped
        })

        $matchingProcesses = @($processInventory | Where-Object {
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'Name') -match $escaped -or
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'ExecutablePath') -match $escaped -or
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'CommandLine') -match $escaped
        })

        $matchingProducts = @($uninstallEntries | Where-Object {
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'DisplayName') -match $escaped -or
            (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'Publisher') -match $escaped
        })

        if ($matchingServices.Count -gt 0 -or $matchingProcesses.Count -gt 0 -or $matchingProducts.Count -gt 0) {
            $services = @($matchingServices | ForEach-Object {
                [pscustomobject]@{
                    Name        = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'Name')
                    DisplayName = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'DisplayName')
                    State       = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'State')
                    StartMode   = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'StartMode')
                    ProcessId   = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'ProcessId')
                }
            })

            $products = @($matchingProducts | ForEach-Object {
                [pscustomobject]@{
                    DisplayName    = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'DisplayName')
                    DisplayVersion = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'DisplayVersion')
                    Publisher      = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'Publisher')
                    InstallDate    = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'InstallDate')
                }
            })
            $processes = @($matchingProcesses | ForEach-Object {
                [pscustomobject]@{
                    Name           = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'Name')
                    ProcessId      = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'ProcessId')
                    ExecutablePath = (Get-ObjectPropertyValueSafe -InputObject $_ -Name 'ExecutablePath')
                }
            })

            $serviceHealthy = if ($services.Count -gt 0) {
                @($services | Where-Object { $_.State -eq 'Running' }).Count -gt 0
            } else {
                $processes.Count -gt 0
            }

            $agents += [pscustomobject]@{
                Pattern        = $pattern
                Detected       = $true
                Healthy        = $serviceHealthy
                Services       = New-ObjectArrayForJson -InputObject $services
                Processes      = New-ObjectArrayForJson -InputObject $processes
                InstalledItems = New-ObjectArrayForJson -InputObject $products
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
        PatternsChecked = New-StringArrayForJson -InputObject $ManagementAgentPatterns
        DetectedCount    = $agents.Count
        Agents           = New-ObjectArrayForJson -InputObject $agents
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
        UpdateServices     = New-ObjectArrayForJson -InputObject $services
        UpdateTasks        = New-ObjectArrayForJson -InputObject $tasks
        UpdatePolicy       = [pscustomobject]$policy
        CrashCount         = $edgeCrashes.Count
        Crashes            = New-ObjectArrayForJson -InputObject $edgeCrashes
    }
}

function Get-HealthSummary {
    $healthy  = @($script:Findings | ForEach-Object { $_ } | Where-Object Severity -eq 'Healthy').Count
    $info     = @($script:Findings | ForEach-Object { $_ } | Where-Object Severity -eq 'Info').Count
    $warnings = @($script:Findings | ForEach-Object { $_ } | Where-Object Severity -eq 'Warning').Count
    $critical = @($script:Findings | ForEach-Object { $_ } | Where-Object Severity -eq 'Critical').Count

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
        Findings      = New-ObjectArrayForJson -InputObject $script:Findings
    }
}

function Write-HealthSnapshotSummary {
    param(
        $OperatingSystem,
        $Performance,
        $Storage,
        $DeviceManager,
        $Defender,
        $Firewall,
        $EventLogs,
        $Services,
        $Agents,
        $Edge,
        [Parameter(Mandatory)]$Summary
    )

    Write-Log -Message ("Health summary: Status={0}; Healthy={1}; Info={2}; Warning={3}; Critical={4}." -f `
        $Summary.Status, $Summary.HealthyCount, $Summary.InfoCount, $Summary.WarningCount, $Summary.CriticalCount)

    if ($OperatingSystem) {
        Write-Log -Message ("Operating system: Build={0}; UptimeDays={1}; PendingReboot={2}; Reasons={3}." -f `
            $OperatingSystem.FullBuild, $OperatingSystem.UptimeDays, $OperatingSystem.PendingReboot, `
            (@($OperatingSystem.PendingRebootReasonNames) -join ','))
    }
    if ($Performance) {
        Write-Log -Message ("Performance: CpuAverage={0}%; CpuMaximum={1}%; MemoryUsed={2}% ({3}/{4} GB)." -f `
            $Performance.CpuAveragePercent, $Performance.CpuMaximumPercent, $Performance.MemoryUsedPercent, `
            $Performance.MemoryUsedGB, $Performance.MemoryTotalGB)
    }
    if ($Storage) {
        foreach ($volume in @($Storage.Volumes)) {
            Write-Log -Message ("Volume: Drive={0}; Used={1}% ({2} GB); Free={3} GB; Size={4} GB; FileSystem={5}." -f `
                $volume.DriveLetter, $volume.UsedPercent, $volume.UsedGB, $volume.FreeGB, $volume.SizeGB, $volume.FileSystem)
        }
        if ($Storage.CapabilityAccessManager) {
            $cam = $Storage.CapabilityAccessManager
            Write-Log -Message ("CapabilityAccessManager.db-wal: Status={0}; Exists={1}; SizeGB={2}; PercentOfSystemDrive={3}; WarningGB={4}; CriticalGB={5}." -f `
                $cam.Status, $cam.Exists, $cam.SizeGB, $cam.PercentOfSystemDrive, $cam.WarningThresholdGB, $cam.CriticalThresholdGB)
        }
    }

    Write-Log -Message ("Operational counts: DeviceProblems={0}; DisabledFirewallProfiles={1}; CriticalEvents={2}; ApplicationCrashes={3}; UnhealthyServices={4}; ManagementAgents={5}; EdgeHealth={6}; EdgeVersion={7}." -f `
        $(if ($DeviceManager) { $DeviceManager.ProblemCount } else { $null }), `
        $(if ($Firewall) { $Firewall.DisabledProfileCount } else { $null }), `
        $(if ($EventLogs) { $EventLogs.CriticalEventCount } else { $null }), `
        $(if ($EventLogs) { $EventLogs.ApplicationCrashCount } else { $null }), `
        $(if ($Services) { $Services.UnhealthyCount } else { $null }), `
        $(if ($Agents) { $Agents.DetectedCount } else { $null }), `
        $(if ($Edge) { $Edge.Health } else { $null }), `
        $(if ($Edge) { $Edge.Version } else { $null }))

    if ($Defender -and $Defender.Available) {
        Write-Log -Message ("Defender: Available={0}; AntivirusEnabled={1}; RealTimeProtection={2}; SignatureVersion={3}; SignatureUpdated={4}; ThreatDetections={5}." -f `
            $Defender.Available, $Defender.AntivirusEnabled, $Defender.RealTimeProtectionEnabled, `
            $Defender.AntivirusSignatureVersion, $Defender.AntivirusSignatureLastUpdated, $Defender.RecentThreatDetectionCount)
    }
    elseif ($Defender) {
        Write-Log -Level 'WARN' -Message 'Defender: status information is unavailable on this endpoint.'
    }

    foreach ($finding in @($Summary.Findings | Where-Object { $_.Severity -in @('Warning','Critical') })) {
        $level = if ($finding.Severity -eq 'Critical') { 'ERROR' } else { 'WARN' }
        Write-Log -Level $level -Message ("Health finding: Severity={0}; Category={1}; Check={2}; Message={3}" -f `
            $finding.Severity, $finding.Category, $finding.Check, $finding.Message)
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

function Write-HealthFindingTelemetry {
    param(
        [Parameter(Mandatory)]$Finding,
        [Parameter(Mandatory)][int]$FindingNumber,
        [Parameter(Mandatory)][int]$FindingCount,
        [Parameter(Mandatory)][datetime]$EventTime,
        $Identity,
        [Parameter(Mandatory)]$Summary
    )

    $findingEvent = [ordered]@{
        '@timestamp'   = $EventTime.ToUniversalTime().ToString('o')
        EventType      = 'endpoint.health.finding'
        SchemaVersion  = '1.0'
        RunId          = $script:RunId

        ComputerName   = $env:COMPUTERNAME
        Domain         = if ($Identity) { $Identity.Domain } else { $env:USERDOMAIN }
        Building       = if ($Identity) { $Identity.Building } else { $null }
        Lab            = if ($Identity) { $Identity.Lab } else { $null }
        DeviceIdentifier = if ($Identity) { $Identity.DeviceIdentifier } else { $null }

        ScriptName     = $script:ScriptName
        ScriptVersion  = $script:ScriptVersion
        FindingNumber  = $FindingNumber
        FindingCount   = $FindingCount

        Finding = [ordered]@{
            Severity     = [string]$Finding.Severity
            Category     = [string]$Finding.Category
            Check        = [string]$Finding.Check
            Message      = [string]$Finding.Message
            ValueType    = [string]$Finding.ValueType
            ValueNumber  = $Finding.ValueNumber
            ValueBoolean = $Finding.ValueBoolean
            ValueText    = $Finding.ValueText
            ValueTexts   = New-StringArrayForJson -InputObject $Finding.ValueTexts
            ValueJson    = $Finding.ValueJson
        }

        HealthSummary = [ordered]@{
            Status        = [string]$Summary.Status
            WarningCount  = [int]$Summary.WarningCount
            CriticalCount = [int]$Summary.CriticalCount
        }
    }

    $jsonCompact = $findingEvent | ConvertTo-Json -Depth 12 -Compress
    Write-MaintenanceTelemetryLine -Path $script:NdjsonPath -JsonLine $jsonCompact
}

Ensure-Directory -Path $LogDirectory
Write-Log -Message ("Starting {0} version {1}. RunId={2}" -f
    $script:ScriptName, $script:ScriptVersion, $script:RunId)
Write-Log -Message ("Active staged text log: {0}" -f $script:LogPath)
Write-Log -Message ("Completed text log publish path: {0}" -f $script:PublishedLogPath)

$identity        = Invoke-Collector -Name 'ComputerIdentity'       -ScriptBlock { Get-ComputerIdentity }
$hardware        = Invoke-Collector -Name 'HardwareInventory'      -ScriptBlock { Get-HardwareInventory }
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
Write-HealthSnapshotSummary -OperatingSystem $operatingSystem -Performance $performance -Storage $storage `
    -DeviceManager $deviceManager -Defender $defender -Firewall $firewall -EventLogs $eventLogs `
    -Services $services -Agents $agents -Edge $edge -Summary $summary

if ($hardware) {
    Write-Log -Message ("Hardware inventory summary: Building={0}; Lab={1}; CPU={2}; Cores={3}; LogicalProcessors={4}; RAM={5}GB; MemorySlots={6}/{7}; GPU={8}; GPUCount={9}; StorageManufacturer={10}; StorageModel={11}; StorageBus={12}; StorageGB={13}." -f `
        $(if ($identity) { $identity.Building } else { $null }),
        $(if ($identity) { $identity.Lab } else { $null }),
        $hardware.CPU.PrimaryModel,
        $hardware.CPU.TotalCores,
        $hardware.CPU.TotalLogicalProcessors,
        $hardware.Memory.TotalInstalledGB,
        $hardware.Memory.PopulatedSlots,
        $hardware.Memory.SlotCount,
        $hardware.GPU.PrimaryModel,
        $hardware.GPU.AdapterCount,
        $(if ($storage) { $storage.PrimaryDiskManufacturer } else { $null }),
        $(if ($storage) { $storage.PrimaryDiskModel } else { $null }),
        $(if ($storage) { $storage.PrimaryDiskBusType } else { $null }),
        $(if ($storage) { $storage.PrimaryDiskSizeGB } else { $null }))
}

$endTime = Get-Date
$failedCollectors = @($script:CollectorResults | ForEach-Object { $_ } | Where-Object Status -eq 'Failed').Count
$systemVolume = if ($storage) {
    @($storage.Volumes | Where-Object { $_.DriveLetter -eq $env:SystemDrive } | Select-Object -First 1)[0]
} else { $null }
$camWal = if ($storage) { $storage.CapabilityAccessManager } else { $null }

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
    Building     = if ($identity) { $identity.Building } else { $null }
    Lab          = if ($identity) { $identity.Lab } else { $null }
    DeviceIdentifier = if ($identity) { $identity.DeviceIdentifier } else { $null }

    ScriptName    = $script:ScriptName
    ScriptVersion = $script:ScriptVersion
    Status        = $executionStatus
    ExitCode      = $exitCode
    StartTime     = $script:StartTime.ToUniversalTime().ToString('o')
    EndTime       = $endTime.ToUniversalTime().ToString('o')
    DurationSeconds = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 3)
    WarningCount  = $script:WarningCount
    ErrorCount    = $script:ErrorCount
    TextLogPath   = $script:PublishedLogPath

    AlertMetrics = [ordered]@{
        CpuAveragePercent = if ($performance) { $performance.CpuAveragePercent } else { $null }
        CpuMaximumPercent = if ($performance) { $performance.CpuMaximumPercent } else { $null }
        MemoryUsedPercent = if ($performance) { $performance.MemoryUsedPercent } else { $null }
        SystemDriveUsedPercent = if ($systemVolume) { $systemVolume.UsedPercent } else { $null }
        SystemDriveFreeGB = if ($systemVolume) { $systemVolume.FreeGB } else { $null }
        CapabilityAccessManagerWalSizeGB = if ($camWal) { $camWal.SizeGB } else { $null }
        CapabilityAccessManagerWalStatus = if ($camWal) { $camWal.Status } else { $null }
        DeviceManagerProblemCount = if ($deviceManager) { $deviceManager.ProblemCount } else { $null }
        DeviceManagerDegradedNonActionableCount = if ($deviceManager) { $deviceManager.DegradedNonActionableCount } else { $null }
        CriticalEventCount = if ($eventLogs) { $eventLogs.CriticalEventCount } else { $null }
        ApplicationCrashCount = if ($eventLogs) { $eventLogs.ApplicationCrashCount } else { $null }
        UnhealthyCriticalServiceCount = if ($services) { $services.UnhealthyCount } else { $null }
        DisabledFirewallProfileCount = if ($firewall) { $firewall.DisabledProfileCount } else { $null }
        DefenderRecentThreatDetectionCount = if ($defender -and $defender.Available) { $defender.RecentThreatDetectionCount } else { $null }
        PendingReboot = if ($operatingSystem) { $operatingSystem.PendingReboot } else { $null }
        PendingRebootReasonNames = if ($operatingSystem) {
            New-StringArrayForJson -InputObject $operatingSystem.PendingRebootReasonNames
        } else {
            [object[]]@()
        }
        HealthWarningCount = $summary.WarningCount
        HealthCriticalCount = $summary.CriticalCount
        CollectorFailureCount = $failedCollectors
    }

    HealthSummary = $summary
    Identity      = $identity
    Hardware      = $hardware
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
    Collectors    = New-ObjectArrayForJson -InputObject $script:CollectorResults
}

try {
    Write-Log -Message ("Elastic endpoint-health summary: Health={0}; Critical={1}; Warnings={2}; CPU={3}%; Memory={4}%; SystemDrive={5}%; PendingReboot={6}; DeviceProblems={7}; FirewallDisabled={8}; CollectorFailures={9}; FindingEvents={10}" -f `
        $summary.Status,
        $summary.CriticalCount,
        $summary.WarningCount,
        $(if ($performance) { $performance.CpuAveragePercent } else { $null }),
        $(if ($performance) { $performance.MemoryUsedPercent } else { $null }),
        $(if ($systemVolume) { $systemVolume.UsedPercent } else { $null }),
        $(if ($operatingSystem) { $operatingSystem.PendingReboot } else { $null }),
        $(if ($deviceManager) { $deviceManager.ProblemCount } else { $null }),
        $(if ($firewall) { $firewall.DisabledProfileCount } else { $null }),
        $failedCollectors,
        @($summary.Findings | Where-Object { $_.Severity -in @('Warning','Critical') }).Count)

    Write-Telemetry -Event $event

    $actionableFindings = @($summary.Findings | Where-Object { $_.Severity -in @('Warning','Critical') })
    for ($findingIndex = 0; $findingIndex -lt $actionableFindings.Count; $findingIndex++) {
        Write-HealthFindingTelemetry -Finding $actionableFindings[$findingIndex] `
            -FindingNumber ($findingIndex + 1) -FindingCount $actionableFindings.Count `
            -EventTime $endTime -Identity $identity -Summary $summary
    }

    Write-Log -Message ("Endpoint snapshot completed. Health={0}; ExitCode={1}; CollectorsFailed={2}; FindingEvents={3}" -f
        $summary.Status, $exitCode, $failedCollectors, $actionableFindings.Count)
}
catch {
    Write-Log -Level 'ERROR' -Message ("Unable to write endpoint telemetry: {0}" -f $_.Exception.Message)
    $exitCode = 4
}

# Final append before the completed immutable file enters C:\Logs.
Write-Log -Message ("Finalizing {0}. Health={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
    $script:ScriptName,
    $summary.Status,
    $exitCode,
    $script:WarningCount,
    $script:ErrorCount) `
    -Level $(if ($exitCode -eq 4) { 'ERROR' } elseif ($exitCode -eq 2) { 'WARN' } else { 'INFO' })

if ($null -ne $script:LogSession) {
    $publishResult = Publish-MaintenanceLog -LogSession $script:LogSession

    if ($publishResult.Published) {
        Write-Host ("Published completed script 14 text log for Elastic: {0}" -f $script:PublishedLogPath) -ForegroundColor Green
    }
    else {
        Write-Warning ("Script 14 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
    }
}

exit $exitCode
