#requires -Version 5.1
#requires -RunAsAdministrator
# ScriptName:    04_Sunday_Lab_Application_Maintenance.ps1
# ScriptVersion: 1.6.0
# LastUpdated:   2026-09-15
<#
.SYNOPSIS
    Runs the Sunday lab application and configuration maintenance in one script.

.DESCRIPTION
    Consolidates the former scripts 11, 12, 13, 15, 17, 18, and 19 and includes
    native Office LTSC migration, Office 2024 installation, and activation
    maintenance. Microsoft 365 Apps installations are intentionally skipped.
    The System Restore and Stellarium
    Location Services code is fully bundled in this file; separately deployed
    scripts 12 and 19 are not needed.
    Each embedded section
    runs in its own 64-bit Windows PowerShell process so duplicate function names,
    strict-mode settings, and a section's final exit statement cannot interfere
    with the remaining sections. A failed section is recorded and the runner
    continues with every other enabled section.

    The original per-section text logs, latest-state JSON files, and NDJSON
    telemetry are preserved. This runner also writes its own summary log and JSON.

.NOTES
    ScriptName:    04_Sunday_Lab_Application_Maintenance.ps1
    ScriptVersion: 1.6.0
    LastUpdated:   2026-09-15
    Requires:      64-bit Windows PowerShell 5.1, Administrator or SYSTEM

    Changes:       v1.6.0 expands Chrome lab policy enforcement to suppress the
                   browser sign-in/onboarding experience and default-browser prompt
                   on every computer, and enriches Honorlock telemetry with
                   Building/Lab/DeviceIdentifier and target-pattern information for
                   scalable Kibana deployment dashboards.
                   v1.5.0 replaces all Gzip/Base64 embedded payloads with directly
                   readable plain-text source blocks. Sections are still staged and
                   executed in isolated child PowerShell processes, preserving the
                   existing behavior while making future maintenance much easier.
                   v1.4.2 normalizes SHARP/PaperCut SuccessRebootRequired (3010)
                   from the child section telemetry so a required reboot is not
                   misclassified as a section failure by the parent runner.
                   v1.4.1 fixes Office registry detection under StrictMode, Honorlock
                   ExtensionSettings creation when the value does not yet exist,
                   and the PowerShell 5.1 generic-list summary serialization error.
                   v1.4.0 detects Microsoft 365, Office 2024, and older Office
                   LTSC/perpetual suites. Microsoft 365 is skipped. Supported
                   older Click-to-Run suites are removed with targeted Office
                   Deployment Tool XML; older MSI suites require the supported
                   <RemoveMSI /> element in office2024config.xml. Office 2024 is
                   staged locally, installed, machine-wide activation is
                   verified, dashboard telemetry is written, and staged files
                   are removed in a finally block.
    Changes:       v1.3.0 embeds script 12 as the independently switchable
                   SystemRestore section. It preserves script 12's restore-point
                   verification, retention, staged logging, and telemetry.
                   v1.2.0 confirms the supplied Stellarium Location Services
                   implementation is bundled as the StellariumLocation section,
                   retains prefix targeting and verification, and removes any
                   runtime dependency on a separately deployed numbered script.

    This single file contains plain-text embedded copies of the seven source scripts.
    They are staged to temporary files only so each section can run in an isolated
    64-bit Windows PowerShell process. No compression or Base64 regeneration is
    required when maintaining an embedded section. The Autologon password and
    Elastic enrollment token remain sensitive; restrict access to the script/share.
#>

[CmdletBinding()]
param(
    # Reinstall Elastic Agent even if an existing installation is healthy.
    [switch]$ForceElasticAgentReinstall
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ============================================================================
# EASY-TO-EDIT COMPUTER TARGETING
# ============================================================================
# Use PowerShell wildcard patterns for Autologon, Honorlock, and Stellarium.
# Examples: 'SSB-122-*', 'SSC-216*', or '*' for every computer.
#
# Elastic Agent uses computer-name PREFIXES, not wildcard patterns.
# Example: 'IB1-103' matches IB1-103-01, IB1-103-02, and so on.
#
# Add or remove quoted entries in the appropriate list. Keep the @(...)
# structure and separate entries with commas.

# --- Autologon and Edge startup targets (former script 13) ------------------
[string[]]$AutologonComputerPatterns = @(
    'SSB-122-*',
    'SSB-114*',
    'SSB-171*'
)

# --- Elastic Agent installation targets (former script 15) -----------------
[string[]]$ElasticAgentComputerPrefixes = @(
    'IB1-103',
	'SSC-216*'
)

# --- Honorlock Chrome extension targets (former script 18) -----------------
# Add one wildcard/prefix per lab that requires Honorlock. The same list is
# passed into the embedded Honorlock section and is also written to telemetry,
# allowing Kibana to show which labs are targeted and which computers verified
# the machine-wide policy. Examples:
#   'SSC-216*'
#   'AHB-146*'
#   'IB1-103*'
# Leave the array empty, @(), ONLY if Honorlock should target every computer.
[string[]]$HonorlockComputerPatterns = @(
    'SSC-216*',
    'AHB-146*'
)

# --- Stellarium Windows Location Services targets (former script 19) --------
[string[]]$StellariumComputerPatterns = @(
    'SSC-216*',
    'MS-203*'
)

# ============================================================================
# SECTION ENABLE/DISABLE SWITCHES
# ============================================================================
# Set a value to $false to retain the section in this combined file but skip it.
[bool]$RunPrinterAndPaperCut = $true
[bool]$RunSystemRestore     = $true
[bool]$RunAutologonAndEdge   = $true
[bool]$RunElasticAgent      = $true
[bool]$RunBrowserHomepage   = $true
[bool]$RunHonorlock         = $true
[bool]$RunStellariumLocation = $true
[bool]$RunOffice2024Maintenance = $true

# ============================================================================
# GENERAL SETTINGS
# ============================================================================
[string]$HomepageUrl = 'https://www.compton.edu'
[string]$Office2024SourcePath = '\\filesvr\Labscripts\Installers\Office2024'
[string]$Office2024ConfigurationFile = 'office2024config.xml'
[string]$LogDirectory = 'C:\Logs'
[string]$RunnerScriptName = '04_Sunday_Lab_Application_Maintenance.ps1'
[string]$RunnerVersion = '1.6.0'
[string]$RunnerLogPath = Join-Path $LogDirectory '04_Sunday_Lab_Application_Maintenance.log'
[string]$RunnerLatestPath = Join-Path $LogDirectory '04_Sunday_Lab_Application_Maintenance.latest.json'
[string]$RunnerTelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
[string]$OfficeMigrationStateDirectory = Join-Path $env:ProgramData 'Compton\Office2024'
[string]$OfficeMigrationStatePath = Join-Path $OfficeMigrationStateDirectory 'MigrationState.json'
$RunnerStartTime = Get-Date
$RunnerRunId = [guid]::NewGuid().Guid

# ============================================================================
# EMBEDDED SECTION SOURCE -- EDIT THESE BLOCKS WHEN MAINTAINING A SECTION
# ============================================================================
$EmbeddedSections = [ordered]@{
    SystemRestore = [ordered]@{
        FileName = 'Embedded_SystemRestore.ps1'
        # Plain-text source is intentionally embedded for maintainability.
        Source = @'
﻿<#
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

    $source = @"
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
"@

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
'@
    }
    PrinterAndPaperCut = [ordered]@{
        FileName = '11_Install_SharpDriver_And_PaperCut.ps1'
        # Plain-text source is intentionally embedded for maintainability.
        Source = @'
<#
.SYNOPSIS
    Installs and maintains the SHARP printer driver, PaperCut Print Deploy client,
    and StudentSecurePrint shared printer connection.

.DESCRIPTION
    Performs idempotent before/after validation of the SHARP driver, PaperCut
    client, and shared printer connection. Driver package deployment is based on
    a deterministic SHA-256 fingerprint. Structured telemetry is written for
    Elastic ingestion while preserving a readable operational log.

.VERSION
    1.2.8

.DATE
    2026-08-17

.CHANGELOG
    1.2.8
    - Fixes empty process-result telemetry so Processes serializes as [] instead of [null].
    - Corrects Windows 11 product-name reporting on builds 22000+ when the legacy registry ProductName still reports Windows 10.
    - Preserves all v1.2.7 Elastic mapping-safety changes and printer/PaperCut compliance telemetry.

    1.2.7
    - Normalizes reboot-reason collections as true JSON string arrays for stable Elastic mappings.
    - Renames reboot reason fields to ReasonNamesBefore/ReasonNamesAfter to avoid collisions with older generic reboot fields.
    - Normalizes process-result telemetry as a true JSON array when empty.
    - Preserves the existing maintenance.printerDeployment namespace and all v1.2.6 compliance/remediation fields.

    1.2.6
    - Replaced mapping-conflicting ChangedComponents telemetry with stable boolean Changes.Driver/PaperCut/PrinterConnection fields.
    - Preserves all v1.2.5 remediation and before/after telemetry.

    1.2.5
    - Preserves immutable pre-remediation snapshots for SHARP, PaperCut, and printer connection telemetry.
    - Adds explicit remediation Required/Attempted/Succeeded/ExitCode/FailureMessage fields.
    - InstallRequired and InstalledThisRun are calculated from immutable before/after snapshots at telemetry time.

    1.2.4
    - Fixed strict-mode failure when Win32_PrinterDriver does not expose DriverVersion.
    - Printer-driver version collection now safely checks Version/DriverVersion and the actual driver file version.

    1.2.3
    - Added explicit DriverInstallRequired/DriverInstalledThisRun and PaperCutInstallRequired/PaperCutInstalledThisRun telemetry.
    - Added retry/fallback SHARP printer-driver verification after Add-PrinterDriver.
    - Normalized generic-list telemetry collections for Windows PowerShell 5.1.

    1.2.2
    - Added Maintenance.Framework v2.4 staged text logging.
    - Completed text logs are published into C:\Logs only after telemetry finishes.
    - Previous script 11 text logs are archived under C:\Logs\Old Logs.

    1.2.1
    - Added Elastic-friendly computer and normalized severity fields to text logs.
    - Added readable before/after SHARP driver, PaperCut, and printer summaries.
    - Added detailed installed printer-driver version and file metadata.
    - Added PaperCut source-MSI product metadata when the package is available.

    1.2.0
    - Added common maintenance execution telemetry and latest-state JSON.
    - Added before/after compliance verification for driver, PaperCut, and printer.
    - Added external-process timeouts and captured exit codes/durations.
    - Reordered source validation so an unavailable share does not block unrelated
      validation when the corresponding component is already compliant.
    - Added PaperCut post-install verification and MSI product details.
    - Added driver package, printer connection, reboot, and Windows build telemetry.
    - Added meaningful final statuses and exit codes.
    - Added atomic deployment-state writes and damaged-state backup handling.

.NOTES
    Designed for elevated 64-bit Windows PowerShell 5.1 and SYSTEM scheduled tasks.
#>

[CmdletBinding()]
param(
    [ValidateRange(60, 7200)]
    [int]$DriverInstallTimeoutSeconds = 900,

    [ValidateRange(60, 7200)]
    [int]$PaperCutInstallTimeoutSeconds = 900,

    [ValidateRange(1, 120)]
    [int]$PrinterVerificationDelaySeconds = 5
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# =========================
# Configuration
# =========================
$ScriptName = '11_Install_SharpDriver_And_PaperCut.ps1'
$ScriptVersion = '1.2.8'
$DriverSourcePath = '\\papercut\Printer Drivers\MFP\win\SH_D31_PCL6_PS_2410a_EnglishUS_64bit'
$PrinterDriverName = 'Sharp BP-70C31 PCL6'
$PaperCutMsiPath = '\\papercut\Print Deploy Clients\win\pc-print-deploy-client[10.2.3.44].msi'
$PrinterSharePath = '\\papercut\StudentSecurePrint'
$LocalDriverStage = 'C:\ProgramData\Compton\Drivers\Sharp'
$StateDirectory = 'C:\ProgramData\Compton\State'
$StatePath = Join-Path $StateDirectory 'SharpDriver-PaperCut-State.json'
$LogDirectory = 'C:\Logs'
$LogPath = $null
$PublishedLogPath = $null
$LogSession = $null
$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestTelemetryPath = Join-Path $LogDirectory '11_Install_SharpDriver_And_PaperCut.latest.json'

$script:RunId = [guid]::NewGuid().Guid
$script:StartTime = Get-Date
$script:WarningCount = 0
$script:ErrorCount = 0
$script:RebootRequired = $false
$script:FailureMessage = $null
$script:ChangedComponents = New-Object System.Collections.Generic.List[string]
$script:ProcessResults = New-Object System.Collections.Generic.List[object]
$script:DriverInstallRequired = $false
$script:DriverInstalledThisRun = $false
$script:PaperCutInstallRequired = $false
$script:PaperCutInstalledThisRun = $false
$script:DriverRemediationAttempted = $false
$script:DriverRemediationSucceeded = $false
$script:DriverRemediationExitCode = $null
$script:DriverRemediationFailureMessage = $null
$script:PaperCutRemediationAttempted = $false
$script:PaperCutRemediationSucceeded = $false
$script:PaperCutRemediationExitCode = $null
$script:PaperCutRemediationFailureMessage = $null



# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 11 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        'Install-SharpDriver-And-PaperCut.log',
        '11_Install_SharpDriver_And_PaperCut.log',
        '*-11_Install_SharpDriver_And_PaperCut-*.log'
    ) | Out-Null

$LogSession = New-MaintenanceStagedLog `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $env:COMPUTERNAME `
    -Timestamp $script:StartTime

$LogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','ACTION','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
    $normalizedLevel = if ($Level -eq 'WARN') { 'WARNING' } else { $Level }
    $line = '{0} [{1}] [{2}] {3}' -f $timestamp, $computerName, $normalizedLevel, $Message
    $color = switch ($Level) {
        'ACTION' { 'Yellow' }
        'SUCCESS' { 'Green' }
        'WARN' { 'DarkYellow' }
        'ERROR' { 'Red' }
        default { 'White' }
    }
    Write-Host $line -ForegroundColor $color
    try {
        $activeLogDirectory = Split-Path -Parent $LogPath
        if (-not (Test-Path -LiteralPath $activeLogDirectory -PathType Container)) {
            New-Item -Path $activeLogDirectory -ItemType Directory -Force | Out-Null
        }
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch { }
}

function Write-Banner {
    $border = '=' * 72
    Write-Host ''
    Write-Host $border -ForegroundColor Cyan
    Write-Host '     SHARP Driver and PaperCut Print Deploy Maintenance' -ForegroundColor Cyan
    Write-Host $border -ForegroundColor Cyan
    Write-Host ''
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-IsSystemAccount {
    try { return ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18') }
    catch { return $false }
}

function Get-WindowsInfo {
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop

        $productName = [string]$cv.ProductName
        $currentBuild = [string]$cv.CurrentBuild
        $buildNumber = 0
        [void][int]::TryParse($currentBuild, [ref]$buildNumber)

        # Windows 11 can retain a legacy "Windows 10" ProductName in this registry key.
        # Build 22000 and later are Windows 11 client builds, so normalize the display name.
        if ($buildNumber -ge 22000 -and $productName -match '^Windows 10') {
            $productName = $productName -replace '^Windows 10', 'Windows 11'
        }

        [ordered]@{
            ProductName = $productName
            EditionID = [string]$cv.EditionID
            DisplayVersion = [string]$cv.DisplayVersion
            CurrentBuild = $currentBuild
            UBR = [int]$cv.UBR
            FullBuild = '{0}.{1}' -f $currentBuild, $cv.UBR
        }
    }
    catch { [ordered]@{} }
}

function Get-PendingRebootState {
    $reasons = New-Object System.Collections.Generic.List[string]
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons.Add('ComponentBasedServicing') }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons.Add('WindowsUpdate') }
    try {
        $session = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop
        if ($session.PendingFileRenameOperations) { $reasons.Add('PendingFileRenameOperations') }
    } catch { }
    [pscustomobject]@{ Pending = ($reasons.Count -gt 0); Reasons = @($reasons) }
}

function Invoke-ExternalProcess {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [int[]]$SuccessExitCodes = @(0),
        [ValidateRange(1,7200)][int]$TimeoutSeconds = 900,
        [Parameter(Mandatory)][string]$Operation
    )
    Write-Log ("Running {0}: {1} {2}" -f $Operation, $FilePath, ($ArgumentList -join ' ')) 'ACTION'
    $started = Get-Date
    $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -WindowStyle Hidden
    $completed = $process.WaitForExit($TimeoutSeconds * 1000)
    if (-not $completed) {
        try { $process.Kill() } catch { }
        $duration = [math]::Round(((Get-Date) - $started).TotalSeconds, 2)
        $result = [pscustomobject]@{ Operation=$Operation; FilePath=$FilePath; ExitCode=$null; DurationSeconds=$duration; TimedOut=$true; Succeeded=$false }
        $script:ProcessResults.Add($result)
        throw "$Operation timed out after $TimeoutSeconds seconds."
    }
    $process.Refresh()
    $duration = [math]::Round(((Get-Date) - $started).TotalSeconds, 2)
    $exitCode = [int]$process.ExitCode
    $succeeded = ($exitCode -in $SuccessExitCodes)
    $result = [pscustomobject]@{ Operation=$Operation; FilePath=$FilePath; ExitCode=$exitCode; DurationSeconds=$duration; TimedOut=$false; Succeeded=$succeeded }
    $script:ProcessResults.Add($result)
    Write-Log ("{0} exit code: {1}; duration: {2}s" -f $Operation, $exitCode, $duration)
    if (-not $succeeded) { throw "$Operation failed with exit code $exitCode." }
    return $exitCode
}

function Get-PrinterInfFiles {
    param([Parameter(Mandatory)][string]$Path)
    $allInfFiles = @(Get-ChildItem -LiteralPath $Path -Filter '*.inf' -File -Recurse -ErrorAction Stop)
    if ($allInfFiles.Count -eq 0) { throw "No INF files were found under: $Path" }
    $printerInfFiles = foreach ($inf in $allInfFiles) {
        $content = Get-Content -LiteralPath $inf.FullName -ErrorAction SilentlyContinue
        if ($content -match '^\s*Class\s*=\s*Printer\s*$') { $inf }
    }
    if (@($printerInfFiles).Count -gt 0) { return @($printerInfFiles) }
    Write-Log 'No INF explicitly declared Class=Printer; all INF files will be staged.' 'WARN'
    return $allInfFiles
}

function Get-DriverPackageFingerprint {
    param([Parameter(Mandatory)][string]$Path)
    $files = @(Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction Stop | Sort-Object FullName)
    if ($files.Count -eq 0) { throw "No files were found in the driver source package: $Path" }
    $root = $Path.TrimEnd('\')
    $manifestLines = foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($root.Length).TrimStart('\')
        $hash = Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop
        '{0}|{1}|{2}' -f $relativePath.ToLowerInvariant(), $file.Length, $hash.Hash
    }
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes(($manifestLines -join "`n"))
        return ([BitConverter]::ToString($sha256.ComputeHash($bytes))).Replace('-','')
    }
    finally { $sha256.Dispose() }
}

function Get-DeploymentState {
    if (-not (Test-Path -LiteralPath $StatePath -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $StatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop) }
    catch {
        $backup = "$StatePath.corrupt.$(Get-Date -Format 'yyyyMMddHHmmss')"
        try { Move-Item -LiteralPath $StatePath -Destination $backup -Force -ErrorAction Stop } catch { }
        Write-Log "Deployment state was invalid and has been backed up: $backup" 'WARN'
        return $null
    }
}

function Save-DeploymentState {
    param([Parameter(Mandatory)][string]$Fingerprint)
    New-Item -Path $StateDirectory -ItemType Directory -Force | Out-Null
    $state = [ordered]@{
        DriverName=$PrinterDriverName
        SourcePath=$DriverSourcePath
        SourceFingerprintSHA256=$Fingerprint
        LastSuccessfulUpdate=(Get-Date).ToUniversalTime().ToString('o')
        ComputerName=$env:COMPUTERNAME
        ScriptVersion=$ScriptVersion
    }
    $temp = "$StatePath.tmp"
    $state | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $temp -Encoding UTF8 -Force
    Move-Item -LiteralPath $temp -Destination $StatePath -Force
}

function Convert-PrinterDriverVersion {
    param($DriverVersion)

    if ($null -eq $DriverVersion) { return $null }
    try {
        $value = [uint64]$DriverVersion
        return '{0}.{1}.{2}.{3}' -f `
            (($value -shr 48) -band 0xFFFF), `
            (($value -shr 32) -band 0xFFFF), `
            (($value -shr 16) -band 0xFFFF), `
            ($value -band 0xFFFF)
    }
    catch { return [string]$DriverVersion }
}

function Get-SafePropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string[]]$PropertyNames
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($propertyName in $PropertyNames) {
        $property = $InputObject.PSObject.Properties[$propertyName]
        if ($null -ne $property) {
            try {
                $value = $property.Value
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    return $value
                }
            }
            catch { }
        }
    }

    return $null
}

function Get-InstalledPrinterDriverState {
    [CmdletBinding()]
    param(
        [int]$RetryCount = 1,
        [int]$RetryDelaySeconds = 0
    )

    $driver = $null
    $cimDriver = $null

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            $driver = Get-PrinterDriver -Name $PrinterDriverName -ErrorAction SilentlyContinue

            if (-not $driver) {
                # Some SHARP packages register a slightly different display name.
                # Fall back to a constrained BP-70C31/PCL6 match rather than
                # declaring the deployment failed immediately.
                $driver = Get-PrinterDriver -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.Name -match '(?i)\bBP-70C31\b' -and
                        $_.Name -match '(?i)\bPCL6\b'
                    } |
                    Select-Object -First 1
            }

            $cimDriver = Get-CimInstance -ClassName Win32_PrinterDriver -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Name -eq $PrinterDriverName -or
                    $_.Name -like "$PrinterDriverName,*" -or
                    ($_.Name -match '(?i)\bBP-70C31\b' -and $_.Name -match '(?i)\bPCL6\b')
                } |
                Select-Object -First 1

            if ($driver -or $cimDriver) {
                break
            }
        }
        catch {
            # Retry below.
        }

        if ($attempt -lt $RetryCount -and $RetryDelaySeconds -gt 0) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    if (-not $driver -and -not $cimDriver) {
        return [pscustomobject][ordered]@{
            Installed        = $false
            Name             = $PrinterDriverName
            Manufacturer     = $null
            Version          = $null
            DriverVersionRaw = $null
            InfPath          = $null
            DriverPath       = $null
            ConfigFile       = $null
            DataFile         = $null
            MajorVersion     = $null
        }
    }

    $cimDriverPath = Get-SafePropertyValue -InputObject $cimDriver -PropertyNames @('DriverPath')
    $cimVersionRaw = Get-SafePropertyValue -InputObject $cimDriver -PropertyNames @('DriverVersion','Version')
    $driverFileVersion = $null

    if ($cimDriverPath -and (Test-Path -LiteralPath ([string]$cimDriverPath) -PathType Leaf)) {
        try {
            $driverFileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo([string]$cimDriverPath).FileVersion
        }
        catch { }
    }

    $driverName = Get-SafePropertyValue -InputObject $driver -PropertyNames @('Name')
    $cimName = Get-SafePropertyValue -InputObject $cimDriver -PropertyNames @('Name')
    $manufacturer = Get-SafePropertyValue -InputObject $driver -PropertyNames @('Manufacturer','ProviderName')
    $infPath = Get-SafePropertyValue -InputObject $driver -PropertyNames @('InfPath','InfName')
    $majorVersion = Get-SafePropertyValue -InputObject $driver -PropertyNames @('MajorVersion')
    $configFile = Get-SafePropertyValue -InputObject $cimDriver -PropertyNames @('ConfigFile')
    $dataFile = Get-SafePropertyValue -InputObject $cimDriver -PropertyNames @('DataFile')

    $normalizedVersion = $null
    if ($driverFileVersion) {
        $normalizedVersion = [string]$driverFileVersion
    }
    elseif ($null -ne $cimVersionRaw) {
        # Win32_PrinterDriver commonly exposes Version rather than DriverVersion.
        # If it is a large packed integer, decode it; otherwise retain the string.
        try {
            $numericVersion = [uint64]$cimVersionRaw
            if ($numericVersion -gt 65535) {
                $normalizedVersion = Convert-PrinterDriverVersion -DriverVersion $numericVersion
            }
            else {
                $normalizedVersion = [string]$cimVersionRaw
            }
        }
        catch {
            $normalizedVersion = [string]$cimVersionRaw
        }
    }

    [pscustomobject][ordered]@{
        Installed        = $true
        Name             = if ($driverName) { [string]$driverName } elseif ($cimName) { [string]$cimName } else { $PrinterDriverName }
        Manufacturer     = if ($manufacturer) { [string]$manufacturer } else { $null }
        Version          = $normalizedVersion
        DriverVersionRaw = if ($null -ne $cimVersionRaw) { [string]$cimVersionRaw } else { $null }
        InfPath          = if ($infPath) { [string]$infPath } else { $null }
        DriverPath       = if ($cimDriverPath) { [string]$cimDriverPath } else { $null }
        ConfigFile       = if ($configFile) { [string]$configFile } else { $null }
        DataFile         = if ($dataFile) { [string]$dataFile } else { $null }
        MajorVersion     = $majorVersion
    }
}


function Get-MsiPackageInfo {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
        $properties = [ordered]@{}
        foreach ($propertyName in @('ProductName','ProductVersion','ProductCode','Manufacturer')) {
            $query = "SELECT `Value` FROM `Property` WHERE `Property`='$propertyName'"
            $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @($query))
            $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
            $properties[$propertyName] = if ($record) { [string]$record.StringData(1) } else { $null }
            $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null) | Out-Null
        }
        return [pscustomobject]$properties
    }
    catch {
        Write-Log "Could not read PaperCut MSI metadata: $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Write-DeploymentStateSummary {
    param(
        [Parameter(Mandatory)][string]$Label,
        $Driver,
        $PaperCut,
        $Printer
    )

    Write-Log ("Driver {0}: Installed={1}; Name={2}; Version={3}; Manufacturer={4}; INF={5}." -f `
        $Label, $Driver.Installed, $Driver.Name, $Driver.Version, $Driver.Manufacturer, $Driver.InfPath)
    Write-Log ("PaperCut {0}: Installed={1}; Name={2}; Version={3}; ProductCode={4}; InstallLocation={5}." -f `
        $Label, $PaperCut.Installed, $PaperCut.DisplayName, $PaperCut.DisplayVersion, $PaperCut.ProductCode, $PaperCut.InstallLocation)
    Write-Log ("Printer connection {0}: Connected={1}; Scope={2}; Share={3}." -f `
        $Label, $Printer.Connected, $Printer.Scope, $Printer.Name)
}

function Get-PaperCutProduct {
    $roots = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $product = Get-ItemProperty -Path $roots -ErrorAction SilentlyContinue | Where-Object {
        $p = $_.PSObject.Properties['DisplayName']
        $null -ne $p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value) -and [string]$p.Value -match 'PaperCut.*Print Deploy|Print Deploy Client'
    } | Select-Object -First 1
    if (-not $product) { return [pscustomobject]@{ Installed=$false; DisplayName=$null; DisplayVersion=$null; ProductCode=$null; InstallLocation=$null } }
    $version = $product.PSObject.Properties['DisplayVersion']
    [pscustomobject]@{
        Installed=$true
        DisplayName=[string]$product.DisplayName
        DisplayVersion=if ($version) { [string]$version.Value } else { $null }
        ProductCode=[string]$product.PSChildName
        InstallLocation=[string]$product.InstallLocation
    }
}

function Test-PerMachinePrinterConnection {
    param([Parameter(Mandatory)][string]$ConnectionPath)
    $connectionsPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections'
    if (-not (Test-Path -LiteralPath $connectionsPath)) { return $false }
    $parts = $ConnectionPath.TrimStart('\') -split '\\',2
    if ($parts.Count -ne 2) { return $false }
    $expected = ',,' + $parts[0] + ',' + $parts[1]
    return $null -ne (Get-ChildItem -LiteralPath $connectionsPath -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -ieq $expected } | Select-Object -First 1)
}

function Get-PrinterConnectionState {
    param([Parameter(Mandatory)][string]$ConnectionPath)
    if (Test-IsSystemAccount) {
        return [pscustomobject]@{ Scope='PerMachine'; Connected=(Test-PerMachinePrinterConnection -ConnectionPath $ConnectionPath); Name=$ConnectionPath }
    }
    $parts = $ConnectionPath.TrimStart('\') -split '\\',2
    $server = if ($parts.Count -eq 2) { $parts[0] } else { $null }
    $share = if ($parts.Count -eq 2) { $parts[1] } else { $null }
    $existing = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $ConnectionPath -or ($_.ComputerName -ieq $server -and $_.ShareName -ieq $share) } | Select-Object -First 1
    [pscustomobject]@{ Scope='CurrentUser'; Connected=($null -ne $existing); Name=$ConnectionPath }
}

function Ensure-SharedPrinterConnection {
    param([Parameter(Mandatory)][string]$ConnectionPath)
    $before = Get-PrinterConnectionState -ConnectionPath $ConnectionPath
    if ($before.Connected) {
        Write-Log "Printer connection is already configured: $ConnectionPath" 'SUCCESS'
        return $before
    }
    if (Test-IsSystemAccount) {
        Invoke-ExternalProcess -FilePath "$env:SystemRoot\System32\rundll32.exe" -ArgumentList @('printui.dll,PrintUIEntry','/ga',"/n`"$ConnectionPath`"") -SuccessExitCodes @(0) -TimeoutSeconds 120 -Operation 'Create per-machine printer connection' | Out-Null
    }
    else {
        try { Add-Printer -ConnectionName $ConnectionPath -ErrorAction Stop }
        catch {
            Write-Log "Add-Printer failed; using PrintUIEntry fallback. $($_.Exception.Message)" 'WARN'
            Invoke-ExternalProcess -FilePath "$env:SystemRoot\System32\rundll32.exe" -ArgumentList @('printui.dll,PrintUIEntry','/in',"/n`"$ConnectionPath`"") -SuccessExitCodes @(0) -TimeoutSeconds 120 -Operation 'Create current-user printer connection' | Out-Null
        }
    }
    Start-Sleep -Seconds $PrinterVerificationDelaySeconds
    $after = Get-PrinterConnectionState -ConnectionPath $ConnectionPath
    if (-not $after.Connected) { throw "Printer connection could not be verified: $ConnectionPath" }
    $script:ChangedComponents.Add('PrinterConnection')
    Write-Log "Printer connection created successfully: $ConnectionPath" 'SUCCESS'
    return $after
}

function Copy-TelemetrySnapshot {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) {
        return $null
    }

    # JSON round-trip deliberately creates a detached snapshot so later
    # remediation-state changes cannot alter the original "Before" object.
    return ($InputObject | ConvertTo-Json -Depth 10 | ConvertFrom-Json)
}

function Convert-ListToArray {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) {
        return @()
    }

    $result = @()
    foreach ($item in $InputObject) {
        $result += $item
    }

    return @($result)
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
        ForEach-Object { $_ }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Write-Telemetry {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [Parameter(Mandatory)]$DriverBefore,
        [Parameter(Mandatory)]$DriverAfter,
        [Parameter(Mandatory)]$PaperCutBefore,
        [Parameter(Mandatory)]$PaperCutAfter,
        [Parameter(Mandatory)]$PrinterBefore,
        [Parameter(Mandatory)]$PrinterAfter,
        $DriverSourceAvailable,
        $PaperCutSourceAvailable,
        $SourceFingerprint,
        $StoredFingerprint,
        $PaperCutSourceMsiInfo,
        [Parameter(Mandatory)]$PendingRebootBefore,
        [Parameter(Mandatory)]$PendingRebootAfter
    )
    $end = Get-Date

    $driverWasMissing = -not [bool]$DriverBefore.Installed
    $driverIsInstalled = [bool]$DriverAfter.Installed
    $paperCutWasMissing = -not [bool]$PaperCutBefore.Installed
    $paperCutIsInstalled = [bool]$PaperCutAfter.Installed

    # Derive these reporting fields from the preserved snapshots instead of
    # trusting mutable run flags.
    $driverInstallRequired = $driverWasMissing
    $driverInstalledThisRun = ($driverWasMissing -and $driverIsInstalled -and $script:DriverRemediationAttempted)
    $paperCutInstallRequired = $paperCutWasMissing
    $paperCutInstalledThisRun = ($paperCutWasMissing -and $paperCutIsInstalled -and $script:PaperCutRemediationAttempted)

    $event = [ordered]@{
        event = [ordered]@{ kind='event'; category=@('configuration','package'); type=@('info'); action='maintenance.execution'; outcome=if ($ExitCode -eq 0 -or $ExitCode -eq 3010) {'success'} else {'failure'} }
        maintenance = [ordered]@{
            execution = [ordered]@{
                ComputerName=$env:COMPUTERNAME
                Domain=$env:USERDOMAIN
                ScriptName=$ScriptName
                ScriptVersion=$ScriptVersion
                RunID=$script:RunId
                Status=$Status
                ExitCode=$ExitCode
                StartTime=$script:StartTime.ToUniversalTime().ToString('o')
                EndTime=$end.ToUniversalTime().ToString('o')
                DurationSeconds=[math]::Round(($end-$script:StartTime).TotalSeconds,2)
                WarningCount=$script:WarningCount
                ErrorCount=$script:ErrorCount
                FailureMessage=$script:FailureMessage
                TextLogPath=$PublishedLogPath
                RunAsSystem=(Test-IsSystemAccount)
                UserName=[Security.Principal.WindowsIdentity]::GetCurrent().Name
            }
            printerDeployment = [ordered]@{
                OverallCompliant=($DriverAfter.Installed -and $PaperCutAfter.Installed -and $PrinterAfter.Connected)
                Changes=[ordered]@{
                Driver=[bool]($script:ChangedComponents -contains 'PrinterDriver')
                PaperCut=[bool]($script:ChangedComponents -contains 'PaperCut')
                PrinterConnection=[bool]($script:ChangedComponents -contains 'PrinterConnection')
            }
                Driver=[ordered]@{
                    Name=$PrinterDriverName
                    InstallRequired=[bool]$driverInstallRequired
                    InstalledThisRun=[bool]$driverInstalledThisRun
                    Before=$DriverBefore
                    After=$DriverAfter
                    Remediation=[ordered]@{
                        Required=[bool]$driverInstallRequired
                        Attempted=[bool]$script:DriverRemediationAttempted
                        Succeeded=[bool]$script:DriverRemediationSucceeded
                        ExitCode=$script:DriverRemediationExitCode
                        FailureMessage=$script:DriverRemediationFailureMessage
                    }
                    SourceAvailable=$DriverSourceAvailable
                    SourcePath=$DriverSourcePath
                    SourceFingerprintSHA256=$SourceFingerprint
                    StoredFingerprintSHA256=$StoredFingerprint
                    StatePath=$StatePath
                }
                PaperCut=[ordered]@{
                    InstallRequired=[bool]$paperCutInstallRequired
                    InstalledThisRun=[bool]$paperCutInstalledThisRun
                    Before=$PaperCutBefore
                    After=$PaperCutAfter
                    Remediation=[ordered]@{
                        Required=[bool]$paperCutInstallRequired
                        Attempted=[bool]$script:PaperCutRemediationAttempted
                        Succeeded=[bool]$script:PaperCutRemediationSucceeded
                        ExitCode=$script:PaperCutRemediationExitCode
                        FailureMessage=$script:PaperCutRemediationFailureMessage
                    }
                    SourceAvailable=$PaperCutSourceAvailable
                    MsiPath=$PaperCutMsiPath
                    SourceMsi=$PaperCutSourceMsiInfo
                }
                PrinterConnection=[ordered]@{ SharePath=$PrinterSharePath; Before=$PrinterBefore; After=$PrinterAfter }
                Processes=if ($script:ProcessResults.Count -gt 0) {
                    [object[]]@($script:ProcessResults.ToArray())
                }
                else {
                    [object[]]@()
                }
                Reboot=[ordered]@{
                    Required=[bool]$script:RebootRequired
                    PendingBefore=[bool]$PendingRebootBefore.Pending
                    ReasonNamesBefore=(New-StringArrayForJson -InputObject $PendingRebootBefore.Reasons)
                    PendingAfter=[bool]$PendingRebootAfter.Pending
                    ReasonNamesAfter=(New-StringArrayForJson -InputObject $PendingRebootAfter.Reasons)
                }
                Windows=(Get-WindowsInfo)
            }
        }
        '@timestamp'=$end.ToUniversalTime().ToString('o')
    }
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    try {
        $pd = $event.maintenance.printerDeployment
        Write-Log ("Elastic printer-deployment summary: OverallCompliant={0}; DriverChanged={1}; PaperCutChanged={2}; PrinterChanged={3}; RebootRequired={4}; RebootReasonsBefore={5}; RebootReasonsAfter={6}" -f `
            $pd.OverallCompliant,
            $pd.Changes.Driver,
            $pd.Changes.PaperCut,
            $pd.Changes.PrinterConnection,
            $pd.Reboot.Required,
            (@($pd.Reboot.ReasonNamesBefore) -join ', '),
            (@($pd.Reboot.ReasonNamesAfter) -join ', ')) 'INFO'
    }
    catch { }

    $compact = $event | ConvertTo-Json -Depth 12 -Compress
    Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compact
    $event | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $LatestTelemetryPath -Encoding UTF8 -Force
}

# =========================
# Main
# =========================
New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
Write-Banner
Write-Log "Starting $ScriptName version $ScriptVersion."
Write-Log "Text log: $LogPath"

$driverBefore = [pscustomobject]@{ Installed=$false; Name=$PrinterDriverName; Manufacturer=$null; Version=$null; DriverVersionRaw=$null; InfPath=$null; DriverPath=$null; ConfigFile=$null; DataFile=$null; MajorVersion=$null }
$driverAfter = $driverBefore
$paperCutBefore = [pscustomobject]@{ Installed=$false; DisplayName=$null; DisplayVersion=$null; ProductCode=$null; InstallLocation=$null }
$paperCutAfter = $paperCutBefore
$printerBefore = [pscustomobject]@{ Scope='Unknown'; Connected=$false; Name=$PrinterSharePath }
$printerAfter = $printerBefore
$sourceFingerprint = $null
$storedFingerprint = $null
$driverSourceAvailable = $false
$paperCutSourceAvailable = $false
$paperCutSourceMsiInfo = $null
$pendingBefore = Get-PendingRebootState
$pendingAfter = $pendingBefore
$finalStatus = 'Failed'
$finalExitCode = 1

try {
    if (-not (Test-IsAdministrator)) { throw 'This script must be run as an administrator.' }
    if ([string]::IsNullOrWhiteSpace($PrinterDriverName) -or $PrinterDriverName -like 'CHANGE ME*') { throw 'PrinterDriverName is not configured.' }
    Import-Module PrintManagement -ErrorAction Stop

    $driverBefore = Copy-TelemetrySnapshot -InputObject (Get-InstalledPrinterDriverState)
    $paperCutBefore = Copy-TelemetrySnapshot -InputObject (Get-PaperCutProduct)
    $printerBefore = Copy-TelemetrySnapshot -InputObject (Get-PrinterConnectionState -ConnectionPath $PrinterSharePath)
    $driverSourceAvailable = Test-Path -LiteralPath $DriverSourcePath -PathType Container
    $paperCutSourceAvailable = Test-Path -LiteralPath $PaperCutMsiPath -PathType Leaf
    if ($paperCutSourceAvailable) {
        $paperCutSourceMsiInfo = Get-MsiPackageInfo -Path $PaperCutMsiPath
        if ($paperCutSourceMsiInfo) {
            Write-Log ("PaperCut source MSI: Name={0}; Version={1}; ProductCode={2}; Manufacturer={3}." -f `
                $paperCutSourceMsiInfo.ProductName, $paperCutSourceMsiInfo.ProductVersion, `
                $paperCutSourceMsiInfo.ProductCode, $paperCutSourceMsiInfo.Manufacturer)
        }
    }
    Write-DeploymentStateSummary -Label 'before' -Driver $driverBefore -PaperCut $paperCutBefore -Printer $printerBefore

    Write-Log ("Install requirements from pre-check: SharpDriverRequired={0}; PaperCutRequired={1}." -f `
        (-not $driverBefore.Installed),
        (-not $paperCutBefore.Installed)) 'INFO'
    $state = Get-DeploymentState
    if ($state) {
        $prop = $state.PSObject.Properties['SourceFingerprintSHA256']
        if ($prop) { $storedFingerprint = [string]$prop.Value }
    }

    $script:DriverInstallRequired = (-not $driverBefore.Installed)
    $script:PaperCutInstallRequired = (-not $paperCutBefore.Installed)

    $driverUpdateRequired = -not $driverBefore.Installed
    if ($driverSourceAvailable) {
        $sourceFingerprint = Get-DriverPackageFingerprint -Path $DriverSourcePath
        if ([string]::IsNullOrWhiteSpace($storedFingerprint) -or $storedFingerprint -ne $sourceFingerprint) { $driverUpdateRequired = $true }
    }
    elseif ($driverBefore.Installed) {
        Write-Log "Driver source is unavailable; retaining the already installed driver: $DriverSourcePath" 'WARN'
    }
    else { throw "SHARP driver is missing and the source folder is unavailable: $DriverSourcePath" }

    if ($driverUpdateRequired) {
        $script:DriverRemediationAttempted = $true
        if (-not $driverSourceAvailable) {
            $script:DriverRemediationFailureMessage = 'Driver deployment is required but the source package is unavailable.'
            throw $script:DriverRemediationFailureMessage
        }
        if (Test-Path -LiteralPath $LocalDriverStage) { Remove-Item -LiteralPath $LocalDriverStage -Recurse -Force -ErrorAction Stop }
        New-Item -Path $LocalDriverStage -ItemType Directory -Force | Out-Null
        Write-Log "Copying driver package to local staging: $LocalDriverStage" 'ACTION'
        Copy-Item -Path (Join-Path $DriverSourcePath '*') -Destination $LocalDriverStage -Recurse -Force -ErrorAction Stop
        $infFiles = @(Get-PrinterInfFiles -Path $LocalDriverStage)
        foreach ($inf in $infFiles) {
            $code = Invoke-ExternalProcess -FilePath "$env:SystemRoot\System32\pnputil.exe" -ArgumentList @('/add-driver',"`"$($inf.FullName)`"") -SuccessExitCodes @(0,3010) -TimeoutSeconds $DriverInstallTimeoutSeconds -Operation ("Stage driver {0}" -f $inf.Name)
            $script:DriverRemediationExitCode = $code
            if ($code -eq 3010) { $script:RebootRequired = $true }
        }
        Add-PrinterDriver -Name $PrinterDriverName -ErrorAction Stop

        Write-Log "Waiting for the Windows print subsystem to register the SHARP driver before verification."
        $driverAfter = Copy-TelemetrySnapshot -InputObject (Get-InstalledPrinterDriverState -RetryCount 6 -RetryDelaySeconds 2)

        if (-not $driverAfter.Installed) {
            $script:DriverRemediationFailureMessage = "Driver could not be verified after staging and Add-PrinterDriver."
            $availableSharpDrivers = @(
                Get-PrinterDriver -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '(?i)Sharp|BP-70C31' } |
                ForEach-Object { $_.Name }
            )

            $availableText = if ($availableSharpDrivers.Count -gt 0) {
                $availableSharpDrivers -join '; '
            }
            else {
                'No SHARP/BP-70C31 printer drivers were returned by Get-PrinterDriver.'
            }

            throw "Driver was staged and Add-PrinterDriver completed, but the installed driver could not be verified after retries. Expected='$PrinterDriverName'. VisibleSHARPDrivers='$availableText'"
        }

        $script:DriverRemediationSucceeded = $true
        if ($null -eq $script:DriverRemediationExitCode) { $script:DriverRemediationExitCode = 0 }
        Save-DeploymentState -Fingerprint $sourceFingerprint
        $script:ChangedComponents.Add('PrinterDriver')
        if ($script:DriverInstallRequired) {
            $script:DriverInstalledThisRun = $true
        }

        Write-Log ("Printer driver deployed successfully: RequestedName={0}; VerifiedName={1}; Version={2}; INF={3}" -f `
            $PrinterDriverName,
            $driverAfter.Name,
            $driverAfter.Version,
            $driverAfter.InfPath) 'SUCCESS'
    }
    else {
        $driverAfter = $driverBefore
        Write-Log "Printer driver is already compliant: $PrinterDriverName" 'SUCCESS'
    }

    if (-not $paperCutBefore.Installed) {
        $script:PaperCutRemediationAttempted = $true
        if (-not $paperCutSourceAvailable) {
            $script:PaperCutRemediationFailureMessage = "PaperCut is not installed and the MSI is unavailable: $PaperCutMsiPath"
            throw $script:PaperCutRemediationFailureMessage
        }
        $msiLogPath = Join-Path $LogDirectory 'PaperCut-Print-Deploy-MSI.log'
        $code = Invoke-ExternalProcess -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList @('/i',"`"$PaperCutMsiPath`"",'/qn','/norestart','REBOOT=ReallySuppress','/L*v',"`"$msiLogPath`"") -SuccessExitCodes @(0,3010) -TimeoutSeconds $PaperCutInstallTimeoutSeconds -Operation 'Install PaperCut Print Deploy'
        $script:PaperCutRemediationExitCode = $code
        if ($code -eq 3010) { $script:RebootRequired = $true }
        Start-Sleep -Seconds 3
        $paperCutAfter = Copy-TelemetrySnapshot -InputObject (Get-PaperCutProduct)
        if (-not $paperCutAfter.Installed) {
            $script:PaperCutRemediationFailureMessage = 'PaperCut MSI returned success, but the product could not be verified in the uninstall registry.'
            throw $script:PaperCutRemediationFailureMessage
        }
        $script:PaperCutRemediationSucceeded = $true
        $script:ChangedComponents.Add('PaperCut')
        if ($script:PaperCutInstallRequired -and $paperCutAfter.Installed) {
            $script:PaperCutInstalledThisRun = $true
        }
        Write-Log "PaperCut installed successfully: $($paperCutAfter.DisplayName) $($paperCutAfter.DisplayVersion)" 'SUCCESS'
    }
    else {
        $paperCutAfter = $paperCutBefore
        Write-Log "PaperCut is already installed: $($paperCutBefore.DisplayName) $($paperCutBefore.DisplayVersion)" 'SUCCESS'
        if (-not $paperCutSourceAvailable) { Write-Log "PaperCut MSI source is currently unavailable, but the client is already installed: $PaperCutMsiPath" 'WARN' }
    }

    $printerAfter = Copy-TelemetrySnapshot -InputObject (Ensure-SharedPrinterConnection -ConnectionPath $PrinterSharePath)
    if (-not $driverAfter.Installed -or -not $paperCutAfter.Installed -or -not $printerAfter.Connected) { throw 'Final compliance verification failed.' }
    Write-DeploymentStateSummary -Label 'after' -Driver $driverAfter -PaperCut $paperCutAfter -Printer $printerAfter

    Write-Log ("Change summary: DriverChanged={0}; PaperCutChanged={1}; PrinterConnectionChanged={2}." -f `
        ($script:ChangedComponents -contains 'PrinterDriver'),
        ($script:ChangedComponents -contains 'PaperCut'),
        ($script:ChangedComponents -contains 'PrinterConnection')) 'INFO'

    Write-Log ("Remediation summary: Sharp Required={0}, Attempted={1}, Succeeded={2}, InstalledThisRun={3}; PaperCut Required={4}, Attempted={5}, Succeeded={6}, InstalledThisRun={7}." -f `
        (-not [bool]$driverBefore.Installed),
        $script:DriverRemediationAttempted,
        $script:DriverRemediationSucceeded,
        ((-not [bool]$driverBefore.Installed) -and [bool]$driverAfter.Installed -and $script:DriverRemediationAttempted),
        (-not [bool]$paperCutBefore.Installed),
        $script:PaperCutRemediationAttempted,
        $script:PaperCutRemediationSucceeded,
        ((-not [bool]$paperCutBefore.Installed) -and [bool]$paperCutAfter.Installed -and $script:PaperCutRemediationAttempted)) 'INFO'

    $pendingAfter = Get-PendingRebootState
    if ($pendingAfter.Pending) { $script:RebootRequired = $true }
    if ($script:RebootRequired) { $finalStatus='SuccessRebootRequired'; $finalExitCode=3010 }
    elseif ($script:WarningCount -gt 0) { $finalStatus='SuccessWithWarnings'; $finalExitCode=0 }
    elseif ($script:ChangedComponents.Count -eq 0) { $finalStatus='AlreadyCompliant'; $finalExitCode=0 }
    else { $finalStatus='Success'; $finalExitCode=0 }
    Write-Log "Final status: $finalStatus" 'SUCCESS'
}
catch {
    $script:FailureMessage = $_.Exception.Message

    if ($script:DriverRemediationAttempted -and -not $script:DriverRemediationSucceeded -and [string]::IsNullOrWhiteSpace($script:DriverRemediationFailureMessage)) {
        $script:DriverRemediationFailureMessage = $script:FailureMessage
    }
    if ($script:PaperCutRemediationAttempted -and -not $script:PaperCutRemediationSucceeded -and [string]::IsNullOrWhiteSpace($script:PaperCutRemediationFailureMessage)) {
        $script:PaperCutRemediationFailureMessage = $script:FailureMessage
    }

    Write-Log $script:FailureMessage 'ERROR'
    $finalStatus='Failed'
    $finalExitCode=1
    try { $driverAfter=Copy-TelemetrySnapshot -InputObject (Get-InstalledPrinterDriverState) } catch { }
    try { $paperCutAfter=Copy-TelemetrySnapshot -InputObject (Get-PaperCutProduct) } catch { }
    try { $printerAfter=Copy-TelemetrySnapshot -InputObject (Get-PrinterConnectionState -ConnectionPath $PrinterSharePath) } catch { }
    try { $pendingAfter=Get-PendingRebootState } catch { }
}
finally {
    try {
        Write-Telemetry -Status $finalStatus -ExitCode $finalExitCode -DriverBefore $driverBefore -DriverAfter $driverAfter -PaperCutBefore $paperCutBefore -PaperCutAfter $paperCutAfter -PrinterBefore $printerBefore -PrinterAfter $printerAfter -DriverSourceAvailable $driverSourceAvailable -PaperCutSourceAvailable $paperCutSourceAvailable -SourceFingerprint $sourceFingerprint -StoredFingerprint $storedFingerprint -PaperCutSourceMsiInfo $paperCutSourceMsiInfo -PendingRebootBefore $pendingBefore -PendingRebootAfter $pendingAfter
        Write-Log "Telemetry written to $LatestTelemetryPath"
    }
    catch {
        Write-Log "Telemetry write failed: $($_.Exception.Message)" 'ERROR'
        if ($finalExitCode -eq 0 -or $finalExitCode -eq 3010) {
            $finalStatus = 'TelemetryFailure'
            $finalExitCode = 4
        }
    }

    # Final text-log append before the immutable completed file enters C:\Logs.
    Write-Log ("Completed {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
        $ScriptName,
        $finalStatus,
        $finalExitCode,
        $script:WarningCount,
        $script:ErrorCount) $(if ($finalExitCode -eq 0 -or $finalExitCode -eq 3010) { 'SUCCESS' } else { 'ERROR' })

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 11 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 11 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $finalExitCode
'@
    }
    AutologonAndEdge = [ordered]@{
        FileName = '13_Configure_Autologon_And_Edge.ps1'
        # Plain-text source is intentionally embedded for maintainability.
        Source = @'
# =====================================================================
# ScriptName: 13_Configure_Autologon_And_Edge.ps1
# ScriptVersion: 2.2.0
# LastUpdated: 2026-08-27
# Changes: v2.2.0 removes ForceAutoLogon so Log off and Switch user remain available for administrator sign-in,
#          while preserving normal CC-Student automatic sign-in at computer startup.
#          v2.1.3 normalizes Windows 11 product naming, makes telemetry collections explicitly JSON-array safe,
#          and adds a concise Elastic configuration summary while preserving password redaction.
#          v2.1.2 uses Maintenance.Framework v2.4 staged text logging.
# Purpose: Configure lab autologon and launch Microsoft Edge InPrivate
#          for approved computer-name patterns, with verification and
#          structured maintenance telemetry. The autologon password is Base64-obfuscated in the script and decoded only at runtime.
# =====================================================================

[CmdletBinding()]
param(
    [string[]]$ComputerNamePatterns = @(
        'SSB-122-*',
        'SSB-114*' ,
	'SSB-171*'
    ),

    [string]$DefaultUserName = 'CC-Student',

    [ValidateNotNullOrEmpty()]
    [string]$DefaultPasswordBase64 = 'Q0MkdHVkM250IQ==',

    [string]$DefaultDomainName = 'Compton.edu',

    [ValidateNotNullOrEmpty()]
    [string]$EdgeUrl = 'https://www.compton.edu',

    [string]$LogDirectory = 'C:\Logs',

    [switch]$AllowHttpEdgeUrl
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptName = '13_Configure_Autologon_And_Edge.ps1'
$ScriptVersion = '2.2.0'
$RunId = [guid]::NewGuid().Guid
$StartTime = Get-Date
$ComputerName = $env:COMPUTERNAME
$DomainName = $env:USERDOMAIN
$RunningAccount = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$LogPath = $null
$PublishedLogPath = $null
$LogSession = $null
$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestTelemetryPath = Join-Path $LogDirectory '13_Configure_Autologon_And_Edge.latest.json'
$WinlogonPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
$RunKeyPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$EdgeRunValueName = 'LaunchComptonEdge'
$AllUsersStartupPath = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup'
$LegacyChromeShortcutPath = Join-Path $AllUsersStartupPath 'Google Chrome.lnk'

$script:WarningCount = 0
$script:ErrorCount = 0
$script:ExitCode = 0
$script:FinalStatus = 'Success'
$script:OverallResult = 'Unknown'
$script:FailureMessage = $null
$script:Targeted = $false
$script:MatchedPattern = $null
$script:ChangesMade = 0
$script:VerificationFailures = New-Object System.Collections.Generic.List[string]
$script:RegistryChanges = New-Object System.Collections.Generic.List[object]
$script:EdgePath = $null
$script:EdgeCommand = $null
$script:LegacyShortcutRemoved = $false
$script:LegacyShortcutPresentBefore = $false
$script:LegacyShortcutPresentAfter = $false
$script:BeforeState = $null
$script:AfterState = $null
$script:DecodedPassword = $null

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 13 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '13_Configure_Autologon_And_Edge.log',
        '*-13_Configure_Autologon_And_Edge-*.log'
    ) | Out-Null

$LogSession = New-MaintenanceStagedLog `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $ComputerName `
    -Timestamp $StartTime

$LogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

function Initialize-LogDirectory {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $normalizedLevel = switch ($Level) {
        'OK'   { 'SUCCESS' }
        'WARN' { 'WARNING' }
        default { $Level }
    }
    $logComputerName = if ($ComputerName) { $ComputerName } else { 'UNKNOWN' }
    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $logComputerName, $normalizedLevel, $Message

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }

    try {
        $activeLogDirectory = Split-Path -Parent $LogPath
        if (-not (Test-Path -LiteralPath $activeLogDirectory -PathType Container)) {
            New-Item -Path $activeLogDirectory -ItemType Directory -Force | Out-Null
        }
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Unable to write to the log file: $($_.Exception.Message)"
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-WindowsInformation {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $currentVersion = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop

        $productName = [string]$currentVersion.ProductName
        $buildText = [string]$os.BuildNumber
        $buildNumber = 0
        [void][int]::TryParse($buildText, [ref]$buildNumber)

        if ($buildNumber -ge 22000 -and $productName -match '^Windows 10') {
            $productName = $productName -replace '^Windows 10', 'Windows 11'
        }

        [pscustomobject]@{
            ProductName    = $productName
            EditionId      = [string]$currentVersion.EditionID
            DisplayVersion = [string]$currentVersion.DisplayVersion
            Version        = [string]$os.Version
            BuildNumber    = $buildText
            UBR            = [int]$currentVersion.UBR
            FullBuild      = '{0}.{1}' -f $buildText, $currentVersion.UBR
        }
    }
    catch {
        Write-Log "Unable to collect Windows version information: $($_.Exception.Message)" 'WARN'
        [pscustomobject]@{
            ProductName = $null; EditionId = $null; DisplayVersion = $null
            Version = $null; BuildNumber = $null; UBR = $null; FullBuild = $null
        }
    }
}

function Get-MatchedComputerPattern {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $Name -like $pattern) {
            return $pattern
        }
    }
    return $null
}

function Get-RegistryStringValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $value = (Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop).$Name
        return [string]$value
    }
    catch {
        return $null
    }
}

function Set-VerifiedRegistryString {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value,
        [switch]$Sensitive
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    $before = Get-RegistryStringValue -Path $Path -Name $Name
    $changed = ([string]$before -cne [string]$Value)

    if ($changed) {
        New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
        $script:ChangesMade++
        if ($Sensitive) {
            Write-Log "Updated protected registry value: $Path\$Name" 'OK'
        }
        else {
            Write-Log "Updated registry value: $Path\$Name" 'OK'
        }
    }
    else {
        Write-Log "Registry value is already correct: $Path\$Name"
    }

    $after = Get-RegistryStringValue -Path $Path -Name $Name
    $verified = ([string]$after -ceq [string]$Value)

    if (-not $verified) {
        $script:VerificationFailures.Add("Registry verification failed: $Path\$Name")
        throw "Registry value verification failed: $Path\$Name"
    }

    $script:RegistryChanges.Add([pscustomobject]@{
        Path          = $Path
        Name          = $Name
        Changed       = $changed
        Verified      = $verified
        Sensitive     = [bool]$Sensitive
        BeforePresent = ($null -ne $before)
        AfterPresent  = ($null -ne $after)
        BeforeValue   = $(if ($Sensitive) { $null } else { $before })
        AfterValue    = $(if ($Sensitive) { $null } else { $after })
    })
}

function Remove-VerifiedRegistryValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    $before = Get-RegistryStringValue -Path $Path -Name $Name
    $changed = ($null -ne $before)

    if ($changed) {
        Remove-ItemProperty -LiteralPath $Path -Name $Name -Force -ErrorAction Stop
        $script:ChangesMade++
        Write-Log "Removed registry value: $Path\$Name" 'OK'
    }
    else {
        Write-Log "Registry value is already absent: $Path\$Name"
    }

    $after = Get-RegistryStringValue -Path $Path -Name $Name
    $verified = ($null -eq $after)

    if (-not $verified) {
        $script:VerificationFailures.Add("Registry removal verification failed: $Path\$Name")
        throw "Registry value removal verification failed: $Path\$Name"
    }

    $script:RegistryChanges.Add([pscustomobject]@{
        Path          = $Path
        Name          = $Name
        Changed       = $changed
        Verified      = $verified
        Sensitive     = $false
        BeforePresent = ($null -ne $before)
        AfterPresent  = ($null -ne $after)
        BeforeValue   = $before
        AfterValue    = $after
    })
}

function Get-DecodedAutologonPassword {
    try {
        $bytes = [Convert]::FromBase64String($DefaultPasswordBase64)
        $password = [Text.Encoding]::UTF8.GetString($bytes)

        if ([string]::IsNullOrWhiteSpace($password)) {
            throw 'The decoded autologon password is blank.'
        }

        return $password
    }
    catch {
        throw "Unable to decode the configured autologon password: $($_.Exception.Message)"
    }
    finally {
        $bytes = $null
    }
}

function Test-ConfigurationInput {
    if (-not $ComputerNamePatterns -or @($ComputerNamePatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0) {
        throw 'At least one nonblank computer-name pattern is required.'
    }
    if ([string]::IsNullOrWhiteSpace($DefaultUserName)) {
        throw 'DefaultUserName cannot be blank.'
    }
    if ([string]::IsNullOrWhiteSpace($DefaultPasswordBase64)) {
        throw 'DefaultPasswordBase64 cannot be blank when autologon is enabled.'
    }

    try {
        $decodedBytes = [Convert]::FromBase64String($DefaultPasswordBase64)
        if ($decodedBytes.Length -eq 0) {
            throw 'The decoded password is empty.'
        }
    }
    catch {
        throw 'DefaultPasswordBase64 is not valid Base64 or decodes to an empty value.'
    }
    finally {
        $decodedBytes = $null
    }
    if ([string]::IsNullOrWhiteSpace($DefaultDomainName)) {
        throw 'DefaultDomainName cannot be blank.'
    }

    $uri = $null
    if (-not [uri]::TryCreate($EdgeUrl, [UriKind]::Absolute, [ref]$uri)) {
        throw "EdgeUrl is not a valid absolute URL: $EdgeUrl"
    }
    if ($uri.Scheme -notin @('https','http')) {
        throw "EdgeUrl must use HTTP or HTTPS: $EdgeUrl"
    }
    if ($uri.Scheme -eq 'http' -and -not $AllowHttpEdgeUrl) {
        throw 'EdgeUrl uses HTTP. Use HTTPS or explicitly supply -AllowHttpEdgeUrl.'
    }
}

function Get-AutologonState {
    $autoAdmin = Get-RegistryStringValue -Path $WinlogonPath -Name 'AutoAdminLogon'
    $user = Get-RegistryStringValue -Path $WinlogonPath -Name 'DefaultUserName'
    $domain = Get-RegistryStringValue -Path $WinlogonPath -Name 'DefaultDomainName'
    $force = Get-RegistryStringValue -Path $WinlogonPath -Name 'ForceAutoLogon'
    $password = Get-RegistryStringValue -Path $WinlogonPath -Name 'DefaultPassword'

    [pscustomobject]@{
        AutoAdminLogonEnabled = ($autoAdmin -eq '1')
        DefaultUserName       = $user
        DefaultDomainName     = $domain
        ForceAutoLogonPresent = ($null -ne $force)
        ForceAutoLogonEnabled = ($force -eq '1')
        PasswordPresent       = (-not [string]::IsNullOrEmpty($password))
        PasswordMatches       = ($null -ne $script:DecodedPassword -and [string]$password -ceq [string]$script:DecodedPassword)
        ConfigurationMatches  = (
            $autoAdmin -eq '1' -and
            $user -ceq $DefaultUserName -and
            $domain -ceq $DefaultDomainName -and
            $null -eq $force -and
            [string]$password -ceq [string]$script:DecodedPassword
        )
    }
}

function Set-AutologonConfiguration {
    Write-Log "Ensuring Windows autologon is configured for '$DefaultDomainName\$DefaultUserName'."

    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'AutoAdminLogon' -Value '1'
    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'DefaultUserName' -Value $DefaultUserName
    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'DefaultPassword' -Value $script:DecodedPassword -Sensitive
    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'DefaultDomainName' -Value $DefaultDomainName
    # ForceAutoLogon=1 immediately signs the default user back in after Log off.
    # Keep standard startup autologon, but remove this override so an administrator can sign in.
    Remove-VerifiedRegistryValue -Path $WinlogonPath -Name 'ForceAutoLogon'

    $state = Get-AutologonState
    if (-not $state.ConfigurationMatches) {
        throw 'Autologon registry values did not pass final verification.'
    }
    Write-Log "Autologon configuration verified for '$DefaultDomainName\$DefaultUserName'." 'OK'
}

function Remove-LegacyChromeStartupShortcut {
    $script:LegacyShortcutPresentBefore = Test-Path -LiteralPath $LegacyChromeShortcutPath -PathType Leaf
    Write-Log 'Checking for the legacy All Users Google Chrome startup shortcut.'

    if ($script:LegacyShortcutPresentBefore) {
        Remove-Item -LiteralPath $LegacyChromeShortcutPath -Force -ErrorAction Stop
        $script:ChangesMade++
        $script:LegacyShortcutRemoved = $true
        Write-Log "Removed legacy Chrome startup shortcut: $LegacyChromeShortcutPath" 'OK'
    }
    else {
        Write-Log 'Legacy Google Chrome startup shortcut is not present.'
    }

    $script:LegacyShortcutPresentAfter = Test-Path -LiteralPath $LegacyChromeShortcutPath -PathType Leaf
    if ($script:LegacyShortcutPresentAfter) {
        $script:VerificationFailures.Add('Legacy Chrome startup shortcut remains present.')
        throw "The legacy Chrome startup shortcut could not be removed: $LegacyChromeShortcutPath"
    }
}

function Get-EdgePath {
    $candidatePaths = @()
    if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
        $candidatePaths += (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe')
    }
    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
        $candidatePaths += (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    }

    return $candidatePaths |
        Select-Object -Unique |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1
}

function Get-EdgeRunState {
    $configuredCommand = Get-RegistryStringValue -Path $RunKeyPath -Name $EdgeRunValueName
    $edgePath = Get-EdgePath
    $edgeVersion = $null
    if ($edgePath) {
        try { $edgeVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($edgePath).ProductVersion } catch { }
    }
    [pscustomobject]@{
        EdgeInstalled      = [bool]$edgePath
        EdgePath           = $edgePath
        EdgeVersion        = $edgeVersion
        RunValuePresent    = (-not [string]::IsNullOrWhiteSpace($configuredCommand))
        RunCommand         = $configuredCommand
        CommandMatches     = ($null -ne $script:EdgeCommand -and $configuredCommand -ceq $script:EdgeCommand)
        UrlPresent         = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand.Contains($EdgeUrl))
        InPrivateEnabled   = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand -match '(?i)--inprivate(?:\s|$)')
        NewWindowEnabled   = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand -match '(?i)--new-window(?:\s|$)')
        MaximizedEnabled   = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand -match '(?i)--start-maximized(?:\s|$)')
    }
}

function Write-ConfigurationStateSummary {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)]$State
    )

    Write-Log ("Autologon {0}: Enabled={1}; User={2}; Domain={3}; ForceAutoLogonPresent={4}; ForceAutoLogonEnabled={5}; PasswordPresent={6}; PasswordMatches={7}; ConfigurationMatches={8}." -f `
        $Label, $State.Autologon.AutoAdminLogonEnabled, $State.Autologon.DefaultUserName, `
        $State.Autologon.DefaultDomainName, $State.Autologon.ForceAutoLogonPresent, $State.Autologon.ForceAutoLogonEnabled, `
        $State.Autologon.PasswordPresent, $State.Autologon.PasswordMatches, $State.Autologon.ConfigurationMatches)

    Write-Log ("Edge startup {0}: Installed={1}; Version={2}; RunValuePresent={3}; CommandMatches={4}; UrlPresent={5}; InPrivate={6}; NewWindow={7}; Maximized={8}." -f `
        $Label, $State.Edge.EdgeInstalled, $State.Edge.EdgeVersion, $State.Edge.RunValuePresent, `
        $State.Edge.CommandMatches, $State.Edge.UrlPresent, $State.Edge.InPrivateEnabled, `
        $State.Edge.NewWindowEnabled, $State.Edge.MaximizedEnabled)
}

function Set-EdgeAutoLaunch {
    $script:EdgePath = Get-EdgePath
    if (-not $script:EdgePath) {
        throw 'Microsoft Edge was not found in either Program Files location.'
    }

    $escapedUrl = $EdgeUrl.Replace('"','')
    $script:EdgeCommand = '"{0}" --inprivate --new-window --start-maximized "{1}"' -f $script:EdgePath, $escapedUrl
    Set-VerifiedRegistryString -Path $RunKeyPath -Name $EdgeRunValueName -Value $script:EdgeCommand

    $edgeState = Get-EdgeRunState
    if (-not $edgeState.CommandMatches -or -not $edgeState.EdgeInstalled) {
        $script:VerificationFailures.Add('Edge auto-launch configuration verification failed.')
        throw 'Edge auto-launch configuration did not pass final verification.'
    }

    Write-Log "Configured and verified Edge auto-launch for every user: $EdgeUrl" 'OK'
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
        ForEach-Object { $_ }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Write-Telemetry {
    param([Parameter(Mandatory)][object]$WindowsInfo)

    try {
        Initialize-LogDirectory
        $endTime = Get-Date
        $duration = [math]::Round(($endTime - $StartTime).TotalSeconds, 3)

        $event = [ordered]@{
            '@timestamp'             = $endTime.ToUniversalTime().ToString('o')
            EventType                = 'maintenance.execution'
            ComputerName             = $ComputerName
            Domain                   = $DomainName
            ScriptName               = $ScriptName
            ScriptVersion            = $ScriptVersion
            RunId                    = $RunId
            StartTime                = $StartTime.ToUniversalTime().ToString('o')
            EndTime                  = $endTime.ToUniversalTime().ToString('o')
            DurationSeconds          = $duration
            Status                   = $script:FinalStatus
            OverallResult            = $script:OverallResult
            ExitCode                 = $script:ExitCode
            ErrorCount               = $script:ErrorCount
            WarningCount             = $script:WarningCount
            FailureMessage           = $script:FailureMessage
            TextLogPath              = $PublishedLogPath
            RunningAccount           = $RunningAccount
            RunningAsSystem          = ($RunningAccount -eq 'NT AUTHORITY\SYSTEM')
            IsAdministrator          = (Test-IsAdministrator)
            Targeted                 = $script:Targeted
            MatchedPattern           = $script:MatchedPattern
            ConfiguredPatterns       = New-StringArrayForJson -InputObject $ComputerNamePatterns
            ChangesMade              = $script:ChangesMade
            VerificationFailureCount = $script:VerificationFailures.Count
            VerificationFailures     = New-StringArrayForJson -InputObject $script:VerificationFailures
            Windows                  = $WindowsInfo
            Configuration            = [ordered]@{
                DefaultUserName   = $DefaultUserName
                DefaultDomainName = $DefaultDomainName
                PasswordProvided  = (-not [string]::IsNullOrWhiteSpace($DefaultPasswordBase64))
                PasswordStorage   = 'Base64ObfuscatedInScript'
                PasswordValueLogged = $false
                ForceAutoLogonDesired = $false
                EdgeUrl           = $EdgeUrl
                EdgeRunValueName  = $EdgeRunValueName
            }
            Before                   = $script:BeforeState
            After                    = $script:AfterState
            RegistryChanges          = New-ObjectArrayForJson -InputObject $script:RegistryChanges
            LegacyChromeShortcut     = [ordered]@{
                Path          = $LegacyChromeShortcutPath
                PresentBefore = $script:LegacyShortcutPresentBefore
                Removed       = $script:LegacyShortcutRemoved
                PresentAfter  = $script:LegacyShortcutPresentAfter
            }
            Edge                     = [ordered]@{
                ExecutablePath = $script:EdgePath
                RunCommand     = $script:EdgeCommand
            }
        }

        try {
            $autologonVerified = $false
            $edgeVerified = $false
            if ($event.After) {
                if ($event.After.Autologon) { $autologonVerified = [bool]$event.After.Autologon.ConfigurationMatches }
                if ($event.After.Edge) { $edgeVerified = [bool]$event.After.Edge.CommandMatches }
            }

            Write-Log ("Elastic configuration summary: Targeted={0}; MatchedPattern={1}; ChangesMade={2}; VerificationFailures={3}; AutologonVerified={4}; EdgeVerified={5}; LegacyChromeRemoved={6}; Result={7}" -f `
                $event.Targeted,
                $event.MatchedPattern,
                $event.ChangesMade,
                $event.VerificationFailureCount,
                $autologonVerified,
                $edgeVerified,
                $event.LegacyChromeShortcut.Removed,
                $event.OverallResult) 'INFO'
        }
        catch { }

        $compactJson = $event | ConvertTo-Json -Depth 10 -Compress
        Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compactJson

        $prettyJson = $event | ConvertTo-Json -Depth 10
        $tempPath = "$LatestTelemetryPath.tmp.$RunId"
        Set-Content -LiteralPath $tempPath -Value $prettyJson -Encoding UTF8
        Move-Item -LiteralPath $tempPath -Destination $LatestTelemetryPath -Force
    }
    catch {
        Write-Warning "Unable to write structured telemetry: $($_.Exception.Message)"
    }
}

$windowsInfo = Get-WindowsInformation

try {
    Initialize-LogDirectory
    Write-Log "===== Autologon and Edge configuration v$ScriptVersion started ====="
    Write-Log "Text log: $LogPath"
    Write-Log "Computer name: $ComputerName"
    Write-Log "Configured computer-name patterns: $($ComputerNamePatterns -join ', ')"

    if (-not (Test-IsAdministrator)) {
        throw 'Please run this script as Administrator or through a SYSTEM scheduled task.'
    }

    Test-ConfigurationInput

    $script:MatchedPattern = Get-MatchedComputerPattern -Name $ComputerName -Patterns $ComputerNamePatterns
    $script:Targeted = (-not [string]::IsNullOrWhiteSpace($script:MatchedPattern))

    if (-not $script:Targeted) {
        $script:BeforeState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
        $script:AfterState = $script:BeforeState
        Write-ConfigurationStateSummary -Label 'unchanged (not targeted)' -State $script:BeforeState
        $script:OverallResult = 'NotTargeted'
        $script:FinalStatus = 'Success'
        Write-Log "Computer '$ComputerName' does not match the configured pattern list. No changes were made."
        Write-Log '===== Script completed: computer not targeted =====' 'OK'
    }
    else {
        Write-Log "Computer '$ComputerName' matches pattern '$($script:MatchedPattern)'." 'OK'

        $script:DecodedPassword = Get-DecodedAutologonPassword

        $script:BeforeState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
        Write-ConfigurationStateSummary -Label 'before' -State $script:BeforeState

        Set-AutologonConfiguration
        Remove-LegacyChromeStartupShortcut
        Set-EdgeAutoLaunch

        $script:AfterState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
        Write-ConfigurationStateSummary -Label 'after' -State $script:AfterState
        Write-Log ("Legacy Chrome startup shortcut: PresentBefore={0}; Removed={1}; PresentAfter={2}." -f `
            $script:LegacyShortcutPresentBefore, $script:LegacyShortcutRemoved, $script:LegacyShortcutPresentAfter)

        if (-not $script:AfterState.Autologon.ConfigurationMatches) {
            $script:VerificationFailures.Add('Final autologon configuration does not match the requested values.')
        }
        if (-not $script:AfterState.Edge.CommandMatches) {
            $script:VerificationFailures.Add('Final Edge Run value does not match the requested command.')
        }
        if ($script:LegacyShortcutPresentAfter) {
            $script:VerificationFailures.Add('Legacy Chrome startup shortcut remains present.')
        }

        if ($script:VerificationFailures.Count -gt 0) {
            throw "One or more final verification checks failed: $($script:VerificationFailures -join '; ')"
        }

        if ($script:ChangesMade -eq 0) {
            $script:OverallResult = 'AlreadyCompliant'
        }
        else {
            $script:OverallResult = 'ConfiguredAndVerified'
        }
        $script:FinalStatus = 'Success'
        Write-Log 'Autologon, legacy Chrome cleanup, and Edge startup settings were verified successfully.' 'OK'
        Write-Log '===== Script completed successfully =====' 'OK'
    }
}
catch {
    $script:ExitCode = 1
    $script:FinalStatus = 'Failed'
    $script:OverallResult = 'Failed'
    $script:FailureMessage = $_.Exception.Message
    Write-Log "Configuration failed: $($_.Exception.Message)" 'ERROR'

    try {
        $script:AfterState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
    }
    catch {
        Write-Log "Unable to collect final configuration state: $($_.Exception.Message)" 'WARN'
    }

    Write-Log '===== Script completed with errors =====' 'ERROR'
}
finally {
    # Clear the decoded credential before final telemetry/log publication.
    $script:DecodedPassword = $null
    [GC]::Collect()

    try {
        Write-Telemetry -WindowsInfo $windowsInfo
    }
    catch {
        Write-Log "Telemetry write failed: $($_.Exception.Message)" 'ERROR'
        if ($script:ExitCode -eq 0) {
            $script:FinalStatus = 'TelemetryFailure'
            $script:OverallResult = 'TelemetryFailure'
            $script:ExitCode = 4
        }
    }

    # Final text-log append before the immutable completed file enters C:\Logs.
    Write-Log ("Finalizing {0}. Status={1}; Result={2}; ExitCode={3}; Warnings={4}; Errors={5}" -f `
        $ScriptName,
        $script:FinalStatus,
        $script:OverallResult,
        $script:ExitCode,
        $script:WarningCount,
        $script:ErrorCount) $(if ($script:ExitCode -eq 0) { 'OK' } else { 'ERROR' })

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 13 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 13 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $script:ExitCode
'@
    }
    ElasticAgent = [ordered]@{
        FileName = '15_Install_Elastic_Agent.ps1'
        # Plain-text source is intentionally embedded for maintainability.
        Source = @'
#requires -Version 5.1
#requires -RunAsAdministrator
# ScriptVersion: 2.1.5

<#
.SYNOPSIS
    Installs and enrolls Elastic Agent only on approved lab computers using
    the configured Fleet Server and enrollment token.

.DESCRIPTION
    - Uses a configurable computer-name prefix allowlist.
    - Defaults to the IB1-103 lab prefix.
    - Installs Elastic Agent from the internal ZIP package, with automatic official Elastic Internet fallback.
    - Uses the current Fleet Server URL and enrollment token from Install.ps1.
    - Performs a fully unattended Fleet installation with no Y/N prompt.
    - Leaves the Elastic Agent service installed and running under Fleet management.
    - Is safe to run repeatedly; already-installed systems are health-checked and skipped when present.

.NOTES
    Script:  15_Install_Elastic_Agent.ps1
    Version: 2.1.5
    LastUpdated: 2026-08-19
    Changes: v2.1.5 adds structured, mapping-safe Elastic Agent deployment telemetry,
             explicit result/failure-stage reporting, package-source/hash/enrollment state,
             existing/final agent health snapshots, and enrollment-token redaction.
             v2.1.4 uses Maintenance.Framework v2.4 staged text logging.
    Requires: Windows PowerShell 5.1, administrative/SYSTEM context
#>

[CmdletBinding()]
param(
    [switch]$ForceReinstall,

    [string[]]$AllowedComputerPrefixes = @(
        'IB1-103'
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '2.1.5'

# ============================================================================
# CONFIGURATION
# ============================================================================

# Keep this version aligned with the ZIP placed on the deployment share.
[string]$ElasticAgentVersion = '9.5.0'

# Fleet enrollment settings copied from the current Install.ps1.
# NOTE: The enrollment token is sensitive. Restrict read access to this script/share.
[string]$FleetServerUrl   = 'https://10.2.12.4:8220'
[string]$EnrollmentToken  = 'X2lseThaOEJ2UFRmNGxFS0k1aEg6dUI2VVFOVGRnWHhmeHI5S05uWDd2dw=='

# Fleet Server is currently using the Quick Start self-signed TLS certificate.
# Keep this $true until Fleet Server is moved to a certificate trusted by the lab PCs.
[bool]$UseInsecureFleetTls = $true

# Preferred and fallback ZIP locations.
[string]$PreferredInstallerPath = "\\filesvr\Labscripts\ElasticAgent\elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"
[string]$FallbackInstallerPath  = "\\10.2.3.30\Labscripts\ElasticAgent\elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"

# Official Elastic download fallback. Used automatically when both internal shares are unavailable.
[string]$DownloadUri = "https://artifacts.elastic.co/downloads/beats/elastic-agent/elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"

# Optional SHA-512 value from Elastic's matching .sha512 file.
# Leave blank to skip package-hash enforcement.
[string]$ExpectedSHA512 = ''

[string]$ElasticServiceName = 'Elastic Agent'
[string]$InstalledAgentPath = 'C:\Program Files\Elastic\Agent\elastic-agent.exe'
[string]$WorkingRoot        = 'C:\ProgramData\Compton\ElasticAgentInstall'
[string]$LogDirectory       = 'C:\Logs'
[string]$LogPath            = $null
[string]$PublishedLogPath   = $null
$LogSession                 = $null
[int]$MaximumLogSizeMB      = 10
[int]$LogRetentionDays      = 60

[string]$TelemetryPath       = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
[string]$LatestTelemetryPath = Join-Path $LogDirectory '15_Install_Elastic_Agent.latest.json'

$script:StartTime    = Get-Date
$script:WarningCount = 0
$script:ErrorCount   = 0
$script:CurrentStage = 'Initialization'
$script:FailureStage = $null
$script:FailureMessage = $null
$script:MatchedPrefix = $null
$script:ExistingHealth = $null
$script:FinalHealth = $null
$script:PackageSource = $null
$script:PackageHashEnforced = -not [string]::IsNullOrWhiteSpace($ExpectedSHA512)
$script:PackageHashVerified = $false
$script:InstallationAttempted = $false
$script:EnrollmentAttempted = $false
$script:InstallerExitCode = $null
$script:VersionResponse = $null
$script:ExistingHealthRestartAttempted = $false
$script:ExistingHealthRestored = $false
$script:ReinstallAttempted = $false
$script:OverallResult = 'Unknown'

$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 15 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName '15_Install_Elastic_Agent.ps1' `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '15_Install_Elastic_Agent.log',
        '*-15_Install_Elastic_Agent-*.log'
    ) | Out-Null

$LogSession = New-MaintenanceStagedLog `
    -ScriptName '15_Install_Elastic_Agent.ps1' `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $env:COMPUTERNAME `
    -Timestamp (Get-Date)

$LogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

# ============================================================================
# FUNCTIONS
# ============================================================================

function Initialize-Directory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Invoke-LogMaintenance {
    [CmdletBinding()]
    param()

    Initialize-Directory -Path $LogDirectory
    # Current and legacy script 15 logs were archived before the staged log was created.
    # Retention of C:\Logs\Old Logs is managed centrally by Maintenance.Framework.
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )

    if ($Level -eq 'WARNING') { $script:WarningCount++ }
    elseif ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:COMPUTERNAME, $Level, $Message
    $activeLogDirectory = Split-Path -Parent $LogPath
    Initialize-Directory -Path $activeLogDirectory
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    Write-Host $line
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

function Convert-AgentHealthForTelemetry {
    [CmdletBinding()]
    param([AllowNull()]$Health)

    if ($null -eq $Health) {
        return $null
    }

    [ordered]@{
        Installed      = [bool]$Health.Installed
        ServiceRunning = [bool]$Health.ServiceRunning
        ServiceStatus  = $Health.ServiceStatus
        StatusHealthy  = [bool]$Health.StatusHealthy
        Healthy        = [bool]$Health.Healthy
        StatusExitCode = $Health.StatusExitCode
        StatusOutput   = New-StringArrayForJson -InputObject $Health.StatusOutput
        Message        = $Health.Message
    }
}

function Write-JsonAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Json
    )

    $temporary = '{0}.{1}.tmp' -f $Path, ([guid]::NewGuid().Guid)
    [IO.File]::WriteAllText($temporary, $Json, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Write-ElasticAgentTelemetry {
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ExitCode)

    $eventTime = Get-Date
    $domain = $env:USERDOMAIN
    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.Domain) { $domain = [string]$computerSystem.Domain }
    }
    catch { }

    $status = if ($ExitCode -ne 0) {
        'Failed'
    }
    elseif ($script:OverallResult -eq 'ExistingUnhealthyNeedsInvestigation' -or
            $script:OverallResult -eq 'InstalledButUnhealthy') {
        'SuccessWithWarnings'
    }
    else {
        'Success'
    }

    $event = [ordered]@{
        '@timestamp'   = $eventTime.ToUniversalTime().ToString('o')
        EventType      = 'maintenance.elastic_agent'
        SchemaVersion  = '1.0'
        ComputerName   = $env:COMPUTERNAME
        Domain         = $domain
        ScriptName     = '15_Install_Elastic_Agent.ps1'
        ScriptVersion  = $ScriptVersion
        Status         = $status
        OverallResult  = $script:OverallResult
        ExitCode       = $ExitCode
        StartTime      = $script:StartTime.ToUniversalTime().ToString('o')
        EndTime        = $eventTime.ToUniversalTime().ToString('o')
        DurationSeconds = [math]::Round(($eventTime - $script:StartTime).TotalSeconds, 3)
        WarningCount   = $script:WarningCount
        ErrorCount     = $script:ErrorCount
        FailureStage   = $script:FailureStage
        FailureMessage = $script:FailureMessage
        TextLogPath    = $PublishedLogPath

        Targeting = [ordered]@{
            Approved        = ($null -ne $script:MatchedPrefix)
            MatchedPrefix   = $script:MatchedPrefix
            AllowedPrefixes = New-StringArrayForJson -InputObject $AllowedComputerPrefixes
        }

        ExistingAgent = Convert-AgentHealthForTelemetry -Health $script:ExistingHealth

        Installation = [ordered]@{
            ForceReinstall          = [bool]$ForceReinstall
            Attempted               = [bool]$script:InstallationAttempted
            ReinstallAttempted      = [bool]$script:ReinstallAttempted
            PackageSource           = $script:PackageSource
            PackageVersion          = $ElasticAgentVersion
            PreferredInstallerPath  = $PreferredInstallerPath
            FallbackInstallerPath   = $FallbackInstallerPath
            DownloadUri             = $DownloadUri
            PackageHashEnforced     = [bool]$script:PackageHashEnforced
            PackageHashVerified     = [bool]$script:PackageHashVerified
            FleetServerUrl          = $FleetServerUrl
            UseInsecureFleetTls     = [bool]$UseInsecureFleetTls
            EnrollmentAttempted     = [bool]$script:EnrollmentAttempted
            EnrollmentTokenProvided = -not [string]::IsNullOrWhiteSpace($EnrollmentToken)
            EnrollmentTokenLogged   = $false
            InstallerExitCode       = $script:InstallerExitCode
        }

        Actions = [ordered]@{
            ExistingAgentDetected          = ($null -ne $script:ExistingHealth -and [bool]$script:ExistingHealth.Installed)
            ExistingHealthRestartAttempted = [bool]$script:ExistingHealthRestartAttempted
            ExistingHealthRestored         = [bool]$script:ExistingHealthRestored
            ReinstallAttempted              = [bool]$script:ReinstallAttempted
            EnrollmentAttempted             = [bool]$script:EnrollmentAttempted
        }

        FinalAgent = Convert-AgentHealthForTelemetry -Health $script:FinalHealth
        VersionResponse = $script:VersionResponse
    }

    Write-Log -Message ("Elastic Agent telemetry summary: Status={0}; Result={1}; Targeted={2}; ExistingHealthy={3}; InstallAttempted={4}; PackageSource={5}; FinalHealthy={6}; ExitCode={7}" -f `
        $event.Status,
        $event.OverallResult,
        $event.Targeting.Approved,
        $(if ($event.ExistingAgent) { $event.ExistingAgent.Healthy } else { $null }),
        $event.Installation.Attempted,
        $event.Installation.PackageSource,
        $(if ($event.FinalAgent) { $event.FinalAgent.Healthy } else { $null }),
        $event.ExitCode)

    $compact = $event | ConvertTo-Json -Depth 12 -Compress
    $pretty  = $event | ConvertTo-Json -Depth 12

    Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compact
    Write-JsonAtomically -Path $LatestTelemetryPath -Json $pretty
}

function Test-ApprovedComputer {
    [CmdletBinding()]
    param()

    foreach ($prefix in $AllowedComputerPrefixes) {
        if (-not [string]::IsNullOrWhiteSpace($prefix) -and
            $env:COMPUTERNAME.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $script:MatchedPrefix = [string]$prefix
            return $true
        }
    }

    return $false
}

function Test-IsAdministrator {
    [CmdletBinding()]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-InstalledElasticAgent {
    [CmdletBinding()]
    param()

    $service = Get-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue
    $exeExists = Test-Path -LiteralPath $InstalledAgentPath -PathType Leaf

    [pscustomobject]@{
        ServiceExists = ($null -ne $service)
        Service       = $service
        ExecutableExists = $exeExists
        IsInstalled   = (($null -ne $service) -and $exeExists)
    }
}


function Get-ElasticAgentHealth {
    [CmdletBinding()]
    param()

    $installed = Get-InstalledElasticAgent

    if (-not $installed.IsInstalled) {
        return [pscustomobject]@{
            Installed      = $false
            ServiceRunning = $false
            StatusHealthy  = $false
            Healthy        = $false
            ServiceStatus  = $null
            StatusExitCode = $null
            StatusOutput   = New-StringArrayForJson -InputObject @()
            Message        = 'Elastic Agent is not fully installed.'
        }
    }

    $service = Get-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue

    if ($service -and $service.Status -ne 'Running') {
        Write-Log -Level WARNING -Message "Elastic Agent service is installed but currently $($service.Status). Attempting to start it."

        try {
            Set-Service -Name $ElasticServiceName -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name $ElasticServiceName -ErrorAction Stop
            $service.WaitForStatus('Running', [timespan]::FromSeconds(20))
        }
        catch {
            Write-Log -Level WARNING -Message "Unable to start Elastic Agent service: $($_.Exception.Message)"
        }

        $service = Get-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue
    }

    $serviceRunning = ($service -and $service.Status -eq 'Running')
    $statusOutput = @()
    $statusExitCode = $null
    $statusHealthy = $false

    if (Test-Path -LiteralPath $InstalledAgentPath -PathType Leaf) {
        try {
            $statusOutput = @(& $InstalledAgentPath status 2>&1 | ForEach-Object { [string]$_ })
            $statusExitCode = $LASTEXITCODE
            $joined = ($statusOutput -join "`n")

            $statusHealthy = (
                $statusExitCode -eq 0 -and
                $joined -match '(?im)\bHEALTHY\b' -and
                $joined -notmatch '(?im)\bFAILED\b|\bDEGRADED\b|\bUNHEALTHY\b'
            )
        }
        catch {
            $statusOutput = @($_.Exception.Message)
            $statusExitCode = -1
        }
    }

    $healthy = ($serviceRunning -and $statusHealthy)

    [pscustomobject]@{
        Installed      = $true
        ServiceRunning = $serviceRunning
        StatusHealthy  = $statusHealthy
        Healthy        = $healthy
        ServiceStatus  = if ($service) { [string]$service.Status } else { 'Missing' }
        StatusExitCode = $statusExitCode
        StatusOutput   = New-StringArrayForJson -InputObject $statusOutput
        Message        = if ($healthy) {
            'Elastic Agent is installed and reporting HEALTHY.'
        }
        elseif (-not $serviceRunning) {
            'Elastic Agent is installed, but its Windows service is not running.'
        }
        else {
            'Elastic Agent service is running, but elastic-agent status is not reporting HEALTHY.'
        }
    }
}

function Wait-ElasticAgentHealth {
    [CmdletBinding()]
    param(
        [ValidateRange(1,20)][int]$Attempts = 8,
        [ValidateRange(1,60)][int]$DelaySeconds = 5
    )

    $lastHealth = $null

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        $lastHealth = Get-ElasticAgentHealth

        if ($lastHealth.Healthy) {
            return $lastHealth
        }

        if ($attempt -lt $Attempts) {
            Write-Log -Message ("Elastic Agent health check {0}/{1} is not yet HEALTHY. ServiceStatus={2}; StatusExitCode={3}. Retrying in {4} seconds." -f `
                $attempt, $Attempts, $lastHealth.ServiceStatus, $lastHealth.StatusExitCode, $DelaySeconds)
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    return $lastHealth
}

function Test-ExistingElasticAgentAndExit {
    [CmdletBinding()]
    param()

    $health = Get-ElasticAgentHealth

    if (-not $health.Installed) {
        return $false
    }

    Write-Log -Message "Existing Elastic Agent detected. ServiceStatus=$($health.ServiceStatus); StatusExitCode=$($health.StatusExitCode)."

    foreach ($line in @($health.StatusOutput | Select-Object -First 15)) {
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            Write-Log -Message "Elastic status: $line"
        }
    }

    if ($health.Healthy) {
        Write-Log -Level SUCCESS -Message 'Elastic Agent is already installed and working properly. No installation action is required.'
        return $true
    }

    Write-Log -Level WARNING -Message $health.Message
    $script:ExistingHealthRestartAttempted = $true

    try {
        Restart-Service -Name $ElasticServiceName -Force -ErrorAction Stop
        (Get-Service -Name $ElasticServiceName -ErrorAction Stop).WaitForStatus('Running', [timespan]::FromSeconds(20))
    }
    catch {
        Write-Log -Level WARNING -Message "Elastic Agent service restart attempt failed: $($_.Exception.Message)"
    }

    $retry = Wait-ElasticAgentHealth -Attempts 6 -DelaySeconds 5

    if ($retry.Healthy) {
        $script:ExistingHealthRestored = $true
        Write-Log -Level SUCCESS -Message 'Elastic Agent was already installed. Health was restored after restarting the service; agent is now working properly.'
        return $true
    }

    Write-Log -Level WARNING -Message 'Elastic Agent is installed but did not verify as healthy after a service restart. Installation will not be duplicated; manual/Fleet investigation is recommended.'
    return $true
}

function Resolve-InstallerPackage {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DestinationPath)

    $source = $null

    if (Test-Path -LiteralPath $PreferredInstallerPath -PathType Leaf -ErrorAction SilentlyContinue) {
        $source = $PreferredInstallerPath
        Write-Log -Level SUCCESS -Message "Elastic Agent package found at preferred share: $PreferredInstallerPath"
    }
    elseif (Test-Path -LiteralPath $FallbackInstallerPath -PathType Leaf -ErrorAction SilentlyContinue) {
        $source = $FallbackInstallerPath
        Write-Log -Level WARNING -Message "Preferred Elastic Agent share was unavailable. Using fallback share: $FallbackInstallerPath"
    }

    if ($null -ne $source) {
        Write-Log -Message "Copying Elastic Agent package from $source"
        Copy-Item -LiteralPath $source -Destination $DestinationPath -Force -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $DestinationPath -PathType Leaf)) {
            throw "Elastic Agent package copy completed without error, but the local ZIP was not created: $DestinationPath"
        }

        return 'Share'
    }

    Write-Log -Level WARNING -Message 'Elastic Agent package was not accessible from either internal share. Falling back to the official Elastic artifact site.'
    Write-Log -Message "Downloading Elastic Agent $ElasticAgentVersion from $DownloadUri"

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        Invoke-WebRequest `
            -Uri $DownloadUri `
            -OutFile $DestinationPath `
            -UseBasicParsing `
            -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $DestinationPath -PathType Leaf)) {
            throw "Elastic Agent Internet download completed without error, but the ZIP was not created: $DestinationPath"
        }

        $downloadedLength = (Get-Item -LiteralPath $DestinationPath -ErrorAction Stop).Length
        if ($downloadedLength -lt 1MB) {
            throw "Elastic Agent Internet download appears incomplete. Downloaded size: $downloadedLength bytes."
        }

        Write-Log -Level SUCCESS -Message "Elastic Agent package downloaded successfully from the official Elastic artifact site. Size=$downloadedLength bytes."
        return 'Internet'
    }
    catch {
        throw "Elastic Agent package was unavailable from both internal shares and the official Internet download also failed: $($_.Exception.Message)"
    }
}

function Test-PackageHash {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($ExpectedSHA512)) {
        Write-Log -Level WARNING -Message 'ExpectedSHA512 is blank; package hash enforcement was skipped.'
        return $false
    }

    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA512).Hash
    if (-not $actual.Equals($ExpectedSHA512.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Elastic Agent package SHA-512 mismatch. Expected $ExpectedSHA512 but found $actual."
    }

    Write-Log -Message 'Elastic Agent package SHA-512 validation passed.'
    return $true
}

# ============================================================================
# MAIN
# ============================================================================

$exitCode = 1

try {
    $script:CurrentStage = 'Initialization'
    Invoke-LogMaintenance
    Write-Log -Message "Elastic Agent installation check started. ScriptVersion=$ScriptVersion."

    $script:CurrentStage = 'PrivilegeCheck'
    if (-not (Test-IsAdministrator)) {
        throw 'This script must run as Administrator or SYSTEM.'
    }

    $script:CurrentStage = 'Targeting'
    if (-not (Test-ApprovedComputer)) {
        Write-Log -Message "Computer is not in an approved lab prefix. Allowed prefixes: $($AllowedComputerPrefixes -join ', '). No changes were made."
        $script:OverallResult = 'NotTargeted'
        $script:FinalHealth = Get-ElasticAgentHealth
        $exitCode = 0
        return
    }

    Write-Log -Message "Computer matched the approved Elastic Agent pilot prefix list."

    $script:CurrentStage = 'ExistingAgentCheck'
    $installed = Get-InstalledElasticAgent
    $script:ExistingHealth = Get-ElasticAgentHealth

    if ($installed.IsInstalled -and -not $ForceReinstall) {
        $existingWasHealthy = [bool]$script:ExistingHealth.Healthy

        if (Test-ExistingElasticAgentAndExit) {
            $script:FinalHealth = Get-ElasticAgentHealth

            if ($existingWasHealthy) {
                $script:OverallResult = 'AlreadyHealthy'
            }
            elseif ($script:FinalHealth.Healthy -and $script:ExistingHealthRestored) {
                $script:OverallResult = 'ExistingHealthRestored'
            }
            else {
                $script:OverallResult = 'ExistingUnhealthyNeedsInvestigation'
            }

            $exitCode = 0
            return
        }
    }

    if (($installed.ServiceExists -or $installed.ExecutableExists) -and $ForceReinstall) {
        $script:CurrentStage = 'UninstallExistingAgent'
        $script:ReinstallAttempted = $true
        Write-Log -Level WARNING -Message 'ForceReinstall was requested. Removing the existing Elastic Agent installation.'

        if (Test-Path -LiteralPath $InstalledAgentPath -PathType Leaf) {
            $uninstallProcess = Start-Process -FilePath $InstalledAgentPath `
                -ArgumentList @('uninstall', '--force', '--skip-fleet-audit') `
                -Wait -PassThru -WindowStyle Hidden

            if ($uninstallProcess.ExitCode -ne 0) {
                throw "Existing Elastic Agent uninstall returned exit code $($uninstallProcess.ExitCode)."
            }
        }
    }

    if (Test-Path -LiteralPath $WorkingRoot) {
        Remove-Item -LiteralPath $WorkingRoot -Recurse -Force
    }
    Initialize-Directory -Path $WorkingRoot

    $zipPath = Join-Path $WorkingRoot "elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"
    $extractRoot = Join-Path $WorkingRoot 'Extracted'

    $script:CurrentStage = 'ResolveInstallerPackage'
    $script:PackageSource = Resolve-InstallerPackage -DestinationPath $zipPath
    Unblock-File -LiteralPath $zipPath -ErrorAction SilentlyContinue

    $script:CurrentStage = 'ValidatePackageHash'
    $script:PackageHashVerified = [bool](Test-PackageHash -Path $zipPath)

    $script:CurrentStage = 'ExtractPackage'
    Initialize-Directory -Path $extractRoot
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force

    $agentExecutable = Get-ChildItem -LiteralPath $extractRoot -Filter 'elastic-agent.exe' -File -Recurse |
        Select-Object -First 1

    if ($null -eq $agentExecutable) {
        throw 'elastic-agent.exe was not found in the extracted package.'
    }

    $agentDirectory = $agentExecutable.Directory.FullName
    Get-ChildItem -LiteralPath $agentDirectory -File -Recurse | Unblock-File -ErrorAction SilentlyContinue

    $script:CurrentStage = 'InstallAndEnroll'
    $script:InstallationAttempted = $true
    $script:EnrollmentAttempted = $true
    Write-Log -Message "Installing Elastic Agent $ElasticAgentVersion and enrolling with Fleet Server $FleetServerUrl."

    # --non-interactive suppresses installation questions/prompts.
    # --force suppresses configuration-overwrite confirmation. This script only
    # reaches this point on a fresh install or after -ForceReinstall removed the
    # previous installation, avoiding duplicate-agent behavior.
    $installArguments = @(
        'install',
        "--url=$FleetServerUrl",
        "--enrollment-token=$EnrollmentToken",
        '--non-interactive',
        '--force'
    )

    # Quick Start Fleet Server uses a self-signed certificate. Remote agents must
    # use --insecure until Fleet Server is configured with a certificate trusted
    # by the Windows endpoints.
    if ($UseInsecureFleetTls) {
        $installArguments += '--insecure'
        Write-Log -Level WARNING -Message 'Fleet Server TLS certificate verification is disabled for enrollment because UseInsecureFleetTls is enabled.'
    }

    # Capture the Elastic Agent install/enrollment output so a failed deployment
    # records the underlying Elastic error instead of only returning exit code 1.
    $installerStdOut = Join-Path $WorkingRoot 'elastic-agent-install.stdout.log'
    $installerStdErr = Join-Path $WorkingRoot 'elastic-agent-install.stderr.log'

    $installProcess = Start-Process -FilePath $agentExecutable.FullName `
        -ArgumentList $installArguments `
        -WorkingDirectory $agentDirectory `
        -Wait -PassThru -WindowStyle Hidden `
        -RedirectStandardOutput $installerStdOut `
        -RedirectStandardError $installerStdErr

    $script:InstallerExitCode = $installProcess.ExitCode
    $capturedInstallerOutput = @()
    if (Test-Path -LiteralPath $installerStdOut -PathType Leaf) {
        $capturedInstallerOutput += Get-Content -LiteralPath $installerStdOut -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $installerStdErr -PathType Leaf) {
        $capturedInstallerOutput += Get-Content -LiteralPath $installerStdErr -ErrorAction SilentlyContinue
    }

    $script:CurrentStage = 'VerifyInstalledAgent'
    $verification = Wait-ElasticAgentHealth -Attempts 8 -DelaySeconds 5
    $script:FinalHealth = $verification

    if ($installProcess.ExitCode -ne 0 -and -not $verification.Healthy) {
        $capturedInstallerOutput | Select-Object -Last 30 | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                Write-Log -Level ERROR -Message "Elastic Agent: $_"
            }
        }
        throw "Elastic Agent installer returned exit code $($installProcess.ExitCode), and the installed agent did not become HEALTHY during verification."
    }

    $capturedInstallerOutput | Select-Object -Last 10 | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_)) {
            Write-Log -Message "Elastic Agent: $_"
        }
    }

    if ($installProcess.ExitCode -ne 0 -and $verification.Healthy) {
        Write-Log -Level WARNING -Message ("Elastic Agent installer returned exit code {0}, but the installed service subsequently reported HEALTHY. The installer result was treated as a transient daemon-startup race." -f $installProcess.ExitCode)
    }

    if (-not $verification.Installed) {
        throw 'Elastic Agent installation verification failed.'
    }

    if (-not $verification.Healthy) {
        Write-Log -Level WARNING -Message "Elastic Agent installation completed, but health verification did not report HEALTHY. ServiceStatus=$($verification.ServiceStatus); StatusExitCode=$($verification.StatusExitCode)."
        foreach ($line in @($verification.StatusOutput | Select-Object -First 15)) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Log -Message "Elastic status: $line"
            }
        }
    }
    else {
        Write-Log -Level SUCCESS -Message 'Elastic Agent installation verified: service is running and elastic-agent status reports HEALTHY.'
    }

    try {
        Set-Service -Name $ElasticServiceName -StartupType Automatic -ErrorAction Stop
        Start-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue
        $serviceState = (Get-Service -Name $ElasticServiceName -ErrorAction Stop).Status
        Write-Log -Message "Elastic Agent service state after installation: $serviceState"
    }
    catch {
        Write-Log -Level WARNING -Message "Elastic Agent installed, but service verification encountered an issue: $($_.Exception.Message)"
    }

    $installedVersion = (& $InstalledAgentPath version 2>$null | Out-String).Trim()
    $script:VersionResponse = $installedVersion
    Write-Log -Level SUCCESS -Message "Elastic Agent installed successfully. Version response: $installedVersion"
    Write-Log -Level SUCCESS -Message 'Fleet enrollment was requested successfully and the Elastic Agent service is left enabled for Fleet management.'

    $script:FinalHealth = Get-ElasticAgentHealth
    if ($script:FinalHealth.Healthy) {
        $script:OverallResult = if ($script:ReinstallAttempted) { 'ReinstalledAndHealthy' } else { 'InstalledAndHealthy' }
    }
    else {
        $script:OverallResult = 'InstalledButUnhealthy'
    }

    $exitCode = 0
}
catch {
    $script:FailureStage = $script:CurrentStage
    $script:FailureMessage = $_.Exception.Message
    $script:OverallResult = 'Failed'
    try {
        $script:FinalHealth = Get-ElasticAgentHealth
    }
    catch { }

    try {
        Write-Log -Level ERROR -Message $_.Exception.Message
    }
    catch {
        Write-Error $_.Exception.Message
    }
    $exitCode = 1
}
finally {
    if (Test-Path -LiteralPath $WorkingRoot -PathType Container) {
        Remove-Item -LiteralPath $WorkingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Emit one structured deployment event for Elastic. The enrollment token itself
    # is never written; telemetry records only whether one was provided.
    try {
        if ($null -eq $script:FinalHealth) {
            $script:FinalHealth = Get-ElasticAgentHealth
        }
        Write-ElasticAgentTelemetry -ExitCode $exitCode
    }
    catch {
        try {
            Write-Log -Level ERROR -Message ("Elastic Agent telemetry finalization failed: {0}" -f $_.Exception.Message)
        }
        catch { }
        if ($exitCode -eq 0) { $exitCode = 4 }
    }

    # Final append before the completed immutable text log enters C:\Logs.
    try {
        Write-Log -Message ("Finalizing 15_Install_Elastic_Agent.ps1. ExitCode={0}" -f $exitCode) `
            -Level $(if ($exitCode -eq 0) { 'SUCCESS' } else { 'ERROR' })
    }
    catch {
        Write-Warning ("Unable to write final script 15 staged log entry: {0}" -f $_.Exception.Message)
    }

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 15 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 15 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $exitCode
'@
    }
    BrowserHomepage = [ordered]@{
        FileName = '17_Set_Browser_Homepage.ps1'
        # Plain-text source is intentionally embedded for maintainability.
        Source = @'
# ScriptName: 17_Set_Browser_Homepage.ps1
# ScriptVersion: 1.1.0
# LastUpdated: 2026-08-27
# Changes: v1.1.0 adds machine-wide Chrome onboarding/default-browser suppression
#          for shared lab computers while retaining browser sign-in disablement.
#          v1.0.1 updates the internal script identity, log names, and telemetry
#          dataset after renaming the file from script 19 to script 17.
# Purpose: Configure Firefox, Google Chrome, and Microsoft Edge to use the
#          Compton College website as the machine-wide homepage and startup
#          page. Disable Chrome browser/profile sign-in and sync, suppress Chrome
#          promotional/first-run sign-in content and default-browser prompts, while
#          preserving normal sign-in to websites.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [ValidatePattern('^https://')]
    [string]$HomepageUrl = 'https://www.compton.edu',

    [string]$LogFolder = 'C:\Logs'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptName = '17_Set_Browser_Homepage.ps1'
$script:ScriptVersion = '1.1.0'
$script:ComputerName = $env:COMPUTERNAME
$script:StartTime = Get-Date
$script:RunId = [guid]::NewGuid().Guid
$script:ChangedCount = 0
$script:VerifiedCount = 0
$script:Status = 'failed'
$script:ExitCode = 1
$script:LogPath = Join-Path $LogFolder '17_Set_Browser_Homepage.log'
$script:LatestTelemetryPath = Join-Path $LogFolder '17_Set_Browser_Homepage.latest.json'
$script:TelemetryPath = Join-Path $LogFolder 'Maintenance-Telemetry.ndjson'

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level = 'INFO'
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '{0} [{1}] [{2}] {3}' -f $timestamp, $script:ComputerName, $Level, $Message
    Write-Host $line -ForegroundColor $(switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        default   { 'Cyan' }
    })
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'This script must run elevated or as Local System.'
    }

    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        throw 'Run this script with 64-bit PowerShell so browser policies are written to the correct registry view.'
    }
}

function Set-PolicyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Browser,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [Parameter(Mandatory)][ValidateSet('String','DWord')][string]$Type
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    $currentValue = $null
    $valueExists = $false
    try {
        $currentValue = Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop
        $valueExists = $true
    }
    catch [System.Management.Automation.PSArgumentException] {}
    catch [System.Management.Automation.ItemNotFoundException] {}

    $expectedValue = if ($Type -eq 'DWord') { [int]$Value } else { [string]$Value }
    $alreadyCorrect = $valueExists -and ([string]$currentValue -eq [string]$expectedValue)

    if (-not $alreadyCorrect) {
        New-ItemProperty -LiteralPath $Path -Name $Name -Value $expectedValue -PropertyType $Type -Force | Out-Null
        $script:ChangedCount++
        Write-Log ("{0}: Set {1} = {2}." -f $Browser, $Name, $expectedValue) 'INFO'
    }
    else {
        Write-Log ("{0}: {1} is already correctly configured." -f $Browser, $Name) 'INFO'
    }

    $verifiedValue = Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop
    if ([string]$verifiedValue -ne [string]$expectedValue) {
        throw ("{0}: Verification failed for {1}. Expected '{2}', found '{3}'." -f $Browser, $Name, $expectedValue, $verifiedValue)
    }

    $script:VerifiedCount++
}

function Remove-UnwantedPolicyValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Browser,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$AllowedNames
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    $properties = @(
        (Get-ItemProperty -LiteralPath $Path -ErrorAction Stop).PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' }
    )

    foreach ($property in $properties) {
        if ($property.Name -notin $AllowedNames) {
            Remove-ItemProperty -LiteralPath $Path -Name $property.Name -Force -ErrorAction Stop
            $script:ChangedCount++
            Write-Log ("{0}: Removed unwanted startup URL policy entry '{1}'." -f $Browser, $property.Name) 'INFO'
        }
    }
}

function Add-SharedTextLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Line,
        [int]$MaximumAttempts = 10
    )

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        $stream = $null
        $writer = $null
        try {
            $stream = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
            $writer.WriteLine($Line)
            $writer.Flush()
            return
        }
        catch [System.IO.IOException] {
            if ($attempt -ge $MaximumAttempts) { throw }
            Start-Sleep -Milliseconds (250 * $attempt)
        }
        finally {
            if ($null -ne $writer) { $writer.Dispose() }
            elseif ($null -ne $stream) { $stream.Dispose() }
        }
    }
}

function Write-Telemetry {
    $endTime = Get-Date
    $event = [pscustomobject][ordered]@{
        EventType       = 'maintenance.execution'
        EventDataset    = 'compton.maintenance.script17'
        ComputerName    = $script:ComputerName
        ScriptName      = $script:ScriptName
        ScriptVersion   = $script:ScriptVersion
        RunId           = $script:RunId
        Status          = $script:Status
        ExitCode        = $script:ExitCode
        StartTime       = $script:StartTime.ToString('o')
        EndTime         = $endTime.ToString('o')
        DurationSeconds = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 2)
        Timestamp       = $endTime.ToUniversalTime().ToString('o')
        HomepageUrl     = $HomepageUrl
        ChangedCount    = $script:ChangedCount
        VerifiedCount   = $script:VerifiedCount
        ChromeSignin                 = 'disabled'
        ChromeSync                   = 'disabled'
        ChromeSigninInterception     = 'disabled'
        ChromePromotionalTabs        = 'disabled'
        ChromeDefaultBrowserPrompt   = 'disabled'
        ChromePolicyScope            = 'HKLM-AllUsers'
        LogPath                      = $script:LogPath
    }

    $json = $event | ConvertTo-Json -Depth 5 -Compress
    Add-SharedTextLine -Path $script:TelemetryPath -Line $json
    $event | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:LatestTelemetryPath -Encoding UTF8 -Force
}

try {
    Assert-Administrator
    Write-Log ("Starting browser homepage configuration. Version={0}; Homepage={1}; RunId={2}" -f $script:ScriptVersion, $HomepageUrl, $script:RunId) 'INFO'

    if (-not $PSCmdlet.ShouldProcess($script:ComputerName, "Set Firefox, Chrome, and Edge homepage policies to $HomepageUrl")) {
        $script:Status = 'what_if'
        $script:ExitCode = 0
        return
    }

    # Google Chrome machine policies.
    $chromePath = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    $chromeStartupPath = Join-Path $chromePath 'RestoreOnStartupURLs'
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'HomepageLocation' -Value $HomepageUrl -Type String
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'HomepageIsNewTabPage' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'ShowHomeButton' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'RestoreOnStartup' -Value 4 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromeStartupPath -Name '1' -Value $HomepageUrl -Type String
    Remove-UnwantedPolicyValues -Browser 'Chrome' -Path $chromeStartupPath -AllowedNames @('1')

    # Disable Chrome browser/profile sign-in and sync prompts. This does not
    # block users from signing in to websites in Chrome.
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'BrowserSignin' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'SyncDisabled' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'SigninInterceptionEnabled' -Value 0 -Type DWord

    # Suppress Chrome's full-tab onboarding/promotional content (including the
    # sign-in / "stay signed out" first-run experience) and stop Chrome from
    # asking users to make it the Windows default browser.
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'PromotionalTabsEnabled' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'DefaultBrowserSettingEnabled' -Value 0 -Type DWord

    # Microsoft Edge machine policies.
    $edgePath = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    $edgeStartupPath = Join-Path $edgePath 'RestoreOnStartupURLs'
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'HomepageLocation' -Value $HomepageUrl -Type String
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'HomepageIsNewTabPage' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'ShowHomeButton' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'RestoreOnStartup' -Value 4 -Type DWord
    Set-PolicyValue -Browser 'Edge' -Path $edgeStartupPath -Name '1' -Value $HomepageUrl -Type String
    Remove-UnwantedPolicyValues -Browser 'Edge' -Path $edgeStartupPath -AllowedNames @('1')

    # Mozilla Firefox machine policies.
    $firefoxHomepagePath = 'HKLM:\SOFTWARE\Policies\Mozilla\Firefox\Homepage'
    Set-PolicyValue -Browser 'Firefox' -Path $firefoxHomepagePath -Name 'URL' -Value $HomepageUrl -Type String
    Set-PolicyValue -Browser 'Firefox' -Path $firefoxHomepagePath -Name 'Locked' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Firefox' -Path $firefoxHomepagePath -Name 'StartPage' -Value 'homepage' -Type String

    $script:Status = 'success'
    $script:ExitCode = 0
    Write-Log ("Browser policies completed successfully. Changed={0}; Verified={1}. Close and reopen each browser to apply the policies." -f $script:ChangedCount, $script:VerifiedCount) 'SUCCESS'
}
catch {
    $script:Status = 'failed'
    $script:ExitCode = 1
    try { Write-Log ("Browser policy configuration failed: {0}" -f $_.Exception.Message) 'ERROR' }
    catch { Write-Error $_.Exception.Message }
}
finally {
    try { Write-Telemetry }
    catch { try { Write-Log ("Telemetry write failed: {0}" -f $_.Exception.Message) 'WARNING' } catch {} }
}

exit $script:ExitCode
'@
    }
    Honorlock = [ordered]@{
        FileName = '18_Install_Honorlock_Chrome_Extension.ps1'
        # Plain-text source is intentionally embedded for maintainability.
        Source = @'
# ScriptName: 18_Install_Honorlock_Chrome_Extension.ps1
# ScriptVersion: 1.2.0
# LastUpdated: 2026-08-27
# Changes: v1.2.0 adds Building/Lab/DeviceIdentifier and target-scope telemetry
#          for scalable Honorlock deployment dashboards as additional labs are added.
#          v1.1.2 updates the internal script identity, log names, and telemetry
#          dataset after renaming the file from script 20 to script 18.
#          v1.1.1 corrects the computer-name pattern array that caused a
#          PowerShell parser error.
# Purpose: Force-install the Honorlock Chrome extension machine-wide through
#          Google Chrome enterprise policy without overwriting other forced
#          extension entries. Deployment can be limited to configured computer
#          name prefixes or wildcard patterns.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [string]$LogFolder = 'C:\Logs',

    # Leave empty to target every computer. Examples:
    # @('IB1-103-*')
    # @('IB1-103-*','SSB-122-*','SSB-114*')
    [string[]]$ComputerNamePatterns = @(
        'SSC-216*',
        'AHB-146*'
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptName = '18_Install_Honorlock_Chrome_Extension.ps1'
$script:ScriptVersion = '1.2.0'
$script:ComputerName = $env:COMPUTERNAME
$script:StartTime = Get-Date
$script:RunId = [guid]::NewGuid().Guid
$script:Status = 'failed'
$script:ExitCode = 1
$script:PolicyChanged = $false
$script:PolicyVerified = $false
$script:PolicyValueName = $null
$script:Targeted = $false
$script:MatchedComputerPattern = $null
$script:OverallResult = 'NotStarted'
$script:Building = $null
$script:Lab = $null
$script:DeviceIdentifier = $null

$script:HonorlockExtensionId = 'hnbmpkmhjackfpkpcbapafmpepgmmddc'
$script:ChromeUpdateUrl = 'https://clients2.google.com/service/update2/crx'
$script:ForceInstallValue = '{0};{1}' -f $script:HonorlockExtensionId, $script:ChromeUpdateUrl
$script:ForceListPath = 'HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist'

$script:LogPath = Join-Path $LogFolder '18_Install_Honorlock_Chrome_Extension.log'
$script:LatestTelemetryPath = Join-Path $LogFolder '18_Install_Honorlock_Chrome_Extension.latest.json'
$script:TelemetryPath = Join-Path $LogFolder 'Maintenance-Telemetry.ndjson'

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level = 'INFO'
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '{0} [{1}] [{2}] {3}' -f $timestamp, $script:ComputerName, $Level, $Message
    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        default   { 'Cyan' }
    }

    Write-Host $line -ForegroundColor $color
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'This script must run elevated or as Local System.'
    }

    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        throw 'Run this script with 64-bit PowerShell so the Chrome policy is written to the correct registry view.'
    }
}

function Get-ComputerLocation {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    # Common Compton naming pattern:
    #   IB1-103-13   -> Building=IB1, Lab=IB1-103, DeviceIdentifier=13
    #   SSC-216-015  -> Building=SSC, Lab=SSC-216, DeviceIdentifier=015
    # Preserve a useful fallback for names that do not follow that pattern.
    if ($ComputerName -match '^(?<Building>[^-]+)-(?<Room>[^-]+)-(?<Device>.+)$') {
        return [pscustomobject][ordered]@{
            Building         = [string]$Matches.Building
            Lab              = '{0}-{1}' -f $Matches.Building, $Matches.Room
            DeviceIdentifier = [string]$Matches.Device
        }
    }

    return [pscustomobject][ordered]@{
        Building         = $null
        Lab              = $null
        DeviceIdentifier = $ComputerName
    }
}

function Get-MatchedComputerPattern {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowEmptyCollection()][string[]]$Patterns
    )

    $configuredPatterns = @(
        $Patterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    # An empty list intentionally targets all computers.
    if ($configuredPatterns.Count -eq 0) {
        return '*'
    }

    foreach ($pattern in $configuredPatterns) {
        if ($Name -like $pattern) {
            return [string]$pattern
        }
    }

    return $null
}

function Get-ForceListProperties {
    if (-not (Test-Path -LiteralPath $script:ForceListPath)) {
        return @()
    }

    return @(
        (Get-ItemProperty -LiteralPath $script:ForceListPath -ErrorAction Stop).PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' }
    )
}

function Set-HonorlockForceInstallPolicy {
    if (-not (Test-Path -LiteralPath $script:ForceListPath)) {
        New-Item -Path $script:ForceListPath -Force | Out-Null
        Write-Log ("Created Chrome force-install policy path: {0}" -f $script:ForceListPath) 'INFO'
    }

    $properties = @(Get-ForceListProperties)
    $extensionPattern = '^{0}(;|$)' -f [regex]::Escape($script:HonorlockExtensionId)
    $existingProperty = @(
        $properties | Where-Object { [string]$_.Value -match $extensionPattern }
    ) | Select-Object -First 1

    if ($null -ne $existingProperty) {
        $script:PolicyValueName = [string]$existingProperty.Name
        if ([string]$existingProperty.Value -ne $script:ForceInstallValue) {
            New-ItemProperty -LiteralPath $script:ForceListPath -Name $script:PolicyValueName `
                -Value $script:ForceInstallValue -PropertyType String -Force | Out-Null
            $script:PolicyChanged = $true
            Write-Log ("Updated existing Honorlock force-install entry '{0}'." -f $script:PolicyValueName) 'INFO'
        }
        else {
            Write-Log ("Honorlock is already present in Chrome's force-install policy as entry '{0}'." -f $script:PolicyValueName) 'INFO'
        }
    }
    else {
        $usedNumbers = @(
            $properties |
            Where-Object { $_.Name -match '^\d+$' } |
            ForEach-Object { [int]$_.Name }
        )

        $nextNumber = 1
        while ($usedNumbers -contains $nextNumber) {
            $nextNumber++
        }

        $script:PolicyValueName = [string]$nextNumber
        New-ItemProperty -LiteralPath $script:ForceListPath -Name $script:PolicyValueName `
            -Value $script:ForceInstallValue -PropertyType String -Force | Out-Null
        $script:PolicyChanged = $true
        Write-Log ("Added Honorlock to Chrome's force-install policy as entry '{0}'." -f $script:PolicyValueName) 'INFO'
    }

    $verifiedProperties = @(Get-ForceListProperties)
    $verified = @(
        $verifiedProperties | Where-Object { [string]$_.Value -eq $script:ForceInstallValue }
    )

    if ($verified.Count -lt 1) {
        throw 'Honorlock force-install policy verification failed after the registry update.'
    }

    $script:PolicyVerified = $true
    Write-Log ("Verified Honorlock force-install policy. ExtensionId={0}; Entry={1}" -f $script:HonorlockExtensionId, $script:PolicyValueName) 'SUCCESS'
}

function Add-SharedTextLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Line,
        [int]$MaximumAttempts = 10
    )

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        $stream = $null
        $writer = $null
        try {
            $stream = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
            $writer.WriteLine($Line)
            $writer.Flush()
            return
        }
        catch [System.IO.IOException] {
            if ($attempt -ge $MaximumAttempts) { throw }
            Start-Sleep -Milliseconds (250 * $attempt)
        }
        finally {
            if ($null -ne $writer) { $writer.Dispose() }
            elseif ($null -ne $stream) { $stream.Dispose() }
        }
    }
}

function Write-Telemetry {
    $endTime = Get-Date
    $installState = switch ($script:OverallResult) {
        'PolicyConfigured' { 'Policy configured; Chrome installs/updates the extension at policy refresh or browser startup.' }
        'NotTargeted'      { 'Skipped because the computer name did not match the configured pattern list.' }
        'WhatIf'           { 'WhatIf evaluation only; no policy changes were made.' }
        'Failed'           { 'Policy deployment failed. Review the text log.' }
        default            { 'No completed deployment result was recorded.' }
    }

    $event = [pscustomobject][ordered]@{
        EventType       = 'maintenance.execution'
        EventDataset    = 'compton.maintenance.script18'
        ComputerName    = $script:ComputerName
        Building        = $script:Building
        Lab             = $script:Lab
        DeviceIdentifier = $script:DeviceIdentifier
        ScriptName      = $script:ScriptName
        ScriptVersion   = $script:ScriptVersion
        RunId           = $script:RunId
        Status          = $script:Status
        ExitCode        = $script:ExitCode
        StartTime       = $script:StartTime.ToString('o')
        EndTime         = $endTime.ToString('o')
        DurationSeconds = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 2)
        Timestamp       = $endTime.ToUniversalTime().ToString('o')
        Browser         = 'Google Chrome'
        ExtensionName   = 'Honorlock'
        ExtensionId     = $script:HonorlockExtensionId
        PolicyChanged   = $script:PolicyChanged
        PolicyVerified  = $script:PolicyVerified
        PolicyValueName = $script:PolicyValueName
        ComputerNamePatterns = @($ComputerNamePatterns)
        TargetPatternCount   = @($ComputerNamePatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
        TargetScope          = if (@($ComputerNamePatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0) { 'AllComputers' } else { 'ConfiguredPatterns' }
        Targeted        = $script:Targeted
        MatchedPattern  = $script:MatchedComputerPattern
        OverallResult   = $script:OverallResult
        InstallState    = $installState
        LogPath         = $script:LogPath
    }

    $json = $event | ConvertTo-Json -Depth 5 -Compress
    Add-SharedTextLine -Path $script:TelemetryPath -Line $json
    $event | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:LatestTelemetryPath -Encoding UTF8 -Force
}

try {
    Write-Log ("Starting Honorlock Chrome extension policy deployment. Version={0}; RunId={1}" -f $script:ScriptVersion, $script:RunId) 'INFO'

    $location = Get-ComputerLocation -ComputerName $script:ComputerName
    $script:Building = $location.Building
    $script:Lab = $location.Lab
    $script:DeviceIdentifier = $location.DeviceIdentifier

    $script:MatchedComputerPattern = Get-MatchedComputerPattern `
        -Name $script:ComputerName `
        -Patterns $ComputerNamePatterns

    if ([string]::IsNullOrWhiteSpace($script:MatchedComputerPattern)) {
        $script:Targeted = $false
        $script:OverallResult = 'NotTargeted'
        $script:Status = 'success'
        $script:ExitCode = 0
        $configuredText = if (@($ComputerNamePatterns).Count -gt 0) {
            @($ComputerNamePatterns) -join ', '
        }
        else {
            '(none)'
        }
        Write-Log ("Computer '{0}' does not match the configured pattern list [{1}]. No changes were made." -f $script:ComputerName, $configuredText) 'INFO'
        return
    }

    $script:Targeted = $true
    Write-Log ("Computer is targeted. Matched pattern: {0}" -f $script:MatchedComputerPattern) 'SUCCESS'
    Assert-Administrator

    if (-not $PSCmdlet.ShouldProcess($script:ComputerName, 'Force-install the Honorlock extension in Google Chrome for all users')) {
        $script:Status = 'what_if'
        $script:ExitCode = 0
        $script:OverallResult = 'WhatIf'
        return
    }

    Set-HonorlockForceInstallPolicy

    $script:Status = 'success'
    $script:ExitCode = 0
    $script:OverallResult = 'PolicyConfigured'
    Write-Log 'Honorlock policy deployment completed. Chrome will install or update the extension when policy refreshes or Chrome next starts.' 'SUCCESS'
}
catch {
    $script:Status = 'failed'
    $script:ExitCode = 1
    $script:OverallResult = 'Failed'
    try { Write-Log ("Honorlock policy deployment failed: {0}" -f $_.Exception.Message) 'ERROR' }
    catch { Write-Error $_.Exception.Message }
}
finally {
    try { Write-Telemetry }
    catch { try { Write-Log ("Telemetry write failed: {0}" -f $_.Exception.Message) 'WARNING' } catch {} }
}

exit $script:ExitCode
'@
    }
    StellariumLocation = [ordered]@{
        FileName = 'Embedded_Stellarium_Location_Services.ps1'
        # Plain-text source is intentionally embedded for maintainability.
        Source = @'
#requires -Version 5.1
#requires -RunAsAdministrator
<#
.SYNOPSIS
    Enables Windows 11 Location Services for Stellarium on approved computers.

.DESCRIPTION
    Runs only when the computer name matches an approved prefix pattern. Clears
    the local policies that disable location and sensors, enables device and
    desktop-application location access, configures the Geolocation Service for
    trigger start, and verifies the resulting registry configuration.

    This script is designed to run as SYSTEM from the Compton College weekly
    maintenance task framework. Machine-level app-location policy is used so
    the setting applies to interactive users rather than the SYSTEM profile.

.NOTES
    ScriptName:    19_Stellarium_Location_Services.ps1
    ScriptVersion: 1.0.0
    LastUpdated:   2026-09-02
    Targets:       SSC-216* and MS-203*
    Designed for:  Windows PowerShell 5.1 / Windows 11
#>

[CmdletBinding()]
param(
    [string[]]$TargetComputerPatterns = @('SSC-216*', 'MS-203*'),
    [string]$LogDirectory = 'C:\Logs'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptName = '19_Stellarium_Location_Services.ps1'
$ScriptVersion = '1.0.0'
$RunId = [guid]::NewGuid().Guid
$StartTime = Get-Date
$ComputerName = [string]$env:COMPUTERNAME
$LogPath = Join-Path $LogDirectory '19_Stellarium_Location_Services.log'
$LatestPath = Join-Path $LogDirectory '19_Stellarium_Location_Services.latest.json'
$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$Changes = New-Object System.Collections.Generic.List[string]
$WarningCount = 0
$ErrorCount = 0
$Status = 'Success'
$ExitCode = 0
$FailureMessage = $null
$MatchedPattern = $null

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'ACTION', 'OK', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $line = '{0} [{1}] [{2}] {3}' -f `
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $ComputerName, $Level, $Message

    try {
        Ensure-Directory -Path $LogDirectory
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {}

    $color = switch ($Level) {
        'ACTION' { 'Yellow' }
        'OK'     { 'Green' }
        'WARN'   { 'DarkYellow' }
        'ERROR'  { 'Red' }
        default  { 'Cyan' }
    }
    Write-Host $line -ForegroundColor $color
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-RequiredRegistryValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [Parameter(Mandatory)][ValidateSet('DWord', 'String')][string]$PropertyType
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        [void]$Changes.Add("Created registry key $Path")
    }

    $currentValue = $null
    $valueExists = $false
    try {
        $currentValue = Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop
        $valueExists = $true
    }
    catch {}

    if (-not $valueExists -or [string]$currentValue -cne [string]$Value) {
        New-ItemProperty `
            -Path $Path `
            -Name $Name `
            -PropertyType $PropertyType `
            -Value $Value `
            -Force | Out-Null

        [void]$Changes.Add("Set $Path\$Name to $Value")
        Write-Log -Level 'ACTION' -Message "Set registry value $Path\$Name to $Value."
    }
    else {
        Write-Log -Message "Registry value is already correct: $Path\$Name"
    }
}

function Test-RegistryValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$ExpectedValue
    )

    try {
        $actualValue = Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop
        return ([string]$actualValue -ceq [string]$ExpectedValue)
    }
    catch {
        return $false
    }
}

function Write-ExecutionRecord {
    $endTime = Get-Date
    $record = [ordered]@{
        EventType        = 'maintenance.stellarium_location_services'
        SchemaVersion    = 1
        ComputerName     = $ComputerName
        ScriptName       = $ScriptName
        ScriptVersion    = $ScriptVersion
        RunId            = $RunId
        Status           = $Status
        ExitCode         = $ExitCode
        StartTime        = $StartTime.ToUniversalTime().ToString('o')
        EndTime          = $endTime.ToUniversalTime().ToString('o')
        DurationSeconds  = [math]::Round(($endTime - $StartTime).TotalSeconds, 3)
        TargetPatterns   = @($TargetComputerPatterns)
        MatchedPattern   = $MatchedPattern
        Changes          = @($Changes | ForEach-Object { $_ })
        WarningCount     = $WarningCount
        ErrorCount       = $ErrorCount
        FailureMessage   = $FailureMessage
        Timestamp        = $endTime.ToUniversalTime().ToString('o')
    }

    try {
        Ensure-Directory -Path $LogDirectory
        $record | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $LatestPath -Encoding UTF8 -Force
        $record | ConvertTo-Json -Depth 6 -Compress | Add-Content -LiteralPath $TelemetryPath -Encoding UTF8
    }
    catch {
        Write-Log -Level 'WARN' -Message "Unable to write JSON telemetry: $($_.Exception.Message)"
    }
}

Write-Log -Message "Starting $ScriptName version $ScriptVersion. RunId=$RunId"

try {
    if ([string]::IsNullOrWhiteSpace($ComputerName)) {
        throw 'The computer name could not be determined.'
    }

    foreach ($pattern in $TargetComputerPatterns) {
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $ComputerName -like $pattern) {
            $MatchedPattern = $pattern
            break
        }
    }

    if ($null -eq $MatchedPattern) {
        $Status = 'SkippedNotTargeted'
        Write-Log -Level 'OK' -Message (
            "Computer is not targeted. Configured patterns: {0}. No changes were made." -f
            ($TargetComputerPatterns -join ', ')
        )
    }
    else {
        Write-Log -Message "Computer matched target pattern '$MatchedPattern'."

        if (-not (Test-IsAdministrator)) {
            throw 'This script must run as Administrator or SYSTEM.'
        }

        $locationPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors'
        $appPrivacyPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy'
        $locationConfigPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration'
        $deviceConsentPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location'
        $desktopConsentPath = Join-Path $deviceConsentPath 'NonPackaged'

        # A value of 0 leaves each "Turn off" policy disabled.
        Set-RequiredRegistryValue -Path $locationPolicyPath -Name 'DisableLocation' -Value 0 -PropertyType DWord
        Set-RequiredRegistryValue -Path $locationPolicyPath -Name 'DisableSensors' -Value 0 -PropertyType DWord
        Set-RequiredRegistryValue -Path $locationPolicyPath -Name 'DisableWindowsLocationProvider' -Value 0 -PropertyType DWord

        # Force-allow location for apps so the setting applies to every user,
        # including when this maintenance script runs under SYSTEM.
        Set-RequiredRegistryValue -Path $appPrivacyPolicyPath -Name 'LetAppsAccessLocation' -Value 1 -PropertyType DWord

        # Enable the Windows location master switch and desktop-app consent.
        Set-RequiredRegistryValue -Path $locationConfigPath -Name 'Status' -Value 1 -PropertyType DWord
        Set-RequiredRegistryValue -Path $deviceConsentPath -Name 'Value' -Value 'Allow' -PropertyType String
        Set-RequiredRegistryValue -Path $desktopConsentPath -Name 'Value' -Value 'Allow' -PropertyType String

        $service = Get-Service -Name 'lfsvc' -ErrorAction Stop
        Set-Service -Name 'lfsvc' -StartupType Manual -ErrorAction Stop
        if ($service.Status -ne 'Running') {
            Start-Service -Name 'lfsvc' -ErrorAction Stop
            [void]$Changes.Add('Started the Geolocation Service (lfsvc)')
            Write-Log -Level 'ACTION' -Message 'Started the Geolocation Service (lfsvc).'
        }
        else {
            Write-Log -Message 'The Geolocation Service (lfsvc) is already running.'
        }

        $verificationChecks = @(
            [pscustomobject]@{ Path=$locationPolicyPath; Name='DisableLocation'; Expected=0 },
            [pscustomobject]@{ Path=$locationPolicyPath; Name='DisableSensors'; Expected=0 },
            [pscustomobject]@{ Path=$locationPolicyPath; Name='DisableWindowsLocationProvider'; Expected=0 },
            [pscustomobject]@{ Path=$appPrivacyPolicyPath; Name='LetAppsAccessLocation'; Expected=1 },
            [pscustomobject]@{ Path=$locationConfigPath; Name='Status'; Expected=1 },
            [pscustomobject]@{ Path=$deviceConsentPath; Name='Value'; Expected='Allow' },
            [pscustomobject]@{ Path=$desktopConsentPath; Name='Value'; Expected='Allow' }
        )

        $failedChecks = @(
            foreach ($check in $verificationChecks) {
                if (-not (Test-RegistryValue -Path $check.Path -Name $check.Name -ExpectedValue $check.Expected)) {
                    "$($check.Path)\$($check.Name)"
                }
            }
        )

        if ($failedChecks.Count -gt 0) {
            throw "Location Services verification failed for: $($failedChecks -join ', ')"
        }

        Write-Log -Level 'OK' -Message "Windows Location Services configuration is enabled and verified. Changes=$($Changes.Count)"
    }
}
catch {
    $Status = 'Failed'
    $ExitCode = 2
    $FailureMessage = $_.Exception.Message
    Write-Log -Level 'ERROR' -Message "Unable to enable Windows Location Services: $FailureMessage"
}
finally {
    Write-ExecutionRecord
    Write-Log -Message "Completed. Status=$Status ExitCode=$ExitCode RunId=$RunId"
}

exit $ExitCode
'@
    }
}

function Write-RunnerLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level = 'INFO'
    )

    if (-not (Test-Path -LiteralPath $LogDirectory -PathType Container)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }

    $computer = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $computer, $Level, $Message
    Add-Content -LiteralPath $RunnerLogPath -Value $line -Encoding UTF8
    Write-Host $line -ForegroundColor $(switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        default   { 'Cyan' }
    })
}

function Test-RunnerAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-ComputerMatchesAnyPattern {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [AllowEmptyCollection()][string[]]$Patterns = @()
    )

    if ($Patterns.Count -eq 0) {
        return $true
    }

    foreach ($pattern in $Patterns) {
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $ComputerName -like $pattern.Trim()) {
            return $true
        }
    }

    return $false
}

function Set-HonorlockChromeMachinePolicy {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][string[]]$ComputerNamePatterns = @()
    )

    $computerName = [string]$env:COMPUTERNAME
    if (-not (Test-ComputerMatchesAnyPattern -ComputerName $computerName -Patterns $ComputerNamePatterns)) {
        Write-RunnerLog -Message "Honorlock machine policy skipped because $computerName does not match the configured patterns."
        return
    }

    # HKLM Chrome policy applies to every Windows user and every Chrome profile.
    # ExtensionSettings is Chrome's current policy; the legacy force list is
    # retained for compatibility with older managed Chrome installations.
    $extensionId = 'hnbmpkmhjackfpkpcbapafmpepgmmddc'
    $updateUrl = 'https://clients2.google.com/service/update2/crx'
    $forceValue = "$extensionId;$updateUrl"
    $chromePolicyPath = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    $forceListPath = Join-Path $chromePolicyPath 'ExtensionInstallForcelist'

    New-Item -Path $chromePolicyPath -Force | Out-Null
    New-Item -Path $forceListPath -Force | Out-Null

    $forceListProperties = Get-ItemProperty -LiteralPath $forceListPath
    $matchingForceEntry = @(
        $forceListProperties.PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS' -and [string]$_.Value -eq $forceValue }
    )

    if ($matchingForceEntry.Count -eq 0) {
        $usedSlots = @(
            $forceListProperties.PSObject.Properties |
                Where-Object { $_.Name -match '^\d+$' } |
                ForEach-Object { [int]$_.Name }
        )
        $slot = 1
        while ($usedSlots -contains $slot) { $slot++ }
        New-ItemProperty -LiteralPath $forceListPath -Name ([string]$slot) -Value $forceValue -PropertyType String -Force | Out-Null
    }

    $existingChromePolicy = Get-ItemProperty -LiteralPath $chromePolicyPath -ErrorAction SilentlyContinue
    $existingSettingsValue = if ($null -ne $existingChromePolicy) {
        $property = $existingChromePolicy.PSObject.Properties['ExtensionSettings']
        if ($null -ne $property) { $property.Value } else { $null }
    }
    else {
        $null
    }

    $existingSettingsJson = [string]$existingSettingsValue
    if ([string]::IsNullOrWhiteSpace($existingSettingsJson)) {
        $extensionSettings = [pscustomobject]@{}
    }
    else {
        try {
            $extensionSettings = $existingSettingsJson | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw "The existing Chrome ExtensionSettings policy is not valid JSON, so it was not overwritten. $($_.Exception.Message)"
        }
        if ($null -eq $extensionSettings) {
            $extensionSettings = [pscustomobject]@{}
        }
    }

    $honorlockSettings = [pscustomobject][ordered]@{
        installation_mode   = 'force_installed'
        update_url          = $updateUrl
        override_update_url = $true
    }
    $extensionSettings | Add-Member -MemberType NoteProperty -Name $extensionId -Value $honorlockSettings -Force
    $updatedSettingsJson = $extensionSettings | ConvertTo-Json -Depth 20 -Compress
    New-ItemProperty -LiteralPath $chromePolicyPath -Name 'ExtensionSettings' -Value $updatedSettingsJson -PropertyType String -Force | Out-Null

    $verifiedForceList = @(
        (Get-ItemProperty -LiteralPath $forceListPath).PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS' -and [string]$_.Value -eq $forceValue }
    ).Count -gt 0
    $verifiedChromePolicy = Get-ItemProperty -LiteralPath $chromePolicyPath -ErrorAction Stop
    $verifiedSettingsJsonProperty = $verifiedChromePolicy.PSObject.Properties['ExtensionSettings']
    if ($null -eq $verifiedSettingsJsonProperty -or [string]::IsNullOrWhiteSpace([string]$verifiedSettingsJsonProperty.Value)) {
        throw 'Honorlock Chrome ExtensionSettings policy was not created.'
    }

    $verifiedSettingsObject = [string]$verifiedSettingsJsonProperty.Value |
        ConvertFrom-Json -ErrorAction Stop
    $verifiedExtensionProperty = $verifiedSettingsObject.PSObject.Properties[$extensionId]
    $verifiedSettings = if ($null -ne $verifiedExtensionProperty) {
        $verifiedExtensionProperty.Value
    }
    else {
        $null
    }

    if (-not $verifiedForceList -or $null -eq $verifiedSettings -or
        [string]$verifiedSettings.installation_mode -ne 'force_installed' -or
        [string]$verifiedSettings.update_url -ne $updateUrl) {
        throw 'Honorlock Chrome machine-policy verification failed.'
    }

    Write-RunnerLog -Message 'Verified Honorlock force-install policy in HKLM for all Windows users and Chrome profiles.' -Level 'SUCCESS'
}

function Write-EmbeddedSectionSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory)][string]$Destination
    )

    if ($null -eq $Payload.Source) {
        throw "Embedded plain-text source is missing for $($Payload.FileName)."
    }

    [System.IO.File]::WriteAllText(
        $Destination,
        [string]$Payload.Source,
        [System.Text.UTF8Encoding]::new($true)
    )
}

function New-SectionBootstrap {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SectionId)

    switch ($SectionId) {
        'AutologonAndEdge' {
            return @'
$settings = $env:COMPTON_SECTION_SETTINGS | ConvertFrom-Json
[string[]]$patterns = @($settings.Patterns | ForEach-Object { [string]$_ })
& $env:COMPTON_SECTION_SCRIPT -ComputerNamePatterns $patterns
'@
        }
        'ElasticAgent' {
            return @'
$settings = $env:COMPTON_SECTION_SETTINGS | ConvertFrom-Json
[string[]]$prefixes = @($settings.Prefixes | ForEach-Object { [string]$_ })
if ([bool]$settings.ForceReinstall) {
    & $env:COMPTON_SECTION_SCRIPT -AllowedComputerPrefixes $prefixes -ForceReinstall
}
else {
    & $env:COMPTON_SECTION_SCRIPT -AllowedComputerPrefixes $prefixes
}
'@
        }
        'BrowserHomepage' {
            return @'
$settings = $env:COMPTON_SECTION_SETTINGS | ConvertFrom-Json
& $env:COMPTON_SECTION_SCRIPT -HomepageUrl ([string]$settings.HomepageUrl)
'@
        }
        'Honorlock' {
            return @'
$settings = $env:COMPTON_SECTION_SETTINGS | ConvertFrom-Json
[string[]]$patterns = @($settings.Patterns | ForEach-Object { [string]$_ })
& $env:COMPTON_SECTION_SCRIPT -ComputerNamePatterns $patterns
'@
        }
        'StellariumLocation' {
            return @'
$settings = $env:COMPTON_SECTION_SETTINGS | ConvertFrom-Json
[string[]]$patterns = @($settings.Patterns | ForEach-Object { [string]$_ })
& $env:COMPTON_SECTION_SCRIPT -TargetComputerPatterns $patterns
'@
        }
        default {
            return '& $env:COMPTON_SECTION_SCRIPT'
        }
    }
}

function Write-RunnerJsonAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Json
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -Path $directory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $temporaryPath = '{0}.{1}.tmp' -f $Path,([guid]::NewGuid().Guid)
    [IO.File]::WriteAllText($temporaryPath, $Json, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force -ErrorAction Stop
}

function Write-RunnerTelemetryLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$JsonLine
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -Path $directory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $lastError = $null
    for ($attempt = 1; $attempt -le 15; $attempt++) {
        $stream = $null
        $writer = $null
        try {
            $stream = [IO.FileStream]::new(
                $Path,
                [IO.FileMode]::Append,
                [IO.FileAccess]::Write,
                [IO.FileShare]::Read
            )
            $writer = [IO.StreamWriter]::new($stream, (New-Object Text.UTF8Encoding($false)))
            $writer.WriteLine($JsonLine)
            $writer.Flush()
            return
        }
        catch {
            $lastError = $_.Exception
            Start-Sleep -Milliseconds 200
        }
        finally {
            if ($null -ne $writer) { $writer.Dispose() }
            elseif ($null -ne $stream) { $stream.Dispose() }
        }
    }

    $lastErrorMessage = if ($null -ne $lastError) { $lastError.Message } else { 'Unknown append error.' }
    throw "Unable to append Office telemetry after 15 attempts: $lastErrorMessage"
}

function Get-SafeObjectPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-OfficeInstallationInventory {
    [CmdletBinding()]
    param()

    $clickToRunPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    [string[]]$productReleaseIds = @()

    if (Test-Path -LiteralPath $clickToRunPath) {
        $releaseText = [string](
            Get-ItemPropertyValue -LiteralPath $clickToRunPath -Name 'ProductReleaseIds' -ErrorAction SilentlyContinue
        )
        $productReleaseIds = @(
            $releaseText -split '[,;\s]+' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        )
    }

    $uninstallPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $uninstallEntries = @(
        Get-ItemProperty -Path $uninstallPaths -ErrorAction SilentlyContinue
    )
    [string[]]$displayNames = @(
        $uninstallEntries |
            ForEach-Object {
                [string](Get-SafeObjectPropertyValue -InputObject $_ -Name 'DisplayName')
            } |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_) -and
                $_ -match '(?i)(Microsoft\s+365|Office\s+365|Microsoft Office|Office LTSC|Office Professional|Office Standard)'
            } |
            Sort-Object -Unique
    )

    [string[]]$microsoft365ProductIds = @(
        $productReleaseIds | Where-Object { $_ -match '(?i)^(?:O365|M365)' }
    )
    [string[]]$microsoft365DisplayNames = @(
        $displayNames | Where-Object { $_ -match '(?i)(?:Microsoft\s+365|Office\s+365)' }
    )

    $suiteIdPattern = '(?i)^(?:ProPlus|Standard|Professional|HomeBusiness|HomeStudent|Personal|Office)'
    [string[]]$office2024ProductIds = @(
        $productReleaseIds | Where-Object {
            $_ -match $suiteIdPattern -and $_ -match '2024' -and $_ -notmatch '(?i)^(?:Project|Visio)'
        }
    )
    [string[]]$office2024DisplayNames = @(
        $displayNames | Where-Object {
            $_ -match '(?i)(?:Microsoft Office|Office LTSC|Office Professional|Office Standard).*2024' -and
            $_ -notmatch '(?i)(?:Project|Visio|Language Pack|Proofing)'
        }
    )

    $legacyProductIdPattern = '(?i)(?:2013|2016|2019|2021)'
    $legacyProductIdsWithoutYear = @(
        'ProPlusRetail','ProfessionalRetail','HomeBusinessRetail','HomeStudentRetail',
        'PersonalRetail','ProPlusVolume','StandardVolume'
    )
    [string[]]$olderClickToRunProductIds = @(
        $productReleaseIds | Where-Object {
            ($_ -match $suiteIdPattern -and $_ -match $legacyProductIdPattern) -or
            ($_ -in $legacyProductIdsWithoutYear)
        } | Sort-Object -Unique
    )
    [string[]]$olderDisplayNames = @(
        $displayNames | Where-Object {
            $_ -match '(?i)(?:Microsoft Office|Office LTSC|Office Professional|Office Standard).*(?:2007|2010|2013|2016|2019|2021)' -and
            $_ -notmatch '(?i)(?:Project|Visio|Language Pack|Proofing|Update)'
        }
    )
    [string[]]$olderMsiDisplayNames = @(
        $uninstallEntries | Where-Object {
            $name = [string](Get-SafeObjectPropertyValue -InputObject $_ -Name 'DisplayName')
            $windowsInstallerValue = Get-SafeObjectPropertyValue -InputObject $_ -Name 'WindowsInstaller'
            $uninstallStringValue = [string](Get-SafeObjectPropertyValue -InputObject $_ -Name 'UninstallString')

            $isOlderSuite =
                -not [string]::IsNullOrWhiteSpace($name) -and
                $name -match '(?i)(?:Microsoft Office|Office LTSC|Office Professional|Office Standard).*(?:2007|2010|2013|2016|2019|2021)' -and
                $name -notmatch '(?i)(?:Project|Visio|Language Pack|Proofing|Update)'

            $isMsi =
                ($null -ne $windowsInstallerValue -and [int]$windowsInstallerValue -eq 1) -or
                ($uninstallStringValue -match '(?i)msiexec(?:\.exe)?')

            $isOlderSuite -and $isMsi
        } | ForEach-Object {
            [string](Get-SafeObjectPropertyValue -InputObject $_ -Name 'DisplayName')
        } | Sort-Object -Unique
    )
    $olderUnknownTypeDetected =
        $olderDisplayNames.Count -gt 0 -and
        $olderClickToRunProductIds.Count -eq 0 -and
        $olderMsiDisplayNames.Count -eq 0

    return [pscustomobject][ordered]@{
        ProductReleaseIds          = @($productReleaseIds)
        OfficeDisplayNames         = @($displayNames)
        Microsoft365Detected       = ($microsoft365ProductIds.Count -gt 0 -or $microsoft365DisplayNames.Count -gt 0)
        Microsoft365ProductIds     = @($microsoft365ProductIds)
        Microsoft365DisplayNames   = @($microsoft365DisplayNames)
        Office2024Detected         = ($office2024ProductIds.Count -gt 0 -or $office2024DisplayNames.Count -gt 0)
        Office2024ProductIds       = @($office2024ProductIds)
        Office2024DisplayNames     = @($office2024DisplayNames)
        OlderOfficeDetected       = ($olderClickToRunProductIds.Count -gt 0 -or $olderDisplayNames.Count -gt 0)
        OlderClickToRunProductIds = @($olderClickToRunProductIds)
        OlderOfficeDisplayNames   = @($olderDisplayNames)
        OlderMsiDisplayNames      = @($olderMsiDisplayNames)
        OlderMsiOrUnknownDetected = ($olderMsiDisplayNames.Count -gt 0 -or $olderUnknownTypeDetected)
    }
}

function Get-OfficeMigrationHistory {
    [CmdletBinding()]
    param()

    $history = [ordered]@{
        SchemaVersion             = '1.0'
        ComputerName              = [string]$env:COMPUTERNAME
        MigratedFromOlderOffice   = $false
        OriginalOlderProducts     = @()
        FirstDetectedOlderUtc     = $null
        LastMigrationAttemptUtc   = $null
        LastMigrationSucceededUtc = $null
        LastFailureStage          = $null
        LastFailureMessage        = $null
    }

    if (Test-Path -LiteralPath $OfficeMigrationStatePath -PathType Leaf) {
        try {
            $saved = Get-Content -LiteralPath $OfficeMigrationStatePath -Raw -ErrorAction Stop |
                ConvertFrom-Json -ErrorAction Stop
            foreach ($name in @($history.Keys)) {
                $property = $saved.PSObject.Properties[$name]
                if ($null -ne $property) { $history[$name] = $property.Value }
            }
        }
        catch {
            Write-RunnerLog -Message "Office migration history could not be read and will be rebuilt: $($_.Exception.Message)" -Level 'WARNING'
        }
    }

    return $history
}

function Save-OfficeMigrationHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary]$History)

    Write-RunnerJsonAtomically `
        -Path $OfficeMigrationStatePath `
        -Json ($History | ConvertTo-Json -Depth 6)
}

function Copy-Office2024InstallerToStage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
        throw "Office 2024 installer source is unavailable: $SourcePath"
    }

    New-Item -Path $DestinationPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    $robocopyPath = Join-Path $env:SystemRoot 'System32\robocopy.exe'
    if (-not (Test-Path -LiteralPath $robocopyPath -PathType Leaf)) {
        throw "Robocopy was not found: $robocopyPath"
    }

    & $robocopyPath $SourcePath $DestinationPath '/E' '/COPY:DAT' '/DCOPY:DAT' '/R:3' '/W:5' '/NFL' '/NDL' '/NJH' '/NJS' '/NP' | Out-Null
    $robocopyExitCode = [int]$LASTEXITCODE
    if ($robocopyExitCode -gt 7) {
        throw "Unable to stage Office 2024 installer files. Robocopy exit code=$robocopyExitCode"
    }

    $setupPath = Join-Path $DestinationPath 'setup.exe'
    $configurationPath = Join-Path $DestinationPath $Office2024ConfigurationFile
    if (-not (Test-Path -LiteralPath $setupPath -PathType Leaf)) {
        throw "The staged Office Deployment Tool was not found: $setupPath"
    }
    if (-not (Test-Path -LiteralPath $configurationPath -PathType Leaf)) {
        throw "The staged Office 2024 configuration file was not found: $configurationPath"
    }

    return [pscustomobject][ordered]@{
        SetupPath        = $setupPath
        ConfigurationPath = $configurationPath
        RobocopyExitCode = $robocopyExitCode
    }
}

function Test-OfficeConfigurationSupportsMsiRemoval {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ConfigurationPath)

    try {
        [xml]$configuration = Get-Content -LiteralPath $ConfigurationPath -Raw -ErrorAction Stop
    }
    catch {
        throw "The Office 2024 configuration file is not valid XML: $ConfigurationPath. $($_.Exception.Message)"
    }

    return ($null -ne $configuration.SelectSingleNode('/Configuration/RemoveMSI'))
}

function New-OlderOfficeRemovalConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$ProductIds,
        [Parameter(Mandatory)][string]$Path
    )

    if ($ProductIds.Count -eq 0) {
        throw 'No older Click-to-Run Office product IDs were supplied for removal.'
    }

    $productLines = @(
        $ProductIds | Sort-Object -Unique | ForEach-Object {
            $escapedId = [Security.SecurityElement]::Escape([string]$_)
            '    <Product ID="{0}" />' -f $escapedId
        }
    )
    $xml = @(
        '<Configuration>'
        '  <Remove>'
        $productLines
        '  </Remove>'
        '  <Display Level="None" AcceptEULA="TRUE" />'
        '</Configuration>'
    ) -join [Environment]::NewLine

    [IO.File]::WriteAllText($Path, $xml, (New-Object Text.UTF8Encoding($false)))
}

function Invoke-OfficeDeploymentTool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SetupPath,
        [Parameter(Mandatory)][string]$ConfigurationPath
    )

    $workingDirectory = Split-Path -Parent $SetupPath
    $configurationName = Split-Path -Leaf $ConfigurationPath
    $process = Start-Process `
        -FilePath $SetupPath `
        -ArgumentList @('/configure', $configurationName) `
        -WorkingDirectory $workingDirectory `
        -Wait `
        -PassThru `
        -ErrorAction Stop

    $exitCode = [int]$process.ExitCode
    if ($exitCode -ne 0 -and $exitCode -ne 3010) {
        throw "Office Deployment Tool failed for $configurationName. Exit code=$exitCode"
    }

    return $exitCode
}

function Get-OfficeOsppPath {
    [CmdletBinding()]
    param()

    $programFilesPaths = @(
        [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles),
        [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($programFilesPath in $programFilesPaths) {
        foreach ($relativePath in @(
            'Microsoft Office\root\Office16\OSPP.VBS',
            'Microsoft Office\Office16\OSPP.VBS'
        )) {
            $candidate = Join-Path $programFilesPath $relativePath
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return $candidate
            }
        }
    }

    return $null
}

function Invoke-OfficeOspp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OsppPath,
        [Parameter(Mandatory)][ValidateSet('/dstatusall','/act')][string]$Argument
    )

    $cscriptPath = Join-Path $env:SystemRoot 'System32\cscript.exe'
    $output = & $cscriptPath '//Nologo' $OsppPath $Argument 2>&1 | Out-String
    return [pscustomobject][ordered]@{
        ExitCode = [int]$LASTEXITCODE
        Output   = [string]$output
    }
}

function Get-Office2024LicenseState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StatusOutput
    )

    $licenseBlocks = @([regex]::Split($StatusOutput, '(?m)^-{20,}\s*$'))
    $office2024Blocks = @(
        $licenseBlocks | Where-Object {
            $_ -match '(?im)^\s*LICENSE (?:NAME|DESCRIPTION):.*(?:Office\s*24|Office\S*2024|Office.*2024)' -and
            $_ -notmatch '(?im)^\s*LICENSE (?:NAME|DESCRIPTION):.*(?:Project|Visio)'
        }
    )
    $licensedBlocks = @(
        $office2024Blocks | Where-Object {
            $_ -match '(?im)^\s*LICENSE STATUS:\s*---LICENSED---\s*$'
        }
    )

    return [pscustomobject][ordered]@{
        LicenseFound = ($office2024Blocks.Count -gt 0)
        IsActivated  = ($licensedBlocks.Count -gt 0)
        LicenseCount = $office2024Blocks.Count
    }
}

function Ensure-Office2024MachineActivation {
    [CmdletBinding()]
    param()

    $osppPath = Get-OfficeOsppPath
    if ([string]::IsNullOrWhiteSpace([string]$osppPath)) {
        throw 'Office 2024 is installed, but OSPP.VBS was not found. This workflow supports volume-licensed Office LTSC 2024.'
    }

    $statusResult = Invoke-OfficeOspp -OsppPath $osppPath -Argument '/dstatusall'
    if ($statusResult.ExitCode -ne 0) {
        throw "Office activation status check failed with exit code $($statusResult.ExitCode)."
    }

    $licenseState = Get-Office2024LicenseState -StatusOutput $statusResult.Output
    if (-not $licenseState.LicenseFound) {
        throw 'Office 2024 is installed, but no Office 2024 volume-license record was returned by OSPP.VBS.'
    }

    if ($licenseState.IsActivated) {
        return [pscustomobject][ordered]@{
            ActivationAttempted = $false
            Activated           = $true
            LicenseCount        = $licenseState.LicenseCount
        }
    }

    Write-RunnerLog -Message 'Office 2024 is not activated. Attempting machine-wide activation for all users.' -Level 'WARNING'
    $activationResult = Invoke-OfficeOspp -OsppPath $osppPath -Argument '/act'
    Start-Sleep -Seconds 5

    $verificationResult = Invoke-OfficeOspp -OsppPath $osppPath -Argument '/dstatusall'
    if ($verificationResult.ExitCode -ne 0) {
        throw "Office activation verification failed with exit code $($verificationResult.ExitCode)."
    }

    $verifiedState = Get-Office2024LicenseState -StatusOutput $verificationResult.Output
    if (-not $verifiedState.IsActivated) {
        $activationError = [regex]::Match(
            [string]$activationResult.Output,
            '(?i)(?:ERROR CODE|error):?\s*(0x[0-9A-F]{8})'
        )
        $errorSuffix = if ($activationError.Success) {
            " Microsoft error code: $($activationError.Groups[1].Value)."
        }
        else {
            " OSPP.VBS exit code: $($activationResult.ExitCode)."
        }
        throw "Office 2024 remained unactivated after the machine-wide activation attempt.$errorSuffix"
    }

    return [pscustomobject][ordered]@{
        ActivationAttempted = $true
        Activated           = $true
        LicenseCount        = $verifiedState.LicenseCount
    }
}

function Invoke-Office2024MaintenanceSection {
    [CmdletBinding()]
    param()

    $sectionStart = Get-Date
    $displayName = 'Office LTSC detection, Office 2024 migration, and activation'
    $exitCode = 0
    $successful = $true
    $skippedNotInstalled = $false
    $skippedMicrosoft365 = $false
    $olderOfficeDetected = $false
    $olderOfficeDetectedThisRun = $false
    $olderOfficeProducts = @()
    $migrationAttempted = $false
    $migrationSucceeded = $false
    $uninstallAttempted = $false
    $uninstallSucceeded = $false
    $installationAttempted = $false
    $installationSucceeded = $false
    $activationAttempted = $false
    $activated = $false
    $office2024Installed = $false
    $microsoft365Detected = $false
    $rebootRequired = $false
    $remediationFailure = $false
    $failureStage = $null
    $failureMessage = $null
    $outcome = 'Unknown'
    $localStagePath = Join-Path ([IO.Path]::GetTempPath()) ("Compton\Office2024\{0}" -f $RunnerRunId)
    $localStageRemoved = $false
    $history = Get-OfficeMigrationHistory

    Write-RunnerLog -Message "Starting section: $displayName"

    try {
        $failureStage = 'DetectOffice'
        $inventory = Get-OfficeInstallationInventory
        $microsoft365Detected = [bool]$inventory.Microsoft365Detected
        $office2024Installed = [bool]$inventory.Office2024Detected
        $olderOfficeDetectedThisRun = [bool]$inventory.OlderOfficeDetected
        $olderOfficeDetected = [bool]$history.MigratedFromOlderOffice -or $olderOfficeDetectedThisRun

        $combinedOlderOfficeProducts = @($inventory.OlderClickToRunProductIds) +
            @($inventory.OlderOfficeDisplayNames) +
            @($history.OriginalOlderProducts)
        $olderOfficeProducts = @(
            $combinedOlderOfficeProducts |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        if ($microsoft365Detected) {
            $skippedMicrosoft365 = $true
            $outcome = 'Microsoft365Skipped'
            Write-RunnerLog -Message 'Microsoft 365 Apps was detected. Office LTSC replacement and volume activation were skipped.' -Level 'SUCCESS'
        }
        elseif ($olderOfficeDetectedThisRun) {
            $migrationAttempted = $true
            $history.MigratedFromOlderOffice = $true
            if ([string]::IsNullOrWhiteSpace([string]$history.FirstDetectedOlderUtc)) {
                $history.FirstDetectedOlderUtc = (Get-Date).ToUniversalTime().ToString('o')
            }
            $history.LastMigrationAttemptUtc = (Get-Date).ToUniversalTime().ToString('o')
            $history.OriginalOlderProducts = @($olderOfficeProducts)
            Save-OfficeMigrationHistory -History $history

            $failureStage = 'StageOffice2024Installer'
            Write-RunnerLog -Message "Staging Office 2024 installation files from $Office2024SourcePath to $localStagePath"
            $stagedInstaller = Copy-Office2024InstallerToStage `
                -SourcePath $Office2024SourcePath `
                -DestinationPath $localStagePath

            $uninstallAttempted = $true
            if ([bool]$inventory.OlderMsiOrUnknownDetected) {
                $failureStage = 'ValidateOffice2024Configuration'
                if (-not (Test-OfficeConfigurationSupportsMsiRemoval -ConfigurationPath $stagedInstaller.ConfigurationPath)) {
                    throw "An older MSI-based Office suite was detected, but $Office2024ConfigurationFile does not contain <RemoveMSI />. Add that element to the Office 2024 configuration before retrying."
                }
                Write-RunnerLog -Message "Older MSI-based Office was detected. It will be removed by the <RemoveMSI /> element during the Office 2024 installation." -Level 'WARNING'
            }

            if (@($inventory.OlderClickToRunProductIds).Count -gt 0) {
                $failureStage = 'UninstallOlderOffice'
                $removalConfigurationPath = Join-Path $localStagePath 'remove-older-office.xml'
                New-OlderOfficeRemovalConfiguration `
                    -ProductIds @($inventory.OlderClickToRunProductIds) `
                    -Path $removalConfigurationPath
                Write-RunnerLog -Message "Removing older Office Click-to-Run product IDs: $(@($inventory.OlderClickToRunProductIds) -join ', ')" -Level 'WARNING'
                $removeExitCode = Invoke-OfficeDeploymentTool `
                    -SetupPath $stagedInstaller.SetupPath `
                    -ConfigurationPath $removalConfigurationPath
                if ($removeExitCode -eq 3010) { $rebootRequired = $true }

                $afterRemoval = Get-OfficeInstallationInventory
                if (@($afterRemoval.OlderClickToRunProductIds).Count -gt 0) {
                    throw "Older Office Click-to-Run products remained after the removal command: $(@($afterRemoval.OlderClickToRunProductIds) -join ', ')"
                }
            }

            $failureStage = 'InstallOffice2024'
            $installationAttempted = $true
            Write-RunnerLog -Message "Installing Office 2024 with: setup.exe /configure $Office2024ConfigurationFile"
            $installExitCode = Invoke-OfficeDeploymentTool `
                -SetupPath $stagedInstaller.SetupPath `
                -ConfigurationPath $stagedInstaller.ConfigurationPath
            if ($installExitCode -eq 3010) { $rebootRequired = $true }

            $afterInstallation = Get-OfficeInstallationInventory
            if (-not $afterInstallation.Office2024Detected) {
                throw 'The Office Deployment Tool completed, but Office 2024 was not detected afterward.'
            }
            if ($afterInstallation.OlderOfficeDetected) {
                throw "Office 2024 was installed, but an older Office suite remained detected: $(@($afterInstallation.OlderOfficeDisplayNames + $afterInstallation.OlderClickToRunProductIds) -join ', ')"
            }
            $uninstallSucceeded = $true
            $installationSucceeded = $true
            $office2024Installed = $true

            $failureStage = 'ActivateOffice2024'
            $activation = Ensure-Office2024MachineActivation
            $activationAttempted = [bool]$activation.ActivationAttempted
            $activated = [bool]$activation.Activated
            $migrationSucceeded = $true
            $outcome = 'OlderOfficeMigrated'
            $history.LastMigrationSucceededUtc = (Get-Date).ToUniversalTime().ToString('o')
            $history.LastFailureStage = $null
            $history.LastFailureMessage = $null
            Save-OfficeMigrationHistory -History $history
            Write-RunnerLog -Message 'Older Office was removed, Office 2024 was installed, and machine-wide activation was verified.' -Level 'SUCCESS'
        }
        elseif (-not $office2024Installed) {
            $skippedNotInstalled = $true
            $outcome = 'NoApplicableOfficeFound'
            Write-RunnerLog -Message 'No Microsoft 365, Office 2024, or supported older Office LTSC/perpetual suite was detected. No change was made.'
        }
        else {
            $failureStage = 'ActivateOffice2024'
            $activation = Ensure-Office2024MachineActivation
            $activationAttempted = [bool]$activation.ActivationAttempted
            $activated = [bool]$activation.Activated
            $outcome = 'Office2024Compliant'
            Write-RunnerLog -Message 'Office 2024 is installed and machine-wide activation is verified.' -Level 'SUCCESS'
        }
    }
    catch {
        $exitCode = 1
        $successful = $false
        $remediationFailure = $true
        $failureMessage = [string]$_.Exception.Message
        $outcome = switch ($failureStage) {
            'ValidateOffice2024Configuration' { 'Office2024ConfigurationInvalid' }
            'UninstallOlderOffice' { 'OlderOfficeUninstallFailed' }
            'InstallOffice2024'    { 'Office2024InstallationFailed' }
            'ActivateOffice2024'   { 'Office2024ActivationFailed' }
            default                { 'Office2024MaintenanceFailed' }
        }
        $history.LastFailureStage = $failureStage
        $history.LastFailureMessage = $failureMessage
        try { Save-OfficeMigrationHistory -History $history } catch {}
        Write-RunnerLog -Message "Office 2024 maintenance failed during $failureStage`: $failureMessage" -Level 'ERROR'
    }
    finally {
        if (Test-Path -LiteralPath $localStagePath -PathType Container) {
            try {
                Remove-Item -LiteralPath $localStagePath -Recurse -Force -ErrorAction Stop
                $localStageRemoved = -not (Test-Path -LiteralPath $localStagePath)
                Write-RunnerLog -Message "Removed local Office 2024 staging directory: $localStagePath" -Level 'SUCCESS'
            }
            catch {
                $cleanupError = "Unable to remove local Office 2024 staging directory $localStagePath`: $($_.Exception.Message)"
                Write-RunnerLog -Message $cleanupError -Level 'ERROR'
                if ($successful) {
                    $successful = $false
                    $exitCode = 1
                    $remediationFailure = $true
                    $failureStage = 'CleanupOffice2024Stage'
                    $failureMessage = $cleanupError
                    $outcome = 'Office2024CleanupFailed'
                }
            }
        }
        else {
            $localStageRemoved = $true
        }
    }

    $duration = [math]::Round(((Get-Date) - $sectionStart).TotalSeconds, 2)
    $eventTime = Get-Date
    $officeEvent = [ordered]@{
        '@timestamp'    = $eventTime.ToUniversalTime().ToString('o')
        EventType       = 'maintenance.office2024_compliance'
        SchemaVersion   = '1.0'
        ComputerName    = [string]$env:COMPUTERNAME
        ScriptName      = $RunnerScriptName
        ScriptVersion   = $RunnerVersion
        RunId           = $RunnerRunId
        Status          = if ($successful) { 'Success' } else { 'Failed' }
        Outcome         = $outcome
        ExitCode        = $exitCode
        StartTime       = $sectionStart.ToUniversalTime().ToString('o')
        EndTime         = $eventTime.ToUniversalTime().ToString('o')
        DurationSeconds = $duration
        FailureStage    = $failureStage
        FailureMessage  = $failureMessage
        Office = [ordered]@{
            Microsoft365Detected       = $microsoft365Detected
            Microsoft365Skipped        = $skippedMicrosoft365
            Office2024Installed        = $office2024Installed
            OlderOfficeDetectedThisRun = $olderOfficeDetectedThisRun
            MigratedFromOlderOffice    = $olderOfficeDetected
            OlderOfficeProducts        = @($olderOfficeProducts)
            MigrationAttempted         = $migrationAttempted
            MigrationSucceeded         = $migrationSucceeded
            UninstallAttempted         = $uninstallAttempted
            UninstallSucceeded         = $uninstallSucceeded
            InstallationAttempted      = $installationAttempted
            InstallationSucceeded      = $installationSucceeded
            ActivationAttempted        = $activationAttempted
            Activated                  = $activated
            ActivationScope            = 'MachineAllUsers'
            RemediationFailure         = $remediationFailure
            LocalStageRemoved          = $localStageRemoved
            SourcePath                 = $Office2024SourcePath
            ConfigurationFile          = $Office2024ConfigurationFile
        }
    }

    try {
        Write-RunnerTelemetryLine `
            -Path $RunnerTelemetryPath `
            -JsonLine ($officeEvent | ConvertTo-Json -Depth 8 -Compress)
        Write-RunnerJsonAtomically `
            -Path (Join-Path $LogDirectory '04_Office2024_Maintenance.latest.json') `
            -Json ($officeEvent | ConvertTo-Json -Depth 8)
    }
    catch {
        Write-RunnerLog -Message "Unable to publish Office 2024 dashboard telemetry: $($_.Exception.Message)" -Level 'ERROR'
        if ($successful) {
            $successful = $false
            $exitCode = 1
            $failureStage = 'PublishOfficeTelemetry'
            $failureMessage = [string]$_.Exception.Message
        }
    }

    Write-RunnerLog `
        -Message "Completed section: $displayName. ExitCode=$exitCode; DurationSeconds=$duration" `
        -Level $(if ($successful) { 'SUCCESS' } else { 'ERROR' })

    return [pscustomobject][ordered]@{
        SectionId           = 'Office2024Maintenance'
        DisplayName         = $displayName
        ExitCode            = $exitCode
        Success             = $successful
        RebootRequired      = $rebootRequired
        DurationSeconds     = $duration
        SkippedNotInstalled = $skippedNotInstalled
        SkippedMicrosoft365 = $skippedMicrosoft365
        Office2024Installed = $office2024Installed
        OlderOfficeDetected = $olderOfficeDetected
        MigrationAttempted  = $migrationAttempted
        MigrationSucceeded  = $migrationSucceeded
        InstallationAttempted = $installationAttempted
        InstallationSucceeded = $installationSucceeded
        ActivationAttempted = $activationAttempted
        Activated           = $activated
        Outcome             = $outcome
        FailureStage        = $failureStage
        FailureMessage      = $failureMessage
    }
}

function Resolve-PrinterSectionExitCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$ProcessExitCode,
        [Parameter(Mandatory)][string]$LogDirectory
    )

    # Script 11 intentionally returns Windows Installer code 3010 when the
    # printer/PaperCut configuration is compliant but a reboot is pending.
    # In some Windows PowerShell 5.1 child-process launches that code can be
    # surfaced to the parent as a generic nonzero value.  The child also
    # writes an authoritative latest.json result, so use that telemetry to
    # preserve SuccessRebootRequired instead of incorrectly marking failure.
    $latestJson = Join-Path $LogDirectory '11_Install_SharpDriver_And_PaperCut.latest.json'

    if (-not (Test-Path -LiteralPath $latestJson -PathType Leaf)) {
        return $ProcessExitCode
    }

    try {
        $result = Get-Content -LiteralPath $latestJson -Raw -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop

        $status = [string]$result.Status
        $reportedExitCode = 0
        $hasReportedExitCode = $false

        if ($null -ne $result.PSObject.Properties['ExitCode']) {
            $reportedExitCode = [int]$result.ExitCode
            $hasReportedExitCode = $true
        }

        $rebootRequired = $false
        if ($null -ne $result.PSObject.Properties['RebootRequired']) {
            $rebootRequired = [bool]$result.RebootRequired
        }

        if (
            $status -eq 'SuccessRebootRequired' -or
            ($hasReportedExitCode -and $reportedExitCode -eq 3010) -or
            ($status -match '^Success' -and $rebootRequired)
        ) {
            return 3010
        }

        if ($status -match '^Success' -and $ProcessExitCode -ne 0) {
            # The child telemetry says the work succeeded and no reboot is
            # required. Normalize a transport/launcher-only nonzero code.
            return 0
        }
    }
    catch {
        Write-RunnerLog `
            -Message "Unable to read Script 11 latest telemetry for exit-code normalization: $($_.Exception.Message)" `
            -Level 'WARNING'
    }

    return $ProcessExitCode
}

function Invoke-MaintenanceSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SectionId,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$StagedScriptPath,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Settings
    )

    $sectionStart = Get-Date
    Write-RunnerLog -Message "Starting section: $DisplayName"

    $previousScriptPath = $env:COMPTON_SECTION_SCRIPT
    $previousSettings = $env:COMPTON_SECTION_SETTINGS
    try {
        if ($SectionId -eq 'Honorlock') {
            Set-HonorlockChromeMachinePolicy -ComputerNamePatterns @($Settings.Patterns)
        }

        $env:COMPTON_SECTION_SCRIPT = $StagedScriptPath
        $env:COMPTON_SECTION_SETTINGS = $Settings | ConvertTo-Json -Depth 5 -Compress
        $bootstrap = New-SectionBootstrap -SectionId $SectionId
        $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($bootstrap))

        $process = Start-Process `
            -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
            -ArgumentList @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-EncodedCommand',$encodedCommand) `
            -Wait `
            -PassThru

        $exitCode = [int]$process.ExitCode

        if ($SectionId -eq 'PrinterAndPaperCut') {
            $normalizedExitCode = Resolve-PrinterSectionExitCode `
                -ProcessExitCode $exitCode `
                -LogDirectory $LogDirectory

            if ($normalizedExitCode -ne $exitCode) {
                Write-RunnerLog `
                    -Message "Normalized SHARP/PaperCut child exit code from $exitCode to $normalizedExitCode based on Script 11 telemetry." `
                    -Level 'INFO'
                $exitCode = $normalizedExitCode
            }
        }
    }
    catch {
        $exitCode = 1
        Write-RunnerLog -Message "Section launch failed: $DisplayName. $($_.Exception.Message)" -Level 'ERROR'
    }
    finally {
        $env:COMPTON_SECTION_SCRIPT = $previousScriptPath
        $env:COMPTON_SECTION_SETTINGS = $previousSettings
    }

    $duration = [math]::Round(((Get-Date) - $sectionStart).TotalSeconds, 2)
    $successful = ($exitCode -eq 0 -or $exitCode -eq 3010)
    Write-RunnerLog `
        -Message "Completed section: $DisplayName. ExitCode=$exitCode; DurationSeconds=$duration" `
        -Level $(if ($successful) { 'SUCCESS' } else { 'ERROR' })

    return [pscustomobject][ordered]@{
        SectionId       = $SectionId
        DisplayName     = $DisplayName
        ExitCode        = $exitCode
        Success         = $successful
        RebootRequired  = ($exitCode -eq 3010)
        DurationSeconds = $duration
    }
}

function Add-SkippedSectionResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.Generic.List[object]]$Results,
        [Parameter(Mandatory)][string]$SectionId,
        [Parameter(Mandatory)][string]$DisplayName
    )

    Write-RunnerLog -Message "Section disabled in runner configuration: $DisplayName" -Level 'WARNING'
    [void]$Results.Add([pscustomobject][ordered]@{
        SectionId       = $SectionId
        DisplayName     = $DisplayName
        ExitCode        = $null
        Success         = $true
        RebootRequired  = $false
        DurationSeconds = 0
        SkippedDisabled = $true
    })
}

$results = New-Object System.Collections.Generic.List[object]
$stagingRoot = Join-Path $env:ProgramData ("Compton\SundayMaintenance\{0}" -f $RunnerRunId)
$finalExitCode = 1
$finalStatus = 'NotStarted'

try {
    if (-not (Test-RunnerAdministrator)) {
        throw 'This script must run as Administrator or SYSTEM.'
    }
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        throw 'Run this script with 64-bit Windows PowerShell.'
    }

    New-Item -Path $stagingRoot -ItemType Directory -Force | Out-Null
    Write-RunnerLog -Message "Starting $RunnerScriptName version $RunnerVersion. RunId=$RunnerRunId"

    foreach ($payloadEntry in $EmbeddedSections.GetEnumerator()) {
        $destination = Join-Path $stagingRoot ([string]$payloadEntry.Value.FileName)
        Write-EmbeddedSectionSource -Payload $payloadEntry.Value -Destination $destination
    }

    $sectionPlan = @(
        [pscustomobject]@{ Id='SystemRestore'; Name='System Restore and verified restore point'; Enabled=$RunSystemRestore; Settings=[ordered]@{} }
        [pscustomobject]@{ Id='PrinterAndPaperCut'; Name='SHARP driver, PaperCut, and printer'; Enabled=$RunPrinterAndPaperCut; Settings=[ordered]@{} }
        [pscustomobject]@{ Id='Office2024Maintenance'; Name='Office LTSC detection, Office 2024 migration, and activation'; Enabled=$RunOffice2024Maintenance; Settings=[ordered]@{} }
        [pscustomobject]@{ Id='AutologonAndEdge'; Name='Autologon and Edge startup'; Enabled=$RunAutologonAndEdge; Settings=[ordered]@{ Patterns=@($AutologonComputerPatterns) } }
        [pscustomobject]@{ Id='ElasticAgent'; Name='Elastic Agent installation'; Enabled=$RunElasticAgent; Settings=[ordered]@{ Prefixes=@($ElasticAgentComputerPrefixes); ForceReinstall=[bool]$ForceElasticAgentReinstall } }
        [pscustomobject]@{ Id='BrowserHomepage'; Name='Browser homepage policies'; Enabled=$RunBrowserHomepage; Settings=[ordered]@{ HomepageUrl=$HomepageUrl } }
        [pscustomobject]@{ Id='Honorlock'; Name='Honorlock Chrome extension'; Enabled=$RunHonorlock; Settings=[ordered]@{ Patterns=@($HonorlockComputerPatterns) } }
        [pscustomobject]@{ Id='StellariumLocation'; Name='Stellarium Location Services'; Enabled=$RunStellariumLocation; Settings=[ordered]@{ Patterns=@($StellariumComputerPatterns) } }
    )

    foreach ($section in $sectionPlan) {
        if (-not [bool]$section.Enabled) {
            Add-SkippedSectionResult -Results $results -SectionId $section.Id -DisplayName $section.Name
            continue
        }

        if ($section.Id -eq 'Office2024Maintenance') {
            [void]$results.Add((Invoke-Office2024MaintenanceSection))
            continue
        }

        $sectionPath = Join-Path $stagingRoot ([string]$EmbeddedSections[$section.Id].FileName)
        $sectionResult = Invoke-MaintenanceSection `
            -SectionId $section.Id `
            -DisplayName $section.Name `
            -StagedScriptPath $sectionPath `
            -Settings $section.Settings
        [void]$results.Add($sectionResult)
    }

    [object[]]$resultArray = $results.ToArray()
    [object[]]$failedSections = @($resultArray | Where-Object { -not $_.Success })
    $rebootRequired = @($resultArray | Where-Object { $_.RebootRequired }).Count -gt 0

    if ($failedSections.Count -gt 0) {
        $finalExitCode = 1
        $finalStatus = 'Failed'
    }
    elseif ($rebootRequired) {
        $finalExitCode = 3010
        $finalStatus = 'SuccessRebootRequired'
    }
    else {
        $finalExitCode = 0
        $finalStatus = 'Success'
    }
}
catch {
    $finalStatus = 'RunnerFailed'
    $finalExitCode = 1
    Write-RunnerLog -Message "Combined runner failed: $($_.Exception.Message)" -Level 'ERROR'
}
finally {
    $runnerEndTime = Get-Date
    $summary = [pscustomobject][ordered]@{
        EventType       = 'maintenance.combined_execution'
        ComputerName    = [string]$env:COMPUTERNAME
        ScriptName      = $RunnerScriptName
        ScriptVersion   = $RunnerVersion
        RunId           = $RunnerRunId
        Status          = $finalStatus
        ExitCode        = $finalExitCode
        StartTime       = $RunnerStartTime.ToUniversalTime().ToString('o')
        EndTime         = $runnerEndTime.ToUniversalTime().ToString('o')
        DurationSeconds = [math]::Round(($runnerEndTime - $RunnerStartTime).TotalSeconds, 2)
        Sections        = $results.ToArray()
        Timestamp       = $runnerEndTime.ToUniversalTime().ToString('o')
    }

    try {
        if (-not (Test-Path -LiteralPath $LogDirectory -PathType Container)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }
        $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $RunnerLatestPath -Encoding UTF8 -Force
    }
    catch {
        Write-Warning "Unable to write the combined runner summary: $($_.Exception.Message)"
        if ($finalExitCode -eq 0 -or $finalExitCode -eq 3010) { $finalExitCode = 4 }
    }

    if (Test-Path -LiteralPath $stagingRoot -PathType Container) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    try {
        Write-RunnerLog `
            -Message "Completed $RunnerScriptName. Status=$finalStatus; ExitCode=$finalExitCode; Sections=$($results.Count)" `
            -Level $(if ($finalExitCode -eq 0 -or $finalExitCode -eq 3010) { 'SUCCESS' } else { 'ERROR' })
    }
    catch {}
}

exit $finalExitCode
