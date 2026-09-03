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
