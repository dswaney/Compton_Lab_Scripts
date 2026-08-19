# ScriptVersion: 2.1.8
# LastUpdated: 2026-08-18
# Changes: v2.1.8 fixes Windows PowerShell 5.1 generic-list array conversion in telemetry.
#          v2.1.7 fixes telemetry parameter binding for ordered dictionaries and
#          routes NDJSON writes through Maintenance.Framework to avoid file-lock issues.
#          v2.1.6 adds Elastic-ready per-application update telemetry: updated count,
#          before/after versions, failures, remaining updates, explicit-target packages,
#          and packages blocked/deferred by WinGet pins.
#          v2.1.5 adds Maintenance.Framework v2.4 staged text logging.

[CmdletBinding()]
param(
    [switch]$IncludeUnknown = $true,
    [switch]$IncludePinned = $false,

    [switch]$AttemptMSStore = $false,
    [switch]$UpdateOffice = $true,
    [int]$OfficeWaitMinutes = 30,
    [string]$LogPath = "$env:SystemDrive\Logs\03_Weekend_Apps_Update.log"
)

# Packages deliberately deferred from unattended targeted retries.
# App Installer contains WinGet itself and can fail when attempting to update
# the servicing component that is currently executing the maintenance run.
$SpecialHandlingPackageIds = @(
    'Microsoft.AppInstaller'
)

$ErrorActionPreference = 'Stop'

$ScriptVersion = '2.1.8'
$ScriptName = '03_Weekend_Apps_Update.ps1'
$ComputerName = $env:COMPUTERNAME
$ExecutionStart = Get-Date
$RunId = [guid]::NewGuid().Guid
$WarningCount = 0
$ErrorCount = 0
$ApplicationResults = New-Object System.Collections.Generic.List[object]
$TelemetryDirectory = Join-Path $env:SystemDrive 'Logs'
$TelemetryNdjsonPath = Join-Path $TelemetryDirectory 'Maintenance-Telemetry.ndjson'
$TelemetryLatestPath = Join-Path $TelemetryDirectory '03_Weekend_Apps_Update.latest.json'

$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $TelemetryDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 03 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

# The legacy -LogPath parameter remains accepted for compatibility, but staged
# logging now controls the active and final per-run text log locations.
Archive-MaintenanceLogs `
    -ScriptName $ScriptName `
    -LogRoot $TelemetryDirectory `
    -AdditionalPatterns @(
        '03_Weekend_Apps_Update.log',
        '*-03_Weekend_Apps_Update-*.log'
    ) | Out-Null

$TextLogSession = New-MaintenanceStagedLog `
    -ScriptName $ScriptName `
    -LogRoot $TelemetryDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $ComputerName `
    -Timestamp $ExecutionStart

$LogPath = [string]$TextLogSession.WorkingPath
$PublishedLogPath = [string]$TextLogSession.PublishedPath

# Writes color-coded status output to the console and appends the same message to the log file.
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','SUCCESS','WARNING')]
        [string]$Level = 'INFO'
    )

    $normalizedLevel = switch ($Level) {
        'OK'   { 'SUCCESS' }
        'WARN' { 'WARNING' }
        default { $Level }
    }

    if ($normalizedLevel -eq 'WARNING') { $script:WarningCount++ }
    if ($normalizedLevel -eq 'ERROR') { $script:ErrorCount++ }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$ComputerName] [$normalizedLevel] $Message"

    switch ($normalizedLevel) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
    }

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    }
    catch {
    }
}

# Confirms the script is running elevated because package updates and system changes require admin rights.
function Test-IsAdministrator {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

# Finds winget.exe from the PATH first, then falls back to the common App Installer locations.
function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $paths = @(
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe",
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    )

    foreach ($pattern in $paths) {
        $match = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($match) {
            return $match.FullName
        }
    }

    return $null
}

# Filters out progress bars, separators, and other noisy winget output that does not help with logging or parsing.
function Test-IsNoiseLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return $true }

    $trimmed = $Line.Trim()

    if ($trimmed -match '^[\-/\\|]+$') { return $true }
    if ($trimmed -match 'Γû') { return $true }
    if ($trimmed -match '^\d+%$') { return $true }
    if ($trimmed -match '^\d+(\.\d+)?\s*(KB|MB|GB)\s*/\s*\d+(\.\d+)?\s*(KB|MB|GB)$') { return $true }

    return $false
}

# Normalizes raw winget output into cleaner line-based text so it can be logged and parsed consistently.
function Get-CleanWingetOutput {
    param([object[]]$RawOutput)

    $clean = @()

    foreach ($line in $RawOutput) {
        if ($null -eq $line) { continue }

        $text = [string]$line
        $text = $text -replace "`r", ''
        $text = $text.TrimEnd()

        if (Test-IsNoiseLine -Line $text) { continue }

        $clean += $text
    }

    return @($clean)
}

# Runs winget with the supplied arguments, captures output, logs it, and returns a structured result object.
function Invoke-Winget {
    param(
        [Parameter(Mandatory)]
        [string[]]$Arguments,
        [switch]$IgnoreExitCode
    )

    $wingetPath = Get-WingetPath
    if (-not $wingetPath) {
        throw "winget.exe was not found on this system."
    }

    $display = $Arguments -join ' '
    Write-Log "Running: winget $display" 'INFO'

    $rawOutput = & $wingetPath @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    $cleanOutput = Get-CleanWingetOutput -RawOutput @($rawOutput)

    foreach ($line in $cleanOutput) {
        Write-Log $line 'INFO'
    }

    if (-not $IgnoreExitCode -and $exitCode -ne 0) {
        throw "winget exited with code $exitCode while running: $display"
    }

    return [PSCustomObject]@{
        ExitCode  = $exitCode
        RawOutput = @($rawOutput)
        Output    = @($cleanOutput)
    }
}

# Enables modern TLS defaults and raises the connection limit to reduce network-related package source issues.
function Initialize-NetworkDefaults {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        catch {}
    }

    try {
        [Net.ServicePointManager]::DefaultConnectionLimit = 64
    }
    catch {}
}

# Classic context-menu configuration was moved to 02_Remove_User_Profiles.ps1.


# Returns the installed version for one package after remediation.
# This is used to record a real "after" version instead of assuming the target version installed.
function Get-InstalledPackageVersion {
    param(
        [Parameter(Mandatory)][string]$Id,
        [string]$Source = 'winget'
    )

    try {
        $args = @('list', '--id', $Id, '--exact', '--disable-interactivity')
        if (-not [string]::IsNullOrWhiteSpace($Source)) {
            $args += @('--source', $Source)
        }

        $result = Invoke-Winget -Arguments $args -IgnoreExitCode
        if ($result.ExitCode -ne 0) { return $null }

        $lines = @($result.Output)
        $header = $null
        $headerIndex = -1

        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ([string]$lines[$i] -match '^Name\s+Id\s+Version(?:\s+Available)?(?:\s+Source)?$') {
                $header = [string]$lines[$i]
                $headerIndex = $i
                break
            }
        }

        if ($null -eq $header -or $headerIndex -lt 0) { return $null }

        $idIndex = $header.IndexOf('Id')
        $versionIndex = $header.IndexOf('Version', $idIndex + 2)
        $availableIndex = $header.IndexOf('Available', $versionIndex + 7)
        $sourceIndex = $header.IndexOf('Source', $versionIndex + 7)

        foreach ($line in $lines[($headerIndex + 1)..($lines.Count - 1)]) {
            $raw = [string]$line
            if ([string]::IsNullOrWhiteSpace($raw) -or $raw.Trim() -match '^-{3,}') { continue }
            if ($raw.Length -le $versionIndex) { continue }

            try {
                $rowId = $raw.Substring($idIndex, $versionIndex - $idIndex).Trim()
                if ($rowId -ne $Id) { continue }

                $endIndex = $raw.Length
                if ($availableIndex -gt $versionIndex) { $endIndex = $availableIndex }
                elseif ($sourceIndex -gt $versionIndex) { $endIndex = $sourceIndex }

                return $raw.Substring($versionIndex, $endIndex - $versionIndex).Trim()
            }
            catch { }
        }
    }
    catch { }

    return $null
}

# Collects all upgradeable packages including pinned packages. Comparing this
# inventory to the normal inventory lets us identify upgrades that WinGet hides
# because they are pinned when -IncludePinned is not enabled.
function Get-UpgradeInventoryIncludingPinned {
    param([string]$Source = 'winget')

    $args = @('upgrade', '--source', $Source, '--include-pinned')
    if ($IncludeUnknown) { $args += '--include-unknown' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

# Refreshes winget package sources before inventory collection and upgrades are attempted.
function Update-WingetSources {
    Write-Log "Refreshing WinGet sources..." 'INFO'
    $null = Invoke-Winget -Arguments @('source', 'update') -IgnoreExitCode
    Write-Log "WinGet source refresh completed." 'OK'
}

# Retrieves the current list of upgradeable packages from a specific winget source.
function Get-UpgradeInventory {
    param([string]$Source = 'winget')

    $args = @('upgrade', '--source', $Source)
    if ($IncludeUnknown) { $args += '--include-unknown' }
    if ($IncludePinned)  { $args += '--include-pinned' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

# Parses winget inventory text into normal upgrades and packages that require explicit targeting.
function Parse-WingetInventory {
    param([string[]]$Lines)

    $mainPackages = @()
    $explicitPackages = @()
    $mode = 'None'
    $columns = $null

    foreach ($line in $Lines) {
        if ($null -eq $line) { continue }

        $raw = [string]$line
        $trimmed = $raw.Trim()

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -match 'require explicit targeting for upgrade' -or
            $trimmed -match 'need to be explicitly upgraded') {
            $mode = 'Explicit'
            $columns = $null
            continue
        }

        # WinGet aligns its output as a fixed-width table. Derive the actual
        # column boundaries from the header instead of assuming two spaces
        # separate every field. This handles package/version values whose
        # spacing changes between WinGet releases.
        if ($trimmed -match '^Name\s+Id\s+Version\s+Available(?:\s+Source)?$') {
            $nameIndex = $raw.IndexOf('Name')
            $idIndex = $raw.IndexOf('Id', $nameIndex + 4)
            $versionIndex = $raw.IndexOf('Version', $idIndex + 2)
            $availableIndex = $raw.IndexOf('Available', $versionIndex + 7)
            $sourceIndex = $raw.IndexOf('Source', $availableIndex + 9)

            if ($nameIndex -ge 0 -and $idIndex -gt $nameIndex -and
                $versionIndex -gt $idIndex -and $availableIndex -gt $versionIndex) {

                $columns = [PSCustomObject]@{
                    Name      = $nameIndex
                    Id        = $idIndex
                    Version   = $versionIndex
                    Available = $availableIndex
                    Source    = $sourceIndex
                }

                if ($mode -eq 'None') {
                    $mode = 'Main'
                }
            }

            continue
        }

        if ($trimmed -match '^-{3,}') { continue }
        if ($trimmed -match '^\d+\s+upgrades available\.?$') { continue }

        if ($trimmed -match '^Installing dependencies:$') { continue }
        if ($trimmed -match '^This package requires the following dependencies:$') { continue }
        if ($trimmed -match '^- Packages$') { continue }
        if ($trimmed -match '^\(\d+/\d+\)\s+Found ') { continue }
        if ($trimmed -match '^Found .+ Version .+$') { continue }
        if ($trimmed -match '^This application is licensed to you by its owner\.$') { continue }
        if ($trimmed -match '^Microsoft is not responsible for, nor does it grant any licenses to, third-party packages\.$') { continue }
        if ($trimmed -match '^Downloading ') { continue }
        if ($trimmed -match '^Successfully verified installer hash$') { continue }
        if ($trimmed -match '^Starting package install\.\.\.$') { continue }
        if ($trimmed -match '^Successfully installed$') { continue }
        if ($trimmed -match '^No installed package found matching input criteria\.$') { continue }
        if ($trimmed -match '^A newer version was found, but the install technology is different from the current version installed\..+$') { continue }

        if ($null -eq $columns -or $mode -eq 'None') {
            continue
        }

        # A package row must extend into the Available column.
        if ($raw.Length -le $columns.Available) {
            continue
        }

        try {
            $name = $raw.Substring(
                $columns.Name,
                $columns.Id - $columns.Name
            ).Trim()

            $id = $raw.Substring(
                $columns.Id,
                $columns.Version - $columns.Id
            ).Trim()

            $version = $raw.Substring(
                $columns.Version,
                $columns.Available - $columns.Version
            ).Trim()

            if ($columns.Source -gt $columns.Available -and
                $raw.Length -gt $columns.Source) {

                $available = $raw.Substring(
                    $columns.Available,
                    $columns.Source - $columns.Available
                ).Trim()

                $source = $raw.Substring($columns.Source).Trim()
            }
            else {
                $available = $raw.Substring($columns.Available).Trim()
                $source = ''
            }

            if ([string]::IsNullOrWhiteSpace($name) -or
                [string]::IsNullOrWhiteSpace($id) -or
                [string]::IsNullOrWhiteSpace($version) -or
                [string]::IsNullOrWhiteSpace($available)) {
                continue
            }

            # Guard against non-package status text being interpreted as a row.
            if ($id -notmatch '^[A-Za-z0-9][A-Za-z0-9._+\-]+$') {
                continue
            }

            $obj = [PSCustomObject]@{
                Name      = $name
                Id        = $id
                Version   = $version
                Available = $available
                Source    = $source
            }

            if ($mode -eq 'Explicit') {
                $explicitPackages += $obj
            }
            elseif ($mode -eq 'Main') {
                $mainPackages += $obj
            }
        }
        catch {
            Write-Log "Unable to parse WinGet inventory row: $trimmed" 'WARN'
        }
    }

    return [PSCustomObject]@{
        Main     = $mainPackages
        Explicit = $explicitPackages
    }
}

# Performs a bulk upgrade pass against the requested winget source using non-interactive switches.
function Invoke-WingetUpgradeAll {
    param([string]$Source = 'winget')

    $args = @(
        'upgrade', '--all',
        '--source', $Source,
        '--silent',
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--disable-interactivity'
    )

    if ($IncludeUnknown) { $args += '--include-unknown' }
    if ($IncludePinned)  { $args += '--include-pinned' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

# Upgrades one package by ID first, then optionally retries by package name if the ID-based attempt fails.
function Invoke-TargetedUpgrade {
    param(
        [Parameter(Mandatory)]$Package,
        [switch]$UseNameFallback
    )

    $args = @(
        'upgrade',
        '--id', $Package.Id,
        '--exact',
        '--silent',
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--disable-interactivity'
    )

    if ($Package.Version -eq 'Unknown' -or $IncludeUnknown) {
        $args += '--include-unknown'
    }

    $targetSource = if (-not [string]::IsNullOrWhiteSpace($Package.Source)) {
        $Package.Source
    }
    else {
        'winget'
    }
    $args += @('--source', $targetSource)

    $result = Invoke-Winget -Arguments $args -IgnoreExitCode

    if ($result.ExitCode -eq 0) {
        return $result
    }

    $text = ($result.Output -join "`n")
    if ($text -match 'install technology is different from the current version installed') {
        Write-Log "Package $($Package.Id) cannot be upgraded in-place because Winget reports the installed technology differs from the new package. Manual uninstall/reinstall is required." 'WARN'
        return $result
    }

    if ($UseNameFallback) {
        Write-Log "ID-targeted upgrade failed for $($Package.Id). Trying name fallback for $($Package.Name)." 'WARN'

        $fallbackArgs = @(
            'upgrade',
            '--name', $Package.Name,
            '--exact',
            '--silent',
            '--accept-source-agreements',
            '--accept-package-agreements',
            '--disable-interactivity'
        )

        if ($Package.Version -eq 'Unknown' -or $IncludeUnknown) {
            $fallbackArgs += '--include-unknown'
        }

        $fallbackArgs += @('--source', $targetSource)

        return (Invoke-Winget -Arguments $fallbackArgs -IgnoreExitCode)
    }

    return $result
}

# Loops through a package collection and performs targeted upgrade attempts while tracking failures.
function Invoke-PackageSet {
    param(
        [object[]]$Packages,
        [string]$Label,
        [switch]$UseNameFallback
    )

    $failures = 0

    if (-not $Packages -or $Packages.Count -eq 0) {
        Write-Log "No packages found for $Label." 'INFO'
        return 0
    }

    foreach ($pkg in $Packages) {
        $packageStart = Get-Date

        if ($pkg.Id -in $SpecialHandlingPackageIds) {
            Write-Log "Special handling: $($pkg.Name) [$($pkg.Id)] is still upgradeable, but targeted unattended self-update is deferred to Microsoft Store/App Installer servicing." 'INFO'

            $script:ApplicationResults.Add([PSCustomObject]@{
                Name            = $pkg.Name
                Id              = $pkg.Id
                VersionBefore   = $pkg.Version
                VersionTarget   = $pkg.Available
                Source          = $pkg.Source
                AttemptType     = $Label
                Status          = 'DeferredSpecialHandling'
                ExitCode        = $null
                DurationSeconds = [math]::Round(((Get-Date) - $packageStart).TotalSeconds, 2)
            }) | Out-Null

            continue
        }

        Write-Log "$($Label): $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        $result = Invoke-TargetedUpgrade -Package $pkg -UseNameFallback:$UseNameFallback
        $status = if ($result.ExitCode -eq 0) { 'Success' } else { 'Failed' }

        if ($result.ExitCode -eq 0) {
            Write-Log "Targeted upgrade completed for $($pkg.Id)." 'OK'
        }
        else {
            Write-Log "Targeted upgrade failed for $($pkg.Id) with exit code $($result.ExitCode)." 'WARN'
            $failures++
        }

        $script:ApplicationResults.Add([PSCustomObject]@{
            Name            = $pkg.Name
            Id              = $pkg.Id
            VersionBefore   = $pkg.Version
            VersionTarget   = $pkg.Available
            Source          = $pkg.Source
            AttemptType     = $Label
            Status          = $status
            ExitCode        = $result.ExitCode
            DurationSeconds = [math]::Round(((Get-Date) - $packageStart).TotalSeconds, 2)
        }) | Out-Null
    }

    return $failures
}

# Starts a Microsoft Office Click-to-Run update and waits up to the configured timeout for completion.
function Update-OfficeClickToRun {
    param([int]$WaitMinutes = 30)

    $officePath = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    if (-not (Test-Path -Path $officePath)) {
        Write-Log "Office Click-to-Run client not found. Skipping Office update." 'INFO'
        return [PSCustomObject]@{ Installed = $false; Attempted = $false; Status = 'NotInstalled'; ExitCode = $null; TimedOut = $false }
    }

    Write-Log "Starting Office Click-to-Run update..." 'INFO'
    $proc = Start-Process -FilePath $officePath `
                          -ArgumentList '/update', 'USER', 'displaylevel=False', 'forceappshutdown=True' `
                          -PassThru `
                          -WindowStyle Hidden

    $completed = $proc.WaitForExit($WaitMinutes * 60 * 1000)
    if (-not $completed) {
        Write-Log "Office update process did not finish within $WaitMinutes minute(s)." 'WARN'
        return [PSCustomObject]@{ Installed = $true; Attempted = $true; Status = 'TimedOut'; ExitCode = $null; TimedOut = $true }
    }

    $status = if ($proc.ExitCode -eq 0) { 'Success' } else { 'Failed' }
    Write-Log "Office Click-to-Run exited with code $($proc.ExitCode)." $(if ($proc.ExitCode -eq 0) { 'OK' } else { 'WARN' })
    return [PSCustomObject]@{ Installed = $true; Attempted = $true; Status = $status; ExitCode = $proc.ExitCode; TimedOut = $false }
}


# HP PowerShell module maintenance was moved to 05_Weekend_HP_Drivers_Update.ps1.
# This application update script no longer installs or updates HP PowerShell modules.

# Checks whether Windows Update reports that a reboot is pending after software maintenance.
function Test-RebootRequired {
    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        return [bool]$sysInfo.RebootRequired
    }
    catch {
        return $false
    }
}

# Removes a defined set of package IDs from a package list so they are not retried unnecessarily.
function Remove-PackagesById {
    param(
        [object[]]$Packages,
        [string[]]$IdsToRemove
    )

    if (-not $Packages) { return @() }
    if (-not $IdsToRemove -or $IdsToRemove.Count -eq 0) { return @($Packages) }

    $idSet = @{}
    foreach ($id in $IdsToRemove) {
        $idSet[$id.ToLowerInvariant()] = $true
    }

    $filtered = @()
    foreach ($pkg in $Packages) {
        if (-not $idSet.ContainsKey($pkg.Id.ToLowerInvariant())) {
            $filtered += $pkg
        }
    }

    return @($filtered)
}


function Get-WingetVersion {
    try {
        $path = Get-WingetPath
        if (-not $path) { return $null }
        $output = & $path --version 2>$null
        if ($LASTEXITCODE -eq 0) { return ([string]$output).Trim() }
    }
    catch {}
    return $null
}

function Write-TelemetryEvent {
    param(
        [Parameter(Mandatory)][int]$ExitCode,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Details
    )

    try {
        if (-not (Test-Path $TelemetryDirectory)) {
            New-Item -Path $TelemetryDirectory -ItemType Directory -Force | Out-Null
        }

        $end = Get-Date
        $event = [ordered]@{
            EventType       = 'maintenance.execution'
            RunId           = $RunId
            ComputerName    = $env:COMPUTERNAME
            Domain          = $env:USERDOMAIN
            ScriptName      = $ScriptName
            ScriptVersion   = $ScriptVersion
            TextLogPath     = $PublishedLogPath
            Status          = $Status
            ExitCode        = $ExitCode
            StartTime       = $ExecutionStart.ToString('o')
            EndTime         = $end.ToString('o')
            DurationSeconds = [math]::Round(($end - $ExecutionStart).TotalSeconds, 2)
            WarningCount    = $WarningCount
            ErrorCount      = $ErrorCount
            Timestamp       = $end.ToUniversalTime().ToString('o')
            Details         = $Details
        }

        $compact = $event | ConvertTo-Json -Depth 10 -Compress
        Write-MaintenanceTelemetryLine -Path $TelemetryNdjsonPath -JsonLine $compact
        $event | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $TelemetryLatestPath -Encoding UTF8
    }
    catch {
        try { Write-Log "Unable to write telemetry: $($_.Exception.GetType().FullName): $($_.Exception.Message)" 'ERROR' } catch {}
    }
}

# Main execution starts here: validate elevation, prepare networking, then process application updates.
$finalExitCode = 1
$finalStatus = 'Failed'
$preMain = @()
$preExplicit = @()
$finalMain = @()
$finalExplicit = @()
$pinnedBefore = @()
$pinnedAfter = @()
$updatedPackageDetails = @()
$failedPackageDetails = @()
$bulkExitCode = $null
$sourceRefreshExitCode = $null
$officeResult = [PSCustomObject]@{ Installed = $false; Attempted = $false; Status = 'NotRun'; ExitCode = $null; TimedOut = $false }
$rebootRequiredBefore = Test-RebootRequired
$rebootRequiredAfter = $false

try {
    if (-not (Test-IsAdministrator)) {
        throw 'This script must be run as Administrator.'
    }

    Initialize-NetworkDefaults
    Write-Log "Initializing application update script..." 'INFO'
    Write-Log "Script version: $ScriptVersion | Last updated: 2026-08-18" 'INFO'
    Write-Log "Active staged text log path: $LogPath" 'INFO'
    Write-Log "Completed text log publish path: $PublishedLogPath" 'INFO'

    Write-Log "Refreshing WinGet sources..." 'INFO'
    $sourceResult = Invoke-Winget -Arguments @('source', 'update') -IgnoreExitCode
    $sourceRefreshExitCode = $sourceResult.ExitCode
    if ($sourceRefreshExitCode -eq 0) { Write-Log "WinGet source refresh completed." 'OK' }
    else { Write-Log "WinGet source refresh completed with exit code $sourceRefreshExitCode." 'WARN' }

    Write-Log "Collecting pre-upgrade inventory..." 'INFO'
    $preInventoryResult = Get-UpgradeInventory -Source 'winget'
    $preParsed = Parse-WingetInventory -Lines $preInventoryResult.Output
    $preMain = @($preParsed.Main)
    $preExplicit = @($preParsed.Explicit)

    # Always perform a separate include-pinned inventory for reporting. Packages
    # visible only in this inventory are blocked/deferred by a WinGet pin.
    $prePinnedInventoryResult = Get-UpgradeInventoryIncludingPinned -Source 'winget'
    $prePinnedParsed = Parse-WingetInventory -Lines $prePinnedInventoryResult.Output
    $normalPreIds = @{}
    foreach ($pkg in @($preMain + $preExplicit)) { $normalPreIds[$pkg.Id.ToLowerInvariant()] = $true }
    $pinnedBefore = @(
        @($prePinnedParsed.Main + $prePinnedParsed.Explicit) |
        Where-Object { -not $normalPreIds.ContainsKey($_.Id.ToLowerInvariant()) }
    )

    Write-Log "Main inventory found $($preMain.Count) upgradeable package(s)." $(if ($preMain.Count -gt 0) { 'OK' } else { 'INFO' })
    Write-Log "Explicit-target inventory found $($preExplicit.Count) package(s)." $(if ($preExplicit.Count -gt 0) { 'WARN' } else { 'INFO' })

    foreach ($pkg in $preExplicit) {
        Write-Log "Pre-scan explicit target: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
    }

    foreach ($pkg in $pinnedBefore) {
        Write-Log "Pinned/deferred upgrade detected: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
    }
    Write-Log "Pinned/deferred upgrade inventory found $($pinnedBefore.Count) package(s)." $(if ($pinnedBefore.Count -gt 0) { 'WARN' } else { 'INFO' })

    $bulkResult = Invoke-WingetUpgradeAll -Source 'winget'
    $bulkExitCode = $bulkResult.ExitCode
    if ($bulkExitCode -ne 0) { Write-Log "Bulk WinGet upgrade returned exit code $bulkExitCode." 'WARN' }

    $explicitFailures = Invoke-PackageSet -Packages $preExplicit -Label 'Explicit target required' -UseNameFallback

    if ($AttemptMSStore) {
        Write-Log "Attempting second pass against msstore source..." 'INFO'
        $storeResult = Invoke-WingetUpgradeAll -Source 'msstore'
        if ($storeResult.ExitCode -ne 0) { Write-Log "Microsoft Store upgrade pass returned exit code $($storeResult.ExitCode)." 'WARN' }
    }

    if ($UpdateOffice) {
        $officeResult = Update-OfficeClickToRun -WaitMinutes $OfficeWaitMinutes
    }

    Write-Log "Collecting post-upgrade inventory..." 'INFO'
    $postInventoryResult = Get-UpgradeInventory -Source 'winget'
    $postParsed = Parse-WingetInventory -Lines $postInventoryResult.Output
    $remainingMain = Remove-PackagesById -Packages @($postParsed.Main) -IdsToRemove ($preExplicit.Id)
    $remainingExplicit = @($postParsed.Explicit)

    $retryFailures = 0
    $retryFailures += Invoke-PackageSet -Packages $remainingMain -Label 'Retry remaining package' -UseNameFallback
    $retryFailures += Invoke-PackageSet -Packages $remainingExplicit -Label 'Retry explicit-target package' -UseNameFallback

    Write-Log "Collecting final inventory..." 'INFO'
    $finalInventoryResult = Get-UpgradeInventory -Source 'winget'
    $finalParsed = Parse-WingetInventory -Lines $finalInventoryResult.Output
    $finalMain = @(Remove-PackagesById -Packages $finalParsed.Main -IdsToRemove ($preExplicit.Id))
    $finalExplicit = @($finalParsed.Explicit)

    $finalPinnedInventoryResult = Get-UpgradeInventoryIncludingPinned -Source 'winget'
    $finalPinnedParsed = Parse-WingetInventory -Lines $finalPinnedInventoryResult.Output
    $normalFinalIds = @{}
    foreach ($pkg in @($finalMain + $finalExplicit)) { $normalFinalIds[$pkg.Id.ToLowerInvariant()] = $true }
    $pinnedAfter = @(
        @($finalPinnedParsed.Main + $finalPinnedParsed.Explicit) |
        Where-Object { -not $normalFinalIds.ContainsKey($_.Id.ToLowerInvariant()) }
    )

    $effectiveFinalMain = @($finalMain | Where-Object { $_.Id -notin $SpecialHandlingPackageIds })
    $deferredFinalMain  = @($finalMain | Where-Object { $_.Id -in $SpecialHandlingPackageIds })

    foreach ($pkg in $effectiveFinalMain) {
        Write-Log "Still remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
    }
    foreach ($pkg in $deferredFinalMain) {
        Write-Log "Deferred special-handling package remains: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'INFO'
    }
    foreach ($pkg in $finalExplicit) {
        Write-Log "Still explicit-target remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
    }

    if ($effectiveFinalMain.Count -eq 0) { Write-Log "Final actionable remaining main packages: 0" 'OK' }
    else { Write-Log "Final actionable remaining main packages: $($effectiveFinalMain.Count)" 'WARN' }

    if ($deferredFinalMain.Count -gt 0) {
        Write-Log "Deferred special-handling packages: $($deferredFinalMain.Count)" 'INFO'
    }

    if ($finalExplicit.Count -eq 0) { Write-Log "Final remaining explicit-target packages: 0" 'OK' }
    else { Write-Log "Final remaining explicit-target packages: $($finalExplicit.Count)" 'WARN' }

    $rebootRequiredAfter = Test-RebootRequired
    if ($rebootRequiredAfter) {
        Write-Log "A reboot is required after application updates." 'WARN'
        $finalExitCode = 3010
        $finalStatus = 'SuccessRebootRequired'
    }
    elseif ($explicitFailures -eq 0 -and $retryFailures -eq 0 -and $effectiveFinalMain.Count -eq 0 -and $finalExplicit.Count -eq 0) {
        Write-Log "Application update script completed successfully." 'OK'
        $finalExitCode = 0
        $finalStatus = if ($WarningCount -gt 0) { 'SuccessWithWarnings' } else { 'Success' }
    }
    else {
        Write-Log "Application update script completed with remaining packages or non-zero package results." 'WARN'
        $finalExitCode = 2
        $finalStatus = 'CompletedWithRemainingUpdates'
    }
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    $finalExitCode = 1
    $finalStatus = 'Failed'
}
finally {
    $initialPackages = @($preMain + $preExplicit)
    $remainingPackages = @($finalMain + $finalExplicit)
    $remainingIds = @{}
    foreach ($pkg in $remainingPackages) { $remainingIds[$pkg.Id.ToLowerInvariant()] = $true }

    $resolvedPackages = @()
    foreach ($pkg in $initialPackages) {
        if (-not $remainingIds.ContainsKey($pkg.Id.ToLowerInvariant())) {
            $resolvedPackages += [PSCustomObject]@{
                Name          = $pkg.Name
                Id            = $pkg.Id
                VersionBefore = $pkg.Version
                VersionTarget = $pkg.Available
                Source        = $pkg.Source
                Status        = 'NoLongerUpgradeable'
            }
        }
    }


    # Build one stable per-package result for Elastic. A package counts as
    # "updated" only when its installed version can be verified as changed.
    $finalUpgradeableById = @{}
    foreach ($pkg in @($remainingPackages + $pinnedAfter)) {
        $finalUpgradeableById[$pkg.Id.ToLowerInvariant()] = $pkg
    }

    $updatedPackageDetails = @()
    $failedPackageDetails = @()

    foreach ($pkg in $initialPackages) {
        $afterVersion = Get-InstalledPackageVersion -Id $pkg.Id -Source $pkg.Source
        $stillUpgradeable = $finalUpgradeableById.ContainsKey($pkg.Id.ToLowerInvariant())
        $changed = (-not [string]::IsNullOrWhiteSpace($afterVersion)) -and ($afterVersion -ne $pkg.Version)

        if ($changed -and -not $stillUpgradeable) {
            $updatedPackageDetails += [PSCustomObject]@{
                Name          = $pkg.Name
                Id            = $pkg.Id
                VersionBefore = $pkg.Version
                VersionAfter  = $afterVersion
                VersionTarget = $pkg.Available
                Source        = $pkg.Source
                Status        = 'Updated'
            }
        }
        elseif ($stillUpgradeable) {
            $remaining = $finalUpgradeableById[$pkg.Id.ToLowerInvariant()]
            $failedPackageDetails += [PSCustomObject]@{
                Name             = $pkg.Name
                Id               = $pkg.Id
                VersionBefore    = $pkg.Version
                VersionAfter     = $afterVersion
                VersionAvailable = $remaining.Available
                Source           = $pkg.Source
                Status           = 'StillUpgradeable'
            }
        }
    }

    $pinnedPackageDetails = @(
        $pinnedAfter | ForEach-Object {
            [PSCustomObject]@{
                Name             = $_.Name
                Id               = $_.Id
                VersionInstalled = $_.Version
                VersionAvailable = $_.Available
                Source           = $_.Source
                Status           = 'PinnedOrDeferred'
            }
        }
    )

    # Include targeted attempts that explicitly failed even when a later inventory
    # could not provide a row. De-duplicate by package ID.
    $failedIds = @{}
    foreach ($item in $failedPackageDetails) { $failedIds[$item.Id.ToLowerInvariant()] = $true }
    foreach ($attempt in ($ApplicationResults | Where-Object Status -eq 'Failed')) {
        if (-not $failedIds.ContainsKey($attempt.Id.ToLowerInvariant())) {
            $failedPackageDetails += [PSCustomObject]@{
                Name             = $attempt.Name
                Id               = $attempt.Id
                VersionBefore    = $attempt.VersionBefore
                VersionAfter     = Get-InstalledPackageVersion -Id $attempt.Id -Source $attempt.Source
                VersionAvailable = $attempt.VersionTarget
                Source           = $attempt.Source
                Status           = 'UpgradeFailed'
                ExitCode         = $attempt.ExitCode
            }
            $failedIds[$attempt.Id.ToLowerInvariant()] = $true
        }
    }

    Write-Log ("Elastic app-update summary: Updated={0}; FailedOrRemaining={1}; PinnedOrDeferred={2}" -f `
        $updatedPackageDetails.Count,
        $failedPackageDetails.Count,
        $pinnedPackageDetails.Count) 'INFO'

    # Windows PowerShell 5.1 can throw "Argument types do not match" when
    # @() is applied directly to System.Collections.Generic.List[object].
    # Materialize every telemetry collection as a normal Object[] first.
    [object[]]$applicationResultsTelemetry = @(
        foreach ($item in $ApplicationResults) { $item }
    )
    [object[]]$updatedApplicationsTelemetry = @(
        foreach ($item in $updatedPackageDetails) { $item }
    )
    [object[]]$failedApplicationsTelemetry = @(
        foreach ($item in $failedPackageDetails) { $item }
    )
    [object[]]$pinnedApplicationsTelemetry = @(
        foreach ($item in $pinnedPackageDetails) { $item }
    )

    $details = [ordered]@{
        WingetAvailable           = [bool](Get-WingetPath)
        WingetVersion             = $(Get-WingetVersion)
        SourceRefreshExitCode     = $sourceRefreshExitCode
        BulkUpgradeExitCode       = $bulkExitCode
        IncludeUnknown            = [bool]$IncludeUnknown
        IncludePinned             = [bool]$IncludePinned
        AttemptMSStore            = [bool]$AttemptMSStore
        UpdateOffice              = [bool]$UpdateOffice
        Office                    = $officeResult
        RebootRequiredBefore      = $rebootRequiredBefore
        RebootRequiredAfter       = $rebootRequiredAfter
        PackagesUpgradeableBefore = $initialPackages.Count
        PackagesResolved          = $resolvedPackages.Count
        PackagesRemaining         = $remainingPackages.Count
        ActionablePackagesRemaining = $effectiveFinalMain.Count + $finalExplicit.Count
        DeferredSpecialPackages  = $deferredFinalMain.Count
        ExplicitTargetsBefore       = $preExplicit.Count

        # Primary Kibana/Elastic reporting fields.
        AppsUpdatedCount            = [int]$updatedPackageDetails.Count
        AppsFailedOrRemainingCount  = [int]$failedPackageDetails.Count
        AppsPinnedOrDeferredCount   = [int]$pinnedPackageDetails.Count
        UpdatedApplications         = $updatedApplicationsTelemetry
        FailedOrRemainingApplications = $failedApplicationsTelemetry
        PinnedOrDeferredApplications  = $pinnedApplicationsTelemetry

        # Lower-level attempt history retained for troubleshooting.
        ApplicationResults          = $applicationResultsTelemetry
        ResolvedPackages          = @($resolvedPackages)
        RemainingPackages         = @($remainingPackages)
    }

    Write-TelemetryEvent -ExitCode $finalExitCode -Status $finalStatus -Details $details

    # Final append before the completed file enters C:\Logs.
    Write-Log ("Completed {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
        $ScriptName,
        $finalStatus,
        $finalExitCode,
        $WarningCount,
        $ErrorCount) $(if ($finalExitCode -eq 0 -or $finalExitCode -eq 3010) { 'OK' } else { 'ERROR' })

    if ($null -ne $TextLogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $TextLogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 03 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 03 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $finalExitCode
