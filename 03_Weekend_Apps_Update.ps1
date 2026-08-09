# ScriptVersion: 2.1.0
# LastUpdated: 2026-07-27

[CmdletBinding()]
param(
    [switch]$IncludeUnknown = $true,
    [switch]$IncludePinned = $false,
    [switch]$AttemptMSStore = $false,
    [switch]$UpdateOffice = $true,
    [int]$OfficeWaitMinutes = 30,
    [string]$LogPath = "$env:SystemDrive\Temp\Weekend-Apps-Update.log"
)

$ErrorActionPreference = 'Stop'

$ScriptVersion = '2.1.0'
$ExecutionStart = Get-Date
$RunId = [guid]::NewGuid().Guid
$WarningCount = 0
$ErrorCount = 0
$ApplicationResults = New-Object System.Collections.Generic.List[object]
$TelemetryDirectory = Join-Path $env:SystemDrive 'Logs'
$TelemetryNdjsonPath = Join-Path $TelemetryDirectory 'Maintenance-Telemetry.ndjson'
$TelemetryLatestPath = Join-Path $TelemetryDirectory '03_Weekend_Apps_Update.latest.json'

# Writes color-coded status output to the console and appends the same message to the log file.
# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
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

    foreach ($line in $Lines) {
        $trimmed = $line.Trim()

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -match 'require explicit targeting for upgrade' -or
            $trimmed -match 'need to be explicitly upgraded') {
            $mode = 'Explicit'
            continue
        }

        if ($trimmed -match '^Name\s+Id\s+Version\s+Available(\s+Source)?$') {
            if ($mode -eq 'None') {
                $mode = 'Main'
            }
            continue
        }

        if ($trimmed -match '^\d+\s+upgrades available\.?$') {
            continue
        }

        if ($trimmed -match '^Installing dependencies:$') { continue }
        if ($trimmed -match '^This package requires the following dependencies:$') { continue }
        if ($trimmed -match '^- Packages$') { continue }
        if ($trimmed -match '^[A-Za-z0-9._+-]+\.[A-Za-z0-9.+-]+$') { continue }

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

        $pattern = '^(?<Name>.+?)\s{2,}(?<Id>[A-Za-z0-9][A-Za-z0-9._-]+)\s{2,}(?<Version>\S+)\s{2,}(?<Available>\S+)(?:\s{2,}(?<Source>\S+))?$'
        if ($trimmed -match $pattern) {
            $obj = [PSCustomObject]@{
                Name      = $matches.Name.Trim()
                Id        = $matches.Id.Trim()
                Version   = $matches.Version.Trim()
                Available = $matches.Available.Trim()
                Source    = if ($matches.Source) { $matches.Source.Trim() } else { '' }
            }

            if ($mode -eq 'Explicit') {
                $explicitPackages += $obj
            }
            elseif ($mode -eq 'Main') {
                $mainPackages += $obj
            }

            continue
        }
    }

    return [PSCustomObject]@{
        Main     = @($mainPackages)
        Explicit = @($explicitPackages)
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

    if (-not [string]::IsNullOrWhiteSpace($Package.Source)) {
        $args += @('--source', $Package.Source)
    }

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

        if (-not [string]::IsNullOrWhiteSpace($Package.Source)) {
            $fallbackArgs += @('--source', $Package.Source)
        }

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
        [Parameter(Mandatory)][hashtable]$Details
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
            ScriptName      = Split-Path -Leaf $PSCommandPath
            ScriptVersion   = $ScriptVersion
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

        $compact = $event | ConvertTo-Json -Depth 8 -Compress
        Write-MaintenanceTelemetryLine -Path $TelemetryNdjsonPath -JsonLine $compact
        $event | ConvertTo-Json -Depth 8 | Set-Content -Path $TelemetryLatestPath -Encoding UTF8
    }
    catch {
        try { Write-Log "Unable to write telemetry: $($_.Exception.Message)" 'ERROR' } catch {}
    }
}

# Main execution starts here: validate elevation, prepare networking, then process application updates.
$finalExitCode = 1
$finalStatus = 'Failed'
$preMain = @()
$preExplicit = @()
$finalMain = @()
$finalExplicit = @()
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
    Write-Log "Script version: $ScriptVersion | Last updated: 2026-07-27" 'INFO'
    Write-Log 'Skipping HP PowerShell module maintenance in 03 script; this is handled by 05_Weekend_HP_Drivers_Update.ps1 after HP vendor detection.' 'INFO'

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

    Write-Log "Main inventory found $($preMain.Count) upgradeable package(s)." $(if ($preMain.Count -gt 0) { 'OK' } else { 'INFO' })
    Write-Log "Explicit-target inventory found $($preExplicit.Count) package(s)." $(if ($preExplicit.Count -gt 0) { 'WARN' } else { 'INFO' })

    foreach ($pkg in $preExplicit) {
        Write-Log "Pre-scan explicit target: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
    }

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

    foreach ($pkg in $finalMain) {
        Write-Log "Still remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
    }
    foreach ($pkg in $finalExplicit) {
        Write-Log "Still explicit-target remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
    }

    if ($finalMain.Count -eq 0) { Write-Log "Final remaining main packages: 0" 'OK' }
    else { Write-Log "Final remaining main packages: $($finalMain.Count)" 'WARN' }
    if ($finalExplicit.Count -eq 0) { Write-Log "Final remaining explicit-target packages: 0" 'OK' }
    else { Write-Log "Final remaining explicit-target packages: $($finalExplicit.Count)" 'WARN' }

    $rebootRequiredAfter = Test-RebootRequired
    if ($rebootRequiredAfter) {
        Write-Log "A reboot is required after application updates." 'WARN'
        $finalExitCode = 3010
        $finalStatus = 'SuccessRebootRequired'
    }
    elseif ($explicitFailures -eq 0 -and $retryFailures -eq 0 -and $finalMain.Count -eq 0 -and $finalExplicit.Count -eq 0) {
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
        ExplicitTargetsBefore     = $preExplicit.Count
        ApplicationResults        = @($ApplicationResults)
        ResolvedPackages          = @($resolvedPackages)
        RemainingPackages         = @($remainingPackages)
    }

    Write-TelemetryEvent -ExitCode $finalExitCode -Status $finalStatus -Details $details
}

exit $finalExitCode
