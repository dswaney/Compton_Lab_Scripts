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
