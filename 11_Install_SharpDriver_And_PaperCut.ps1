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
    1.2.0

.DATE
    2026-07-27

.CHANGELOG
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
$ScriptVersion = '1.2.0'
$DriverSourcePath = '\\papercut\Printer Drivers\MFP\win\SH_D31_PCL6_PS_2410a_EnglishUS_64bit'
$PrinterDriverName = 'Sharp BP-70C31 PCL6'
$PaperCutMsiPath = '\\papercut\Print Deploy Clients\win\pc-print-deploy-client[10.2.3.44].msi'
$PrinterSharePath = '\\papercut\StudentSecurePrint'
$LocalDriverStage = 'C:\ProgramData\Compton\Drivers\Sharp'
$StateDirectory = 'C:\ProgramData\Compton\State'
$StatePath = Join-Path $StateDirectory 'SharpDriver-PaperCut-State.json'
$LogDirectory = 'C:\Logs'
$LogPath = Join-Path $LogDirectory 'Install-SharpDriver-And-PaperCut.log'
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

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','ACTION','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '[{0}] [{1}] {2}' -f $timestamp, $Level, $Message
    $color = switch ($Level) {
        'ACTION' { 'Yellow' }
        'SUCCESS' { 'Green' }
        'WARN' { 'DarkYellow' }
        'ERROR' { 'Red' }
        default { 'White' }
    }
    Write-Host $line -ForegroundColor $color
    try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch { }
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
        [ordered]@{
            ProductName = [string]$cv.ProductName
            EditionID = [string]$cv.EditionID
            DisplayVersion = [string]$cv.DisplayVersion
            CurrentBuild = [string]$cv.CurrentBuild
            UBR = [int]$cv.UBR
            FullBuild = '{0}.{1}' -f $cv.CurrentBuild, $cv.UBR
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

function Get-InstalledPrinterDriverState {
    try {
        $driver = Get-PrinterDriver -Name $PrinterDriverName -ErrorAction SilentlyContinue
        if (-not $driver) { return [pscustomobject]@{ Installed=$false; Name=$PrinterDriverName; Manufacturer=$null; InfPath=$null; MajorVersion=$null } }
        [pscustomobject]@{ Installed=$true; Name=[string]$driver.Name; Manufacturer=[string]$driver.Manufacturer; InfPath=[string]$driver.InfPath; MajorVersion=$driver.MajorVersion }
    }
    catch { [pscustomobject]@{ Installed=$false; Name=$PrinterDriverName; Manufacturer=$null; InfPath=$null; MajorVersion=$null } }
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
        [Parameter(Mandatory)]$PendingRebootBefore,
        [Parameter(Mandatory)]$PendingRebootAfter
    )
    $end = Get-Date
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
                RunAsSystem=(Test-IsSystemAccount)
                UserName=[Security.Principal.WindowsIdentity]::GetCurrent().Name
            }
            printerDeployment = [ordered]@{
                OverallCompliant=($DriverAfter.Installed -and $PaperCutAfter.Installed -and $PrinterAfter.Connected)
                ChangedComponents=@($script:ChangedComponents)
                Driver=[ordered]@{ Name=$PrinterDriverName; Before=$DriverBefore; After=$DriverAfter; SourceAvailable=$DriverSourceAvailable; SourcePath=$DriverSourcePath; SourceFingerprintSHA256=$SourceFingerprint; StoredFingerprintSHA256=$StoredFingerprint; StatePath=$StatePath }
                PaperCut=[ordered]@{ Before=$PaperCutBefore; After=$PaperCutAfter; SourceAvailable=$PaperCutSourceAvailable; MsiPath=$PaperCutMsiPath }
                PrinterConnection=[ordered]@{ SharePath=$PrinterSharePath; Before=$PrinterBefore; After=$PrinterAfter }
                Processes=@($script:ProcessResults)
                Reboot=[ordered]@{ Required=$script:RebootRequired; PendingBefore=$PendingRebootBefore.Pending; ReasonsBefore=@($PendingRebootBefore.Reasons); PendingAfter=$PendingRebootAfter.Pending; ReasonsAfter=@($PendingRebootAfter.Reasons) }
                Windows=(Get-WindowsInfo)
            }
        }
        '@timestamp'=$end.ToUniversalTime().ToString('o')
    }
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
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

$driverBefore = [pscustomobject]@{ Installed=$false; Name=$PrinterDriverName; Manufacturer=$null; InfPath=$null; MajorVersion=$null }
$driverAfter = $driverBefore
$paperCutBefore = [pscustomobject]@{ Installed=$false; DisplayName=$null; DisplayVersion=$null; ProductCode=$null; InstallLocation=$null }
$paperCutAfter = $paperCutBefore
$printerBefore = [pscustomobject]@{ Scope='Unknown'; Connected=$false; Name=$PrinterSharePath }
$printerAfter = $printerBefore
$sourceFingerprint = $null
$storedFingerprint = $null
$driverSourceAvailable = $false
$paperCutSourceAvailable = $false
$pendingBefore = Get-PendingRebootState
$pendingAfter = $pendingBefore
$finalStatus = 'Failed'
$finalExitCode = 1

try {
    if (-not (Test-IsAdministrator)) { throw 'This script must be run as an administrator.' }
    if ([string]::IsNullOrWhiteSpace($PrinterDriverName) -or $PrinterDriverName -like 'CHANGE ME*') { throw 'PrinterDriverName is not configured.' }
    Import-Module PrintManagement -ErrorAction Stop

    $driverBefore = Get-InstalledPrinterDriverState
    $paperCutBefore = Get-PaperCutProduct
    $printerBefore = Get-PrinterConnectionState -ConnectionPath $PrinterSharePath
    $driverSourceAvailable = Test-Path -LiteralPath $DriverSourcePath -PathType Container
    $paperCutSourceAvailable = Test-Path -LiteralPath $PaperCutMsiPath -PathType Leaf
    $state = Get-DeploymentState
    if ($state) {
        $prop = $state.PSObject.Properties['SourceFingerprintSHA256']
        if ($prop) { $storedFingerprint = [string]$prop.Value }
    }

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
        if (-not $driverSourceAvailable) { throw 'Driver deployment is required but the source package is unavailable.' }
        if (Test-Path -LiteralPath $LocalDriverStage) { Remove-Item -LiteralPath $LocalDriverStage -Recurse -Force -ErrorAction Stop }
        New-Item -Path $LocalDriverStage -ItemType Directory -Force | Out-Null
        Write-Log "Copying driver package to local staging: $LocalDriverStage" 'ACTION'
        Copy-Item -Path (Join-Path $DriverSourcePath '*') -Destination $LocalDriverStage -Recurse -Force -ErrorAction Stop
        $infFiles = @(Get-PrinterInfFiles -Path $LocalDriverStage)
        foreach ($inf in $infFiles) {
            $code = Invoke-ExternalProcess -FilePath "$env:SystemRoot\System32\pnputil.exe" -ArgumentList @('/add-driver',"`"$($inf.FullName)`"") -SuccessExitCodes @(0,3010) -TimeoutSeconds $DriverInstallTimeoutSeconds -Operation ("Stage driver {0}" -f $inf.Name)
            if ($code -eq 3010) { $script:RebootRequired = $true }
        }
        Add-PrinterDriver -Name $PrinterDriverName -ErrorAction Stop
        $driverAfter = Get-InstalledPrinterDriverState
        if (-not $driverAfter.Installed) { throw "Driver was staged but could not be verified: $PrinterDriverName" }
        Save-DeploymentState -Fingerprint $sourceFingerprint
        $script:ChangedComponents.Add('PrinterDriver')
        Write-Log "Printer driver deployed successfully: $PrinterDriverName" 'SUCCESS'
    }
    else {
        $driverAfter = $driverBefore
        Write-Log "Printer driver is already compliant: $PrinterDriverName" 'SUCCESS'
    }

    if (-not $paperCutBefore.Installed) {
        if (-not $paperCutSourceAvailable) { throw "PaperCut is not installed and the MSI is unavailable: $PaperCutMsiPath" }
        $msiLogPath = Join-Path $LogDirectory 'PaperCut-Print-Deploy-MSI.log'
        $code = Invoke-ExternalProcess -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList @('/i',"`"$PaperCutMsiPath`"",'/qn','/norestart','REBOOT=ReallySuppress','/L*v',"`"$msiLogPath`"") -SuccessExitCodes @(0,3010) -TimeoutSeconds $PaperCutInstallTimeoutSeconds -Operation 'Install PaperCut Print Deploy'
        if ($code -eq 3010) { $script:RebootRequired = $true }
        Start-Sleep -Seconds 3
        $paperCutAfter = Get-PaperCutProduct
        if (-not $paperCutAfter.Installed) { throw 'PaperCut MSI returned success, but the product could not be verified in the uninstall registry.' }
        $script:ChangedComponents.Add('PaperCut')
        Write-Log "PaperCut installed successfully: $($paperCutAfter.DisplayName) $($paperCutAfter.DisplayVersion)" 'SUCCESS'
    }
    else {
        $paperCutAfter = $paperCutBefore
        Write-Log "PaperCut is already installed: $($paperCutBefore.DisplayName) $($paperCutBefore.DisplayVersion)" 'SUCCESS'
        if (-not $paperCutSourceAvailable) { Write-Log "PaperCut MSI source is currently unavailable, but the client is already installed: $PaperCutMsiPath" 'WARN' }
    }

    $printerAfter = Ensure-SharedPrinterConnection -ConnectionPath $PrinterSharePath
    if (-not $driverAfter.Installed -or -not $paperCutAfter.Installed -or -not $printerAfter.Connected) { throw 'Final compliance verification failed.' }

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
    Write-Log $script:FailureMessage 'ERROR'
    $finalStatus='Failed'
    $finalExitCode=1
    try { $driverAfter=Get-InstalledPrinterDriverState } catch { }
    try { $paperCutAfter=Get-PaperCutProduct } catch { }
    try { $printerAfter=Get-PrinterConnectionState -ConnectionPath $PrinterSharePath } catch { }
    try { $pendingAfter=Get-PendingRebootState } catch { }
}
finally {
    try {
        Write-Telemetry -Status $finalStatus -ExitCode $finalExitCode -DriverBefore $driverBefore -DriverAfter $driverAfter -PaperCutBefore $paperCutBefore -PaperCutAfter $paperCutAfter -PrinterBefore $printerBefore -PrinterAfter $printerAfter -DriverSourceAvailable $driverSourceAvailable -PaperCutSourceAvailable $paperCutSourceAvailable -SourceFingerprint $sourceFingerprint -StoredFingerprint $storedFingerprint -PendingRebootBefore $pendingBefore -PendingRebootAfter $pendingAfter
        Write-Log "Telemetry written to $LatestTelemetryPath"
    }
    catch { Write-Log "Telemetry write failed: $($_.Exception.Message)" 'ERROR' }
}

exit $finalExitCode
