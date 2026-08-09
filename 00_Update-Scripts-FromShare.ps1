#requires -version 5.1
# =====================================================================
# ScriptName: 00_Update-Scripts-FromShare.ps1
# ScriptVersion: 4.1.0
# LastUpdated: 2026-07-27
# Purpose:
#   Manifest-driven, self-bootstrapping updater for C:\Scripts.
#   - Uses \\filesvr\Labscripts with an IP fallback.
#   - Downloads Maintenance.Framework.psm1 when missing or changed.
#   - Validates source and staged files against SHA-256 hashes.
#   - Parses all PowerShell files before installation.
#   - Creates rollback copies before replacement.
#   - Updates itself last and safely relaunches.
#   - Reconciles scheduled tasks after a successful synchronization.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$Relaunched
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$ScriptName = '00_Update-Scripts-FromShare.ps1'
$ScriptVersion = '4.1.0'
$PreferredSourceRoot = '\\filesvr\Labscripts'
$FallbackSourceRoot = '\\10.2.3.30\Labscripts'
$ManifestName = 'DeploymentManifest.json'
$FrameworkName = 'Maintenance.Framework.psm1'
$RegisterTasksName = 'Register-Tasks_SYSTEM.ps1'
$LocalRoot = 'C:\Scripts'
$LogRoot = 'C:\Logs'
$RollbackRoot = 'C:\Scripts\Rollback'
$StagingRoot = 'C:\Scripts\.staging'
$LogPath = Join-Path $LogRoot '00_Update-Scripts-FromShare.log'
$LatestPath = Join-Path $LogRoot '00_Update-Scripts-FromShare.latest.json'
$TelemetryPath = Join-Path $LogRoot 'Maintenance-Telemetry.ndjson'
$PowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$RunId = [guid]::NewGuid().ToString('N')
$Started = Get-Date
$RunStamp = $Started.ToString('yyyyMMdd_HHmmss')
$RollbackRunRoot = Join-Path $RollbackRoot $RunStamp
$StagingRunRoot = Join-Path $StagingRoot $RunId
$WarningCount = 0
$ErrorCount = 0
$SourceRootUsed = $null
$FrameworkImported = $false
$Results = New-Object System.Collections.Generic.List[object]

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
}

function Write-UpdaterStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','ACTION','OK','WARN','ERROR')][string]$Level = 'INFO'
    )
    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }
    $line = '[{0}] [{1,-6}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $color = switch ($Level) {
        'ACTION' { 'Yellow' }
        'OK' { 'Green' }
        'WARN' { 'DarkYellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host $line -ForegroundColor $color
    try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Get-Sha256 {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
}

function Test-PowerShellFile {
    param([Parameter(Mandatory)][string]$Path)
    $extension = [IO.Path]::GetExtension($Path)
    if ($extension -notin @('.ps1','.psm1','.psd1')) {
        return [pscustomobject]@{ Valid = $true; Errors = @() }
    }
    $tokens = $null
    $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors) | Out-Null
    $details = @(
        foreach ($parseError in @($parseErrors)) {
            [pscustomobject]@{
                Line = $parseError.Extent.StartLineNumber
                Column = $parseError.Extent.StartColumnNumber
                ErrorId = $parseError.ErrorId
                Message = $parseError.Message
                Code = $parseError.Extent.Text
            }
        }
    )
    [pscustomobject]@{ Valid = ($details.Count -eq 0); Errors = $details }
}

function Get-FileVersionText {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return 'Missing' }
    try {
        $header = Select-String -LiteralPath $Path -Pattern '^\s*#\s*ScriptVersion\s*:\s*(.+?)\s*$' -ErrorAction Stop | Select-Object -First 1
        if ($header) { return $header.Matches[0].Groups[1].Value.Trim() }
        $moduleVersion = Select-String -LiteralPath $Path -Pattern '^\s*\$script:FrameworkVersion\s*=\s*["'']([^"'']+)["'']' -ErrorAction Stop | Select-Object -First 1
        if ($moduleVersion) { return $moduleVersion.Matches[0].Groups[1].Value.Trim() }
    } catch { }
    'Unknown'
}

function Get-ActiveSourceRoot {
    foreach ($candidate in @($PreferredSourceRoot,$FallbackSourceRoot)) {
        try {
            if (Test-Path -LiteralPath $candidate -PathType Container -ErrorAction Stop) {
                $script:SourceRootUsed = $candidate
                return $candidate
            }
        } catch { }
    }
    throw "Neither source share is available. Preferred: $PreferredSourceRoot | Fallback: $FallbackSourceRoot"
}

function Read-DeploymentManifest {
    param([Parameter(Mandatory)][string]$SourceRoot)
    $path = Join-Path $SourceRoot $ManifestName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Deployment manifest is missing: $path" }
    try { $manifest = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { throw "Deployment manifest is invalid JSON: $($_.Exception.Message)" }
    if ($null -eq $manifest.ManifestVersion -or $null -eq $manifest.Files) { throw 'Deployment manifest is missing ManifestVersion or Files.' }
    $files = @($manifest.Files)
    if ($files.Count -eq 0) { throw 'Deployment manifest contains no files.' }
    $duplicateNames = $files | Group-Object Name | Where-Object Count -gt 1
    if ($duplicateNames) { throw "Deployment manifest contains duplicate file names: $($duplicateNames.Name -join ', ')" }
    foreach ($entry in $files) {
        if ([string]::IsNullOrWhiteSpace([string]$entry.Name)) { throw 'A manifest file entry has no Name.' }
        if ([IO.Path]::GetFileName([string]$entry.Name) -ne [string]$entry.Name) { throw "Manifest entry must be a file name without a path: $($entry.Name)" }
        if ([string]$entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$') { throw "Manifest entry has an invalid SHA-256 value: $($entry.Name)" }
    }
    if (-not ($files | Where-Object Name -eq $FrameworkName)) { throw "Manifest does not contain required framework: $FrameworkName" }
    if (-not ($files | Where-Object Name -eq $ScriptName)) { throw "Manifest does not contain updater: $ScriptName" }
    if (-not ($files | Where-Object Name -eq $RegisterTasksName)) { throw "Manifest does not contain task-registration script: $RegisterTasksName" }
    $manifest
}

function Add-Result {
    param(
        [string]$Name,[string]$Role,[string]$Status,[string]$LocalVersion,[string]$ShareVersion,
        [string]$LocalHash,[string]$ShareHash,[string]$BackupPath,[string]$Message
    )
    [void]$Results.Add([ordered]@{
        Name=$Name;Role=$Role;Status=$Status;LocalVersion=$LocalVersion;ShareVersion=$ShareVersion
        LocalHash=$LocalHash;ShareHash=$ShareHash;BackupPath=$BackupPath;Message=$Message
    })
}

function Install-ManifestEntry {
    param(
        [Parameter(Mandatory)]$Entry,
        [Parameter(Mandatory)][string]$SourceRoot
    )
    $name = [string]$Entry.Name
    $role = if ($Entry.Role) { [string]$Entry.Role } else { 'ManagedFile' }
    $expectedHash = ([string]$Entry.SHA256).ToUpperInvariant()
    $shareVersion = if ($Entry.Version) { [string]$Entry.Version } else { 'Unknown' }
    $sourcePath = Join-Path $SourceRoot $name
    $destinationPath = Join-Path $LocalRoot $name
    $stagedPath = Join-Path $StagingRunRoot $name
    $localExists = Test-Path -LiteralPath $destinationPath -PathType Leaf
    $localHash = Get-Sha256 -Path $destinationPath
    $localVersion = Get-FileVersionText -Path $destinationPath

    if ($localExists -and $localHash -eq $expectedHash) {
        Write-UpdaterStatus "$name is current." 'OK'
        Add-Result -Name $name -Role $role -Status 'Current' -LocalVersion $localVersion -ShareVersion $shareVersion -LocalHash $localHash -ShareHash $expectedHash -BackupPath $null -Message $null
        return [pscustomobject]@{ Changed=$false; SelfUpdated=$false }
    }
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "Manifest file is missing from source share: $sourcePath" }

    $sourceHash = Get-Sha256 -Path $sourcePath
    if ($sourceHash -ne $expectedHash) { throw "Source SHA-256 does not match the manifest for $name. Expected $expectedHash; found $sourceHash." }

    Copy-Item -LiteralPath $sourcePath -Destination $stagedPath -Force -ErrorAction Stop
    $stagedHash = Get-Sha256 -Path $stagedPath
    if ($stagedHash -ne $expectedHash) { throw "Staged SHA-256 verification failed for $name." }

    $parserResult = Test-PowerShellFile -Path $stagedPath
    if (-not $parserResult.Valid) {
        $summary = @($parserResult.Errors | ForEach-Object { "Line $($_.Line), column $($_.Column): $($_.Message)" }) -join ' | '
        throw "Parser validation failed for $name. $summary"
    }

    $backupPath = $null
    if ($localExists) {
        Ensure-Directory -Path $RollbackRunRoot
        $backupPath = Join-Path $RollbackRunRoot $name
        Copy-Item -LiteralPath $destinationPath -Destination $backupPath -Force -ErrorAction Stop
    }

    try {
        Move-Item -LiteralPath $stagedPath -Destination $destinationPath -Force -ErrorAction Stop
        $installedHash = Get-Sha256 -Path $destinationPath
        if ($installedHash -ne $expectedHash) { throw "Post-install SHA-256 verification failed for $name." }
        $installedParser = Test-PowerShellFile -Path $destinationPath
        if (-not $installedParser.Valid) { throw "Post-install parser validation failed for $name." }
    }
    catch {
        $installError = $_.Exception.Message
        try {
            if ($backupPath -and (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
                Copy-Item -LiteralPath $backupPath -Destination $destinationPath -Force -ErrorAction Stop
                Write-UpdaterStatus "Rolled back $name after installation failure." 'WARN'
            } elseif (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
                Remove-Item -LiteralPath $destinationPath -Force -ErrorAction SilentlyContinue
            }
        } catch {
            $installError += " Rollback also failed: $($_.Exception.Message)"
        }
        throw $installError
    }

    $status = if ($localExists) { 'Updated' } else { 'Installed' }
    Write-UpdaterStatus "$status and verified: $name" 'OK'
    Add-Result -Name $name -Role $role -Status $status -LocalVersion $localVersion -ShareVersion $shareVersion -LocalHash (Get-Sha256 -Path $destinationPath) -ShareHash $expectedHash -BackupPath $backupPath -Message $null
    [pscustomobject]@{ Changed=$true; SelfUpdated=($name -ieq $ScriptName) }
}

function Import-RequiredFramework {
    $frameworkPath = Join-Path $LocalRoot $FrameworkName
    if (-not (Test-Path -LiteralPath $frameworkPath -PathType Leaf)) { throw "Framework bootstrap failed; file is missing: $frameworkPath" }
    Import-Module -Name $frameworkPath -Force -ErrorAction Stop
    $script:FrameworkImported = $true
    $configuration = Get-MaintenanceConfiguration -ScriptRoot $LocalRoot -LogRoot $LogRoot
    Write-UpdaterStatus "Framework imported. Version: $($configuration.FrameworkVersion)" 'OK'
}

function Write-ExecutionRecord {
    param([string]$Status,[int]$ExitCode,[string]$FailureMessage)
    $ended = Get-Date
    $record = [ordered]@{
        EventType='maintenance.updater';SchemaVersion=2;ComputerName=$env:COMPUTERNAME
        ScriptName=$ScriptName;ScriptVersion=$ScriptVersion;RunId=$RunId;Status=$Status;ExitCode=$ExitCode
        StartTime=$Started.ToString('o');EndTime=$ended.ToString('o');DurationSeconds=[math]::Round(($ended-$Started).TotalSeconds,3)
        SourceShare=$SourceRootUsed;ManifestName=$ManifestName;Relaunched=[bool]$Relaunched
        WarningCount=$WarningCount;ErrorCount=$ErrorCount;Results=@($Results);FailureMessage=$FailureMessage
        Timestamp=$ended.ToUniversalTime().ToString('o')
    }
    try {
        if ($FrameworkImported) {
            Write-MaintenanceTelemetry -Record $record -Path $TelemetryPath
            $policyPath = Join-Path $LocalRoot 'Maintenance.Policy.json'
            if (Test-Path -LiteralPath $policyPath -PathType Leaf) {
                $policy = Get-MaintenancePolicy -Path $policyPath
                $fleetRecord = [ordered]@{ SchemaVersion=1;ComputerName=$env:COMPUTERNAME;ScriptName=$ScriptName;RunId=$RunId;CorrelationId=$RunId;Status=$Status;ExitCode=$ExitCode;LastCheckIn=$ended.ToUniversalTime().ToString('o');FrameworkVersion=(Get-MaintenanceConfiguration).FrameworkVersion;PolicyVersion=[string]$policy.PolicyVersion;FailureMessage=$FailureMessage }
                Publish-MaintenanceFleetStatus -Record $fleetRecord -StatusRoots @($policy.FleetStatusRoots) | Out-Null
                Write-MaintenanceEvent -EventId $(if($ExitCode -eq 0){1001}else{1900}) -EntryType $(if($ExitCode -eq 0){'Information'}else{'Error'}) -Message "$ScriptName completed. Status=$Status ExitCode=$ExitCode RunId=$RunId" | Out-Null
            }
        }
        $record | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $LatestPath -Encoding UTF8 -Force
    } catch {
        Write-UpdaterStatus "Unable to write updater telemetry: $($_.Exception.Message)" 'WARN'
    }
}

function Invoke-TaskReconciliation {
    $path = Join-Path $LocalRoot $RegisterTasksName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Task-registration script is missing: $path" }
    Write-UpdaterStatus 'Reconciling scheduled maintenance tasks.' 'ACTION'
    $process = Start-Process -FilePath $PowerShellExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$path`"" -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) { throw "Task reconciliation returned exit code $($process.ExitCode)." }
    Write-UpdaterStatus 'Scheduled task reconciliation completed.' 'OK'
}

try {
    Ensure-Directory -Path $LocalRoot
    Ensure-Directory -Path $LogRoot
    Ensure-Directory -Path $RollbackRoot
    Ensure-Directory -Path $StagingRunRoot
    Write-UpdaterStatus "Starting manifest-driven update. RunId: $RunId" 'INFO'

    $sourceRoot = Get-ActiveSourceRoot
    Write-UpdaterStatus "Using source share: $sourceRoot" 'OK'
    $manifest = Read-DeploymentManifest -SourceRoot $sourceRoot
    Write-UpdaterStatus "Loaded manifest version $($manifest.ManifestVersion), package version $($manifest.PackageVersion), containing $(@($manifest.Files).Count) file(s)." 'INFO'

    $frameworkEntry = @($manifest.Files | Where-Object Name -eq $FrameworkName)[0]
    [void](Install-ManifestEntry -Entry $frameworkEntry -SourceRoot $sourceRoot)
    Import-RequiredFramework

    $orderedEntries = @(
        $manifest.Files | Where-Object { $_.Name -notin @($FrameworkName,$ScriptName) } | Sort-Object @{Expression={ if ($_.Name -eq $RegisterTasksName) { 1 } else { 0 } }}, Name
    )
    foreach ($entry in $orderedEntries) { [void](Install-ManifestEntry -Entry $entry -SourceRoot $sourceRoot) }

    $selfEntry = @($manifest.Files | Where-Object Name -eq $ScriptName)[0]
    $selfResult = Install-ManifestEntry -Entry $selfEntry -SourceRoot $sourceRoot
    if ($selfResult.SelfUpdated -and -not $Relaunched) {
        Write-UpdaterStatus 'Updater was replaced successfully. Relaunching the new version.' 'ACTION'
        Write-ExecutionRecord -Status 'Relaunching' -ExitCode 0 -FailureMessage $null
        & $PowerShellExe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $LocalRoot $ScriptName) -Relaunched
        $childExit = if ($null -eq $LASTEXITCODE) { 1 } else { [int]$LASTEXITCODE }
        exit $childExit
    }

    Invoke-TaskReconciliation
    Write-UpdaterStatus 'Manifest synchronization completed successfully.' 'OK'
    Write-ExecutionRecord -Status 'Success' -ExitCode 0 -FailureMessage $null
    exit 0
}
catch {
    $message = $_.Exception.Message
    Write-UpdaterStatus "Fatal updater error: $message" 'ERROR'
    Write-ExecutionRecord -Status 'Failed' -ExitCode 2 -FailureMessage $message
    exit 2
}
finally {
    try {
        if (Test-Path -LiteralPath $StagingRunRoot -PathType Container) { Remove-Item -LiteralPath $StagingRunRoot -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $StagingRoot -PathType Container) {
            Get-ChildItem -LiteralPath $StagingRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object LastWriteTime -lt (Get-Date).AddDays(-2) |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $RollbackRoot -PathType Container) {
            Get-ChildItem -LiteralPath $RollbackRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object LastWriteTime -lt (Get-Date).AddDays(-30) |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}
