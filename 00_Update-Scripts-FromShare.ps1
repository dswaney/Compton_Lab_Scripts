#requires -version 5.1
# =====================================================================
# ScriptName: 00_Update-Scripts-FromShare.ps1
# ScriptVersion: 4.5.0
# LastUpdated: 2026-09-03
# Changes: v4.5.0 deploys the combined script 04 and safely retires the six scripts it replaces.
# Changes: v4.4.0 approves and supplementally deploys script 19 for Stellarium Location Services.
# Changes: v4.3.0 approves scripts 17 and 18 for manifest-managed deployment.
# Purpose:
#   Manifest-driven, self-bootstrapping updater for C:\Scripts.
#   - Uses standardized Elastic-friendly text log format: timestamp, computer, level, message.
#   - Uses \\filesvr\Labscripts with an IP fallback.
#   - Downloads Maintenance.Framework.psm1 when missing or changed.
#   - Validates PowerShell syntax and updater structure before deployment.
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
$ScriptVersion = '4.5.0'
$PreferredSourceRoot = '\\filesvr\Labscripts'
$FallbackSourceRoot = '\\10.2.3.30\Labscripts'
$ManifestName = 'DeploymentManifest.json'
$FrameworkName = 'Maintenance.Framework.psm1'
$RegisterTasksName = 'Register-Tasks_SYSTEM.ps1'
$CombinedSundayScriptName = '04_Sunday_Lab_Application_Maintenance.ps1'
$DeepFreezeStatusScriptName = '16_Check_Deep_Freeze_Status.ps1'

# These scripts are now embedded in 04_Sunday_Lab_Application_Maintenance.ps1.
# They are removed from C:\Scripts only after the combined replacement has been
# deployed and passed PowerShell parser validation. A rollback copy is retained.
[string[]]$RetiredMaintenanceFiles = @(
    '11_Install_SharpDriver_And_PaperCut.ps1',
    '13_Configure_Autologon_And_Edge.ps1',
    '15_Install_Elastic_Agent.ps1',
    '17_Set_Browser_Homepage.ps1',
    '18_Install_Honorlock_Chrome_Extension.ps1',
    '19_Stellarium_Location_Services.ps1'
)

# Supplemental files are deployed directly from the active source share even
# when they are not yet listed in DeploymentManifest.json.
[string[]]$SupplementalManagedFiles = @(
    $CombinedSundayScriptName,
    $DeepFreezeStatusScriptName
)

# Only these files are permitted to deploy into C:\Scripts.
# Files present on the share but not listed here are intentionally ignored.
[string[]]$ApprovedMaintenanceFiles = @(
    '00_Update-Scripts-FromShare.ps1',
    '01_Enable_Windows_Update_Services.ps1',
    '02_Remove_User_Profiles.ps1',
    '03_Weekend_Apps_Update.ps1',
    '04_Sunday_Lab_Application_Maintenance.ps1',
    '05_Weekend_HP_Drivers_Update.ps1',
    '06_Weekend_Windows_Updates.ps1',
    '07_Force_Reboot_Install_Updates.ps1',
    '08_System_Repair.ps1',
    '09_Disable_Windows_Update_Services.ps1',
    '10_Sync_System_Time.ps1',
    '12_Enable-SystemRestore-And-Create-RestorePoint.ps1',
    '14_Endpoint_Health_Inventory.ps1',
    '16_Check_Deep_Freeze_Status.ps1',
    'Get-MaintenanceFleetStatus.ps1',
    'Invoke-MaintenanceScript.ps1',
    'Maintenance.Framework.psm1',
    'Maintenance.Policy.json',
    'Register-Tasks_SYSTEM.ps1'
)
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
    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:COMPUTERNAME, $Level, $Message
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
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateRange(1,20)][int]$MaxReadAttempts = 8,
        [ValidateRange(50,5000)][int]$InitialRetryDelayMilliseconds = 250
    )
    $extension = [IO.Path]::GetExtension($Path)
    if ($extension -notin @('.ps1','.psm1','.psd1')) {
        return [pscustomobject]@{ Valid = $true; Errors = @() }
    }

    $parseErrors = @()
    for ($attempt = 1; $attempt -le $MaxReadAttempts; $attempt++) {
        $tokens = $null
        $currentParseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$currentParseErrors) | Out-Null
        $parseErrors = @($currentParseErrors)

        $fileReadErrors = @(
            $parseErrors |
            Where-Object {
                $_.Extent.StartLineNumber -eq 0 -and
                $_.Message -match '(?i)file could not be read|cannot access the file|being used by another process|sharing violation'
            }
        )

        if ($fileReadErrors.Count -eq 0 -or $attempt -eq $MaxReadAttempts) {
            break
        }

        $delay = [math]::Min(
            $InitialRetryDelayMilliseconds * [math]::Pow(2, $attempt - 1),
            2000
        )
        Write-UpdaterStatus -Message ("Parser read attempt {0}/{1} was blocked for {2}; retrying in {3} ms." -f `
            $attempt, $MaxReadAttempts, [IO.Path]::GetFileName($Path), [int]$delay) -Level INFO
        Start-Sleep -Milliseconds ([int]$delay)
    }

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
        # SHA256 remains optional manifest metadata only; update decisions use the actual source-file hash.
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


function Test-UpdaterSourceStructure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $requiredFunctions = @(
        'Ensure-Directory',
        'Write-UpdaterStatus',
        'Get-Sha256',
        'Test-PowerShellFile',
        'Get-FileVersionText',
        'Get-ActiveSourceRoot',
        'Read-DeploymentManifest',
        'Add-Result',
        'Install-ManifestEntry',
        'Remove-RetiredMaintenanceFiles',
        'Import-RequiredFramework',
        'Write-ExecutionRecord',
        'Invoke-TaskReconciliation'
    )

    $parserResult = Test-PowerShellFile -Path $Path
    if (-not $parserResult.Valid) {
        $summary = @(
            $parserResult.Errors |
            ForEach-Object { "Line $($_.Line), column $($_.Column): $($_.Message)" }
        ) -join ' | '

        return [pscustomobject]@{
            Valid   = $false
            Message = "PowerShell parser validation failed. $summary"
        }
    }

    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $missingFunctions = @()

    foreach ($functionName in $requiredFunctions) {
        $pattern = '(?m)^\s*function\s+{0}\s*\{{' -f [regex]::Escape($functionName)

        if ($content -notmatch $pattern) {
            $missingFunctions += $functionName
        }
    }

    if ($missingFunctions.Count -gt 0) {
        return [pscustomobject]@{
            Valid   = $false
            Message = "Missing required function(s): $($missingFunctions -join ', ')"
        }
    }

    return [pscustomobject]@{
        Valid   = $true
        Message = 'Updater source structure is valid.'
    }
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

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "Manifest file is missing from source share: $sourcePath" }

    # The active network share is authoritative. Compare the actual source file
    # directly with the local file instead of using the manifest SHA as the
    # update decision. This prevents stale manifest hashes from blocking updates.
    $sourceHash = Get-Sha256 -Path $sourcePath

    if ($localExists -and $localHash -eq $sourceHash) {
        Write-UpdaterStatus "$name is current." 'OK'
        Add-Result -Name $name -Role $role -Status 'Current' -LocalVersion $localVersion -ShareVersion $shareVersion -LocalHash $localHash -ShareHash $sourceHash -BackupPath $null -Message $null
        return [pscustomobject]@{ Changed=$false; SelfUpdated=$false }
    }

    if ($name -ieq $ScriptName) {
        $selfValidation = Test-UpdaterSourceStructure -Path $sourcePath

        if (-not $selfValidation.Valid) {
            Write-UpdaterStatus `
                -Message "The source copy of $ScriptName is incomplete and will not replace the working local updater. $($selfValidation.Message)" `
                -Level WARN

            Add-Result `
                -Name $name `
                -Role $role `
                -Status 'SkippedInvalidUpdaterSource' `
                -LocalVersion $localVersion `
                -ShareVersion $shareVersion `
                -LocalHash $localHash `
                -ShareHash $expectedHash `
                -BackupPath $null `
                -Message $selfValidation.Message

            return [pscustomobject]@{
                Changed     = $false
                SelfUpdated = $false
            }
        }

        Write-UpdaterStatus -Message 'Source updater passed structural validation.' -Level OK
    }


    Copy-Item -LiteralPath $sourcePath -Destination $stagedPath -Force -ErrorAction Stop
    $stagedHash = Get-Sha256 -Path $stagedPath

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
    Add-Result -Name $name -Role $role -Status $status -LocalVersion $localVersion -ShareVersion $shareVersion -LocalHash (Get-Sha256 -Path $destinationPath) -ShareHash $sourceHash -BackupPath $backupPath -Message $null
    [pscustomobject]@{ Changed=$true; SelfUpdated=($name -ieq $ScriptName) }
}


function Install-SupplementalManagedFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceRoot
    )

    foreach ($fileName in $SupplementalManagedFiles) {
        if ($fileName -notin $ApprovedMaintenanceFiles) {
            Write-UpdaterStatus -Message "Supplemental file is not approved for maintenance deployment; skipping: $fileName" -Level WARN
            continue
        }

        $sourcePath = Join-Path $SourceRoot $fileName

        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            Write-UpdaterStatus -Message "Supplemental managed file is not present on the source share; skipping: $fileName" -Level INFO
            continue
        }

        $sourceVersion = Get-FileVersionText -Path $sourcePath

        $entry = [pscustomobject]@{
            Name    = $fileName
            Role    = 'SupplementalManagedFile'
            Version = $sourceVersion
            SHA256  = ''
        }

        Write-UpdaterStatus -Message "Checking supplemental managed file: $fileName" -Level INFO
        [void](Install-ManifestEntry -Entry $entry -SourceRoot $SourceRoot)
    }
}

function Remove-RetiredMaintenanceFiles {
    [CmdletBinding()]
    param()

    $combinedPath = Join-Path $LocalRoot $CombinedSundayScriptName
    if (-not (Test-Path -LiteralPath $combinedPath -PathType Leaf)) {
        throw "The combined replacement is missing; retired scripts will not be removed: $combinedPath"
    }

    $combinedParserResult = Test-PowerShellFile -Path $combinedPath
    if (-not $combinedParserResult.Valid) {
        throw "The combined replacement failed parser validation; retired scripts will not be removed: $combinedPath"
    }

    foreach ($fileName in $RetiredMaintenanceFiles) {
        $retiredPath = Join-Path $LocalRoot $fileName
        if (-not (Test-Path -LiteralPath $retiredPath -PathType Leaf)) {
            Write-UpdaterStatus -Message "Retired maintenance script is already absent: $fileName" -Level INFO
            continue
        }

        $retiredBackupRoot = Join-Path $RollbackRunRoot 'Retired'
        Ensure-Directory -Path $retiredBackupRoot
        $backupPath = Join-Path $retiredBackupRoot $fileName
        Move-Item -LiteralPath $retiredPath -Destination $backupPath -Force -ErrorAction Stop

        Write-UpdaterStatus -Message "Removed retired maintenance script from C:\Scripts and retained a rollback copy: $fileName" -Level OK
        Add-Result `
            -Name $fileName `
            -Role 'RetiredMaintenanceFile' `
            -Status 'Retired' `
            -LocalVersion 'Retired' `
            -ShareVersion 'ReplacedByScript04' `
            -LocalHash $null `
            -ShareHash $null `
            -BackupPath $backupPath `
            -Message "Replaced by $CombinedSundayScriptName"
    }
}

function Import-RequiredFramework {
    $frameworkPath = Join-Path $LocalRoot $FrameworkName
    if (-not (Test-Path -LiteralPath $frameworkPath -PathType Leaf)) { throw "Framework bootstrap failed; file is missing: $frameworkPath" }
    Import-Module -Name $frameworkPath -Force -DisableNameChecking -ErrorAction Stop
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
        WarningCount=$WarningCount;ErrorCount=$ErrorCount;Results=@($Results | ForEach-Object { $_ });FailureMessage=$FailureMessage
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

    $ignoredManifestEntries = @(
        $manifest.Files |
        Where-Object { $_.Name -notin $ApprovedMaintenanceFiles }
    )

    foreach ($ignoredEntry in $ignoredManifestEntries) {
        Write-UpdaterStatus -Message "Ignoring non-maintenance manifest entry: $($ignoredEntry.Name)" -Level INFO
    }

    $approvedManifestEntries = @(
        $manifest.Files |
        Where-Object { $_.Name -in $ApprovedMaintenanceFiles }
    )

    $frameworkEntry = @($approvedManifestEntries | Where-Object Name -eq $FrameworkName)[0]
    [void](Install-ManifestEntry -Entry $frameworkEntry -SourceRoot $sourceRoot)
    Import-RequiredFramework

    $orderedEntries = @(
        $approvedManifestEntries | Where-Object { $_.Name -notin (@($FrameworkName,$ScriptName) + $SupplementalManagedFiles) } | Sort-Object @{Expression={ if ($_.Name -eq $RegisterTasksName) { 1 } else { 0 } }}, Name
    )
    foreach ($entry in $orderedEntries) { [void](Install-ManifestEntry -Entry $entry -SourceRoot $sourceRoot) }

    # Deploy supplemental scripts that are intentionally managed outside the
    # current DeploymentManifest.json, including combined script 04 and script 16.
    Install-SupplementalManagedFiles -SourceRoot $sourceRoot

    # Remove the six standalone scripts only after their combined replacement
    # exists locally and has passed parser validation.
    Remove-RetiredMaintenanceFiles

    $selfEntry = @($approvedManifestEntries | Where-Object Name -eq $ScriptName)[0]
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
