# ============================================================================
# ModuleVersion: 2.4.0
# DateUpdated: 2026-08-17
# ============================================================================

#requires -version 5.1
<#
.SYNOPSIS
    Shared production framework for the Compton College maintenance scripts.
.DESCRIPTION
    Defines the canonical C:\Scripts installation root, C:\Logs output root,
    common status/logging helpers, standardized exit codes, administrative
    checks, run context creation, file resolution, and concurrency-safe NDJSON
    telemetry with size-based rotation and retention. Also provides staged
    text logging so active logs remain outside Elastic's watched C:\Logs path
    until the maintenance run is complete.
#>

Set-StrictMode -Version 2.0

$script:FrameworkVersion = '2.4.0'
$script:ProductionScriptRoot = 'C:\Scripts'
$script:DefaultLogRoot = 'C:\Logs'
$script:DefaultLogStagingRoot = 'C:\Temp\Compton-Maintenance-Logs'
$script:DefaultOldLogsFolderName = 'Old Logs'
$script:DefaultTelemetryMaxSizeMB = 50
$script:DefaultTelemetryRetentionDays = 60
$script:DefaultTelemetryMutexTimeoutSeconds = 30
$script:DefaultTelemetryMutexName = 'Global\Compton_MaintenanceTelemetry'

$script:ExitCodes = [ordered]@{
    Success             = 0
    SuccessWithWarnings = 1
    RecoverableFailure  = 2
    CriticalFailure     = 3
    TelemetryFailure    = 4
}

function Get-MaintenanceFrameworkVersion {
    [CmdletBinding()]
    param()
    $script:FrameworkVersion
}

function Get-MaintenanceExitCodes {
    [CmdletBinding()]
    param()
    [pscustomobject]$script:ExitCodes
}

function Get-MaintenanceConfiguration {
    [CmdletBinding()]
    param(
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [string]$LogRoot = $script:DefaultLogRoot
    )

    if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { throw 'ScriptRoot cannot be empty.' }
    if ([string]::IsNullOrWhiteSpace($LogRoot)) { throw 'LogRoot cannot be empty.' }

    [pscustomobject]@{
        FrameworkVersion             = $script:FrameworkVersion
        ScriptRoot                   = [IO.Path]::GetFullPath($ScriptRoot)
        LogRoot                      = [IO.Path]::GetFullPath($LogRoot)
        LogStagingRoot               = [IO.Path]::GetFullPath($script:DefaultLogStagingRoot)
        OldLogsRoot                  = Join-Path ([IO.Path]::GetFullPath($LogRoot)) $script:DefaultOldLogsFolderName
        TelemetryPath                = Join-Path $LogRoot 'Maintenance-Telemetry.ndjson'
        TelemetryMaxSizeMB           = $script:DefaultTelemetryMaxSizeMB
        TelemetryRetentionDays       = $script:DefaultTelemetryRetentionDays
        TelemetryMutexTimeoutSeconds = $script:DefaultTelemetryMutexTimeoutSeconds
        TelemetryMutexName           = $script:DefaultTelemetryMutexName
        ExitCodes                    = Get-MaintenanceExitCodes
    }
}

function Initialize-MaintenanceDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName
}

function Initialize-MaintenanceEnvironment {
    [CmdletBinding()]
    param(
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [string]$LogRoot = $script:DefaultLogRoot,
        [switch]$RequireScriptRoot
    )
    $configuration = Get-MaintenanceConfiguration -ScriptRoot $ScriptRoot -LogRoot $LogRoot
    if ($RequireScriptRoot -and -not (Test-Path -LiteralPath $configuration.ScriptRoot -PathType Container)) {
        throw "Required script root is missing: $($configuration.ScriptRoot)"
    }
    Initialize-MaintenanceDirectory -Path $configuration.LogRoot | Out-Null
    $configuration
}

function Test-MaintenanceAdministrator {
    [CmdletBinding()]
    param()
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { $false }
}


function Archive-MaintenanceLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptName,
        [string]$LogRoot = $script:DefaultLogRoot,
        [string]$OldLogsRoot = (Join-Path $LogRoot $script:DefaultOldLogsFolderName),
        [string[]]$AdditionalPatterns,
        [string[]]$ExcludeNames = @('Maintenance-Telemetry.ndjson'),
        [switch]$IncludeJson,
        [switch]$IncludeYaml
    )

    Initialize-MaintenanceDirectory -Path $LogRoot | Out-Null
    Initialize-MaintenanceDirectory -Path $OldLogsRoot | Out-Null

    $baseName = [IO.Path]::GetFileNameWithoutExtension($ScriptName)

    $patterns = @(
        "$baseName.log",
        "$baseName-*.log",
        "*-$baseName-*.log",
        "${baseName}_*.log"
    )

    if ($IncludeJson) {
        $patterns += @(
            "$baseName.json",
            "$baseName.latest.json",
            "$baseName-*.json",
            "*-$baseName-*.json"
        )
    }

    if ($IncludeYaml) {
        $patterns += @(
            "$baseName.yml",
            "$baseName.yaml",
            "$baseName-*.yml",
            "$baseName-*.yaml",
            "*-$baseName-*.yml",
            "*-$baseName-*.yaml"
        )
    }

    if ($AdditionalPatterns) {
        $patterns += @($AdditionalPatterns)
    }

    $seen = @{}
    $results = @()

    foreach ($pattern in ($patterns | Where-Object { $_ } | Sort-Object -Unique)) {
        foreach ($file in @(Get-ChildItem -LiteralPath $LogRoot -File -Filter $pattern -ErrorAction SilentlyContinue)) {
            if ($ExcludeNames -contains $file.Name) { continue }
            if ($seen.ContainsKey($file.FullName)) { continue }

            $seen[$file.FullName] = $true
            $destination = Join-Path $OldLogsRoot $file.Name

            if (Test-Path -LiteralPath $destination -PathType Leaf) {
                $base = [IO.Path]::GetFileNameWithoutExtension($file.Name)
                $extension = [IO.Path]::GetExtension($file.Name)
                $destination = Join-Path $OldLogsRoot (
                    '{0}_{1}{2}' -f $base, (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $extension
                )
            }

            try {
                Move-Item -LiteralPath $file.FullName -Destination $destination -Force -ErrorAction Stop
                $results += [pscustomobject][ordered]@{
                    Source      = $file.FullName
                    Destination = $destination
                    Status      = 'Archived'
                    Error       = $null
                }
            }
            catch {
                # A file already opened by Elastic must not fail the maintenance run.
                $results += [pscustomobject][ordered]@{
                    Source      = $file.FullName
                    Destination = $destination
                    Status      = 'SkippedLockedOrUnavailable'
                    Error       = $_.Exception.Message
                }

                Write-Warning (
                    "Unable to archive maintenance log '{0}': {1}" -f
                    $file.FullName, $_.Exception.Message
                )
            }
        }
    }

    return @($results)
}

function New-MaintenanceStagedLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptName,
        [string]$LogRoot = $script:DefaultLogRoot,
        [string]$StagingRoot = $script:DefaultLogStagingRoot,
        [string]$ComputerName = $env:COMPUTERNAME,
        [datetime]$Timestamp = (Get-Date)
    )

    Initialize-MaintenanceDirectory -Path $LogRoot | Out-Null
    Initialize-MaintenanceDirectory -Path $StagingRoot | Out-Null

    $baseName = [IO.Path]::GetFileNameWithoutExtension($ScriptName)
    $safeComputer = if ([string]::IsNullOrWhiteSpace($ComputerName)) {
        'UNKNOWN'
    }
    else {
        $ComputerName -replace '[^A-Za-z0-9_.-]', '_'
    }

    $stamp = $Timestamp.ToString('yyyy-MM-dd_HHmmss')
    $fileName = '{0}-{1}-{2}.log' -f $safeComputer, $baseName, $stamp

    $finalPath = Join-Path $LogRoot $fileName
    $workingPath = Join-Path $StagingRoot $fileName

    if (Test-Path -LiteralPath $workingPath -PathType Leaf) {
        Remove-Item -LiteralPath $workingPath -Force -ErrorAction SilentlyContinue
    }

    [pscustomobject][ordered]@{
        ScriptName   = $ScriptName
        FileName     = $fileName
        WorkingPath  = $workingPath
        PublishedPath= $finalPath
        StagingRoot  = $StagingRoot
        LogRoot      = $LogRoot
        Published    = $false
    }
}

function Publish-MaintenanceLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$LogSession
    )

    $workingPath = [string]$LogSession.WorkingPath
    $publishedPath = [string]$LogSession.PublishedPath

    if (-not (Test-Path -LiteralPath $workingPath -PathType Leaf)) {
        return [pscustomobject][ordered]@{
            Published = $false
            Path      = $publishedPath
            Reason    = 'Working log did not exist.'
        }
    }

    Initialize-MaintenanceDirectory -Path (Split-Path -Parent $publishedPath) | Out-Null

    try {
        Move-Item `
            -LiteralPath $workingPath `
            -Destination $publishedPath `
            -Force `
            -ErrorAction Stop

        if ($LogSession.PSObject.Properties['Published']) {
            $LogSession.Published = $true
        }

        return [pscustomobject][ordered]@{
            Published = $true
            Path      = $publishedPath
            Reason    = $null
        }
    }
    catch {
        # Publishing failure is nonfatal. Leave the complete file in staging so
        # it can be recovered manually or by a later cleanup/publish workflow.
        Write-Warning (
            "Unable to publish completed maintenance log '{0}' to '{1}': {2}" -f
            $workingPath, $publishedPath, $_.Exception.Message
        )

        return [pscustomobject][ordered]@{
            Published = $false
            Path      = $workingPath
            Reason    = $_.Exception.Message
        }
    }
}

function New-MaintenanceRunContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptName,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptVersion,
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [string]$LogRoot = $script:DefaultLogRoot,
        [switch]$UseStagedLogging,
        [switch]$ArchivePreviousLogs,
        [string[]]$AdditionalArchivePatterns
    )

    $config = Initialize-MaintenanceEnvironment -ScriptRoot $ScriptRoot -LogRoot $LogRoot
    $logSession = $null

    if ($ArchivePreviousLogs) {
        Archive-MaintenanceLogs `
            -ScriptName $ScriptName `
            -LogRoot $config.LogRoot `
            -OldLogsRoot $config.OldLogsRoot `
            -AdditionalPatterns $AdditionalArchivePatterns | Out-Null
    }

    if ($UseStagedLogging) {
        $logSession = New-MaintenanceStagedLog `
            -ScriptName $ScriptName `
            -LogRoot $config.LogRoot `
            -StagingRoot $config.LogStagingRoot

        $logPath = $logSession.WorkingPath
        $publishedLogPath = $logSession.PublishedPath
    }
    else {
        $logPath = Join-Path $config.LogRoot ('{0}.log' -f [IO.Path]::GetFileNameWithoutExtension($ScriptName))
        $publishedLogPath = $logPath
    }

    [pscustomobject]@{
        ScriptName       = $ScriptName
        ScriptVersion    = $ScriptVersion
        ScriptRoot       = $config.ScriptRoot
        LogRoot          = $config.LogRoot
        LogPath          = $logPath
        WorkingLogPath   = $logPath
        PublishedLogPath = $publishedLogPath
        UseStagedLogging = [bool]$UseStagedLogging
        LogSession       = $logSession
        RunId            = [guid]::NewGuid().Guid
        StartTime        = Get-Date
        WarningCount     = 0
        ErrorCount       = 0
        ExitCodes        = $config.ExitCodes
        Configuration    = $config
    }
}


function Write-MaintenanceStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','SUCCESS','WARNING')][string]$Level = 'INFO',
        [string]$LogPath,
        $Context,
        [string]$ScriptName,
        [string]$RunId,
        [ValidateRange(1,1024)][int]$MaxLogSizeMB = 20,
        [ValidateRange(1,3650)][int]$RetentionDays = 60
    )
    if ($Context) {
        if (-not $LogPath -and $Context.PSObject.Properties['LogPath']) { $LogPath = [string]$Context.LogPath }
        if (-not $ScriptName -and $Context.PSObject.Properties['ScriptName']) { $ScriptName = [string]$Context.ScriptName }
        if (-not $RunId -and $Context.PSObject.Properties['RunId']) { $RunId = [string]$Context.RunId }
    }

    $canonicalLevel = switch ($Level) {
        'OK'      { 'SUCCESS' }
        'SUCCESS' { 'SUCCESS' }
        'WARN'    { 'WARNING' }
        'WARNING' { 'WARNING' }
        default   { $Level }
    }
    if ($Context -and $canonicalLevel -eq 'WARNING' -and $Context.PSObject.Properties['WarningCount']) { $Context.WarningCount++ }
    if ($Context -and $canonicalLevel -eq 'ERROR' -and $Context.PSObject.Properties['ErrorCount']) { $Context.ErrorCount++ }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
    $metadata = @()
    if ($ScriptName) { $metadata += "Script=$ScriptName" }
    if ($RunId) { $metadata += "RunId=$RunId" }
    $metadataText = if ($metadata.Count -gt 0) { ' [' + ($metadata -join ' ') + ']' } else { '' }
    $line = "$timestamp [$computerName] [$canonicalLevel]$metadataText $Message"
    $color = switch ($canonicalLevel) { 'INFO' {'Cyan'} 'SUCCESS' {'Green'} 'WARNING' {'Yellow'} 'ERROR' {'Red'} default {'White'} }
    Write-Host $line -ForegroundColor $color
    if ($LogPath) {
        Initialize-MaintenanceDirectory -Path (Split-Path -Parent $LogPath) | Out-Null
        Rotate-MaintenanceLog -Path $LogPath -MaxSizeMB $MaxLogSizeMB -RetentionDays $RetentionDays
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    $line
}

function Start-MaintenanceScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptName,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptVersion,
        [switch]$RequireAdministrator,
        [switch]$UseStagedLogging,
        [switch]$ArchivePreviousLogs,
        [string[]]$AdditionalArchivePatterns
    )

    if ($RequireAdministrator -and -not (Test-MaintenanceAdministrator)) {
        throw "$ScriptName must run from an elevated administrator session or as SYSTEM."
    }

    $context = New-MaintenanceRunContext `
        -ScriptName $ScriptName `
        -ScriptVersion $ScriptVersion `
        -UseStagedLogging:$UseStagedLogging `
        -ArchivePreviousLogs:$ArchivePreviousLogs `
        -AdditionalArchivePatterns $AdditionalArchivePatterns

    Write-MaintenanceStatus `
        -Message "Starting $ScriptName version $ScriptVersion." `
        -Level INFO `
        -Context $context | Out-Null

    if ($context.UseStagedLogging) {
        Write-MaintenanceStatus `
            -Message ("Active text log is staged at: {0}" -f $context.WorkingLogPath) `
            -Level INFO `
            -Context $context | Out-Null

        Write-MaintenanceStatus `
            -Message ("Completed text log will be published to: {0}" -f $context.PublishedLogPath) `
            -Level INFO `
            -Context $context | Out-Null
    }

    $context
}


function Stop-MaintenanceScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [ValidateSet('Success','SuccessWithWarnings','RecoverableFailure','CriticalFailure','TelemetryFailure')]
        [string]$Status = 'Success'
    )

    $end = Get-Date
    $duration = [math]::Round(($end - $Context.StartTime).TotalSeconds,2)
    $exitCode = [int]$Context.ExitCodes.$Status

    # This completion entry must be the final write to the staged text log.
    Write-MaintenanceStatus `
        -Message "Completed $($Context.ScriptName). Status=$Status; ExitCode=$exitCode; DurationSeconds=$duration" `
        -Level $(if ($exitCode -eq 0) {'SUCCESS'} elseif ($exitCode -eq 1) {'WARNING'} else {'ERROR'}) `
        -Context $Context | Out-Null

    $publishResult = $null

    if (
        $Context.PSObject.Properties['UseStagedLogging'] -and
        $Context.UseStagedLogging -and
        $Context.PSObject.Properties['LogSession'] -and
        $null -ne $Context.LogSession
    ) {
        $publishResult = Publish-MaintenanceLog -LogSession $Context.LogSession

        if ($publishResult.Published) {
            $Context.LogPath = [string]$Context.PublishedLogPath
        }
    }

    [pscustomobject]@{
        Status          = $Status
        ExitCode        = $exitCode
        EndTime         = $end
        DurationSeconds = $duration
        LogPath         = if ($Context.PSObject.Properties['PublishedLogPath']) { [string]$Context.PublishedLogPath } else { [string]$Context.LogPath }
        LogPublished    = if ($null -ne $publishResult) { [bool]$publishResult.Published } else { $true }
        PublishReason   = if ($null -ne $publishResult) { $publishResult.Reason } else { $null }
    }
}


function Resolve-MaintenanceFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Name,
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [switch]$Require
    )
    $path = Join-Path $ScriptRoot $Name
    if ($Require -and -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required maintenance file is missing: $path" }
    $path
}

function Write-MaintenanceTelemetryLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$JsonLine,
        [ValidateRange(1,4096)][int]$MaxSizeMB = $script:DefaultTelemetryMaxSizeMB,
        [ValidateRange(1,3650)][int]$RetentionDays = $script:DefaultTelemetryRetentionDays,
        [ValidateRange(1,300)][int]$MutexTimeoutSeconds = $script:DefaultTelemetryMutexTimeoutSeconds,
        [ValidateNotNullOrEmpty()][string]$MutexName = $script:DefaultTelemetryMutexName
    )
    $directory = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($directory)) { throw "Telemetry path must include a directory: $Path" }
    Initialize-MaintenanceDirectory -Path $directory | Out-Null

    $mutex = $null; $acquired = $false
    try {
        $createdNew = $false
        $mutex = New-Object System.Threading.Mutex($false,$MutexName,[ref]$createdNew)
        try { $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds($MutexTimeoutSeconds)) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw "Timed out after $MutexTimeoutSeconds seconds waiting for telemetry mutex '$MutexName'." }

        $encoding = New-Object System.Text.UTF8Encoding($false)
        $line = $JsonLine.TrimEnd("`r","`n") + [Environment]::NewLine
        $incoming = $encoding.GetByteCount($line)
        $maxBytes = [int64]$MaxSizeMB * 1MB

        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            $item = Get-Item -LiteralPath $Path -ErrorAction Stop
            if (($item.Length + $incoming) -gt $maxBytes -and $item.Length -gt 0) {
                $base = [IO.Path]::GetFileNameWithoutExtension($Path)
                $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
                $rotated = Join-Path $directory ("{0}_{1}.ndjson" -f $base,$stamp)
                $archive = [IO.Path]::ChangeExtension($rotated,'.zip')
                Move-Item -LiteralPath $Path -Destination $rotated -Force -ErrorAction Stop
                try {
                    Compress-Archive -LiteralPath $rotated -DestinationPath $archive -CompressionLevel Optimal -Force -ErrorAction Stop
                    Remove-Item -LiteralPath $rotated -Force -ErrorAction Stop
                }
                catch { Write-Warning "Telemetry rotated to '$rotated', but compression failed: $($_.Exception.Message)" }
            }
        }

        [IO.File]::AppendAllText($Path,$line,$encoding)
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        $pattern = ([IO.Path]::GetFileNameWithoutExtension($Path)) + '_*'
        Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like $pattern -and $_.Extension -in @('.zip','.ndjson') -and $_.LastWriteTime -lt $cutoff } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
    finally {
        if ($acquired -and $null -ne $mutex) { try { $mutex.ReleaseMutex() } catch {} }
        if ($null -ne $mutex) { $mutex.Dispose() }
    }
}

function Write-MaintenanceTelemetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Record,
        [string]$Path = (Join-Path $script:DefaultLogRoot 'Maintenance-Telemetry.ndjson'),
        [int]$Depth = 12
    )
    $json = $Record | ConvertTo-Json -Compress -Depth $Depth
    Write-MaintenanceTelemetryLine -Path $Path -JsonLine $json
}

function Write-MaintenanceJsonAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [Parameter(Mandatory)]$Record,
        [ValidateRange(2,100)][int]$Depth = 12,
        [switch]$Compress
    )

    $directory = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($directory)) { throw "JSON path must include a directory: $Path" }
    Initialize-MaintenanceDirectory -Path $directory | Out-Null
    $json = if ($Compress) { $Record | ConvertTo-Json -Depth $Depth -Compress } else { $Record | ConvertTo-Json -Depth $Depth }
    $temporary = '{0}.{1}.tmp' -f $Path, ([guid]::NewGuid().Guid)
    try {
        [IO.File]::WriteAllText($temporary, $json, (New-Object Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $temporary -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $temporary -PathType Leaf) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue }
    }
    $Path
}

function New-MaintenanceTelemetryRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [datetime]$EndTime = (Get-Date),
        [string]$FailureMessage,
        [hashtable]$AdditionalFields
    )

    $domain = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { $env:USERDOMAIN }
    $record = [ordered]@{
        '@timestamp'    = $EndTime.ToUniversalTime().ToString('o')
        EventType       = 'maintenance.execution'
        ComputerName    = $env:COMPUTERNAME
        Domain          = $domain
        ScriptName      = [string]$Context.ScriptName
        ScriptVersion   = [string]$Context.ScriptVersion
        FrameworkVersion= $script:FrameworkVersion
        RunId           = [string]$Context.RunId
        Status          = $Status
        ExitCode        = $ExitCode
        StartTime       = $Context.StartTime.ToUniversalTime().ToString('o')
        EndTime         = $EndTime.ToUniversalTime().ToString('o')
        DurationSeconds = [math]::Round(($EndTime - $Context.StartTime).TotalSeconds, 3)
        WarningCount    = [int]$Context.WarningCount
        ErrorCount      = [int]$Context.ErrorCount
        FailureMessage  = $FailureMessage
        TextLogPath     = if ($Context.PSObject.Properties['PublishedLogPath']) {
            [string]$Context.PublishedLogPath
        }
        elseif ($Context.PSObject.Properties['LogPath']) {
            [string]$Context.LogPath
        }
        else {
            $null
        }
    }
    if ($AdditionalFields) {
        foreach ($key in $AdditionalFields.Keys) { $record[$key] = $AdditionalFields[$key] }
    }
    [pscustomobject]$record
}



function Rotate-MaintenanceLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateRange(1,1024)][int]$MaxSizeMB = 20,
        [ValidateRange(1,3650)][int]$RetentionDays = 60
    )
    $directory = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($directory)) { throw "Log path must include a directory: $Path" }
    Initialize-MaintenanceDirectory -Path $directory | Out-Null
    if ((Test-Path -LiteralPath $Path -PathType Leaf) -and ((Get-Item -LiteralPath $Path).Length -ge ([int64]$MaxSizeMB * 1MB))) {
        $base = [IO.Path]::GetFileNameWithoutExtension($Path)
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $rotated = Join-Path $directory ("{0}_{1}{2}" -f $base,$stamp,[IO.Path]::GetExtension($Path))
        $archive = [IO.Path]::ChangeExtension($rotated,'.zip')
        Move-Item -LiteralPath $Path -Destination $rotated -Force
        try {
            Compress-Archive -LiteralPath $rotated -DestinationPath $archive -CompressionLevel Optimal -Force
            Remove-Item -LiteralPath $rotated -Force
        } catch { Write-Warning "Log rotation compression failed for '$rotated': $($_.Exception.Message)" }
    }
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $basePattern = ([IO.Path]::GetFileNameWithoutExtension($Path)) + '_*'
    Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like $basePattern -and $_.LastWriteTime -lt $cutoff } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Get-MaintenanceCorrelationId {
    [CmdletBinding()]
    param([string]$CorrelationId)
    if (-not [string]::IsNullOrWhiteSpace($CorrelationId)) { return $CorrelationId }
    if (-not [string]::IsNullOrWhiteSpace($env:COMPTON_MAINTENANCE_CORRELATION_ID)) { return $env:COMPTON_MAINTENANCE_CORRELATION_ID }
    return [guid]::NewGuid().Guid
}

function Enter-MaintenanceExecutionLock {
    [CmdletBinding()]
    param(
        [string]$Name = 'Global\Compton_MaintenanceExecution',
        [ValidateRange(1,43200)][int]$TimeoutSeconds = 60
    )
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($false,$Name,[ref]$createdNew)
    try {
        try { $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds)) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { $mutex.Dispose(); throw "Timed out waiting for maintenance execution lock '$Name'." }
        [pscustomobject]@{ Name=$Name; Mutex=$mutex; Acquired=$true; AcquiredAt=Get-Date }
    } catch { if ($mutex) { $mutex.Dispose() }; throw }
}

function Exit-MaintenanceExecutionLock {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Lock)
    if ($Lock -and $Lock.Mutex) {
        try { if ($Lock.Acquired) { $Lock.Mutex.ReleaseMutex() } } catch {}
        try { $Lock.Mutex.Dispose() } catch {}
    }
}

function Get-MaintenancePolicy {
    [CmdletBinding()]
    param([string]$Path = 'C:\Scripts\Maintenance.Policy.json')
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Maintenance policy is missing: $Path" }
    Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}

function Test-MaintenanceWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        $Policy = (Get-MaintenancePolicy),
        [datetime]$At = (Get-Date)
    )
    $entry = @($Policy.Scripts | Where-Object { $_.Name -ieq $ScriptName }) | Select-Object -First 1
    if (-not $entry -or -not $entry.Window -or $entry.Window.Enforce -eq $false) {
        return [pscustomobject]@{ Allowed=$true; Reason='No enforced window'; ScriptName=$ScriptName }
    }
    $days = @($entry.Window.Days)
    if ($days.Count -gt 0 -and $At.DayOfWeek.ToString() -notin $days) {
        return [pscustomobject]@{ Allowed=$false; Reason="Day $($At.DayOfWeek) is outside the approved window"; ScriptName=$ScriptName }
    }
    $start = [TimeSpan]::Parse([string]$entry.Window.Start)
    $end = [TimeSpan]::Parse([string]$entry.Window.End)
    $time = $At.TimeOfDay
    $allowed = if ($end -ge $start) { $time -ge $start -and $time -le $end } else { $time -ge $start -or $time -le $end }
    [pscustomobject]@{ Allowed=$allowed; Reason=if($allowed){'Within approved window'}else{"Time $($At.ToString('HH:mm')) is outside $($entry.Window.Start)-$($entry.Window.End)"}; ScriptName=$ScriptName }
}

function Test-MaintenanceDependencies {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        $Policy = (Get-MaintenancePolicy)
    )
    $entry = @($Policy.Scripts | Where-Object { $_.Name -ieq $ScriptName }) | Select-Object -First 1
    $results = @()
    foreach ($dependency in @($entry.Dependencies)) {
        $ok=$false; $detail=$null
        switch ([string]$dependency.Type) {
            'File'    { $ok=Test-Path -LiteralPath ([string]$dependency.Value) -PathType Leaf; $detail=[string]$dependency.Value }
            'Folder'  { $ok=Test-Path -LiteralPath ([string]$dependency.Value) -PathType Container; $detail=[string]$dependency.Value }
            'Command' { $ok=$null -ne (Get-Command ([string]$dependency.Value) -ErrorAction SilentlyContinue); $detail=[string]$dependency.Value }
            'Module'  { $ok=$null -ne (Get-Module -ListAvailable -Name ([string]$dependency.Value) | Select-Object -First 1); $detail=[string]$dependency.Value }
            'Service' { $svc=Get-Service -Name ([string]$dependency.Value) -ErrorAction SilentlyContinue; $ok=$null -ne $svc; $detail=if($svc){$svc.Status}else{'Missing'} }
            default   { $ok=$false; $detail="Unsupported dependency type: $($dependency.Type)" }
        }
        $results += [pscustomobject]@{ Type=[string]$dependency.Type; Value=[string]$dependency.Value; Required=($dependency.Required -ne $false); Passed=$ok; Detail=$detail }
    }
    $failedRequired=@($results | Where-Object { $_.Required -and -not $_.Passed })
    [pscustomobject]@{ Passed=($failedRequired.Count -eq 0); Results=@($results); FailedRequired=$failedRequired }
}

function Initialize-MaintenanceEventLog {
    [CmdletBinding()]
    param([string]$LogName='Compton Maintenance',[string]$Source='Compton-Maintenance')
    if (-not [System.Diagnostics.EventLog]::SourceExists($Source)) {
        if (-not (Test-MaintenanceAdministrator)) { return $false }
        New-EventLog -LogName $LogName -Source $Source -ErrorAction Stop
    }
    $true
}

function Write-MaintenanceEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Information','Warning','Error')][string]$EntryType='Information',
        [ValidateRange(1,65535)][int]$EventId=1000,
        [string]$LogName='Compton Maintenance',[string]$Source='Compton-Maintenance'
    )
    try {
        if (Initialize-MaintenanceEventLog -LogName $LogName -Source $Source) {
            try {
                Write-EventLog -LogName $LogName -Source $Source -EventId $EventId -EntryType $EntryType -Message $Message -ErrorAction Stop
            }
            catch {
                Write-Warning ("Unable to write Windows Event Log telemetry to '{0}' using source '{1}': {2}" -f $LogName, $Source, $_.Exception.Message)
            }
            return $true
        }
    } catch { Write-Warning "Unable to write maintenance event: $($_.Exception.Message)" }
    $false
}

function Publish-MaintenanceFleetStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Record,
        [string[]]$StatusRoots=@('\\filesvr\MaintenanceStatus','\\10.2.3.30\MaintenanceStatus'),
        [string]$LocalFallback='C:\Logs\FleetStatus-Pending'
    )
    $json=$Record | ConvertTo-Json -Depth 12
    $fileName='{0}.json' -f $env:COMPUTERNAME
    $attempts = @()
    foreach($root in $StatusRoots) {
        try {
            if(Test-Path -LiteralPath $root -PathType Container) {
                $temp=Join-Path $root ('.{0}.{1}.tmp' -f $env:COMPUTERNAME,[guid]::NewGuid().Guid)
                $dest=Join-Path $root $fileName
                [IO.File]::WriteAllText($temp,$json,(New-Object Text.UTF8Encoding($false)))
                Move-Item -LiteralPath $temp -Destination $dest -Force
                $attempts += [pscustomobject]@{ Root=$root; Status='Published'; Error=$null }
                return [pscustomobject]@{ Published=$true; Path=$dest; Pending=$false; Attempts=@($attempts) }
            }
            $attempts += [pscustomobject]@{ Root=$root; Status='Unavailable'; Error='Path was not accessible or did not exist.' }
        }
        catch {
            $attempts += [pscustomobject]@{ Root=$root; Status='Failed'; Error=$_.Exception.Message }
            Write-Warning ("Unable to publish maintenance status to '{0}': {1}" -f $root, $_.Exception.Message)
        }
    }
    Initialize-MaintenanceDirectory -Path $LocalFallback | Out-Null
    $pending=Join-Path $LocalFallback ('{0}_{1}.json' -f $env:COMPUTERNAME,(Get-Date -Format 'yyyyMMdd_HHmmss'))
    [IO.File]::WriteAllText($pending,$json,(New-Object Text.UTF8Encoding($false)))
    [pscustomobject]@{ Published=$false; Path=$pending; Pending=$true; Attempts=@($attempts) }
}

Export-ModuleMember -Function @(
    'Get-MaintenanceFrameworkVersion','Get-MaintenanceExitCodes','Get-MaintenanceConfiguration','Initialize-MaintenanceDirectory',
    'Initialize-MaintenanceEnvironment','Test-MaintenanceAdministrator','New-MaintenanceRunContext',
    'Write-MaintenanceStatus','Start-MaintenanceScript','Stop-MaintenanceScript',
    'Archive-MaintenanceLogs','New-MaintenanceStagedLog','Publish-MaintenanceLog',
    'Resolve-MaintenanceFile','Write-MaintenanceTelemetryLine','Write-MaintenanceTelemetry',
    'Write-MaintenanceJsonAtomically','New-MaintenanceTelemetryRecord',
    'Rotate-MaintenanceLog','Get-MaintenanceCorrelationId','Enter-MaintenanceExecutionLock',
    'Exit-MaintenanceExecutionLock','Get-MaintenancePolicy','Test-MaintenanceWindow',
    'Test-MaintenanceDependencies','Initialize-MaintenanceEventLog','Write-MaintenanceEvent',
    'Publish-MaintenanceFleetStatus'
)
﻿#requires -version 5.1
<#
.SYNOPSIS
    Shared production framework for the Compton College maintenance scripts.
.DESCRIPTION
    Defines the canonical C:\Scripts installation root, C:\Logs output root,
    common status/logging helpers, standardized exit codes, administrative
    checks, run context creation, file resolution, and concurrency-safe NDJSON
    telemetry with size-based rotation and retention.
#>

Set-StrictMode -Version 2.0

$script:FrameworkVersion = '2.2.0'
$script:ProductionScriptRoot = 'C:\Scripts'
$script:DefaultLogRoot = 'C:\Logs'
$script:DefaultTelemetryMaxSizeMB = 50
$script:DefaultTelemetryRetentionDays = 60
$script:DefaultTelemetryMutexTimeoutSeconds = 30
$script:DefaultTelemetryMutexName = 'Global\Compton_MaintenanceTelemetry'

$script:ExitCodes = [ordered]@{
    Success             = 0
    SuccessWithWarnings = 1
    RecoverableFailure  = 2
    CriticalFailure     = 3
    TelemetryFailure    = 4
}

function Get-MaintenanceExitCodes {
    [CmdletBinding()]
    param()
    [pscustomobject]$script:ExitCodes
}

function Get-MaintenanceConfiguration {
    [CmdletBinding()]
    param(
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [string]$LogRoot = $script:DefaultLogRoot
    )

    if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { throw 'ScriptRoot cannot be empty.' }
    if ([string]::IsNullOrWhiteSpace($LogRoot)) { throw 'LogRoot cannot be empty.' }

    [pscustomobject]@{
        FrameworkVersion             = $script:FrameworkVersion
        ScriptRoot                   = [IO.Path]::GetFullPath($ScriptRoot)
        LogRoot                      = [IO.Path]::GetFullPath($LogRoot)
        TelemetryPath                = Join-Path $LogRoot 'Maintenance-Telemetry.ndjson'
        TelemetryMaxSizeMB           = $script:DefaultTelemetryMaxSizeMB
        TelemetryRetentionDays       = $script:DefaultTelemetryRetentionDays
        TelemetryMutexTimeoutSeconds = $script:DefaultTelemetryMutexTimeoutSeconds
        TelemetryMutexName           = $script:DefaultTelemetryMutexName
        ExitCodes                    = Get-MaintenanceExitCodes
    }
}

function Initialize-MaintenanceDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName
}

function Initialize-MaintenanceEnvironment {
    [CmdletBinding()]
    param(
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [string]$LogRoot = $script:DefaultLogRoot,
        [switch]$RequireScriptRoot
    )
    $configuration = Get-MaintenanceConfiguration -ScriptRoot $ScriptRoot -LogRoot $LogRoot
    if ($RequireScriptRoot -and -not (Test-Path -LiteralPath $configuration.ScriptRoot -PathType Container)) {
        throw "Required script root is missing: $($configuration.ScriptRoot)"
    }
    Initialize-MaintenanceDirectory -Path $configuration.LogRoot | Out-Null
    $configuration
}

function Test-MaintenanceAdministrator {
    [CmdletBinding()]
    param()
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { $false }
}

function New-MaintenanceRunContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptName,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptVersion,
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [string]$LogRoot = $script:DefaultLogRoot
    )
    $config = Initialize-MaintenanceEnvironment -ScriptRoot $ScriptRoot -LogRoot $LogRoot
    [pscustomobject]@{
        ScriptName    = $ScriptName
        ScriptVersion = $ScriptVersion
        ScriptRoot    = $config.ScriptRoot
        LogRoot       = $config.LogRoot
        RunId         = [guid]::NewGuid().Guid
        StartTime     = Get-Date
        WarningCount  = 0
        ErrorCount    = 0
        ExitCodes     = $config.ExitCodes
        Configuration = $config
    }
}

function Write-MaintenanceStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO',
        [string]$LogPath
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $label = switch ($Level) { 'INFO' {'INFO '} 'OK' {'OK   '} 'WARN' {'WARN '} 'ERROR' {'ERROR'} }
    $line = "[$timestamp] [$label] $Message"
    $color = switch ($Level) { 'INFO' {'Cyan'} 'OK' {'Green'} 'WARN' {'Yellow'} 'ERROR' {'Red'} }
    Write-Host $line -ForegroundColor $color
    if ($LogPath) {
        Initialize-MaintenanceDirectory -Path (Split-Path -Parent $LogPath) | Out-Null
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    $line
}

function Start-MaintenanceScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptName,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ScriptVersion,
        [switch]$RequireAdministrator
    )
    if ($RequireAdministrator -and -not (Test-MaintenanceAdministrator)) {
        throw "$ScriptName must run from an elevated administrator session or as SYSTEM."
    }
    $context = New-MaintenanceRunContext -ScriptName $ScriptName -ScriptVersion $ScriptVersion
    Write-MaintenanceStatus -Message "Starting $ScriptName version $ScriptVersion. Run ID: $($context.RunId)" -Level INFO | Out-Null
    $context
}

function Stop-MaintenanceScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [ValidateSet('Success','SuccessWithWarnings','RecoverableFailure','CriticalFailure','TelemetryFailure')]
        [string]$Status = 'Success'
    )
    $end = Get-Date
    $duration = [math]::Round(($end - $Context.StartTime).TotalSeconds,2)
    $exitCode = [int]$Context.ExitCodes.$Status
    Write-MaintenanceStatus -Message "Completed $($Context.ScriptName). Status=$Status; ExitCode=$exitCode; DurationSeconds=$duration" -Level $(if ($exitCode -eq 0) {'OK'} elseif ($exitCode -eq 1) {'WARN'} else {'ERROR'}) | Out-Null
    [pscustomobject]@{ Status=$Status; ExitCode=$exitCode; EndTime=$end; DurationSeconds=$duration }
}

function Resolve-MaintenanceFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Name,
        [string]$ScriptRoot = $script:ProductionScriptRoot,
        [switch]$Require
    )
    $path = Join-Path $ScriptRoot $Name
    if ($Require -and -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required maintenance file is missing: $path" }
    $path
}

function Write-MaintenanceTelemetryLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$JsonLine,
        [ValidateRange(1,4096)][int]$MaxSizeMB = $script:DefaultTelemetryMaxSizeMB,
        [ValidateRange(1,3650)][int]$RetentionDays = $script:DefaultTelemetryRetentionDays,
        [ValidateRange(1,300)][int]$MutexTimeoutSeconds = $script:DefaultTelemetryMutexTimeoutSeconds,
        [ValidateNotNullOrEmpty()][string]$MutexName = $script:DefaultTelemetryMutexName
    )
    $directory = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($directory)) { throw "Telemetry path must include a directory: $Path" }
    Initialize-MaintenanceDirectory -Path $directory | Out-Null

    $mutex = $null; $acquired = $false
    try {
        $createdNew = $false
        $mutex = New-Object System.Threading.Mutex($false,$MutexName,[ref]$createdNew)
        try { $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds($MutexTimeoutSeconds)) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw "Timed out after $MutexTimeoutSeconds seconds waiting for telemetry mutex '$MutexName'." }

        $encoding = New-Object System.Text.UTF8Encoding($false)
        $line = $JsonLine.TrimEnd("`r","`n") + [Environment]::NewLine
        $incoming = $encoding.GetByteCount($line)
        $maxBytes = [int64]$MaxSizeMB * 1MB

        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            $item = Get-Item -LiteralPath $Path -ErrorAction Stop
            if (($item.Length + $incoming) -gt $maxBytes -and $item.Length -gt 0) {
                $base = [IO.Path]::GetFileNameWithoutExtension($Path)
                $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
                $rotated = Join-Path $directory ("{0}_{1}.ndjson" -f $base,$stamp)
                $archive = [IO.Path]::ChangeExtension($rotated,'.zip')
                Move-Item -LiteralPath $Path -Destination $rotated -Force -ErrorAction Stop
                try {
                    Compress-Archive -LiteralPath $rotated -DestinationPath $archive -CompressionLevel Optimal -Force -ErrorAction Stop
                    Remove-Item -LiteralPath $rotated -Force -ErrorAction Stop
                }
                catch { Write-Warning "Telemetry rotated to '$rotated', but compression failed: $($_.Exception.Message)" }
            }
        }

        [IO.File]::AppendAllText($Path,$line,$encoding)
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        $pattern = ([IO.Path]::GetFileNameWithoutExtension($Path)) + '_*'
        Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like $pattern -and $_.Extension -in @('.zip','.ndjson') -and $_.LastWriteTime -lt $cutoff } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
    finally {
        if ($acquired -and $null -ne $mutex) { try { $mutex.ReleaseMutex() } catch {} }
        if ($null -ne $mutex) { $mutex.Dispose() }
    }
}

function Write-MaintenanceTelemetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Record,
        [string]$Path = (Join-Path $script:DefaultLogRoot 'Maintenance-Telemetry.ndjson'),
        [int]$Depth = 12
    )
    $json = $Record | ConvertTo-Json -Compress -Depth $Depth
    Write-MaintenanceTelemetryLine -Path $Path -JsonLine $json
}



function Rotate-MaintenanceLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateRange(1,1024)][int]$MaxSizeMB = 20,
        [ValidateRange(1,3650)][int]$RetentionDays = 60
    )
    $directory = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($directory)) { throw "Log path must include a directory: $Path" }
    Initialize-MaintenanceDirectory -Path $directory | Out-Null
    if ((Test-Path -LiteralPath $Path -PathType Leaf) -and ((Get-Item -LiteralPath $Path).Length -ge ([int64]$MaxSizeMB * 1MB))) {
        $base = [IO.Path]::GetFileNameWithoutExtension($Path)
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $rotated = Join-Path $directory ("{0}_{1}{2}" -f $base,$stamp,[IO.Path]::GetExtension($Path))
        $archive = [IO.Path]::ChangeExtension($rotated,'.zip')
        Move-Item -LiteralPath $Path -Destination $rotated -Force
        try {
            Compress-Archive -LiteralPath $rotated -DestinationPath $archive -CompressionLevel Optimal -Force
            Remove-Item -LiteralPath $rotated -Force
        } catch { Write-Warning "Log rotation compression failed for '$rotated': $($_.Exception.Message)" }
    }
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $basePattern = ([IO.Path]::GetFileNameWithoutExtension($Path)) + '_*'
    Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like $basePattern -and $_.LastWriteTime -lt $cutoff } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Get-MaintenanceCorrelationId {
    [CmdletBinding()]
    param([string]$CorrelationId)
    if (-not [string]::IsNullOrWhiteSpace($CorrelationId)) { return $CorrelationId }
    if (-not [string]::IsNullOrWhiteSpace($env:COMPTON_MAINTENANCE_CORRELATION_ID)) { return $env:COMPTON_MAINTENANCE_CORRELATION_ID }
    return [guid]::NewGuid().Guid
}

function Enter-MaintenanceExecutionLock {
    [CmdletBinding()]
    param(
        [string]$Name = 'Global\Compton_MaintenanceExecution',
        [ValidateRange(1,43200)][int]$TimeoutSeconds = 60
    )
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($false,$Name,[ref]$createdNew)
    try {
        try { $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds)) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { $mutex.Dispose(); throw "Timed out waiting for maintenance execution lock '$Name'." }
        [pscustomobject]@{ Name=$Name; Mutex=$mutex; Acquired=$true; AcquiredAt=Get-Date }
    } catch { if ($mutex) { $mutex.Dispose() }; throw }
}

function Exit-MaintenanceExecutionLock {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Lock)
    if ($Lock -and $Lock.Mutex) {
        try { if ($Lock.Acquired) { $Lock.Mutex.ReleaseMutex() } } catch {}
        try { $Lock.Mutex.Dispose() } catch {}
    }
}

function Get-MaintenancePolicy {
    [CmdletBinding()]
    param([string]$Path = 'C:\Scripts\Maintenance.Policy.json')
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Maintenance policy is missing: $Path" }
    Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}

function Test-MaintenanceWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        $Policy = (Get-MaintenancePolicy),
        [datetime]$At = (Get-Date)
    )
    $entry = @($Policy.Scripts | Where-Object { $_.Name -ieq $ScriptName }) | Select-Object -First 1
    if (-not $entry -or -not $entry.Window -or $entry.Window.Enforce -eq $false) {
        return [pscustomobject]@{ Allowed=$true; Reason='No enforced window'; ScriptName=$ScriptName }
    }
    $days = @($entry.Window.Days)
    if ($days.Count -gt 0 -and $At.DayOfWeek.ToString() -notin $days) {
        return [pscustomobject]@{ Allowed=$false; Reason="Day $($At.DayOfWeek) is outside the approved window"; ScriptName=$ScriptName }
    }
    $start = [TimeSpan]::Parse([string]$entry.Window.Start)
    $end = [TimeSpan]::Parse([string]$entry.Window.End)
    $time = $At.TimeOfDay
    $allowed = if ($end -ge $start) { $time -ge $start -and $time -le $end } else { $time -ge $start -or $time -le $end }
    [pscustomobject]@{ Allowed=$allowed; Reason=if($allowed){'Within approved window'}else{"Time $($At.ToString('HH:mm')) is outside $($entry.Window.Start)-$($entry.Window.End)"}; ScriptName=$ScriptName }
}

function Test-MaintenanceDependencies {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        $Policy = (Get-MaintenancePolicy)
    )
    $entry = @($Policy.Scripts | Where-Object { $_.Name -ieq $ScriptName }) | Select-Object -First 1
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($dependency in @($entry.Dependencies)) {
        $ok=$false; $detail=$null
        switch ([string]$dependency.Type) {
            'File'    { $ok=Test-Path -LiteralPath ([string]$dependency.Value) -PathType Leaf; $detail=[string]$dependency.Value }
            'Folder'  { $ok=Test-Path -LiteralPath ([string]$dependency.Value) -PathType Container; $detail=[string]$dependency.Value }
            'Command' { $ok=$null -ne (Get-Command ([string]$dependency.Value) -ErrorAction SilentlyContinue); $detail=[string]$dependency.Value }
            'Module'  { $ok=$null -ne (Get-Module -ListAvailable -Name ([string]$dependency.Value) | Select-Object -First 1); $detail=[string]$dependency.Value }
            'Service' { $svc=Get-Service -Name ([string]$dependency.Value) -ErrorAction SilentlyContinue; $ok=$null -ne $svc; $detail=if($svc){$svc.Status}else{'Missing'} }
            default   { $ok=$false; $detail="Unsupported dependency type: $($dependency.Type)" }
        }
        [void]$results.Add([pscustomobject]@{ Type=[string]$dependency.Type; Value=[string]$dependency.Value; Required=($dependency.Required -ne $false); Passed=$ok; Detail=$detail })
    }
    $failedRequired=@($results | Where-Object { $_.Required -and -not $_.Passed })
    [pscustomobject]@{ Passed=($failedRequired.Count -eq 0); Results=@($results); FailedRequired=$failedRequired }
}

function Initialize-MaintenanceEventLog {
    [CmdletBinding()]
    param([string]$LogName='Compton Maintenance',[string]$Source='Compton-Maintenance')
    if (-not [System.Diagnostics.EventLog]::SourceExists($Source)) {
        if (-not (Test-MaintenanceAdministrator)) { return $false }
        New-EventLog -LogName $LogName -Source $Source -ErrorAction Stop
    }
    $true
}

function Write-MaintenanceEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Information','Warning','Error')][string]$EntryType='Information',
        [ValidateRange(1,65535)][int]$EventId=1000,
        [string]$LogName='Compton Maintenance',[string]$Source='Compton-Maintenance'
    )
    try {
        if (Initialize-MaintenanceEventLog -LogName $LogName -Source $Source) {
            Write-EventLog -LogName $LogName -Source $Source -EventId $EventId -EntryType $EntryType -Message $Message
            return $true
        }
    } catch { Write-Warning "Unable to write maintenance event: $($_.Exception.Message)" }
    $false
}

function Publish-MaintenanceFleetStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Record,
        [string[]]$StatusRoots=@('\\filesvr\MaintenanceStatus','\\10.2.3.30\MaintenanceStatus'),
        [string]$LocalFallback='C:\Logs\FleetStatus-Pending'
    )
    $json=$Record | ConvertTo-Json -Depth 12
    $fileName='{0}.json' -f $env:COMPUTERNAME
    foreach($root in $StatusRoots) {
        try {
            if(Test-Path -LiteralPath $root -PathType Container) {
                $temp=Join-Path $root ('.{0}.{1}.tmp' -f $env:COMPUTERNAME,[guid]::NewGuid().Guid)
                $dest=Join-Path $root $fileName
                [IO.File]::WriteAllText($temp,$json,(New-Object Text.UTF8Encoding($false)))
                Move-Item -LiteralPath $temp -Destination $dest -Force
                return [pscustomobject]@{ Published=$true; Path=$dest; Pending=$false }
            }
        } catch {}
    }
    Initialize-MaintenanceDirectory -Path $LocalFallback | Out-Null
    $pending=Join-Path $LocalFallback ('{0}_{1}.json' -f $env:COMPUTERNAME,(Get-Date -Format 'yyyyMMdd_HHmmss'))
    [IO.File]::WriteAllText($pending,$json,(New-Object Text.UTF8Encoding($false)))
    [pscustomobject]@{ Published=$false; Path=$pending; Pending=$true }
}

Export-ModuleMember -Function @(
    'Get-MaintenanceExitCodes','Get-MaintenanceConfiguration','Initialize-MaintenanceDirectory',
    'Initialize-MaintenanceEnvironment','Test-MaintenanceAdministrator','New-MaintenanceRunContext',
    'Write-MaintenanceStatus','Start-MaintenanceScript','Stop-MaintenanceScript',
    'Resolve-MaintenanceFile','Write-MaintenanceTelemetryLine','Write-MaintenanceTelemetry',
    'Rotate-MaintenanceLog','Get-MaintenanceCorrelationId','Enter-MaintenanceExecutionLock',
    'Exit-MaintenanceExecutionLock','Get-MaintenancePolicy','Test-MaintenanceWindow',
    'Test-MaintenanceDependencies','Initialize-MaintenanceEventLog','Write-MaintenanceEvent',
    'Publish-MaintenanceFleetStatus'
)
