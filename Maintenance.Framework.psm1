#requires -version 5.1
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
