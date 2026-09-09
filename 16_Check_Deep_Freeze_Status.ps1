#requires -Version 5.1
<#
.SYNOPSIS
    Checks Faronics Deep Freeze state and writes one Elastic-friendly JSON event file per run.
.NOTES
    Script: 16_Check_Deep_Freeze_Status.ps1
    Version: 1.3.0
    Updated: 2026-09-09
    Changes: 1.3.0 tracks the first observation of the current state and publishes
             ThawedSinceUtc, ThawedForDays, ThawedForHours, PreviousState, and
             StateChanged fields for an Elastic/Kibana latest-state dashboard.
             Systems without DFC.exe publish Installed=false/NotInstalled so a
             latest-state dashboard can remove stale prior records, while its
             Installed=true filter excludes those computers from all totals.
             1.2.2 adds unified Maintenance-Telemetry.ndjson and latest-event telemetry while preserving the existing atomic Deep Freeze JSON events.
             Adds mapping-safe DFC output, explicit overall result/failure stage, and alert fields for Frozen/Thawed/Unknown state.
             1.2.1 adds 14-day cleanup of completed Deep Freeze JSON event files.
             Per-run JSON event files remain atomic and are not converted to staged text logging.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptName = '16_Check_Deep_Freeze_Status.ps1'
$ScriptVersion = '1.3.0'
$LogDirectory = 'C:\Logs'
$DeepFreezeLogDirectory = Join-Path $LogDirectory 'DeepFreeze'
$DeepFreezeStateDirectory = Join-Path $env:ProgramData 'Compton\DeepFreeze'
$DeepFreezeStatePath = Join-Path $DeepFreezeStateDirectory '16_Check_Deep_Freeze_Status.state.json'
$ComputerName = $env:COMPUTERNAME
$RunId = [guid]::NewGuid().ToString()
$DeepFreezeLogRetentionDays = 14
$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestTelemetryPath = Join-Path $LogDirectory '16_Check_Deep_Freeze_Status.latest.json'
$StartTime = Get-Date
$script:CurrentStage = 'Initialization'
$dfcPath = $null


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

function Update-DeepFreezeStateTracking {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Frozen','Thawed','Unknown','NotInstalled')]
        [string]$State,

        [Parameter(Mandatory)]
        [datetime]$ObservedAt
    )

    $previousState = $null
    $previousStateSince = $null

    if (Test-Path -LiteralPath $DeepFreezeStatePath -PathType Leaf) {
        try {
            $prior = Get-Content -LiteralPath $DeepFreezeStatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $previousState = [string]$prior.State

            $parsedStateSince = [datetimeoffset]::MinValue
            if (-not [string]::IsNullOrWhiteSpace([string]$prior.StateSinceUtc) -and
                [datetimeoffset]::TryParse([string]$prior.StateSinceUtc, [ref]$parsedStateSince)) {
                $previousStateSince = $parsedStateSince.UtcDateTime
            }
        }
        catch {
            # A damaged tracking file starts a new observation period. The status
            # audit itself must continue and publish the current DFC result.
            $previousState = $null
            $previousStateSince = $null
        }
    }

    $stateChanged = [string]::IsNullOrWhiteSpace($previousState) -or $previousState -ne $State
    $stateSince = $ObservedAt.ToUniversalTime()

    if (-not $stateChanged -and $null -ne $previousStateSince -and $previousStateSince -le $ObservedAt.ToUniversalTime()) {
        $stateSince = $previousStateSince
    }

    $thawedSinceUtc = $null
    $thawedForDays = $null
    $thawedForHours = $null

    if ($State -eq 'Thawed') {
        $thawedSinceUtc = $stateSince.ToString('o')
        $elapsed = $ObservedAt.ToUniversalTime() - $stateSince
        $thawedForDays = [int][math]::Floor([math]::Max(0, $elapsed.TotalDays))
        $thawedForHours = [math]::Round([math]::Max(0, $elapsed.TotalHours), 1)
    }

    if (-not (Test-Path -LiteralPath $DeepFreezeStateDirectory -PathType Container)) {
        New-Item -Path $DeepFreezeStateDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $stateRecord = [ordered]@{
        SchemaVersion  = '1.0'
        ComputerName   = $ComputerName
        State          = $State
        StateSinceUtc  = $stateSince.ToString('o')
        LastObservedUtc = $ObservedAt.ToUniversalTime().ToString('o')
    }
    Write-JsonAtomically -Path $DeepFreezeStatePath -Json ($stateRecord | ConvertTo-Json -Depth 4)

    return [pscustomobject]@{
        PreviousState   = $previousState
        StateChanged    = $stateChanged
        StateSinceUtc   = $stateSince.ToString('o')
        ThawedSinceUtc  = $thawedSinceUtc
        ThawedForDays   = $thawedForDays
        ThawedForHours  = $thawedForHours
    }
}

function Remove-ExpiredDeepFreezeLogs {
    [CmdletBinding()]
    param()

    if (-not (Test-Path -LiteralPath $DeepFreezeLogDirectory -PathType Container)) {
        return
    }

    $cutoff = (Get-Date).AddDays(-$DeepFreezeLogRetentionDays)

    Get-ChildItem -LiteralPath $DeepFreezeLogDirectory -File -Filter '16_DeepFreeze_*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Write-DeepFreezeLog {
    param(
        [ValidateSet('INFO','WARN','ERROR')][string]$Level,
        [ValidateSet('Frozen','Thawed','Unknown','NotInstalled')][string]$State,
        [string]$Message,
        [bool]$Installed = $true,
        [Nullable[int]]$DfcExitCode = $null,
        [string]$DfcPath = $null,
        [AllowNull()]$DfcOutput = $null,
        [ValidateSet('Frozen','Thawed','NotInstalled','QueryFailed','UnhandledError')][string]$OverallResult = 'QueryFailed',
        [int]$ScriptExitCode = 0,
        [string]$FailureStage = $null
    )

    if (-not (Test-Path -LiteralPath $DeepFreezeLogDirectory -PathType Container)) {
        New-Item -Path $DeepFreezeLogDirectory -ItemType Directory -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $LogDirectory -PathType Container)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }

    $eventTime = Get-Date
    $isFrozen = if ($State -eq 'Frozen') { $true } elseif ($State -eq 'Thawed') { $false } else { $null }
    $alertRequired = ($State -eq 'Thawed' -or $State -eq 'Unknown')
    $stateTracking = Update-DeepFreezeStateTracking -State $State -ObservedAt $eventTime

    if ($State -eq 'Thawed') {
        $Message = '{0} First observed thawed at {1}; {2} complete day(s) and {3} hour(s) thawed as of this check.' -f `
            $Message,$stateTracking.ThawedSinceUtc,$stateTracking.ThawedForDays,$stateTracking.ThawedForHours
    }

    # Preserve the original simple per-run Deep Freeze JSON event.
    $record = [ordered]@{
        timestamp             = $eventTime.ToUniversalTime().ToString('o')
        computer_name         = $ComputerName
        script_name           = $ScriptName
        script_version        = $ScriptVersion
        run_id                = $RunId
        log_level             = $Level
        event_type            = 'deep_freeze_status'
        deep_freeze_state     = $State
        deep_freeze_is_frozen = $isFrozen
        deep_freeze_installed = $Installed
        deep_freeze_previous_state = $stateTracking.PreviousState
        deep_freeze_state_changed = $stateTracking.StateChanged
        deep_freeze_state_since_utc = $stateTracking.StateSinceUtc
        deep_freeze_thawed_since_utc = $stateTracking.ThawedSinceUtc
        deep_freeze_thawed_for_days = $stateTracking.ThawedForDays
        deep_freeze_thawed_for_hours = $stateTracking.ThawedForHours
        alert_required        = $alertRequired
        dfc_exit_code         = $DfcExitCode
        dfc_path              = $DfcPath
        message               = $Message
    }

    $timestampForFile = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $safeComputerName = ($ComputerName -replace '[^A-Za-z0-9_.-]', '_')
    $eventFileName = '16_DeepFreeze_{0}_{1}_{2}.json' -f $safeComputerName, $timestampForFile, $RunId
    $eventFilePath = Join-Path $DeepFreezeLogDirectory $eventFileName

    $json = $record | ConvertTo-Json -Compress
    Set-Content -LiteralPath $eventFilePath -Value $json -Encoding UTF8 -Force -ErrorAction Stop

    # Write one schema-stable event into the common maintenance telemetry stream.
    $status = if ($ScriptExitCode -ne 0) {
        'Failed'
    }
    elseif ($State -eq 'Thawed') {
        'SuccessWithWarnings'
    }
    else {
        'Success'
    }

    $domain = $env:USERDOMAIN
    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.Domain) { $domain = [string]$computerSystem.Domain }
    }
    catch { }

    $maintenanceEvent = [ordered]@{
        '@timestamp'    = $eventTime.ToUniversalTime().ToString('o')
        EventType       = 'maintenance.deep_freeze'
        SchemaVersion   = '1.0'
        ComputerName    = $ComputerName
        Domain          = $domain
        ScriptName      = $ScriptName
        ScriptVersion   = $ScriptVersion
        RunId           = $RunId
        Status          = $status
        OverallResult   = $OverallResult
        ExitCode        = $ScriptExitCode
        StartTime       = $StartTime.ToUniversalTime().ToString('o')
        EndTime         = $eventTime.ToUniversalTime().ToString('o')
        DurationSeconds = [math]::Round(($eventTime - $StartTime).TotalSeconds, 3)
        FailureStage    = $FailureStage
        FailureMessage  = if ($ScriptExitCode -ne 0) { $Message } else { $null }

        DeepFreeze = [ordered]@{
            Installed     = $Installed
            State         = $State
            IsFrozen      = $isFrozen
            PreviousState = $stateTracking.PreviousState
            StateChanged  = $stateTracking.StateChanged
            StateSinceUtc = $stateTracking.StateSinceUtc
            ThawedSinceUtc = $stateTracking.ThawedSinceUtc
            ThawedForDays = $stateTracking.ThawedForDays
            ThawedForHours = $stateTracking.ThawedForHours
            AlertRequired = $alertRequired
            DfcExitCode   = $DfcExitCode
            DfcPath       = $DfcPath
            DfcOutput     = New-StringArrayForJson -InputObject $DfcOutput
            Message       = $Message
        }

        EventFilePath = $eventFilePath
    }

    $compact = $maintenanceEvent | ConvertTo-Json -Depth 10 -Compress
    $pretty  = $maintenanceEvent | ConvertTo-Json -Depth 10

    Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compact
    Write-JsonAtomically -Path $LatestTelemetryPath -Json $pretty

    Write-Host ('{0} [{1}] [{2}] [{3}] {4}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$ComputerName,$Level,$State,$Message)
    Write-Host ('Elastic event file: {0}' -f $eventFilePath)
    Write-Host ('Maintenance telemetry: {0}' -f $TelemetryPath)
}

function Find-DfcExecutable {
    foreach ($candidate in @("$env:WINDIR\SysWOW64\DFC.exe","$env:WINDIR\System32\DFC.exe")) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    $cmd = Get-Command 'DFC.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

try {
    # Treat the presence of DFC.exe as the local Deep Freeze installation check.
    # An explicit NotInstalled event supersedes any older Frozen/Thawed event for
    # this hostname. The dashboard filters Installed=false computers out.
    $script:CurrentStage = 'DetectDeepFreeze'
    $dfcPath = Find-DfcExecutable

    $script:CurrentStage = 'LoadMaintenanceFramework'
    $MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
    Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop

    $requiredFrameworkVersion = [version]'2.4.0'
    $currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)
    if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
        throw "Script 16 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
    }

    if ([string]::IsNullOrWhiteSpace($dfcPath)) {
        Write-DeepFreezeLog -Level INFO -State 'NotInstalled' -Installed $false `
            -OverallResult 'NotInstalled' -ScriptExitCode 0 `
            -Message 'Deep Freeze is not installed; this computer is excluded from Deep Freeze compliance totals.'
        exit 0
    }

    # Deep Freeze is installed. Remove only this script's completed JSON event
    # files older than the configured retention period.
    Remove-ExpiredDeepFreezeLogs

    $script:CurrentStage = 'QueryDeepFreezeState'
    $dfcOutput = @(& $dfcPath get /ISFROZEN 2>&1 | ForEach-Object { [string]$_ })
    $dfcExitCode = $LASTEXITCODE

    switch ($dfcExitCode) {
        0 {
            Write-DeepFreezeLog -Level WARN -State 'Thawed' -DfcExitCode 0 -DfcPath $dfcPath -DfcOutput $dfcOutput `
                -OverallResult 'Thawed' -ScriptExitCode 0 `
                -Message 'ALERT: Deep Freeze is installed but this workstation is THAWED (unfrozen).'
            exit 0
        }
        1 {
            Write-DeepFreezeLog -Level INFO -State 'Frozen' -DfcExitCode 1 -DfcPath $dfcPath -DfcOutput $dfcOutput `
                -OverallResult 'Frozen' -ScriptExitCode 0 `
                -Message 'Deep Freeze is installed and the workstation is Frozen.'
            exit 0
        }
        default {
            $outputText = ($dfcOutput -join ' | ').Trim()
            if ([string]::IsNullOrWhiteSpace($outputText)) { $outputText = 'No DFC output returned.' }
            $scriptExitCode = if ($dfcExitCode -eq 0) { 1 } else { [int]$dfcExitCode }
            Write-DeepFreezeLog -Level ERROR -State 'Unknown' -DfcExitCode $dfcExitCode -DfcPath $dfcPath -DfcOutput $dfcOutput `
                -OverallResult 'QueryFailed' -ScriptExitCode $scriptExitCode -FailureStage $script:CurrentStage `
                -Message "Deep Freeze status query failed. DFC exit code=$dfcExitCode. Output=$outputText"
            exit $scriptExitCode
        }
    }
}
catch {
    $failureMessage = "Unhandled Deep Freeze status check error: $($_.Exception.Message)"
    try {
        if (Get-Command Write-MaintenanceTelemetryLine -ErrorAction SilentlyContinue) {
            Write-DeepFreezeLog -Level ERROR -State 'Unknown' -Installed (-not [string]::IsNullOrWhiteSpace($dfcPath)) `
                -DfcPath $dfcPath -DfcOutput @() `
                -OverallResult 'UnhandledError' -ScriptExitCode 1 -FailureStage $script:CurrentStage `
                -Message $failureMessage
        }
        else {
            Write-Host ('{0} [{1}] [ERROR] [Unknown] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$ComputerName,$failureMessage)
        }
    }
    catch {
        Write-Host ('{0} [{1}] [ERROR] [Unknown] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$ComputerName,$failureMessage)
    }
    exit 1
}
