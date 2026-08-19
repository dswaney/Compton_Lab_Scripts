# =====================================================================
# ScriptName: 08_System_Repair.ps1
# ScriptVersion: 4.5.5
# LastUpdated: 2026-08-18
# Changes: v4.5.5 moves Script 08 staging out of C:\Temp so TempCleanup cannot delete the active runtime log.
#          Uses C:\ProgramData\Compton\Maintenance-Logs\Staging for the active log, then publishes the completed immutable log to C:\Logs.
#          v4.5.4 converts the active text log to Maintenance.Framework staged logging so Elastic only sees the completed immutable log.
#          Archives legacy/current Script 08 text logs before each run and publishes the completed timestamped log at finalization.
#          v4.5.3 adds mapping-safe Elastic repair summaries, distinguishes pre-existing vs newly-created reboot requirements,
#          reports DISM detection/remediation state explicitly, and fixes PowerShell 5.1 Disk Cleanup job cancellation.
#          v4.5.2 adds 14-day retention cleanup for C:\Logs\DeepFreeze JSON event files.
# Previous: v4.3.0 adds Elastic-compatible execution and repair telemetry, before/after health snapshots, detailed reboot reasons, repair-result extraction, safe finalization on all exit paths, and corrects status messaging and exit handling.
# Previous: v4.2 fixes Windows PowerShell 5.1 generic-list Count warnings in RPC and Explorer diagnostics.
# Previous: v4.1 prevents service-stop hangs by placing the stop request under its own timeout; if a service remains running after the service-state timeout, the script force-terminates its backing process when requested.
# Previous: v4.0 changes safety defaults to prevent automatic repairs, disables HP driver remediation by default, and removes Windows.old/System.sav cleanup targets while investigating boot/SSD issues.
# Previous: v3.9 adds Tier 3 HP Image Assistant driver-only remediation that runs only when CBS/SFC driver corruption is detected after DISM/SFC repair.
# Previous: v3.8 adds SFC/CBS corruption extraction and a DISM+SFC repair workflow with grep-friendly CBS/SFC markers.
# Previous: v3.7 adds explicit RPC root-cause classification output and ties EventLogSummary RPC errors to the RPC diagnostics assessment.
# Previous: v3.6 adds RPC diagnostics to separate local client/service failures from remote endpoint/network RPC failures, with grep-friendly RPC_* markers.
# Previous: v3.5 adds Explorer hang/crash diagnostics, shell extension inventory, network path checks, Explorer cache cleanup, Search service restart, and grep-friendly Explorer instability markers.
# Previous: v3.4 adds grep-friendly SSD/NVMe health logging, flags wear greater than 50%, improves Repair-Volume scan parsing, and verifies the volume again after SpotFix before considering offline repair.
# Previous: v3.3 deletes C:\HP_Bios_Config during temp cleanup and attempts Repair-Volume -SpotFix first when disk corruption is detected to avoid scheduling a reboot when possible.
# Previous: v3.2 adds native SSD/NVMe health checks, SMART failure prediction reporting, reliability counter logging, and optional smartctl reporting when available.
# Previous: v3.1 removes the storage optimization block and related SSD-detection helper.
# Previous: v3.0 rebuilds SoftwareDistribution.bak* cleanup with resumable state tracking, capped parallel deletion, timeout-safe jobs, and orphan job cleanup for Task Scheduler windows.
# Previous: v2.9 adds capped parallel SoftwareDistribution.bak* deletion, a cleanup time budget, robocopy mirror cleanup per folder, and safer Task Scheduler overlap protection.
# Previous: v2.8 adds guarded SoftwareDistribution deletion with timeout-protected robocopy/rmdir fallbacks and immediate YAML log writes before cleanup steps.
# Previous: v2.7 adds Action1 Agent validation and silent MSI installation from \\filesvr\Labscripts when C:\Windows\Action1\action1_agent.exe is missing.
# Previous: v2.6 replaces SoftwareDistribution cleanup with robocopy empty-folder mirror cleanup for the active folder and SoftwareDistribution.bak* folders.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$AutoRepairOnDetection = $false,
    [switch]$AllowWmiRepair = $true,
    [switch]$AllowNetworkReset = $false,
    [switch]$AllowWindowsUpdateReset = $false,
    [switch]$AllowOfflineDiskRepair = $false,
    [switch]$AllowFirewallReset = $false,
    [switch]$AllowIconCacheRebuild = $false,
    [switch]$AllowExplorerHealthRepair = $true,
    [switch]$AllowSearchServiceRestart = $true,
    [switch]$AllowHpDriverRepairFromCbs = $false,
    [string]$HpImageAssistantSourcePath = '\\filesvr\Labscripts\HPImageAssistant',
    [string]$HpImageAssistantLocalPath = 'C:\ProgramData\SystemRepair\HPImageAssistant',
    [switch]$AllowCopilotRemoval = $false,
    [switch]$AggressiveCleanup = $false,
    [switch]$ClearEventLogs = $false,
    [switch]$AutoRebootIfNeeded = $false,
    [int]$AutoRebootDelaySeconds = 60,
    [int]$MaxParallelSoftwareDistributionDeletes = 2,
    [int]$SoftwareDistributionCleanupTimeLimitMinutes = 40,
    [int]$SoftwareDistributionFolderJobTimeoutMinutes = 15,
    [string]$SoftwareDistributionCleanupStatePath = 'C:\ProgramData\SystemRepair\SoftwareDistributionCleanupState.json',
    [string]$LogDirectory = 'C:\Logs'
)

$ErrorActionPreference = 'Stop'

$script:ScriptName = '08_System_Repair.ps1'
$script:ScriptVersion = '4.5.5'
$script:RunId = [guid]::NewGuid().Guid
$script:TelemetryNdjsonPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestTelemetryPath = Join-Path $LogDirectory '08_System_Repair.latest.json'
$script:LegacyTextLogPath = Join-Path $LogDirectory '08_System_Repair.log'
$script:DedicatedStagingRoot = 'C:\ProgramData\Compton\Maintenance-Logs\Staging'
$script:LogSession = $null
$script:TextLogPath = $null
$script:PublishedTextLogPath = $null
$script:LogPublished = $false
$script:TelemetryWritten = $false
$script:FinalExitCode = 0
$script:FinalStatus = 'Running'
$script:FailureMessage = $null
$script:PendingRebootBefore = $null
$script:PendingRebootAfter = $null
$script:DiskSpaceBefore = $null
$script:DiskSpaceAfter = $null
$script:RebootInitiated = $false

$script:RunStart = Get-Date
$script:ComputerName = $env:COMPUTERNAME
$script:TimestampForFile = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
$script:BaseFileName = "{0}-SystemRepair-{1}" -f $script:ComputerName, $script:TimestampForFile
$script:YamlLogPath = Join-Path $LogDirectory ($script:BaseFileName + '.yaml')

$script:Summary = [ordered]@{
    ComputerName                 = $script:ComputerName
    StartTime                    = $script:RunStart
    EndTime                      = $null
    StepsSucceeded               = 0
    StepsFailed                  = 0
    Warnings                     = 0
    RebootRequired               = $false
    PendingRebootDetected        = $false
    PendingRebootReasonsBefore   = New-Object System.Collections.Generic.List[string]
    PendingRebootReasonsAfter    = New-Object System.Collections.Generic.List[string]
    DiskCorruptionSuspected      = $false
    DismCorruptionDetected       = $false
    SfcIntegrityViolations       = $false
    SfcRepairAttempted           = $false
    SfcRepairSucceeded           = $false
    CbsCorruptFileCount          = 0
    CbsCorruptionExtractPath     = $null
    CbsDriverCorruptionDetected  = $false
    CbsDriverCorruptionRemaining = $false
    CbsCorruptDriverFiles        = New-Object System.Collections.Generic.List[string]
    HpDriverRepairAttempted      = $false
    HpDriverRepairSucceeded      = $false
    HpImageAssistantPath         = $null
    WmiRepositoryInconsistent    = $false
    StorageHealthWarnings        = 0
    StorageFailurePredicted      = $false
    StorageWearOver50Detected     = $false
    ExplorerHangCount             = 0
    ExplorerCrashCount            = 0
    ExplorerInstabilityDetected   = $false
    NonMicrosoftShellExtensionCount = 0
    BrokenNetworkPathCount        = 0
    RpcClientSideIssueDetected    = $false
    RpcRemoteSideIssueDetected    = $false
    RpcDiagnosticsStatus          = 'NotRun'
    RepairVolumeSpotFixSucceeded = $false
    RepairsAttempted             = New-Object System.Collections.Generic.List[string]
    Notes                        = New-Object System.Collections.Generic.List[string]
}

$script:DetailedResults = New-Object System.Collections.Generic.List[object]

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Ensure-LogDirectory {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }
}

function ConvertTo-YamlScalar {
    param(
        [Parameter(Mandatory)][AllowNull()]$Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }

    if ($Value -is [datetime]) {
        return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'"
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [string]$Value
    }

    $text = [string]$Value
    $text = $text -replace "'", "''"
    return "'" + $text + "'"
}

function Add-DetailedResult {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [AllowNull()]$Data = $null
    )

    $script:DetailedResults.Add([PSCustomObject]@{
        Timestamp = Get-Date
        Step      = $Step
        Status    = $Status
        Message   = $Message
        Data      = $Data
    }) | Out-Null
}


function Initialize-SystemRepairLogging {
    Ensure-LogDirectory

    # Archive any legacy fixed-name log and prior timestamped Script 08 logs.
    # Maintenance.Framework places archived files under C:\Logs\Old Logs.
    Archive-MaintenanceLogs `
        -ScriptName $script:ScriptName `
        -LogRoot $LogDirectory `
        -AdditionalPatterns @(
            '08_System_Repair.log',
            '*-08_System_Repair-*.log'
        ) | Out-Null

    if (-not (Test-Path -LiteralPath $script:DedicatedStagingRoot -PathType Container)) {
        New-Item -Path $script:DedicatedStagingRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $script:LogSession = New-MaintenanceStagedLog `
        -ScriptName $script:ScriptName `
        -LogRoot $LogDirectory `
        -StagingRoot $script:DedicatedStagingRoot `
        -ComputerName $script:ComputerName `
        -Timestamp $script:RunStart

    $script:TextLogPath = [string]$script:LogSession.WorkingPath
    $script:PublishedTextLogPath = [string]$script:LogSession.PublishedPath

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines(
        $script:TextLogPath,
        @(
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$($script:ComputerName)] [INFO] Script 08 staged runtime log initialized.",
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$($script:ComputerName)] [INFO] Run ID: $($script:RunId)",
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$($script:ComputerName)] [INFO] Completed log publish path: $($script:PublishedTextLogPath)"
        ),
        $utf8
    )
}

function Publish-SystemRepairLog {
    if ($script:LogPublished -or $null -eq $script:LogSession) {
        return
    }

    try {
        $publishResult = Publish-MaintenanceLog -LogSession $script:LogSession
        $script:LogPublished = [bool]$publishResult.Published

        if ($publishResult.Published) {
            Write-Host ("Published completed script 08 text log for Elastic: {0}" -f $script:PublishedTextLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 08 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
    catch {
        Write-Warning "Failed to publish completed Script 08 text log: $($_.Exception.Message)"
    }
}

function Write-YamlLog {
    try {
        Ensure-LogDirectory

        $lines = New-Object System.Collections.Generic.List[string]

        $lines.Add('run:') | Out-Null
        $lines.Add("  computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("  start_time: $(ConvertTo-YamlScalar $script:Summary.StartTime)") | Out-Null
        $lines.Add("  end_time: $(ConvertTo-YamlScalar $script:Summary.EndTime)") | Out-Null
        $lines.Add("  yaml_log_path: $(ConvertTo-YamlScalar $script:YamlLogPath)") | Out-Null
        $lines.Add("  text_log_path: $(ConvertTo-YamlScalar $script:TextLogPath)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('settings:') | Out-Null
        $lines.Add("  auto_repair_on_detection: $(ConvertTo-YamlScalar $AutoRepairOnDetection)") | Out-Null
        $lines.Add("  allow_wmi_repair: $(ConvertTo-YamlScalar $AllowWmiRepair)") | Out-Null
        $lines.Add("  allow_network_reset: $(ConvertTo-YamlScalar $AllowNetworkReset)") | Out-Null
        $lines.Add("  allow_windows_update_reset: $(ConvertTo-YamlScalar $AllowWindowsUpdateReset)") | Out-Null
        $lines.Add("  allow_offline_disk_repair: $(ConvertTo-YamlScalar $AllowOfflineDiskRepair)") | Out-Null
        $lines.Add("  allow_firewall_reset: $(ConvertTo-YamlScalar $AllowFirewallReset)") | Out-Null
        $lines.Add("  allow_icon_cache_rebuild: $(ConvertTo-YamlScalar $AllowIconCacheRebuild)") | Out-Null
        $lines.Add("  allow_explorer_health_repair: $(ConvertTo-YamlScalar $AllowExplorerHealthRepair)") | Out-Null
        $lines.Add("  allow_search_service_restart: $(ConvertTo-YamlScalar $AllowSearchServiceRestart)") | Out-Null
        $lines.Add("  allow_hp_driver_repair_from_cbs: $(ConvertTo-YamlScalar $AllowHpDriverRepairFromCbs)") | Out-Null
        $lines.Add("  hp_image_assistant_source_path: $(ConvertTo-YamlScalar $HpImageAssistantSourcePath)") | Out-Null
        $lines.Add("  hp_image_assistant_local_path: $(ConvertTo-YamlScalar $HpImageAssistantLocalPath)") | Out-Null
        $lines.Add("  allow_copilot_removal: $(ConvertTo-YamlScalar $AllowCopilotRemoval)") | Out-Null
        $lines.Add("  aggressive_cleanup: $(ConvertTo-YamlScalar $AggressiveCleanup)") | Out-Null
        $lines.Add("  clear_event_logs: $(ConvertTo-YamlScalar $ClearEventLogs)") | Out-Null
        $lines.Add("  auto_reboot_if_needed: $(ConvertTo-YamlScalar $AutoRebootIfNeeded)") | Out-Null
        $lines.Add("  auto_reboot_delay_seconds: $(ConvertTo-YamlScalar $AutoRebootDelaySeconds)") | Out-Null
        $lines.Add("  max_parallel_softwaredistribution_deletes: $(ConvertTo-YamlScalar $MaxParallelSoftwareDistributionDeletes)") | Out-Null
        $lines.Add("  softwaredistribution_cleanup_time_limit_minutes: $(ConvertTo-YamlScalar $SoftwareDistributionCleanupTimeLimitMinutes)") | Out-Null
        $lines.Add("  softwaredistribution_folder_job_timeout_minutes: $(ConvertTo-YamlScalar $SoftwareDistributionFolderJobTimeoutMinutes)") | Out-Null
        $lines.Add("  softwaredistribution_cleanup_state_path: $(ConvertTo-YamlScalar $SoftwareDistributionCleanupStatePath)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('summary:') | Out-Null
        $lines.Add("  steps_succeeded: $(ConvertTo-YamlScalar $script:Summary.StepsSucceeded)") | Out-Null
        $lines.Add("  steps_failed: $(ConvertTo-YamlScalar $script:Summary.StepsFailed)") | Out-Null
        $lines.Add("  warnings: $(ConvertTo-YamlScalar $script:Summary.Warnings)") | Out-Null
        $lines.Add("  reboot_required: $(ConvertTo-YamlScalar $script:Summary.RebootRequired)") | Out-Null
        $lines.Add("  pending_reboot_detected: $(ConvertTo-YamlScalar $script:Summary.PendingRebootDetected)") | Out-Null
        $lines.Add("  disk_corruption_suspected: $(ConvertTo-YamlScalar $script:Summary.DiskCorruptionSuspected)") | Out-Null
        $lines.Add("  dism_corruption_detected: $(ConvertTo-YamlScalar $script:Summary.DismCorruptionDetected)") | Out-Null
        $lines.Add("  sfc_integrity_violations: $(ConvertTo-YamlScalar $script:Summary.SfcIntegrityViolations)") | Out-Null
        $lines.Add("  sfc_repair_attempted: $(ConvertTo-YamlScalar $script:Summary.SfcRepairAttempted)") | Out-Null
        $lines.Add("  sfc_repair_succeeded: $(ConvertTo-YamlScalar $script:Summary.SfcRepairSucceeded)") | Out-Null
        $lines.Add("  cbs_corrupt_file_count: $(ConvertTo-YamlScalar $script:Summary.CbsCorruptFileCount)") | Out-Null
        $lines.Add("  cbs_corruption_extract_path: $(ConvertTo-YamlScalar $script:Summary.CbsCorruptionExtractPath)") | Out-Null
        $lines.Add("  cbs_driver_corruption_detected: $(ConvertTo-YamlScalar $script:Summary.CbsDriverCorruptionDetected)") | Out-Null
        $lines.Add("  cbs_driver_corruption_remaining: $(ConvertTo-YamlScalar $script:Summary.CbsDriverCorruptionRemaining)") | Out-Null
        $lines.Add("  cbs_corrupt_driver_files: $(ConvertTo-YamlScalar (($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ','))") | Out-Null
        $lines.Add("  hp_driver_repair_attempted: $(ConvertTo-YamlScalar $script:Summary.HpDriverRepairAttempted)") | Out-Null
        $lines.Add("  hp_driver_repair_succeeded: $(ConvertTo-YamlScalar $script:Summary.HpDriverRepairSucceeded)") | Out-Null
        $lines.Add("  hp_image_assistant_path: $(ConvertTo-YamlScalar $script:Summary.HpImageAssistantPath)") | Out-Null
        $lines.Add("  wmi_repository_inconsistent: $(ConvertTo-YamlScalar $script:Summary.WmiRepositoryInconsistent)") | Out-Null
        $lines.Add("  storage_health_warnings: $(ConvertTo-YamlScalar $script:Summary.StorageHealthWarnings)") | Out-Null
        $lines.Add("  storage_failure_predicted: $(ConvertTo-YamlScalar $script:Summary.StorageFailurePredicted)") | Out-Null
        $lines.Add("  storage_wear_over_50_detected: $(ConvertTo-YamlScalar $script:Summary.StorageWearOver50Detected)") | Out-Null
        $lines.Add("  explorer_hang_count: $(ConvertTo-YamlScalar $script:Summary.ExplorerHangCount)") | Out-Null
        $lines.Add("  explorer_crash_count: $(ConvertTo-YamlScalar $script:Summary.ExplorerCrashCount)") | Out-Null
        $lines.Add("  explorer_instability_detected: $(ConvertTo-YamlScalar $script:Summary.ExplorerInstabilityDetected)") | Out-Null
        $lines.Add("  non_microsoft_shell_extension_count: $(ConvertTo-YamlScalar $script:Summary.NonMicrosoftShellExtensionCount)") | Out-Null
        $lines.Add("  broken_network_path_count: $(ConvertTo-YamlScalar $script:Summary.BrokenNetworkPathCount)") | Out-Null
        $lines.Add("  rpc_client_side_issue_detected: $(ConvertTo-YamlScalar $script:Summary.RpcClientSideIssueDetected)") | Out-Null
        $lines.Add("  rpc_remote_side_issue_detected: $(ConvertTo-YamlScalar $script:Summary.RpcRemoteSideIssueDetected)") | Out-Null
        $lines.Add("  rpc_diagnostics_status: $(ConvertTo-YamlScalar $script:Summary.RpcDiagnosticsStatus)") | Out-Null
        $lines.Add("  repair_volume_spotfix_succeeded: $(ConvertTo-YamlScalar $script:Summary.RepairVolumeSpotFixSucceeded)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('repairs_attempted:') | Out-Null
        if ($script:Summary.RepairsAttempted.Count -gt 0) {
            foreach ($repair in $script:Summary.RepairsAttempted) {
                $lines.Add("  - $(ConvertTo-YamlScalar $repair)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }
        $lines.Add('') | Out-Null

        $lines.Add('notes:') | Out-Null
        if ($script:Summary.Notes.Count -gt 0) {
            foreach ($note in $script:Summary.Notes) {
                $lines.Add("  - $(ConvertTo-YamlScalar $note)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }
        $lines.Add('') | Out-Null

        $lines.Add('detailed_results:') | Out-Null
        if ($script:DetailedResults.Count -gt 0) {
            foreach ($entry in $script:DetailedResults) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    timestamp: $(ConvertTo-YamlScalar $entry.Timestamp)") | Out-Null
                $lines.Add("    step: $(ConvertTo-YamlScalar $entry.Step)") | Out-Null
                $lines.Add("    status: $(ConvertTo-YamlScalar $entry.Status)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $entry.Message)") | Out-Null

                if ($null -eq $entry.Data) {
                    $lines.Add('    data: null') | Out-Null
                }
                elseif ($entry.Data -is [System.Collections.IDictionary]) {
                    $lines.Add('    data:') | Out-Null
                    foreach ($key in $entry.Data.Keys) {
                        $lines.Add("      $key`: $(ConvertTo-YamlScalar $entry.Data[$key])") | Out-Null
                    }
                }
                else {
                    $lines.Add("    data: $(ConvertTo-YamlScalar $entry.Data)") | Out-Null
                }
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        Set-Content -Path $script:YamlLogPath -Value $lines -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','SUCCESS','WARNING')][string]$Level = 'INFO'
    )

    $normalizedLevel = switch ($Level) {
        'OK'   { 'SUCCESS' }
        'WARN' { 'WARNING' }
        default { $Level }
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$($script:ComputerName)] [$normalizedLevel] $Message"

    switch ($normalizedLevel) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
    }

    try {
        Ensure-LogDirectory
        Add-Content -LiteralPath $script:TextLogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write text log '$($script:TextLogPath)': $($_.Exception.Message)"
    }
}

function Add-Note {
    param([string]$Message)
    $script:Summary.Notes.Add($Message) | Out-Null
}

function Add-RepairAttempt {
    param([string]$Message)
    $script:Summary.RepairsAttempted.Add($Message) | Out-Null
}

function Complete-Step {
    param([string]$Name)
    $script:Summary.StepsSucceeded++
    Write-Log "$Name completed." 'OK'
    Add-DetailedResult -Step $Name -Status 'Succeeded' -Message "$Name completed successfully."
}

function Fail-Step {
    param(
        [string]$Name,
        [string]$Reason
    )
    $script:Summary.StepsFailed++
    Write-Log "$Name failed: $Reason" 'ERROR'
    Add-Note "$Name failed: $Reason"
    Add-DetailedResult -Step $Name -Status 'Failed' -Message $Reason
}

function Warn-Step {
    param(
        [string]$Name,
        [string]$Reason
    )
    $script:Summary.Warnings++
    Write-Log "$Name warning: $Reason" 'WARN'
    Add-Note "$Name warning: $Reason"
    Add-DetailedResult -Step $Name -Status 'Warning' -Message $Reason
}

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}


function Set-ClassicContextMenuForHive {
    param(
        [Parameter(Mandatory)][string]$RootKey
    )

    $basePath = Join-Path $RootKey 'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
    $inprocPath = Join-Path $basePath 'InprocServer32'

    if (-not (Test-Path -LiteralPath $basePath)) {
        New-Item -Path $basePath -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $inprocPath)) {
        New-Item -Path $inprocPath -Force | Out-Null
    }

    New-ItemProperty -Path $inprocPath -Name '(default)' -Value '' -PropertyType String -Force | Out-Null
    Write-Log "Classic context menu registry value set for hive: $RootKey" 'OK'
    Add-DetailedResult -Step 'ClassicContextMenuRegistry' -Status 'Info' -Message "Classic context menu registry value set." -Data @{
        RootKey = $RootKey
        RegistryPath = $inprocPath
    }
}

function Enable-ClassicContextMenuAllUsers {
    Write-Log 'Applying classic Windows 10-style context menu for all users...' 'INFO'

    Set-ClassicContextMenuForHive -RootKey 'Registry::HKEY_CURRENT_USER'

    $userSids = Get-ChildItem Registry::HKEY_USERS |
        Where-Object {
            $_.PSChildName -match '^S-1-5-21-' -and
            $_.PSChildName -notmatch '_Classes$'
        } |
        Select-Object -ExpandProperty PSChildName

    foreach ($sid in $userSids) {
        Set-ClassicContextMenuForHive -RootKey "Registry::HKEY_USERS\$sid"
    }

    $defaultHiveName = 'HKU\DefaultTemp'
    $defaultHivePsPath = 'Registry::HKEY_USERS\DefaultTemp'
    $defaultUserNtUserDat = 'C:\Users\Default\NTUSER.DAT'

    if (Test-Path -LiteralPath $defaultUserNtUserDat) {
        $hiveLoaded = $false

        try {
            $loadResult = & reg.exe load $defaultHiveName $defaultUserNtUserDat
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to load Default User hive: $($loadResult -join ' ')"
            }

            $hiveLoaded = $true
            Start-Sleep -Milliseconds 750

            Set-ClassicContextMenuForHive -RootKey $defaultHivePsPath

            Start-Sleep -Milliseconds 750
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 750
        }
        catch {
            Write-Log "Failed to update Default User profile: $($_.Exception.Message)" 'ERROR'
            Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Failed' -Message $_.Exception.Message
        }
        finally {
            if ($hiveLoaded) {
                $unloaded = $false

                foreach ($attempt in 1..5) {
                    $unloadResult = & reg.exe unload $defaultHiveName
                    if ($LASTEXITCODE -eq 0) {
                        $unloaded = $true
                        Write-Log 'Applied classic context menu to Default User profile.' 'OK'
                        Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Succeeded' -Message 'Applied classic context menu to Default User profile.'
                        break
                    }

                    Start-Sleep -Seconds 1
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }

                if (-not $unloaded) {
                    Write-Log 'Classic context menu was written to Default User profile, but unloading the hive failed. A reboot may be required before the hive is released.' 'WARN'
                    Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Warning' -Message 'Classic context menu was written to Default User profile, but unloading the hive failed. A reboot may be required before the hive is released.'
                }
            }
        }
    }
    else {
        Write-Log 'Default User NTUSER.DAT not found; future new users were not updated.' 'WARN'
        Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Warning' -Message 'Default User NTUSER.DAT not found; future new users were not updated.'
    }

    Write-Log 'Classic context menu registry changes applied. Users may need to sign out and back in.' 'INFO'
    Add-DetailedResult -Step 'ClassicContextMenuAllUsers' -Status 'Info' -Message 'Classic context menu registry changes applied for current, loaded, and default user profiles.'
}

function Invoke-Safely {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [switch]$WarnOnly
    )

    try {
        & $ScriptBlock
        Complete-Step -Name $Name
        return $true
    }
    catch {
        if ($WarnOnly) {
            Warn-Step -Name $Name -Reason $_.Exception.Message
        }
        else {
            Fail-Step -Name $Name -Reason $_.Exception.Message
        }
        return $false
    }
}

function Ensure-Action1Agent {
    [CmdletBinding()]
    param(
        [string]$AgentPath = 'C:\Windows\Action1\action1_agent.exe',
        [string]$InstallerPath = '\\filesvr\Labscripts\action1_agent(Compton_College).msi'
    )

    Write-Log 'Checking for Action1 Agent...' 'INFO'

    if (Test-Path -LiteralPath $AgentPath) {
        Write-Log 'Action1 Agent already exists. Skipping installation.' 'OK'
        Add-DetailedResult -Step 'Action1Agent' -Status 'Skipped' -Message 'Action1 Agent already installed.' -Data @{
            AgentPath = $AgentPath
        }
        return
    }

    Write-Log "Action1 Agent not found at $AgentPath. Preparing silent install..." 'WARN'

    if (-not (Test-Path -LiteralPath $InstallerPath)) {
        throw "Action1 Agent installer not found: $InstallerPath"
    }

    $arguments = "/i `"$InstallerPath`" /quiet /qn"
    Write-Log "Running Action1 Agent installer from $InstallerPath" 'INFO'

    $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
    $exitCode = [int]$process.ExitCode

    Add-DetailedResult -Step 'Action1AgentInstall' -Status 'Info' -Message 'Action1 Agent installer finished.' -Data @{
        InstallerPath = $InstallerPath
        AgentPath     = $AgentPath
        ExitCode      = $exitCode
    }

    if ($exitCode -eq 3010 -or $exitCode -eq 1641) {
        $script:Summary.RebootRequired = $true
        Write-Log "Action1 Agent installer returned exit code $exitCode. Reboot required." 'WARN'
    }
    elseif ($exitCode -ne 0) {
        throw "Action1 Agent MSI install failed with exit code $exitCode"
    }

    Start-Sleep -Seconds 3

    if (Test-Path -LiteralPath $AgentPath) {
        Write-Log 'Action1 Agent installed successfully.' 'OK'
        Add-DetailedResult -Step 'Action1Agent' -Status 'Succeeded' -Message 'Action1 Agent detected after installation.' -Data @{
            AgentPath = $AgentPath
        }
    }
    else {
        throw 'Action1 Agent installer completed, but C:\Windows\Action1\action1_agent.exe was not detected afterward.'
    }
}

function Get-PendingRebootState {
    $result = [ordered]@{
        CBServicing_RebootPending         = $false
        WindowsUpdate_RebootRequired      = $false
        SessionManager_PendingFileRename  = $false
        SessionManager_PendingFileRename2 = $false
        UpdateExeVolatile                 = $false
        PackagesPending                   = $false
        WUAU_RebootRequired_COM           = $false
        AnyPendingReboot                  = $false
    }

    $cbsRebootPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    $wuRebootRequired = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    $packagesPending  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    $sessionMgr       = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    $updateExe        = 'HKLM:\SOFTWARE\Microsoft\Updates'

    try { $result.CBServicing_RebootPending = Test-Path -LiteralPath $cbsRebootPending } catch {}
    try { $result.WindowsUpdate_RebootRequired = Test-Path -LiteralPath $wuRebootRequired } catch {}
    try { $result.PackagesPending = Test-Path -LiteralPath $packagesPending } catch {}

    try {
        $pendingRename = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations' -ErrorAction Stop).PendingFileRenameOperations
        if ($null -ne $pendingRename -and $pendingRename.Count -gt 0) {
            $result.SessionManager_PendingFileRename = $true
        }
    }
    catch {}

    try {
        $pendingRename2 = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations2' -ErrorAction Stop).PendingFileRenameOperations2
        if ($null -ne $pendingRename2 -and $pendingRename2.Count -gt 0) {
            $result.SessionManager_PendingFileRename2 = $true
        }
    }
    catch {}

    try {
        $uev = (Get-ItemProperty -Path $updateExe -Name 'UpdateExeVolatile' -ErrorAction Stop).UpdateExeVolatile
        if ($null -ne $uev -and [int]$uev -ne 0) {
            $result.UpdateExeVolatile = $true
        }
    }
    catch {}

    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        if ($sysInfo.RebootRequired) {
            $result.WUAU_RebootRequired_COM = $true
        }
    }
    catch {}

    if ($result.Values -contains $true) {
        $result.AnyPendingReboot = $true
    }

    Add-DetailedResult -Step 'PendingRebootCheckData' -Status 'Info' -Message 'Collected pending reboot state.' -Data $result
    [PSCustomObject]$result
}


function Invoke-DismCommand {
    param([string[]]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\dism.exe"
    $psi.Arguments = $Arguments -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
    }

    Add-DetailedResult -Step 'DISM' -Status 'Info' -Message ("Executed DISM: " + ($Arguments -join ' ')) -Data @{
        ExitCode = $proc.ExitCode
    }

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Invoke-SfcCommand {
    param([string]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\sfc.exe"
    $psi.Arguments = $Arguments
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
    }

    Add-DetailedResult -Step 'SFC' -Status 'Info' -Message ("Executed SFC: " + $Arguments) -Data @{
        ExitCode = $proc.ExitCode
    }

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}



function Get-StorageSmartFailurePrediction {
    [CmdletBinding()]
    param()

    try {
        Get-CimInstance -Namespace 'root\wmi' -ClassName 'MSStorageDriver_FailurePredictStatus' -ErrorAction Stop |
            Select-Object InstanceName, PredictFailure, Reason
    }
    catch {
        Write-Log "SMART failure prediction status is not available through root\wmi on this system: $($_.Exception.Message)" 'WARN'
        Add-DetailedResult -Step 'StorageSmartFailurePrediction' -Status 'Warning' -Message 'SMART failure prediction status is not available through root\wmi.' -Data @{
            Error = $_.Exception.Message
        }
        @()
    }
}

function Get-StorageReliabilitySnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$PhysicalDisk
    )

    try {
        $counter = $PhysicalDisk | Get-StorageReliabilityCounter -ErrorAction Stop
        if ($null -eq $counter) {
            return $null
        }

        [PSCustomObject]@{
            Temperature          = $counter.Temperature
            Wear                 = $counter.Wear
            PowerOnHours         = $counter.PowerOnHours
            ReadErrorsTotal      = $counter.ReadErrorsTotal
            WriteErrorsTotal     = $counter.WriteErrorsTotal
            MediaErrors          = $counter.MediaErrors
            LoadUnloadCycleCount = $counter.LoadUnloadCycleCount
        }
    }
    catch {
        Write-Log "Reliability counters are not available for $($PhysicalDisk.FriendlyName): $($_.Exception.Message)" 'WARN'
        [PSCustomObject]@{
            Temperature          = $null
            Wear                 = $null
            PowerOnHours         = $null
            ReadErrorsTotal      = $null
            WriteErrorsTotal     = $null
            MediaErrors          = $null
            LoadUnloadCycleCount = $null
            Error                = $_.Exception.Message
        }
    }
}

function Invoke-SmartCtlOptionalReport {
    [CmdletBinding()]
    param()

    $smartCtl = Get-Command -Name 'smartctl.exe' -ErrorAction SilentlyContinue
    if (-not $smartCtl) {
        Add-DetailedResult -Step 'SmartCtlOptionalReport' -Status 'Skipped' -Message 'smartctl.exe is not installed or not in PATH. Native Windows storage checks were used.'
        return
    }

    Write-Log "smartctl.exe detected at $($smartCtl.Source). Capturing optional SMART summary." 'INFO'

    try {
        $scanOutput = & $smartCtl.Source --scan 2>&1
        Add-DetailedResult -Step 'SmartCtlScan' -Status 'Info' -Message 'smartctl scan output captured.' -Data @{
            SmartCtlPath = $smartCtl.Source
            ScanOutput   = ($scanOutput -join ' | ')
        }

        foreach ($line in @($scanOutput)) {
            if ($line -match '^(?<device>\S+)\s') {
                $device = $Matches.device
                $healthOutput = & $smartCtl.Source -H $device 2>&1
                Add-DetailedResult -Step 'SmartCtlHealth' -Status 'Info' -Message "smartctl health output captured for $device." -Data @{
                    Device = $device
                    Output = ($healthOutput -join ' | ')
                }
            }
        }
    }
    catch {
        Write-Log "smartctl optional report failed: $($_.Exception.Message)" 'WARN'
        Add-DetailedResult -Step 'SmartCtlOptionalReport' -Status 'Warning' -Message 'smartctl optional report failed.' -Data @{
            Error = $_.Exception.Message
        }
    }
}

function Invoke-StorageHealthCheck {
    [CmdletBinding()]
    param()

    Write-Log 'Checking SSD/NVMe/M.2 health with native Windows storage APIs...' 'INFO'

    $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop)
    if ($physicalDisks.Count -eq 0) {
        throw 'No physical disks were returned by Get-PhysicalDisk.'
    }

    $diskSummaries = New-Object System.Collections.Generic.List[object]
    $warningMessages = New-Object System.Collections.Generic.List[string]

    foreach ($disk in $physicalDisks) {
        $reliability = Get-StorageReliabilitySnapshot -PhysicalDisk $disk
        $mediaType = [string]$disk.MediaType
        $busType = if ($disk.PSObject.Properties.Name -contains 'BusType') { [string]$disk.BusType } else { 'Unknown' }
        $isSolidState = ($mediaType -match 'SSD|SCM') -or ($busType -match 'NVMe')

        $summary = [ordered]@{
            FriendlyName          = [string]$disk.FriendlyName
            SerialNumber          = [string]$disk.SerialNumber
            MediaType             = $mediaType
            BusType               = $busType
            HealthStatus          = [string]$disk.HealthStatus
            OperationalStatus     = (@($disk.OperationalStatus) -join ', ')
            SizeGB                = [math]::Round(([double]$disk.Size / 1GB), 2)
            IsSolidStateOrNVMe    = [bool]$isSolidState
            Temperature           = $reliability.Temperature
            Wear                  = $reliability.Wear
            PowerOnHours          = $reliability.PowerOnHours
            ReadErrorsTotal       = $reliability.ReadErrorsTotal
            WriteErrorsTotal      = $reliability.WriteErrorsTotal
            MediaErrors           = $reliability.MediaErrors
            LoadUnloadCycleCount  = $reliability.LoadUnloadCycleCount
        }

        $diskSummaries.Add([PSCustomObject]$summary) | Out-Null

        Write-Log ("Drive: {0} | Media: {1} | Bus: {2} | Health: {3} | Operational: {4} | Temp: {5}C | Wear: {6}% | PowerOnHours: {7}" -f $summary.FriendlyName, $summary.MediaType, $summary.BusType, $summary.HealthStatus, $summary.OperationalStatus, $summary.Temperature, $summary.Wear, $summary.PowerOnHours) 'INFO'
        Write-Log ("STORAGE_HEALTH_DATA|Computer={0}|Drive={1}|Serial={2}|Media={3}|Bus={4}|Health={5}|Operational={6}|SizeGB={7}|TemperatureC={8}|WearPercent={9}|PowerOnHours={10}|MediaErrors={11}|ReadErrorsTotal={12}|WriteErrorsTotal={13}" -f $script:ComputerName, $summary.FriendlyName, $summary.SerialNumber, $summary.MediaType, $summary.BusType, $summary.HealthStatus, $summary.OperationalStatus, $summary.SizeGB, $summary.Temperature, $summary.Wear, $summary.PowerOnHours, $summary.MediaErrors, $summary.ReadErrorsTotal, $summary.WriteErrorsTotal) 'INFO'

        if ($summary.HealthStatus -and $summary.HealthStatus -notin @('Healthy','Unknown')) {
            $warningMessages.Add("$($summary.FriendlyName) reports HealthStatus $($summary.HealthStatus).") | Out-Null
        }

        if ($summary.OperationalStatus -and $summary.OperationalStatus -notmatch 'OK|Unknown') {
            $warningMessages.Add("$($summary.FriendlyName) reports OperationalStatus $($summary.OperationalStatus).") | Out-Null
        }

        if ($null -ne $summary.Temperature -and [int]$summary.Temperature -ge 70) {
            $warningMessages.Add("$($summary.FriendlyName) temperature is $($summary.Temperature)C.") | Out-Null
        }

        if ($null -ne $summary.Wear -and [int]$summary.Wear -gt 50) {
            $script:Summary.StorageWearOver50Detected = $true
            Write-Log ("STORAGE_WEAR_GT_50|Computer={0}|Drive={1}|Serial={2}|Media={3}|Bus={4}|WearPercent={5}|Health={6}|Operational={7}" -f $script:ComputerName, $summary.FriendlyName, $summary.SerialNumber, $summary.MediaType, $summary.BusType, $summary.Wear, $summary.HealthStatus, $summary.OperationalStatus) 'WARN'
            $warningMessages.Add("$($summary.FriendlyName) wear is $($summary.Wear) percent, which is greater than the 50 percent review threshold. Review vendor interpretation and plan replacement if appropriate.") | Out-Null
        }
        elseif ($null -ne $summary.Wear -and [int]$summary.Wear -ge 80) {
            $warningMessages.Add("$($summary.FriendlyName) wear is $($summary.Wear) percent. Review vendor interpretation and plan replacement if appropriate.") | Out-Null
        }

        if (($null -ne $summary.MediaErrors -and [int64]$summary.MediaErrors -gt 0) -or
            ($null -ne $summary.ReadErrorsTotal -and [int64]$summary.ReadErrorsTotal -gt 0) -or
            ($null -ne $summary.WriteErrorsTotal -and [int64]$summary.WriteErrorsTotal -gt 0)) {
            $warningMessages.Add("$($summary.FriendlyName) reports storage reliability errors. MediaErrors=$($summary.MediaErrors), ReadErrors=$($summary.ReadErrorsTotal), WriteErrors=$($summary.WriteErrorsTotal).") | Out-Null
        }
    }

    $smartPredictions = @(Get-StorageSmartFailurePrediction)
    foreach ($prediction in $smartPredictions) {
        Write-Log ("SMART PredictFailure: {0} | Reason: {1} | Instance: {2}" -f $prediction.PredictFailure, $prediction.Reason, $prediction.InstanceName) 'INFO'
        if ([bool]$prediction.PredictFailure) {
            $script:Summary.StorageFailurePredicted = $true
            $warningMessages.Add("SMART predicts a possible storage failure for instance $($prediction.InstanceName). Reason=$($prediction.Reason).") | Out-Null
        }
    }

    if ($warningMessages.Count -gt 0) {
        $script:Summary.StorageHealthWarnings += $warningMessages.Count
        $script:Summary.Warnings += $warningMessages.Count
        foreach ($warning in $warningMessages) {
            Write-Log $warning 'WARN'
            Add-Note "Storage health warning: $warning"
        }
    }
    else {
        Write-Log 'No SSD/NVMe health warnings detected by native Windows storage checks.' 'OK'
    }

    Add-DetailedResult -Step 'StorageHealthCheck' -Status 'Info' -Message 'SSD/NVMe/M.2 health check completed.' -Data @{
        DiskSummaryJson       = (($diskSummaries | ConvertTo-Json -Depth 5 -Compress) -replace "'", "''")
        SmartPredictionJson   = (($smartPredictions | ConvertTo-Json -Depth 5 -Compress) -replace "'", "''")
        WarningCount          = $warningMessages.Count
        Warnings              = ($warningMessages -join ' | ')
    }

    Invoke-SmartCtlOptionalReport
}

function Get-DiskSpaceInfo {
    param([string]$Path)

    try {
        $driveRoot = Split-Path -Path $Path -Qualifier
        if ([string]::IsNullOrWhiteSpace($driveRoot)) {
            $driveRoot = $env:SystemDrive + '\'
        }

        $drive = [System.IO.DriveInfo]::new($driveRoot)
        return @{
            FreeSpace  = [int64]$drive.AvailableFreeSpace
            TotalSize  = [int64]$drive.TotalSize
            UsedSpace  = [int64]($drive.TotalSize - $drive.AvailableFreeSpace)
        }
    }
    catch {
        return $null
    }
}

function Test-SafeCleanupPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    $normalized = $Path.TrimEnd('\')

    $blockedPaths = @(
        'C:\Windows\System32',
        'C:\Windows\SysWOW64',
        'C:\Program Files',
        'C:\Program Files (x86)',
        'C:\Windows\explorer.exe',
        'C:\Windows\System32\drivers'
    )

    foreach ($blocked in $blockedPaths) {
        if ($normalized -ieq $blocked -or $normalized -like ($blocked + '\*')) {
            return $false
        }
    }

    $allowedPatterns = @(
        'C:\Windows\Temp*',
        'C:\Windows.old*',
        'C:\Temp*',
        'C:\SWSetup*',
        'C:\Lab Update Scripts*',
        'C:\ProgramData\Win11UpgradeStage*',
        'C:\Windows\SoftwareDistribution.bak*',
        'C:\SoftwareDistribution.bak*',
        'C:\Windows\SoftwareDistribution\Download*',
        'C:\Windows\Prefetch*',
        'C:\Windows\Logs\CBS*',
        'C:\ProgramData\Microsoft\Windows\WER\ReportQueue*',
        "$env:TEMP*",
        "$env:LOCALAPPDATA\Temp*",
        "$env:LOCALAPPDATA\Microsoft\Windows\INetCache*",
        "$env:LOCALAPPDATA\Microsoft\Windows\WebCache*",
        "$env:LOCALAPPDATA\CrashDumps*",
        "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache*",
        "$env:LOCALAPPDATA\D3DSCache*",
        "$env:LOCALAPPDATA\NVIDIA\DXCache*",
        "$env:LOCALAPPDATA\NVIDIA\GLCache*"
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($pattern in $allowedPatterns) {
        if ($normalized -like $pattern) {
            return $true
        }
    }

    return $false
}

function Remove-FolderContents {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Description,
        [switch]$ContentsOnly
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{
            Success    = $true
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = 'Path does not exist'
        }
    }

    if (-not (Test-SafeCleanupPath -Path $Path)) {
        return @{
            Success    = $false
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = 'Path blocked for security'
        }
    }

    try {
        $items = if ($ContentsOnly) {
            @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)
        }
        else {
            @(Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue)
        }

        $itemCount = $items.Count
        [int64]$sizeBefore = 0

        if ($itemCount -gt 0) {
            $files = if ($ContentsOnly) {
                $items | Where-Object { -not $_.PSIsContainer }
            }
            else {
                $items
            }

            if ($files.Count -gt 0) {
                $sizeSum = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -ne $sizeSum) {
                    $sizeBefore = [int64]$sizeSum
                }
            }
        }

        if ($itemCount -eq 0) {
            return @{
                Success    = $true
                SpaceFreed = [int64]0
                ItemCount  = 0
                Message    = 'Folder is empty'
            }
        }

        if ($ContentsOnly) {
            Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                }
                catch {
                    Write-Log "Could not remove $($_.FullName): $($_.Exception.Message)" 'WARN'
                }
            }
        }
        else {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        }

        return @{
            Success    = $true
            SpaceFreed = $sizeBefore
            ItemCount  = $itemCount
            Message    = 'Successfully cleaned'
        }
    }
    catch {
        return @{
            Success    = $false
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = $_.Exception.Message
        }
    }
}




function Get-FolderSizeInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $info = [ordered]@{
        Path       = $Path
        Exists     = $false
        ItemCount  = 0
        FileCount  = 0
        FolderCount = 0
        SizeBytes  = [int64]0
        SizeMB     = [double]0
        SizeGB     = [double]0
        Message    = $null
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $info.Message = 'Path does not exist'
        return [PSCustomObject]$info
    }

    $info.Exists = $true

    try {
        $items = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue)
        $files = @($items | Where-Object { -not $_.PSIsContainer })
        $folders = @($items | Where-Object { $_.PSIsContainer })
        $sizeBytes = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if ($null -eq $sizeBytes) { $sizeBytes = 0 }

        $info.ItemCount = $items.Count
        $info.FileCount = $files.Count
        $info.FolderCount = $folders.Count
        $info.SizeBytes = [int64]$sizeBytes
        $info.SizeMB = [math]::Round(([double]$info.SizeBytes / 1MB), 2)
        $info.SizeGB = [math]::Round(([double]$info.SizeBytes / 1GB), 3)
        $info.Message = 'Size calculated successfully'
    }
    catch {
        $info.Message = $_.Exception.Message
        Write-Log "Could not calculate folder size for $Path`: $($info.Message)" 'WARN'
    }

    return [PSCustomObject]$info
}

function Stop-WindowsUpdateLockingProcesses {
    [CmdletBinding()]
    param()

    Write-Log 'Checking for Windows Update processes that may lock SoftwareDistribution...' 'INFO'

    $processNames = @(
        'MoUsoCoreWorker',
        'TiWorker',
        'TrustedInstaller',
        'UsoClient',
        'MusNotification',
        'MusNotificationUx',
        'SIHClient'
    )

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($name in $processNames) {
        $processes = @(Get-Process -Name $name -ErrorAction SilentlyContinue)

        if ($processes.Count -eq 0) {
            $results.Add([PSCustomObject]@{
                Name    = $name
                Action  = 'NotRunning'
                Success = $true
                Message = 'Process not running'
            }) | Out-Null
            continue
        }

        foreach ($proc in $processes) {
            try {
                Write-Log "Stopping possible Windows Update lock process: $($proc.ProcessName) PID $($proc.Id)" 'WARN'
                Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                Start-Sleep -Milliseconds 500

                $stillRunning = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
                if ($null -eq $stillRunning) {
                    Write-Log "Stopped process $($proc.ProcessName) PID $($proc.Id)." 'OK'
                    $results.Add([PSCustomObject]@{
                        Name    = $proc.ProcessName
                        ProcessId = $proc.Id
                        Action  = 'Stopped'
                        Success = $true
                        Message = 'Stopped successfully'
                    }) | Out-Null
                }
                else {
                    Write-Log "Process $($proc.ProcessName) PID $($proc.Id) is still running after stop attempt." 'WARN'
                    $results.Add([PSCustomObject]@{
                        Name    = $proc.ProcessName
                        ProcessId = $proc.Id
                        Action  = 'StopAttempted'
                        Success = $false
                        Message = 'Still running after Stop-Process'
                    }) | Out-Null
                }
            }
            catch {
                Write-Log "Could not stop $($proc.ProcessName) PID $($proc.Id): $($_.Exception.Message)" 'WARN'
                $results.Add([PSCustomObject]@{
                    Name    = $proc.ProcessName
                    ProcessId = $proc.Id
                    Action  = 'FailedToStop'
                    Success = $false
                    Message = $_.Exception.Message
                }) | Out-Null
            }
        }
    }

    return @($results)
}

function Get-FolderSizeBytesSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
            return [int64]0
        }

        $sum = (Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer } |
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum

        if ($null -eq $sum) { return [int64]0 }
        return [int64]$sum
    }
    catch {
        Write-Log "Unable to calculate size for ${Path}: $($_.Exception.Message)" 'WARN'
        return [int64]0
    }
}

function Stop-SoftwareDistributionBackupLockingServices {
    [CmdletBinding()]
    param()

    $services = @(
        'wuauserv',
        'bits',
        'cryptsvc',
        'dosvc',
        'UsoSvc',
        'WaaSMedicSvc',
        'TrustedInstaller',
        'msiserver'
    )

    foreach ($svcName in $services) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($null -eq $svc) {
                Write-Log "Service not found while releasing SoftwareDistribution locks: $svcName" 'INFO'
                continue
            }

            if ($svc.Status -ne 'Stopped') {
                Write-Log "Stopping service to release SoftwareDistribution locks: $svcName ($($svc.Status))" 'INFO'
                [void](Stop-ServiceWithValidation -Name $svcName -TimeoutSeconds 30 -KillProcessOnTimeout)
                Start-Sleep -Seconds 2
            }

            $svc.Refresh()
            Write-Log "Service state after stop attempt: $svcName = $($svc.Status)" 'INFO'
        }
        catch {
            Write-Log "Could not stop service ${svcName}: $($_.Exception.Message)" 'WARN'
        }
    }
}

function Stop-SoftwareDistributionBackupLockingProcesses {
    [CmdletBinding()]
    param()

    $processNames = @(
        'TiWorker',
        'TrustedInstaller',
        'MoUsoCoreWorker',
        'UsoClient',
        'wuauclt',
        'bitsadmin',
        'msiexec',
        'MusNotification',
        'MusNotificationUx',
        'SIHClient'
    )

    foreach ($procName in $processNames) {
        try {
            $procs = @(Get-Process -Name $procName -ErrorAction SilentlyContinue)
            foreach ($proc in $procs) {
                Write-Log "Stopping process to release SoftwareDistribution locks: $($proc.ProcessName) PID $($proc.Id)" 'WARN'
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log "Could not stop process ${procName}: $($_.Exception.Message)" 'WARN'
        }
    }
}

function Start-SoftwareDistributionBackupUpdateServices {
    [CmdletBinding()]
    param()

    $services = @('cryptsvc', 'bits', 'wuauserv', 'dosvc', 'UsoSvc')

    foreach ($svcName in $services) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($null -eq $svc) { continue }

            if ($svc.Status -ne 'Running') {
                Write-Log "Restarting update-related service after SoftwareDistribution backup cleanup: $svcName" 'INFO'
                Start-Service -Name $svcName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
            }

            $svc.Refresh()
            Write-Log "Service state after restart attempt: $svcName = $($svc.Status)" 'INFO'
        }
        catch {
            Write-Log "Could not restart service ${svcName}: $($_.Exception.Message)" 'WARN'
        }
    }
}


function New-EmptyMirrorFolderForCleanup {
    [CmdletBinding()]
    param()

    $emptyPath = Join-Path $env:TEMP ('EmptyMirror_{0}' -f ([guid]::NewGuid().ToString('N')))
    New-Item -Path $emptyPath -ItemType Directory -Force | Out-Null
    return $emptyPath
}


function Invoke-NativeProcessWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$ArgumentList,
        [int]$TimeoutSeconds = 60,
        [string]$Description = 'Native process'
    )

    $outputPath = Join-Path $env:TEMP ('NativeProcessOutput_{0}.log' -f ([guid]::NewGuid().ToString('N')))
    $errorPath = Join-Path $env:TEMP ('NativeProcessError_{0}.log' -f ([guid]::NewGuid().ToString('N')))

    $result = [ordered]@{
        FilePath       = $FilePath
        Arguments      = ($ArgumentList -join ' ')
        TimedOut       = $false
        ExitCode       = $null
        Output         = @()
        ErrorOutput    = @()
        Message        = ''
    }

    try {
        Write-Log "$Description starting with timeout ${TimeoutSeconds}s: $FilePath $($ArgumentList -join ' ')" 'INFO'
        Add-DetailedResult -Step 'NativeProcessTimeoutGuard' -Status 'Started' -Message "$Description started." -Data @{
            FilePath = $FilePath
            Arguments = ($ArgumentList -join ' ')
            TimeoutSeconds = $TimeoutSeconds
        }
        Write-YamlLog

        $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -WindowStyle Hidden -RedirectStandardOutput $outputPath -RedirectStandardError $errorPath -PassThru -ErrorAction Stop

        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            $result.TimedOut = $true
            $result.Message = "$Description timed out after ${TimeoutSeconds}s and was terminated."
            Write-Log $result.Message 'WARN'

            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Log "Could not terminate timed-out process $($proc.Id): $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            $result.ExitCode = $proc.ExitCode
            $result.Message = "$Description completed with exit code $($proc.ExitCode)."
            Write-Log $result.Message 'INFO'
        }
    }
    catch {
        $result.Message = "$Description failed to start or monitor: $($_.Exception.Message)"
        Write-Log $result.Message 'WARN'
    }
    finally {
        try {
            if (Test-Path -LiteralPath $outputPath) {
                $result.Output = @(Get-Content -LiteralPath $outputPath -ErrorAction SilentlyContinue)
                Remove-Item -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }

        try {
            if (Test-Path -LiteralPath $errorPath) {
                $result.ErrorOutput = @(Get-Content -LiteralPath $errorPath -ErrorAction SilentlyContinue)
                Remove-Item -LiteralPath $errorPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }

        Add-DetailedResult -Step 'NativeProcessTimeoutGuard' -Status $(if ($result.TimedOut) { 'TimedOut' } elseif ($null -ne $result.ExitCode -and $result.ExitCode -le 7) { 'Completed' } else { 'Warning' }) -Message $result.Message -Data @{
            FilePath = $FilePath
            Arguments = ($ArgumentList -join ' ')
            ExitCode = $result.ExitCode
            TimedOut = $result.TimedOut
        }
        Write-YamlLog
    }

    return [PSCustomObject]$result
}

function Invoke-RobocopyMirrorDeleteFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Description,
        [Parameter(Mandatory = $true)][string]$EmptyMirrorPath,
        [switch]$RecreateAfterDelete,
        [int]$RobocopyTimeoutSeconds = 120,
        [int]$RemoveTimeoutSeconds = 60
    )

    $result = [ordered]@{
        Path             = $Path
        Description      = $Description
        ExistsBefore     = $false
        Deleted          = $false
        Recreated        = $false
        SizeBytesBefore  = [int64]0
        SizeMBBefore     = [double]0
        SizeGBBefore     = [double]0
        ItemCountBefore  = 0
        SpaceFreedBytes  = [int64]0
        SpaceFreedMB     = [double]0
        SpaceFreedGB     = [double]0
        RobocopyExitCode = $null
        RobocopyTimedOut = $false
        RmdirTimedOut    = $false
        Message          = ''
    }

    try {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            $result.Message = 'Path was blank.'
            return [PSCustomObject]$result
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            $result.Message = 'Path did not exist.'
            if ($RecreateAfterDelete) {
                New-Item -Path $Path -ItemType Directory -Force | Out-Null
                $result.Recreated = $true
                $result.Message = 'Path did not exist and was recreated.'
            }
            return [PSCustomObject]$result
        }

        $result.ExistsBefore = $true
        $sizeBytes = Get-FolderSizeBytesSafe -Path $Path
        $result.SizeBytesBefore = [int64]$sizeBytes
        $result.SizeMBBefore = [math]::Round(([double]$sizeBytes / 1MB), 2)
        $result.SizeGBBefore = [math]::Round(([double]$sizeBytes / 1GB), 3)
        $result.ItemCountBefore = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue).Count

        Write-Log "Preparing to clean $Description`: $Path | Size before cleanup: $($result.SizeGBBefore) GB ($($result.SizeMBBefore) MB) | Items: $($result.ItemCountBefore)" 'WARN'
        Add-DetailedResult -Step 'FolderCleanup' -Status 'Started' -Message "Preparing to clean $Description." -Data @{
            Path = $Path
            Description = $Description
            SizeMBBefore = $result.SizeMBBefore
            ItemCountBefore = $result.ItemCountBefore
        }
        Write-YamlLog

        try {
            $attribResult = Invoke-NativeProcessWithTimeout -FilePath 'cmd.exe' -ArgumentList @('/d','/c',"attrib -r -s -h `"$Path`" /s /d") -TimeoutSeconds 45 -Description "Attribute clear for $Description"
            if ($attribResult.TimedOut) {
                Write-Log "Attribute clear timed out for $Path. Continuing to deletion fallback steps." 'WARN'
            }
        }
        catch {
            Write-Log "Could not clear attributes on ${Path}: $($_.Exception.Message)" 'WARN'
        }

        if ($result.ItemCountBefore -eq 0) {
            Write-Log "Folder appears empty; attempting fast empty-directory delete first: $Path" 'INFO'
            try {
                [System.IO.Directory]::Delete($Path, $false)
            }
            catch {
                Write-Log "Fast empty-directory delete did not complete for ${Path}: $($_.Exception.Message). Continuing with guarded robocopy/rmdir cleanup." 'WARN'
            }
        }

        if (Test-Path -LiteralPath $Path) {
            $robocopyArgs = @(
                $EmptyMirrorPath,
                $Path,
                '/MIR',
                '/R:0',
                '/W:0',
                '/XJ',
                '/NP',
                '/NFL',
                '/NDL'
            )

            $robocopyResult = Invoke-NativeProcessWithTimeout -FilePath 'robocopy.exe' -ArgumentList $robocopyArgs -TimeoutSeconds $RobocopyTimeoutSeconds -Description "Robocopy mirror cleanup for $Description"
            $result.RobocopyExitCode = $robocopyResult.ExitCode
            $result.RobocopyTimedOut = [bool]$robocopyResult.TimedOut

            foreach ($line in @($robocopyResult.Output + $robocopyResult.ErrorOutput)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
                    Write-Log "robocopy: $line" 'INFO'
                }
            }

            if ($result.RobocopyTimedOut) {
                Write-Log "Robocopy timed out for $Path. Continuing to guarded rmdir fallback." 'WARN'
            }
            elseif ($null -ne $result.RobocopyExitCode -and $result.RobocopyExitCode -gt 7) {
                Write-Log "Robocopy returned failure exit code $($result.RobocopyExitCode) for $Path. Continuing to guarded rmdir fallback." 'WARN'
            }
        }

        if (Test-Path -LiteralPath $Path) {
            Write-Log "Trying guarded cmd.exe rmdir fallback for: $Path" 'WARN'
            $rmdirResult = Invoke-NativeProcessWithTimeout -FilePath 'cmd.exe' -ArgumentList @('/d','/c',"rmdir /s /q `"$Path`"") -TimeoutSeconds $RemoveTimeoutSeconds -Description "rmdir fallback for $Description"
            $result.RmdirTimedOut = [bool]$rmdirResult.TimedOut
        }

        if (Test-Path -LiteralPath $Path) {
            Write-Log "Trying final PowerShell Remove-Item fallback for: $Path" 'WARN'
            try {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            }
            catch {
                Write-Log "Final Remove-Item fallback reported an issue for ${Path}: $($_.Exception.Message)" 'WARN'
            }
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            $result.Deleted = $true
            $result.SpaceFreedBytes = [int64]$sizeBytes
            $result.SpaceFreedMB = [math]::Round(([double]$sizeBytes / 1MB), 2)
            $result.SpaceFreedGB = [math]::Round(([double]$sizeBytes / 1GB), 3)
            $result.Message = 'Deleted with guarded cleanup logic.'
            Write-Log "Deleted $Description`: $Path | Estimated freed: $($result.SpaceFreedGB) GB ($($result.SpaceFreedMB) MB)" 'OK'
        }
        else {
            $result.Message = 'Folder still exists after guarded robocopy, guarded rmdir, and Remove-Item fallback. Continuing script instead of hanging.'
            Write-Log "FAILED to delete $Description but continuing script: $Path" 'ERROR'
        }

        Add-DetailedResult -Step 'FolderCleanup' -Status $(if ($result.Deleted) { 'Success' } else { 'Warning' }) -Message $result.Message -Data @{
            Path = $Path
            Description = $Description
            Deleted = $result.Deleted
            RobocopyExitCode = $result.RobocopyExitCode
            RobocopyTimedOut = $result.RobocopyTimedOut
            RmdirTimedOut = $result.RmdirTimedOut
            SpaceFreedMB = $result.SpaceFreedMB
        }
        Write-YamlLog

        if ($RecreateAfterDelete) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
            $result.Recreated = $true
            Write-Log "Recreated required folder: $Path" 'OK'
            Write-YamlLog
        }
    }
    catch {
        $result.Message = $_.Exception.Message
        Write-Log "Guarded folder cleanup failed for ${Path}: $($_.Exception.Message). Continuing script." 'ERROR'
        Add-DetailedResult -Step 'FolderCleanup' -Status 'Error' -Message $result.Message -Data @{
            Path = $Path
            Description = $Description
        }
        Write-YamlLog
    }

    return [PSCustomObject]$result
}

function Remove-SoftwareDistributionMainFolder {
    [CmdletBinding()]
    param()

    Write-Log 'Cleaning active C:\Windows\SoftwareDistribution folder with robocopy mirror method.' 'INFO'

    $emptyPath = $null
    try {
        $emptyPath = New-EmptyMirrorFolderForCleanup
        return Invoke-RobocopyMirrorDeleteFolder -Path (Join-Path $env:WINDIR 'SoftwareDistribution') -Description 'Windows Update SoftwareDistribution folder' -EmptyMirrorPath $emptyPath -RecreateAfterDelete
    }
    finally {
        if ($emptyPath -and (Test-Path -LiteralPath $emptyPath)) {
            Remove-Item -LiteralPath $emptyPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Remove-SoftwareDistributionBakFolders {
    [CmdletBinding()]
    param(
        [int]$MaxParallelDeletes = $MaxParallelSoftwareDistributionDeletes,
        [int]$CleanupTimeLimitMinutes = $SoftwareDistributionCleanupTimeLimitMinutes,
        [int]$FolderJobTimeoutMinutes = $SoftwareDistributionFolderJobTimeoutMinutes,
        [string]$StatePath = $SoftwareDistributionCleanupStatePath
    )

    Write-Log 'Cleaning SoftwareDistribution.bak* folders with v3.0 resumable capped parallel cleanup.' 'INFO'

    if ($MaxParallelDeletes -lt 1) { $MaxParallelDeletes = 1 }
    if ($MaxParallelDeletes -gt 4) {
        Write-Log "MaxParallelDeletes was set to $MaxParallelDeletes. Capping at 4 to avoid disk thrashing." 'WARN'
        $MaxParallelDeletes = 4
    }
    if ($CleanupTimeLimitMinutes -lt 5) {
        Write-Log "CleanupTimeLimitMinutes was set to $CleanupTimeLimitMinutes. Raising to 5 minutes." 'WARN'
        $CleanupTimeLimitMinutes = 5
    }
    if ($FolderJobTimeoutMinutes -lt 3) {
        Write-Log "FolderJobTimeoutMinutes was set to $FolderJobTimeoutMinutes. Raising to 3 minutes." 'WARN'
        $FolderJobTimeoutMinutes = 3
    }

    $cleanupStart = Get-Date
    $deadline = $cleanupStart.AddMinutes($CleanupTimeLimitMinutes)
    $jobTimeoutSeconds = [int]($FolderJobTimeoutMinutes * 60)
    $robocopyTimeoutSeconds = [Math]::Max(120, [int]($jobTimeoutSeconds - 60))
    $rmdirTimeoutSeconds = 120

    $stateDirectory = Split-Path -Path $StatePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($stateDirectory) -and -not (Test-Path -LiteralPath $stateDirectory)) {
        New-Item -Path $stateDirectory -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    }

    function New-SoftwareDistributionCleanupState {
        return [PSCustomObject]@{
            Version = 3
            ComputerName = $env:COMPUTERNAME
            LastUpdated = (Get-Date).ToString('o')
            Pending = @()
            Completed = @()
            Failed = @()
            Deferred = @()
        }
    }

    function Read-SoftwareDistributionCleanupState {
        param([string]$Path)
        try {
            if (Test-Path -LiteralPath $Path) {
                $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $loaded = $raw | ConvertFrom-Json -ErrorAction Stop
                    if ($null -ne $loaded) { return $loaded }
                }
            }
        }
        catch {
            Write-Log "Could not read SoftwareDistribution cleanup state file. A new state file will be created. Error: $($_.Exception.Message)" 'WARN'
        }
        return New-SoftwareDistributionCleanupState
    }

    function Save-SoftwareDistributionCleanupState {
        param(
            [Parameter(Mandatory = $true)]$State,
            [Parameter(Mandatory = $true)][string]$Path
        )
        try {
            $State.LastUpdated = (Get-Date).ToString('o')
            $State | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8 -Force
        }
        catch {
            Write-Log "Could not write SoftwareDistribution cleanup state file: $($_.Exception.Message)" 'WARN'
        }
    }

    function Get-StateArray {
        param($Value)
        if ($null -eq $Value) { return @() }
        return @($Value)
    }

    function Add-StatePathUnique {
        param(
            [Parameter(Mandatory = $true)]$State,
            [Parameter(Mandatory = $true)][string]$PropertyName,
            [Parameter(Mandatory = $true)][string]$Path
        )
        $existing = @(Get-StateArray $State.$PropertyName | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($existing -notcontains $Path) { $existing += $Path }
        $State.$PropertyName = @($existing | Sort-Object -Unique)
    }

    function Remove-StatePath {
        param(
            [Parameter(Mandatory = $true)]$State,
            [Parameter(Mandatory = $true)][string]$PropertyName,
            [Parameter(Mandatory = $true)][string]$Path
        )
        $State.$PropertyName = @(Get-StateArray $State.$PropertyName | Where-Object { [string]$_ -ne $Path })
    }

    $state = Read-SoftwareDistributionCleanupState -Path $StatePath
    $basePaths = @('C:\Windows', 'C:\')
    $discoveredTargets = @()

    foreach ($basePath in $basePaths) {
        try {
            if (Test-Path -LiteralPath $basePath) {
                $discoveredTargets += @(
                    Get-ChildItem -LiteralPath $basePath -Directory -Force -Filter 'SoftwareDistribution.bak*' -ErrorAction Stop |
                        Where-Object { $_.Name -match '^SoftwareDistribution\.bak.*$' } |
                        Select-Object -ExpandProperty FullName
                )
            }
        }
        catch {
            Write-Log "Could not scan $basePath for SoftwareDistribution.bak* folders: $($_.Exception.Message)" 'WARN'
        }
    }

    $pendingFromState = @(Get-StateArray $state.Pending)
    $deferredFromState = @(Get-StateArray $state.Deferred)
    $candidateTargets = @($discoveredTargets + $pendingFromState + $deferredFromState |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        Sort-Object -Unique)

    $targets = @()
    foreach ($candidate in $candidateTargets) {
        if (Test-Path -LiteralPath $candidate) {
            $targets += $candidate
        }
        else {
            Remove-StatePath -State $state -PropertyName 'Pending' -Path $candidate
            Remove-StatePath -State $state -PropertyName 'Deferred' -Path $candidate
            Add-StatePathUnique -State $state -PropertyName 'Completed' -Path $candidate
        }
    }
    $targets = @($targets | Sort-Object -Unique)

    $state.Pending = @($targets)
    $state.Deferred = @()
    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath

    if ($targets.Count -eq 0) {
        Write-Log 'No SoftwareDistribution.bak* folders were found or all tracked folders were already removed.' 'INFO'
        Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status 'Success' -Message 'No SoftwareDistribution.bak* folders found.' -Data @{ StatePath = $StatePath }
        Write-YamlLog
        return [PSCustomObject]@{
            Success         = $true
            SpaceFreed      = [int64]0
            SpaceFreedBytes = [int64]0
            SpaceFreedMB    = [double]0
            SpaceFreedGB    = [double]0
            ItemCount       = 0
            DeletedCount    = 0
            FailedCount     = 0
            TimedOutCount   = 0
            DeferredCount   = 0
            Message         = 'No SoftwareDistribution.bak* folders found'
        }
    }

    Write-Log "Found/tracked $($targets.Count) SoftwareDistribution.bak* folder(s). Max parallel deletes: $MaxParallelDeletes. Cleanup budget: $CleanupTimeLimitMinutes minute(s). Per-folder job timeout: $FolderJobTimeoutMinutes minute(s)." 'WARN'
    Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status 'Started' -Message 'Starting v3.0 resumable capped parallel SoftwareDistribution.bak* cleanup.' -Data @{
        TargetCount = $targets.Count
        MaxParallelDeletes = $MaxParallelDeletes
        CleanupTimeLimitMinutes = $CleanupTimeLimitMinutes
        FolderJobTimeoutMinutes = $FolderJobTimeoutMinutes
        StatePath = $StatePath
        Deadline = $deadline
    }
    Write-YamlLog

    [int64]$totalFreedBytes = 0
    [int]$deletedCount = 0
    [int]$failedCount = 0
    [int]$timedOutCount = 0
    [int]$deferredCount = 0
    [int]$totalItems = 0
    $failureMessages = @()
    $script:SDv3TotalFreedBytes = [int64]0
    $script:SDv3DeletedCount = 0
    $script:SDv3FailedCount = 0
    $script:SDv3TimedOutCount = 0
    $script:SDv3DeferredCount = 0
    $script:SDv3TotalItems = 0
    $script:SDv3FailureMessages = @()
    $jobs = @()
    $jobPathById = @{}
    $jobStartedById = @{}
    $queue = New-Object System.Collections.Queue
    foreach ($target in $targets) { [void]$queue.Enqueue($target) }

    Stop-SoftwareDistributionBackupLockingServices
    Stop-SoftwareDistributionBackupLockingProcesses

    $jobScript = {
        param(
            [string]$TargetPath,
            [int]$RobocopyTimeoutSeconds,
            [int]$RmdirTimeoutSeconds
        )

        $result = [ordered]@{
            Path = $TargetPath
            Deleted = $false
            TimedOut = $false
            RobocopyTimedOut = $false
            RmdirTimedOut = $false
            RobocopyExitCode = $null
            SizeBytesBefore = [int64]0
            SizeMBBefore = [double]0
            SizeGBBefore = [double]0
            ItemCountBefore = 0
            Message = ''
            Started = (Get-Date).ToString('o')
            Finished = $null
        }

        function Invoke-NativeProcessWithTimeoutForSdCleanup {
            param(
                [Parameter(Mandatory = $true)][string]$FilePath,
                [Parameter(Mandatory = $true)][string[]]$ArgumentList,
                [Parameter(Mandatory = $true)][int]$TimeoutSeconds
            )

            $p = $null
            $outputFile = Join-Path $env:TEMP ('sdcleanup_out_' + [guid]::NewGuid().ToString('N') + '.txt')
            $errorFile = Join-Path $env:TEMP ('sdcleanup_err_' + [guid]::NewGuid().ToString('N') + '.txt')
            $nativeResult = [ordered]@{ ExitCode = $null; TimedOut = $false; ErrorMessage = '' }

            try {
                $p = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -NoNewWindow -PassThru -RedirectStandardOutput $outputFile -RedirectStandardError $errorFile
                if (-not $p.WaitForExit($TimeoutSeconds * 1000)) {
                    $nativeResult.TimedOut = $true
                    try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { }
                }
                else {
                    $nativeResult.ExitCode = $p.ExitCode
                }
            }
            catch {
                $nativeResult.ErrorMessage = $_.Exception.Message
            }
            finally {
                Remove-Item -LiteralPath $outputFile -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $errorFile -Force -ErrorAction SilentlyContinue
            }

            return [PSCustomObject]$nativeResult
        }

        try {
            if ([string]::IsNullOrWhiteSpace($TargetPath)) {
                $result.Message = 'Target path was blank.'
                return [PSCustomObject]$result
            }

            if (-not (Test-Path -LiteralPath $TargetPath)) {
                $result.Deleted = $true
                $result.Message = 'Target folder was already gone.'
                return [PSCustomObject]$result
            }

            try {
                $files = @(Get-ChildItem -LiteralPath $TargetPath -Force -File -ErrorAction SilentlyContinue)
                foreach ($file in $files) { $result.SizeBytesBefore += [int64]$file.Length; $result.ItemCountBefore++ }
                $result.SizeMBBefore = [math]::Round(([double]$result.SizeBytesBefore / 1MB), 2)
                $result.SizeGBBefore = [math]::Round(([double]$result.SizeBytesBefore / 1GB), 3)
            }
            catch { }

            $emptyMirror = Join-Path $env:TEMP ('EmptyDeleteSource_' + [guid]::NewGuid().ToString('N'))
            New-Item -Path $emptyMirror -ItemType Directory -Force | Out-Null

            try {
                $roboArgs = @($emptyMirror, $TargetPath, '/MIR', '/R:0', '/W:0', '/XJ', '/NP', '/NFL', '/NDL', '/NJH', '/NJS', '/NC', '/NS')
                $robo = Invoke-NativeProcessWithTimeoutForSdCleanup -FilePath 'robocopy.exe' -ArgumentList $roboArgs -TimeoutSeconds $RobocopyTimeoutSeconds
                $result.RobocopyExitCode = $robo.ExitCode
                $result.RobocopyTimedOut = [bool]$robo.TimedOut
                if ($robo.TimedOut) { $result.TimedOut = $true }

                if (Test-Path -LiteralPath $TargetPath) {
                    $rm = Invoke-NativeProcessWithTimeoutForSdCleanup -FilePath 'cmd.exe' -ArgumentList @('/d','/c', ('rmdir /s /q "{0}"' -f $TargetPath)) -TimeoutSeconds $RmdirTimeoutSeconds
                    $result.RmdirTimedOut = [bool]$rm.TimedOut
                    if ($rm.TimedOut) { $result.TimedOut = $true }
                }

                if (Test-Path -LiteralPath $TargetPath) {
                    try { [System.IO.Directory]::Delete($TargetPath, $true) } catch { }
                }
            }
            finally {
                Remove-Item -LiteralPath $emptyMirror -Recurse -Force -ErrorAction SilentlyContinue
            }

            if (-not (Test-Path -LiteralPath $TargetPath)) {
                $result.Deleted = $true
                $result.Message = 'Deleted by v3.0 resumable parallel cleanup job.'
            }
            elseif ($result.TimedOut) {
                $result.Message = 'Timed out before the folder could be fully deleted. It will be retried next run.'
            }
            else {
                $result.Message = 'Folder still exists after cleanup attempt. It will be retried next run.'
            }
        }
        catch {
            $result.Message = $_.Exception.Message
        }
        finally {
            $result.Finished = (Get-Date).ToString('o')
        }

        return [PSCustomObject]$result
    }

    function Receive-FinishedSdCleanupJobs {
        param([switch]$Force)

        $finishedJobs = @()
        foreach ($job in @($jobs)) {
            if ($Force -or $job.State -ne 'Running') { $finishedJobs += $job }
        }

        foreach ($job in $finishedJobs) {
            $targetPath = ''
            if ($jobPathById.ContainsKey($job.Id)) { $targetPath = [string]$jobPathById[$job.Id] }
            try {
                if ($job.State -eq 'Running' -and $Force) {
                    Stop-Job -Job $job -ErrorAction SilentlyContinue
                    $script:SDv3TimedOutCount++
                    $script:SDv3DeferredCount++
                    if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $targetPath
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                        Write-Log "Stopped SoftwareDistribution cleanup job due to cleanup window ending: $targetPath" 'WARN'
                    }
                    continue
                }

                $jobResults = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
                if ($jobResults.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($targetPath)) {
                    $script:SDv3FailedCount++
                    $script:SDv3FailureMessages += "No result returned for $targetPath"
                    Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                    Write-Log "SoftwareDistribution cleanup job returned no result: $targetPath" 'WARN'
                    continue
                }

                foreach ($jobResult in $jobResults) {
                    if ($null -eq $jobResult) { continue }
                    $path = [string]$jobResult.Path
                    if ([string]::IsNullOrWhiteSpace($path)) { $path = $targetPath }

                    if ($jobResult.Deleted) {
                        $script:SDv3DeletedCount++
                        $script:SDv3TotalFreedBytes += [int64]$jobResult.SizeBytesBefore
                        $script:SDv3TotalItems += [int]$jobResult.ItemCountBefore
                        Remove-StatePath -State $state -PropertyName 'Pending' -Path $path
                        Remove-StatePath -State $state -PropertyName 'Deferred' -Path $path
                        Add-StatePathUnique -State $state -PropertyName 'Completed' -Path $path
                        Write-Log "Deleted SoftwareDistribution backup folder: $path" 'OK'
                    }
                    else {
                        if ($jobResult.TimedOut -or $jobResult.RobocopyTimedOut -or $jobResult.RmdirTimedOut) { $script:SDv3TimedOutCount++ }
                        $script:SDv3FailedCount++
                        $script:SDv3FailureMessages += "Failed to delete ${path}: $($jobResult.Message)"
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $path
                        Write-Log "SoftwareDistribution backup folder remains and will be retried: $path | $($jobResult.Message)" 'WARN'
                    }

                    $statusText = 'Warning'
                    if ($jobResult.Deleted) { $statusText = 'Success' }
                    Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status $statusText -Message $jobResult.Message -Data @{
                        Path = $path
                        Deleted = $jobResult.Deleted
                        TimedOut = $jobResult.TimedOut
                        RobocopyExitCode = $jobResult.RobocopyExitCode
                        RobocopyTimedOut = $jobResult.RobocopyTimedOut
                        RmdirTimedOut = $jobResult.RmdirTimedOut
                        StatePath = $StatePath
                    }
                    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                    Write-YamlLog
                }
            }
            catch {
                $script:SDv3FailedCount++
                $script:SDv3FailureMessages += "Failed to receive delete job result for ${targetPath}: $($_.Exception.Message)"
                if (-not [string]::IsNullOrWhiteSpace($targetPath)) { Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath }
                Write-Log "Failed to receive SoftwareDistribution delete job result: $($_.Exception.Message)" 'WARN'
            }
            finally {
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                $jobPathById.Remove($job.Id)
                $jobStartedById.Remove($job.Id)
            }
        }

        $remainingJobs = @()
        foreach ($job in @($jobs)) {
            if ($finishedJobs -notcontains $job) { $remainingJobs += $job }
        }
        $script:__SdCleanupRemainingJobs = $remainingJobs
    }

    try {
        while ($true) {
            Receive-FinishedSdCleanupJobs
            $jobs = @($script:__SdCleanupRemainingJobs)

            $now = Get-Date
            foreach ($job in @($jobs)) {
                if ($jobStartedById.ContainsKey($job.Id)) {
                    $ageSeconds = ($now - [datetime]$jobStartedById[$job.Id]).TotalSeconds
                    if ($ageSeconds -gt $jobTimeoutSeconds) {
                        $targetPath = [string]$jobPathById[$job.Id]
                        Stop-Job -Job $job -ErrorAction SilentlyContinue
                        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                        $jobs = @($jobs | Where-Object { $_.Id -ne $job.Id })
                        $jobPathById.Remove($job.Id)
                        $jobStartedById.Remove($job.Id)
                        $script:SDv3TimedOutCount++
                        $script:SDv3DeferredCount++
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $targetPath
                        Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                        Write-Log "SoftwareDistribution delete job exceeded $FolderJobTimeoutMinutes minute timeout and was stopped: $targetPath" 'WARN'
                    }
                }
            }

            if ((Get-Date) -ge $deadline) {
                while ($queue.Count -gt 0) {
                    $deferredPath = [string]$queue.Dequeue()
                    if (Test-Path -LiteralPath $deferredPath) {
                        $script:SDv3DeferredCount++
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $deferredPath
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $deferredPath
                    }
                }

                foreach ($job in @($jobs)) {
                    $runningPath = ''
                    if ($jobPathById.ContainsKey($job.Id)) { $runningPath = [string]$jobPathById[$job.Id] }
                    if (-not [string]::IsNullOrWhiteSpace($runningPath)) {
                        $script:SDv3DeferredCount++
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $runningPath
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $runningPath
                    }
                    Stop-Job -Job $job -ErrorAction SilentlyContinue
                    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                }
                $jobs = @()
                Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                Write-Log "SoftwareDistribution cleanup time budget reached. Deferred folders will resume next run. Deferred count this run: $($script:SDv3DeferredCount)" 'WARN'
                break
            }

            $runningCount = @($jobs | Where-Object { $_.State -eq 'Running' }).Count
            while ($queue.Count -gt 0 -and $runningCount -lt $MaxParallelDeletes) {
                $target = [string]$queue.Dequeue()
                if (-not (Test-Path -LiteralPath $target)) {
                    Remove-StatePath -State $state -PropertyName 'Pending' -Path $target
                    Add-StatePathUnique -State $state -PropertyName 'Completed' -Path $target
                    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                    Write-Log "SoftwareDistribution.bak* folder no longer exists: $target" 'INFO'
                    continue
                }

                Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $target
                Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                Write-Log "Starting v3.0 timeout-safe delete job for: $target" 'INFO'
                Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status 'Queued' -Message 'Started timeout-safe resumable delete job.' -Data @{ Path = $target; StatePath = $StatePath }
                Write-YamlLog

                $jobName = 'SDCleanup_' + ([guid]::NewGuid().ToString('N'))
                $newJob = Start-Job -Name $jobName -ScriptBlock $jobScript -ArgumentList $target, $robocopyTimeoutSeconds, $rmdirTimeoutSeconds
                $jobs += $newJob
                $jobPathById[$newJob.Id] = $target
                $jobStartedById[$newJob.Id] = Get-Date
                $runningCount = @($jobs | Where-Object { $_.State -eq 'Running' }).Count
            }

            $runningCount = @($jobs | Where-Object { $_.State -eq 'Running' }).Count
            if ($queue.Count -eq 0 -and $runningCount -eq 0) { break }
            Start-Sleep -Seconds 2
        }

        Receive-FinishedSdCleanupJobs -Force
        $jobs = @($script:__SdCleanupRemainingJobs)
    }
    finally {
        foreach ($job in @($jobs)) {
            try {
                $targetPath = ''
                if ($jobPathById.ContainsKey($job.Id)) { $targetPath = [string]$jobPathById[$job.Id] }
                if ($job.State -eq 'Running') { Stop-Job -Job $job -ErrorAction SilentlyContinue }
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
                    Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                    Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $targetPath
                }
            }
            catch { }
        }
        Remove-Variable -Name __SdCleanupRemainingJobs -Scope Script -ErrorAction SilentlyContinue
        Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
        Start-SoftwareDistributionBackupUpdateServices
    }

    $remainingPending = @(Get-StateArray $state.Pending | Where-Object { Test-Path -LiteralPath ([string]$_) })
    $state.Pending = @($remainingPending | Sort-Object -Unique)
    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath

    $totalFreedBytes = [int64]$script:SDv3TotalFreedBytes
    $deletedCount = [int]$script:SDv3DeletedCount
    $failedCount = [int]$script:SDv3FailedCount
    $timedOutCount = [int]$script:SDv3TimedOutCount
    $deferredCount = [int]$script:SDv3DeferredCount
    $totalItems = [int]$script:SDv3TotalItems
    $failureMessages = @($script:SDv3FailureMessages)
    $freedMB = [math]::Round(([double]$totalFreedBytes / 1MB), 2)
    $freedGB = [math]::Round(([double]$totalFreedBytes / 1GB), 3)
    if ($state.Pending.Count -gt 0 -and $deferredCount -eq 0) { $deferredCount = $state.Pending.Count }
    $success = ($failedCount -eq 0 -and $state.Pending.Count -eq 0)

    if ($success) {
        Write-Log "SoftwareDistribution.bak* v3.0 parallel cleanup completed. Deleted folders: $deletedCount. Remaining pending: 0." 'OK'
    }
    else {
        Write-Log "SoftwareDistribution.bak* v3.0 cleanup completed with retryable items. Deleted: $deletedCount. Failed attempts: $failedCount. Timed out: $timedOutCount. Pending for next run: $($state.Pending.Count)." 'WARN'
    }

    $finalStatus = 'Warning'
    if ($success) { $finalStatus = 'Success' }
    Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status $finalStatus -Message 'SoftwareDistribution.bak* v3.0 resumable parallel cleanup finished.' -Data @{
        DeletedCount = $deletedCount
        FailedCount = $failedCount
        TimedOutCount = $timedOutCount
        DeferredCount = $deferredCount
        PendingCount = $state.Pending.Count
        SpaceFreedMB = $freedMB
        SpaceFreedGB = $freedGB
        StatePath = $StatePath
    }
    Write-YamlLog

    $message = 'SoftwareDistribution.bak* v3.0 parallel cleanup completed'
    if (-not $success) {
        $message = (($failureMessages + @("PendingCount=$($state.Pending.Count)", "TimedOutCount=$timedOutCount", "StatePath=$StatePath")) -join '; ')
    }

    return [PSCustomObject]@{
        Success         = $success
        SpaceFreed      = $totalFreedBytes
        SpaceFreedBytes = $totalFreedBytes
        SpaceFreedMB    = $freedMB
        SpaceFreedGB    = $freedGB
        ItemCount       = $totalItems
        DeletedCount    = $deletedCount
        FailedCount     = $failedCount
        TimedOutCount   = $timedOutCount
        DeferredCount   = $deferredCount
        PendingCount    = $state.Pending.Count
        StatePath       = $StatePath
        Message         = $message
    }
}

function Invoke-WindowsCleanup {
    param([int]$TimeoutSec = 300)

    try {
        $cleanmgrPath = Join-Path $env:SystemRoot 'System32\cleanmgr.exe'
        if (-not (Test-Path -LiteralPath $cleanmgrPath)) {
            throw 'Windows Disk Cleanup utility not found.'
        }

        Write-Log "Starting Windows Disk Cleanup (timeout: ${TimeoutSec}s)..." 'INFO'

        $job = Start-Job -ScriptBlock {
            Start-Process -FilePath "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList '/SAGERUN:1','/VERYLOWDISK' -NoNewWindow -Wait -PassThru
        }

        $result = Wait-Job -Job $job -Timeout $TimeoutSec

        if ($null -eq $result -or $job.State -eq 'Running') {
            Stop-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
            Remove-Job -Job $job -Force | Out-Null
            throw "Disk Cleanup timed out after $TimeoutSec seconds."
        }

        $proc = Receive-Job -Job $job
        Remove-Job -Job $job -Force | Out-Null

        return @{
            Success  = $true
            ExitCode = $proc.ExitCode
        }
    }
    catch {
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}


function Get-DirectorySizeBytes {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            return [int64]0
        }

        $sum = (
            Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum
        ).Sum

        if ($null -eq $sum) { return [int64]0 }
        return [int64]$sum
    }
    catch {
        Write-Log -Level WARNING -Message "Unable to calculate directory size for $Path : $($_.Exception.Message)"
        return [int64]0
    }
}

function Convert-BytesToReadableSize {
    [CmdletBinding()]
    param([Parameter(Mandatory)][int64]$Bytes)

    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N2} KB' -f ($Bytes / 1KB)) }
    return "$Bytes bytes"
}

function Get-LargestChildDirectories {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootPath,
        [ValidateRange(1,100)][int]$Top = 20
    )

    if (-not (Test-Path -LiteralPath $RootPath -PathType Container)) {
        return @()
    }

    $rows = foreach ($directory in @(
        Get-ChildItem -LiteralPath $RootPath -Directory -Force -ErrorAction SilentlyContinue
    )) {
        $sizeBytes = Get-DirectorySizeBytes -Path $directory.FullName

        [pscustomobject]@{
            RootPath      = $RootPath
            Directory     = $directory.FullName
            SizeBytes     = $sizeBytes
            SizeReadable  = Convert-BytesToReadableSize -Bytes $sizeBytes
            LastWriteTime = $directory.LastWriteTime
        }
    }

    return @(
        $rows |
        Sort-Object SizeBytes -Descending |
        Select-Object -First $Top
    )
}

function Invoke-DiskUsageInventory {
    [CmdletBinding()]
    param([ValidateRange(1,100)][int]$TopDirectories = 20)

    Write-Log -Message "Collecting pre-cleanup disk-usage inventory. Top $TopDirectories child directories per root will be recorded."

    $inventoryRoots = @(
        'C:\',
        'C:\Windows',
        'C:\ProgramData',
        'C:\Users'
    )

    foreach ($root in $inventoryRoots) {
        Write-Log -Message "Scanning disk usage under: $root"

        try {
            $largest = @(Get-LargestChildDirectories -RootPath $root -Top $TopDirectories)

            foreach ($item in $largest) {
                Write-Log -Message ("Disk usage: {0} | {1}" -f $item.SizeReadable, $item.Directory)
            }

            Add-DetailedResult `
                -Step ('DiskUsageInventory:' + $root) `
                -Status 'Info' `
                -Message "Collected top directory usage for $root" `
                -Data @{
                    RootPath = $root
                    TopCount = $TopDirectories
                    Results  = @($largest | ForEach-Object { $_ })
                }
        }
        catch {
            Write-Log -Level WARNING -Message "Disk-usage inventory failed for $root : $($_.Exception.Message)"
            Add-DetailedResult `
                -Step ('DiskUsageInventory:' + $root) `
                -Status 'Warning' `
                -Message "Disk-usage inventory failed for $root" `
                -Data @{
                    RootPath = $root
                    Error    = $_.Exception.Message
                }
        }
    }

    try {
        $systemDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop

        if ($systemDrive) {
            $size = [int64]$systemDrive.Size
            $free = [int64]$systemDrive.FreeSpace
            $used = [int64]($size - $free)
            $freePercent = if ($size -gt 0) { [math]::Round(($free / $size) * 100, 2) } else { 0 }

            Write-Log -Message ("C: drive before cleanup: Used={0}; Free={1}; FreePercent={2}%" -f `
                (Convert-BytesToReadableSize -Bytes $used),
                (Convert-BytesToReadableSize -Bytes $free),
                $freePercent)

            Add-DetailedResult `
                -Step 'DiskUsageInventory:DriveSummary' `
                -Status 'Info' `
                -Message 'Captured pre-cleanup system drive usage.' `
                -Data @{
                    DeviceID      = 'C:'
                    SizeBytes     = $size
                    UsedBytes     = $used
                    FreeBytes     = $free
                    SizeReadable  = Convert-BytesToReadableSize -Bytes $size
                    UsedReadable  = Convert-BytesToReadableSize -Bytes $used
                    FreeReadable  = Convert-BytesToReadableSize -Bytes $free
                    FreePercent   = $freePercent
                }
        }
    }
    catch {
        Write-Log -Level WARNING -Message "Unable to capture C: drive summary before cleanup: $($_.Exception.Message)"
    }

    Write-Log -Level SUCCESS -Message 'Pre-cleanup disk-usage inventory completed.'
}


function Remove-ExpiredDeepFreezeLogs {
    [CmdletBinding()]
    param(
        [string]$Path = 'C:\Logs\DeepFreeze',
        [ValidateRange(1,3650)]
        [int]$RetentionDays = 14
    )

    $result = [ordered]@{
        Path            = $Path
        RetentionDays   = $RetentionDays
        FilesExamined   = 0
        FilesDeleted    = 0
        FilesFailed     = 0
        SpaceFreedBytes = [int64]0
        Cutoff          = (Get-Date).AddDays(-$RetentionDays)
        Message         = ''
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        $result.Message = 'Deep Freeze log directory does not exist; nothing to clean.'
        Write-Log $result.Message 'INFO'
        Add-DetailedResult -Step 'DeepFreezeLogRetention' -Status 'Info' -Message $result.Message -Data @{
            Path          = $Path
            RetentionDays = $RetentionDays
            FilesDeleted  = 0
            SpaceFreedMB  = 0
        }
        return [pscustomobject]$result
    }

    $expiredFiles = @(
        Get-ChildItem `
            -LiteralPath $Path `
            -File `
            -Filter '*.json' `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LastWriteTime -lt $result.Cutoff
        }
    )

    $result.FilesExamined = $expiredFiles.Count

    if ($expiredFiles.Count -eq 0) {
        $result.Message = "No Deep Freeze JSON event files older than $RetentionDays days were found."
        Write-Log $result.Message 'INFO'
        Add-DetailedResult -Step 'DeepFreezeLogRetention' -Status 'Success' -Message $result.Message -Data @{
            Path          = $Path
            RetentionDays = $RetentionDays
            FilesDeleted  = 0
            SpaceFreedMB  = 0
        }
        return [pscustomobject]$result
    }

    Write-Log ("Deep Freeze retention cleanup found {0} JSON event file(s) older than {1} days in {2}." -f `
        $expiredFiles.Count, $RetentionDays, $Path) 'INFO'

    foreach ($file in $expiredFiles) {
        try {
            $length = [int64]$file.Length
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop

            $result.FilesDeleted++
            $result.SpaceFreedBytes += $length
        }
        catch {
            $result.FilesFailed++
            Write-Log ("Unable to delete expired Deep Freeze event file '{0}': {1}" -f `
                $file.FullName, $_.Exception.Message) 'WARN'
        }
    }

    $spaceFreedMB = [math]::Round(([double]$result.SpaceFreedBytes / 1MB), 3)

    $result.Message = (
        "Deep Freeze log retention completed. Deleted={0}; Failed={1}; RetentionDays={2}; SpaceFreedMB={3}." -f
        $result.FilesDeleted,
        $result.FilesFailed,
        $RetentionDays,
        $spaceFreedMB
    )

    Write-Log $result.Message $(if ($result.FilesFailed -gt 0) { 'WARN' } else { 'OK' })

    Add-DetailedResult `
        -Step 'DeepFreezeLogRetention' `
        -Status $(if ($result.FilesFailed -gt 0) { 'Warning' } else { 'Success' }) `
        -Message $result.Message `
        -Data @{
            Path            = $Path
            RetentionDays   = $RetentionDays
            Cutoff          = $result.Cutoff.ToString('o')
            FilesExamined   = $result.FilesExamined
            FilesDeleted    = $result.FilesDeleted
            FilesFailed     = $result.FilesFailed
            SpaceFreedBytes = $result.SpaceFreedBytes
            SpaceFreedMB    = $spaceFreedMB
        }

    return [pscustomobject]$result
}

function Invoke-TempCleanup {
    [int64]$totalSpaceFreed = 0
    $cleanupResults = New-Object System.Collections.Generic.List[object]
    $initialSpace = Get-DiskSpaceInfo -Path $env:SystemDrive

    Write-Log 'Cleaning temporary files and caches...' 'INFO'

    $deepFreezeCleanup = Remove-ExpiredDeepFreezeLogs -Path 'C:\Logs\DeepFreeze' -RetentionDays 14
    if ($deepFreezeCleanup -and $deepFreezeCleanup.SpaceFreedBytes -gt 0) {
        $totalSpaceFreed += [int64]$deepFreezeCleanup.SpaceFreedBytes
    }

    $cleanupResults.Add([PSCustomObject]@{
        Path        = 'C:\Logs\DeepFreeze'
        Description = 'Deep Freeze JSON Event Logs Older Than 14 Days'
        ItemCount   = if ($deepFreezeCleanup) { [int]$deepFreezeCleanup.FilesDeleted } else { 0 }
        SpaceFreed  = if ($deepFreezeCleanup) { [int64]$deepFreezeCleanup.SpaceFreedBytes } else { [int64]0 }
        Status      = if ($deepFreezeCleanup -and $deepFreezeCleanup.FilesFailed -gt 0) { 'Warning' } else { 'Success' }
        Message     = if ($deepFreezeCleanup) { $deepFreezeCleanup.Message } else { 'Deep Freeze retention cleanup returned no result.' }
    }) | Out-Null

    $windowsCleanup = Invoke-WindowsCleanup -TimeoutSec 300
    if ($windowsCleanup.Success) {
        Write-Log 'Windows Disk Cleanup completed successfully.' 'OK'
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'cleanmgr.exe'
            Description = 'Windows Disk Cleanup'
            ItemCount   = 0
            SpaceFreed  = [int64]0
            Status      = 'Success'
            Message     = "Exit code $($windowsCleanup.ExitCode)"
        }) | Out-Null
    }
    else {
        Warn-Step -Name 'TempCleanup' -Reason "Windows Disk Cleanup failed: $($windowsCleanup.Error)"
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'cleanmgr.exe'
            Description = 'Windows Disk Cleanup'
            ItemCount   = 0
            SpaceFreed  = [int64]0
            Status      = 'Warning'
            Message     = $windowsCleanup.Error
        }) | Out-Null
    }

    $cleanupTargets = New-Object System.Collections.Generic.List[object]

    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Lab Update Scripts'; Description = 'Lab Update Scripts'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\ProgramData\Win11UpgradeStage'; Description = 'Windows 11 Upgrade Staging'; ContentsOnly = $false })
    # Remove HP SoftPaq/HPIA extraction staging after maintenance.
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\SWSetup'; Description = 'HP Software Setup'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\HP_Bios_Config'; Description = 'HP BIOS Config Staging'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Temp'; Description = 'System Temp'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Temp'; Description = 'Windows Temp'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = $env:TEMP; Description = 'User Temp'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Temp"; Description = 'Local Temp'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\SoftwareDistribution\Download'; Description = 'Windows Update Cache'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Prefetch'; Description = 'Windows Prefetch'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Logs\CBS'; Description = 'CBS Logs'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"; Description = 'Internet Cache'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache"; Description = 'Web Cache'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\ProgramData\Microsoft\Windows\WER\ReportQueue'; Description = 'Error Report Queue'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\CrashDumps"; Description = 'Crash Dumps'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"; Description = 'Delivery Optimization Cache'; ContentsOnly = $true })

    if ($AggressiveCleanup) {
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\D3DSCache"; Description = 'Direct3D Shader Cache'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\NVIDIA\DXCache"; Description = 'NVIDIA DX Cache'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\NVIDIA\GLCache"; Description = 'NVIDIA GL Cache'; ContentsOnly = $true })
    }

    foreach ($target in $cleanupTargets) {
        if ($null -eq $target -or [string]::IsNullOrWhiteSpace($target.Path)) {
            Write-Log 'Skipping cleanup target because the path is blank.' 'WARN'
            continue
        }

        Write-Log "Cleaning $($target.Description) at $($target.Path)" 'INFO'
        $result = Remove-FolderContents -Path $target.Path -Description $target.Description -ContentsOnly:([bool]$target.ContentsOnly)

        if ($result.Success) {
            if ($result.SpaceFreed -gt 0) {
                $totalSpaceFreed += [int64]$result.SpaceFreed
                $sizeText = if ($result.SpaceFreed -ge 1GB) {
                    '{0} GB' -f [math]::Round($result.SpaceFreed / 1GB, 2)
                }
                else {
                    '{0} MB' -f [math]::Round($result.SpaceFreed / 1MB, 1)
                }
                Write-Log "Cleaned $($target.Description): $sizeText freed across $($result.ItemCount) item(s)." 'OK'
            }
            else {
                Write-Log "$($target.Description): $($result.Message)" 'INFO'
            }

            $cleanupResults.Add([PSCustomObject]@{
                Path        = $target.Path
                Description = $target.Description
                ItemCount   = $result.ItemCount
                SpaceFreed  = [int64]$result.SpaceFreed
                Status      = 'Success'
                Message     = $result.Message
            }) | Out-Null
        }
        else {
            Warn-Step -Name 'TempCleanup' -Reason "$($target.Description) failed: $($result.Message)"
            $cleanupResults.Add([PSCustomObject]@{
                Path        = $target.Path
                Description = $target.Description
                ItemCount   = 0
                SpaceFreed  = [int64]0
                Status      = 'Failed'
                Message     = $result.Message
            }) | Out-Null
        }
    }

    Write-Log 'Cleaning SoftwareDistribution backup folders at C:\Windows and C:\ root backup variants' 'INFO'
    $sdBackupResult = Remove-SoftwareDistributionBakFolders
    if ($sdBackupResult.Success) {
        if ($sdBackupResult.SpaceFreed -gt 0) {
            $totalSpaceFreed += [int64]$sdBackupResult.SpaceFreed
            Write-Log "Cleaned SoftwareDistribution Backup Folders: $($sdBackupResult.SpaceFreedGB) GB ($($sdBackupResult.SpaceFreedMB) MB) freed across $($sdBackupResult.ItemCount) item(s)." 'OK'
        }
        else {
            Write-Log "SoftwareDistribution Backup Folders: $($sdBackupResult.Message)" 'INFO'
        }

        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'C:\Windows and C:\ SoftwareDistribution backup variants'
            Description = 'SoftwareDistribution Backup Folders'
            ItemCount   = $sdBackupResult.ItemCount
            SpaceFreed  = [int64]$sdBackupResult.SpaceFreed
            Status      = 'Success'
            Message     = $sdBackupResult.Message
        }) | Out-Null
    }
    else {
        Warn-Step -Name 'TempCleanup' -Reason "SoftwareDistribution Backup Folders failed: $($sdBackupResult.Message)"
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'C:\Windows and C:\ SoftwareDistribution backup variants'
            Description = 'SoftwareDistribution Backup Folders'
            ItemCount   = $sdBackupResult.ItemCount
            SpaceFreed  = [int64]0
            Status      = 'Warning'
            Message     = $sdBackupResult.Message
        }) | Out-Null
    }

    $finalSpace = Get-DiskSpaceInfo -Path $env:SystemDrive
    $actualFreed = [int64]0
    if ($initialSpace -and $finalSpace) {
        $actualFreed = [int64]($finalSpace.FreeSpace - $initialSpace.FreeSpace)
    }

    Add-DetailedResult -Step 'TempCleanup' -Status 'Info' -Message 'Enhanced temporary file cleanup completed.' -Data @{
        EstimatedSpaceFreedMB = [math]::Round($totalSpaceFreed / 1MB, 2)
        ActualSpaceFreedMB    = [math]::Round($actualFreed / 1MB, 2)
        TargetsProcessed      = $cleanupResults.Count
        ResultsJson           = ($cleanupResults | ForEach-Object {
            [ordered]@{
                Path         = $_.Path
                Description  = $_.Description
                ItemCount    = $_.ItemCount
                SpaceFreedMB = [math]::Round(([double]$_.SpaceFreed) / 1MB, 2)
                Status       = $_.Status
                Message      = $_.Message
            }
        } | ConvertTo-Json -Compress)
    }
}

function Test-RepairVolumeScanIndicatesCorruption {
    [CmdletBinding()]
    param(
        [AllowNull()]$ScanResult
    )

    if ($null -eq $ScanResult) {
        return $false
    }

    $resultText = ($ScanResult | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($resultText)) {
        return $false
    }

    # Repair-Volume -Scan commonly returns status strings such as NoErrorsFound, ScanNeeded,
    # SpotFixNeeded, or FullRepairNeeded. Do not treat NoErrorsFound as a match simply
    # because it contains the word "Errors".
    if ($resultText -match '(?i)NoErrorsFound|No errors found|No problems found|No corruption detected') {
        return $false
    }

    if ($resultText -match '(?i)SpotFixNeeded|FullRepairNeeded|ScanNeeded|corrupt|corruption|repair needed|needs repair|errors found|file system errors') {
        return $true
    }

    return $false
}

function Invoke-RepairVolumeScan {
    [CmdletBinding()]
    param(
        [string]$StepName = 'RepairVolumeScan',
        [switch]$UpdateSummary = $true
    )

    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Running Repair-Volume scan on $($env:SystemDrive)" 'INFO'
    $result = Repair-Volume -DriveLetter $systemDrive -Scan -ErrorAction Stop

    $resultText = $null
    $corruptionDetected = $false
    if ($null -ne $result) {
        $resultText = ($result | Out-String).Trim()
        if (-not [string]::IsNullOrWhiteSpace($resultText)) {
            Write-Log "Repair-Volume scan result for $($env:SystemDrive): $resultText" 'INFO'
        }
        $corruptionDetected = Test-RepairVolumeScanIndicatesCorruption -ScanResult $result
    }

    if ($corruptionDetected) {
        if ($UpdateSummary) {
            $script:Summary.DiskCorruptionSuspected = $true
        }
        Warn-Step -Name $StepName -Reason 'Disk scan output indicates scan, spot fix, full repair, errors, or corruption may still exist.'
    }
    else {
        if ($UpdateSummary) {
            $script:Summary.DiskCorruptionSuspected = $false
        }
        Write-Log "Repair-Volume scan did not report actionable file system corruption on $($env:SystemDrive)." 'OK'
    }

    Add-DetailedResult -Step $StepName -Status 'Info' -Message 'Repair-Volume scan completed.' -Data @{
        Output             = $resultText
        CorruptionDetected = $corruptionDetected
    }

    return [PSCustomObject]@{
        CorruptionDetected = $corruptionDetected
        Output             = $resultText
    }
}

function Invoke-RepairVolumeSpotFix {
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Attempting online spot fix on $($env:SystemDrive) with Repair-Volume -SpotFix." 'WARN'

    try {
        Repair-Volume -DriveLetter $systemDrive -SpotFix -ErrorAction Stop | Out-Null
        $script:Summary.RepairVolumeSpotFixSucceeded = $true
        Add-RepairAttempt 'Repair-Volume -SpotFix'
        Add-DetailedResult -Step 'RepairVolumeSpotFix' -Status 'Success' -Message 'Repair-Volume -SpotFix completed successfully.'
        Write-Log 'Repair-Volume spot fix completed successfully without scheduling an offline repair.' 'OK'
        return $true
    }
    catch {
        $script:Summary.RepairVolumeSpotFixSucceeded = $false
        Warn-Step -Name 'RepairVolumeSpotFix' -Reason "Repair-Volume -SpotFix could not complete online: $($_.Exception.Message)"
        Add-DetailedResult -Step 'RepairVolumeSpotFix' -Status 'Warning' -Message 'Repair-Volume -SpotFix could not complete online.' -Data @{
            Error = $_.Exception.Message
        }
        return $false
    }
}

function Invoke-RepairVolumeOfflineFix {
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Running offline disk repair on $($env:SystemDrive)" 'WARN'
    Repair-Volume -DriveLetter $systemDrive -OfflineScanAndFix -ErrorAction Stop | Out-Null
    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Repair-Volume -OfflineScanAndFix'
    Add-DetailedResult -Step 'OfflineDiskRepair' -Status 'Info' -Message 'Offline disk repair was started.'
}

function Invoke-DismDetection {
    Write-Log "Running DISM CheckHealth..." 'INFO'
    $check = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/CheckHealth')

    Write-Log "Running DISM ScanHealth..." 'INFO'
    $scan = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/ScanHealth')

    $combined = (($check.StdOut, $scan.StdOut, $check.StdErr, $scan.StdErr) -join "`n")

    if ($check.ExitCode -ne 0 -or $scan.ExitCode -ne 0) {
        $script:Summary.DismCorruptionDetected = $true
        Warn-Step -Name 'DISMDetection' -Reason 'DISM detection returned a non-zero exit code.'
        return
    }

    # Avoid false positives from phrases like "No component store corruption detected."
    if ($combined -match '(?i)No component store corruption detected|No component store corruption was detected|The component store is repairable\s*:\s*No') {
        $script:Summary.DismCorruptionDetected = $false
        Write-Log 'DISM did not detect component store corruption.' 'OK'
        return
    }

    if ($combined -match '(?i)The component store is repairable|component store is repairable|repairable\s*:\s*Yes|corruption detected|component store corruption detected') {
        $script:Summary.DismCorruptionDetected = $true
        Warn-Step -Name 'DISMDetection' -Reason 'DISM detected component store corruption.'
    }
    else {
        $script:Summary.DismCorruptionDetected = $false
        Write-Log 'DISM detection completed without confirmed corruption.' 'OK'
    }
}

function Invoke-DismRepair {
    Write-Log "Running DISM RestoreHealth..." 'WARN'
    $repair = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/RestoreHealth')
    if ($repair.ExitCode -ne 0) {
        throw "DISM RestoreHealth exited with code $($repair.ExitCode)"
    }

    Write-Log "Running DISM StartComponentCleanup..." 'INFO'
    $cleanup = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/StartComponentCleanup')
    if ($cleanup.ExitCode -ne 0) {
        throw "DISM StartComponentCleanup exited with code $($cleanup.ExitCode)"
    }

    Add-RepairAttempt 'DISM RestoreHealth + StartComponentCleanup'
}

function Invoke-SfcDetection {
    Write-Log "Running SFC verify-only scan..." 'INFO'
    $result = Invoke-SfcCommand -Arguments '/verifyonly'

    $combined = (($result.StdOut, $result.StdErr) -join "`n")

    if ($combined -match '(?i)Windows Resource Protection found integrity violations|found integrity violations') {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason 'SFC detected integrity violations.'
        return
    }

    if ($combined -match '(?i)Windows Resource Protection did not find any integrity violations|did not find any integrity violations') {
        $script:Summary.SfcIntegrityViolations = $false
        Write-Log 'SFC did not detect integrity violations.' 'OK'
        return
    }

    if ($combined -match '(?i)Windows Resource Protection found corrupt files and successfully repaired them|Windows Resource Protection found corrupt files but was unable to fix some of them') {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason 'SFC reported corrupt files.'
        return
    }

    if ($result.ExitCode -notin 0,1) {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason "SFC verify returned unusual exit code $($result.ExitCode)."
    }
    else {
        Write-Log 'SFC detection completed without confirmed integrity violations.' 'OK'
    }
}

function Invoke-SfcRepair {
    Write-Log "Running SFC /SCANNOW..." 'WARN'
    $result = Invoke-SfcCommand -Arguments '/scannow'

    if ($result.ExitCode -notin 0,1) {
        throw "SFC /SCANNOW exited with code $($result.ExitCode)"
    }

    Add-RepairAttempt 'SFC /SCANNOW'
}

function Export-CbsCorruptionDetails {
    param(
        [int]$TailLines = 20000
    )

    $cbsPath = Join-Path $env:SystemRoot 'Logs\CBS\CBS.log'
    $extractPath = Join-Path $LogDirectory ($script:BaseFileName + '-CBS-SFC-CorruptionExtract.txt')

    if (-not (Test-Path -LiteralPath $cbsPath)) {
        Write-Log "CBS_CORRUPTION_EXTRACT|Status=MissingCBSLog|Path=$cbsPath" 'WARN'
        Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Warning' -Message 'CBS.log was not found.' -Data @{ Path = $cbsPath }
        return @()
    }

    try {
        $lines = @(Get-Content -LiteralPath $cbsPath -Tail $TailLines -ErrorAction Stop)
        $matches = @($lines | Where-Object {
            $_ -match '\[SR\]' -or
            $_ -match 'Corrupt file:' -or
            $_ -match 'Cannot repair member file' -or
            $_ -match 'Repairing corrupted file' -or
            $_ -match 'Repaired file' -or
            $_ -match 'could not reproject corrupted file' -or
            $_ -match 'Hash mismatch' -or
            $_ -match 'CSI.*corrupt'
        })

        if ($matches.Count -gt 0) {
            $header = @(
                "# CBS/SFC corruption extract",
                "# Computer: $($script:ComputerName)",
                "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
                "# Source: $cbsPath",
                "# TailLinesAnalyzed: $TailLines",
                ""
            )
            Ensure-LogDirectory
            $header + $matches | Set-Content -LiteralPath $extractPath -Encoding UTF8 -Force
            $script:Summary.CbsCorruptionExtractPath = $extractPath

            $corruptFileLines = @($matches | Where-Object { $_ -match 'Corrupt file:|Cannot repair member file|Hash mismatch|could not reproject corrupted file' })
            $script:Summary.CbsCorruptFileCount = $corruptFileLines.Count

            Write-Log "CBS_CORRUPTION_EXTRACT|Status=Created|Path=$extractPath|MatchingLines=$($matches.Count)|CorruptFileLines=$($corruptFileLines.Count)" 'WARN'

            foreach ($line in ($corruptFileLines | Select-Object -First 25)) {
                $safeLine = ([string]$line) -replace '\s+', ' '
                Write-Log "CBS_CORRUPTION_DETAIL|$safeLine" 'WARN'
            }

            $driverCorruptionLines = @($corruptFileLines | Where-Object { $_ -match '\\System32\\drivers\\|\.sys' })
            if ($driverCorruptionLines.Count -gt 0) {
                $script:Summary.CbsDriverCorruptionDetected = $true
                foreach ($driverLine in $driverCorruptionLines) {
                    $driverMatches = [regex]::Matches([string]$driverLine, '(?i)[A-Za-z0-9_.-]+\.sys')
                    foreach ($driverMatch in $driverMatches) {
                        $driverName = $driverMatch.Value
                        if (-not [string]::IsNullOrWhiteSpace($driverName) -and -not $script:Summary.CbsCorruptDriverFiles.Contains($driverName)) {
                            $script:Summary.CbsCorruptDriverFiles.Add($driverName) | Out-Null
                        }
                    }
                }
                $driverList = (($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ',')
                Write-Log "CBS_DRIVER_CORRUPTION_DETECTED|Drivers=$driverList|Action=DISM_SFC_FIRST_THEN_HPIA_DRIVER_ONLY_IF_STILL_CORRUPT" 'WARN'
                Add-Note 'CBS corruption references driver files. If DISM/SFC does not clear it, HP Image Assistant driver-only remediation can be run.'
            }

            Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Warning' -Message 'CBS/SFC corruption details were extracted.' -Data @{
                SourcePath = $cbsPath
                ExtractPath = $extractPath
                MatchingLines = $matches.Count
                CorruptFileLines = $corruptFileLines.Count
            }
        }
        else {
            Write-Log "CBS_CORRUPTION_EXTRACT|Status=NoMatchingCorruptionLines|Path=$cbsPath|TailLines=$TailLines" 'INFO'
            Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Info' -Message 'No CBS/SFC corruption detail lines found in recent CBS.log tail.' -Data @{
                SourcePath = $cbsPath
                TailLines = $TailLines
            }
        }

        return $matches
    }
    catch {
        Write-Log "CBS_CORRUPTION_EXTRACT|Status=Failed|Error=$($_.Exception.Message)" 'WARN'
        Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Warning' -Message 'Failed to extract CBS/SFC corruption details.' -Data @{ Error = $_.Exception.Message }
        return @()
    }
}

function Invoke-SystemFileRepairWorkflow {
    Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Status=Starting|Actions=ExtractCBS,DISMRestoreHealth,SFCScannow,VerifyOnly' 'WARN'
    $script:Summary.SfcRepairAttempted = $true

    Export-CbsCorruptionDetails | Out-Null

    try {
        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=DISMRestoreHealth|Status=Starting' 'WARN'
        $dismRestore = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/RestoreHealth')
        if ($dismRestore.ExitCode -ne 0) {
            throw "DISM RestoreHealth exited with code $($dismRestore.ExitCode)"
        }
        Add-RepairAttempt 'DISM RestoreHealth for SFC/CBS corruption'
        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=DISMRestoreHealth|Status=Completed' 'OK'

        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=SFCScannow|Status=Starting' 'WARN'
        $sfcRepair = Invoke-SfcCommand -Arguments '/scannow'
        if ($sfcRepair.ExitCode -notin 0,1) {
            throw "SFC /SCANNOW exited with code $($sfcRepair.ExitCode)"
        }
        Add-RepairAttempt 'SFC /SCANNOW for CBS corruption'
        Write-Log "SFC_CBS_REPAIR_WORKFLOW|Step=SFCScannow|Status=Completed|ExitCode=$($sfcRepair.ExitCode)" 'OK'

        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=VerifyOnlyAfterRepair|Status=Starting' 'INFO'
        $verify = Invoke-SfcCommand -Arguments '/verifyonly'
        $verifyText = (($verify.StdOut, $verify.StdErr) -join "`n")

        if ($verifyText -match '(?i)did not find any integrity violations') {
            $script:Summary.SfcIntegrityViolations = $false
            $script:Summary.SfcRepairSucceeded = $true
            $script:Summary.CbsDriverCorruptionRemaining = $false
            Write-Log 'SFC_REPAIR_RESULT|Status=CleanAfterRepair|IntegrityViolations=False' 'OK'
        }
        elseif ($verifyText -match '(?i)found integrity violations|found corrupt files|unable to fix') {
            $script:Summary.SfcIntegrityViolations = $true
            $script:Summary.SfcRepairSucceeded = $false
            Write-Log 'SFC_REPAIR_RESULT|Status=StillDetectedAfterRepair|IntegrityViolations=True|NextStep=If CBS driver corruption remains, Tier3 HPIA driver-only remediation may run' 'WARN'
            Export-CbsCorruptionDetails | Out-Null
            $script:Summary.CbsDriverCorruptionRemaining = [bool]$script:Summary.CbsDriverCorruptionDetected
        }
        else {
            $script:Summary.SfcRepairSucceeded = $false
            Write-Log "SFC_REPAIR_RESULT|Status=UnknownAfterRepair|ExitCode=$($verify.ExitCode)" 'WARN'
        }

        Add-DetailedResult -Step 'SfcCbsRepairWorkflow' -Status 'Info' -Message 'SFC/CBS repair workflow completed.' -Data @{
            SfcRepairSucceeded = $script:Summary.SfcRepairSucceeded
            SfcIntegrityViolations = $script:Summary.SfcIntegrityViolations
            CbsCorruptFileCount = $script:Summary.CbsCorruptFileCount
            CbsCorruptionExtractPath = $script:Summary.CbsCorruptionExtractPath
        }
    }
    catch {
        $script:Summary.SfcRepairSucceeded = $false
        Write-Log "SFC_REPAIR_RESULT|Status=Failed|Error=$($_.Exception.Message)" 'ERROR'
        Add-DetailedResult -Step 'SfcCbsRepairWorkflow' -Status 'Failed' -Message $_.Exception.Message -Data @{
            CbsCorruptionExtractPath = $script:Summary.CbsCorruptionExtractPath
        }
        throw
    }
}

function Test-IsHpSystem {
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $manufacturer = [string]$cs.Manufacturer
        if ($manufacturer -match '(?i)\bHP\b|Hewlett-Packard|Hewlett Packard') {
            Write-Log "HP_PLATFORM_DETECTED|Manufacturer=$(Convert-ExplorerSafeString $manufacturer)|Model=$(Convert-ExplorerSafeString $cs.Model)" 'INFO'
            return $true
        }

        Write-Log "HP_PLATFORM_DETECTED|Status=False|Manufacturer=$(Convert-ExplorerSafeString $manufacturer)|Model=$(Convert-ExplorerSafeString $cs.Model)" 'INFO'
        return $false
    }
    catch {
        Write-Log "HP_PLATFORM_DETECTED|Status=Unknown|Error=$($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Find-HpImageAssistantExecutable {
    [CmdletBinding()]
    param()

    $candidatePaths = @(
        'C:\Program Files\HP\HP Image Assistant\HPImageAssistant.exe',
        'C:\Program Files (x86)\HP\HP Image Assistant\HPImageAssistant.exe',
        (Join-Path $HpImageAssistantLocalPath 'HPImageAssistant.exe')
    )

    foreach ($candidate in $candidatePaths) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    try {
        $localFound = Get-ChildItem -LiteralPath $HpImageAssistantLocalPath -Filter 'HPImageAssistant.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($localFound) { return $localFound.FullName }
    }
    catch { }

    return $null
}

function Ensure-HpImageAssistantAvailable {
    [CmdletBinding()]
    param()

    $existing = Find-HpImageAssistantExecutable
    if ($existing) {
        $script:Summary.HpImageAssistantPath = $existing
        Write-Log "HPIA_AVAILABLE|Status=AlreadyPresent|Path=$existing" 'OK'
        return $existing
    }

    Write-Log "HPIA_AVAILABLE|Status=NotFound|Action=StageFromShare|Source=$HpImageAssistantSourcePath|Destination=$HpImageAssistantLocalPath" 'WARN'

    if (-not (Test-Path -LiteralPath $HpImageAssistantSourcePath)) {
        throw "HP Image Assistant source path was not found: $HpImageAssistantSourcePath"
    }

    if (-not (Test-Path -LiteralPath $HpImageAssistantLocalPath)) {
        New-Item -Path $HpImageAssistantLocalPath -ItemType Directory -Force | Out-Null
    }

    $sourceExe = $null
    try {
        $sourceExe = Get-ChildItem -LiteralPath $HpImageAssistantSourcePath -Filter 'HPImageAssistant.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    catch { }

    if ($sourceExe) {
        Write-Log "HPIA_STAGE|Mode=CopyExtractedFolder|Source=$($sourceExe.DirectoryName)|Destination=$HpImageAssistantLocalPath" 'INFO'
        Copy-Item -LiteralPath (Join-Path $sourceExe.DirectoryName '*') -Destination $HpImageAssistantLocalPath -Recurse -Force -ErrorAction Stop
        $stagedExe = Find-HpImageAssistantExecutable
        if ($stagedExe) {
            $script:Summary.HpImageAssistantPath = $stagedExe
            Write-Log "HPIA_AVAILABLE|Status=Staged|Path=$stagedExe" 'OK'
            return $stagedExe
        }
    }

    $installer = $null
    try {
        $installer = Get-ChildItem -LiteralPath $HpImageAssistantSourcePath -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '(?i)hp.*image.*assistant|hpia|sp\d+' -and $_.Extension -match '(?i)\.exe|\.msi' } |
            Sort-Object Length -Descending |
            Select-Object -First 1
    }
    catch { }

    if (-not $installer) {
        throw "Could not find HPImageAssistant.exe or an HPIA installer under $HpImageAssistantSourcePath"
    }

    Write-Log "HPIA_STAGE|Mode=Installer|Installer=$($installer.FullName)|Destination=$HpImageAssistantLocalPath" 'INFO'

    if ($installer.Extension -ieq '.msi') {
        $args = "/i `"$($installer.FullName)`" /qn /norestart"
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -notin 0,3010,1641) {
            throw "HPIA MSI installer failed with exit code $($proc.ExitCode)"
        }
    }
    else {
        $extractArgs = "/s /e /f `"$HpImageAssistantLocalPath`""
        $proc = Start-Process -FilePath $installer.FullName -ArgumentList $extractArgs -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -notin 0,1168,3010,1641) {
            throw "HPIA installer/extractor failed with exit code $($proc.ExitCode)"
        }
    }

    $afterInstall = Find-HpImageAssistantExecutable
    if (-not $afterInstall) {
        throw "HPImageAssistant.exe was not found after staging/installing HPIA."
    }

    $script:Summary.HpImageAssistantPath = $afterInstall
    Write-Log "HPIA_AVAILABLE|Status=InstalledOrExtracted|Path=$afterInstall" 'OK'
    return $afterInstall
}

function Invoke-HpiaDriverOnlyRepairFromCbs {
    [CmdletBinding()]
    param()

    Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Evaluating|Trigger=CBS_DRIVER_CORRUPTION_REMAINING' 'INFO'

    if (-not $AllowHpDriverRepairFromCbs) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=AllowHpDriverRepairFromCbsDisabled' 'INFO'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'HP driver remediation is disabled by parameter.'
        return
    }

    if (-not $script:Summary.CbsDriverCorruptionDetected) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=NoCBSDriverCorruptionDetected' 'INFO'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'No CBS driver corruption marker was detected.'
        return
    }

    if ($script:Summary.SfcRepairSucceeded -and -not $script:Summary.SfcIntegrityViolations) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=DISMAndSfcClearedIntegrityViolations' 'OK'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'DISM/SFC cleared the issue; HPIA driver remediation not needed.'
        return
    }

    if (-not (Test-IsHpSystem)) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=NotAnHPSystem' 'WARN'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'System manufacturer is not HP/Hewlett-Packard.'
        return
    }

    $script:Summary.HpDriverRepairAttempted = $true
    $driverList = (($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ',')
    if ([string]::IsNullOrWhiteSpace($driverList)) { $driverList = 'UnknownDriverFilesFromCBS' }

    Write-Log "HPIA_DRIVER_REMEDIATION|Status=Starting|Trigger=CBS_DRIVER_CORRUPTION|Drivers=$driverList|Category=DriversOnly|Excluded=BIOS,Firmware" 'WARN'

    $hpiaExe = Ensure-HpImageAssistantAvailable
    $hpiaRoot = Join-Path 'C:\ProgramData\SystemRepair' 'HPIA-CBS-DriverRepair'
    $hpiaReportFolder = Join-Path $hpiaRoot 'Reports'
    $hpiaDownloadFolder = Join-Path $hpiaRoot 'SoftPaqs'
    New-Item -Path $hpiaReportFolder -ItemType Directory -Force | Out-Null
    New-Item -Path $hpiaDownloadFolder -ItemType Directory -Force | Out-Null

    $args = @(
        '/Operation:Analyze',
        '/Action:Install',
        '/Selection:All',
        '/Category:Drivers',
        '/Silent',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`""
    )

    Write-Log "HPIA_DRIVER_REMEDIATION|Command=$hpiaExe $($args -join ' ')" 'INFO'

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $hpiaExe
    $psi.Arguments = $args -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 80 | ForEach-Object { Write-Log "HPIA_STDOUT|$($_.Trim())" 'INFO' }
    }
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 80 | ForEach-Object { Write-Log "HPIA_STDERR|$($_.Trim())" 'WARN' }
    }

    $reportFiles = @()
    try { $reportFiles = @(Get-ChildItem -LiteralPath $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName) } catch { }

    if ($proc.ExitCode -in 0,256,3010,3020,4096) {
        if ($proc.ExitCode -in 3010,3020) { $script:Summary.RebootRequired = $true }
        $script:Summary.HpDriverRepairSucceeded = $true
        Add-RepairAttempt 'HP Image Assistant driver-only remediation for CBS driver corruption'
        Write-Log "HPIA_DRIVER_REMEDIATION|Status=Completed|ExitCode=$($proc.ExitCode)|ReportFolder=$hpiaReportFolder|SoftpaqDownloadFolder=$hpiaDownloadFolder|RebootRequired=$($script:Summary.RebootRequired)" 'OK'
    }
    else {
        $script:Summary.HpDriverRepairSucceeded = $false
        Write-Log "HPIA_DRIVER_REMEDIATION|Status=Failed|ExitCode=$($proc.ExitCode)|ReportFolder=$hpiaReportFolder|SoftpaqDownloadFolder=$hpiaDownloadFolder" 'ERROR'
    }

    if ($script:Summary.HpDriverRepairSucceeded) {
        Write-Log 'HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=Starting' 'INFO'
        try {
            $postHpiaVerify = Invoke-SfcCommand -Arguments '/verifyonly'
            $postHpiaVerifyText = (($postHpiaVerify.StdOut, $postHpiaVerify.StdErr) -join "`n")
            if ($postHpiaVerifyText -match '(?i)did not find any integrity violations') {
                $script:Summary.SfcIntegrityViolations = $false
                $script:Summary.CbsDriverCorruptionRemaining = $false
                Write-Log 'HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=CleanAfterDriverRepair|IntegrityViolations=False' 'OK'
            }
            elseif ($postHpiaVerifyText -match '(?i)found integrity violations|found corrupt files|unable to fix') {
                $script:Summary.SfcIntegrityViolations = $true
                $script:Summary.CbsDriverCorruptionRemaining = $true
                Write-Log 'HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=StillCorruptAfterDriverRepair|IntegrityViolations=True|NextStep=Review CBS extract and consider in-place repair install' 'WARN'
                Export-CbsCorruptionDetails | Out-Null
            }
            else {
                Write-Log "HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=Unknown|ExitCode=$($postHpiaVerify.ExitCode)" 'WARN'
            }
        }
        catch {
            Write-Log "HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=Failed|Error=$($_.Exception.Message)" 'WARN'
        }
    }

    Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Info' -Message 'HPIA driver-only remediation attempted because CBS driver corruption was detected.' -Data @{
        ExitCode = $proc.ExitCode
        HpiaPath = $hpiaExe
        ReportFolder = $hpiaReportFolder
        SoftpaqDownloadFolder = $hpiaDownloadFolder
        DriverFilesFromCbs = $driverList
        ReportFiles = ($reportFiles -join ';')
        PostRepairSfcIntegrityViolations = $script:Summary.SfcIntegrityViolations
    }

    if (-not $script:Summary.HpDriverRepairSucceeded) {
        throw "HPIA driver-only remediation failed with exit code $($proc.ExitCode)"
    }
}

function Invoke-WmiCheck {
    Write-Log "Checking WMI repository consistency..." 'INFO'
    $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /verifyrepository 2>&1
    $text = ($output | Out-String).Trim()

    if ($text) {
        $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    Add-DetailedResult -Step 'WmiRepositoryCheck' -Status 'Info' -Message 'WMI verify completed.' -Data @{
        Output = $text
    }

    if ($text -match 'inconsistent') {
        $script:Summary.WmiRepositoryInconsistent = $true
        Warn-Step -Name 'WmiRepositoryCheck' -Reason 'WMI repository reported as inconsistent.'
    }
}

function Invoke-WmiRepair {
    Write-Log "Repairing WMI repository..." 'WARN'
    $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /salvagerepository 2>&1
    $text = ($output | Out-String).Trim()

    if ($text) {
        $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    Add-RepairAttempt 'winmgmt /salvagerepository'
    Add-DetailedResult -Step 'WmiRepair' -Status 'Info' -Message 'WMI salvage completed.' -Data @{
        Output = $text
    }
}

function Invoke-NetworkReset {
    Write-Log "Flushing DNS cache..." 'INFO'
    ipconfig /flushdns | Out-Null

    Write-Log "Resetting Winsock..." 'WARN'
    netsh winsock reset | Out-Null

    Write-Log "Resetting TCP/IP stack..." 'WARN'
    netsh int ip reset | Out-Null

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Winsock/TCPIP reset'
    Add-DetailedResult -Step 'NetworkReset' -Status 'Info' -Message 'Network reset completed.'
}

function Invoke-DnsFlushOnly {
    Write-Log "Flushing DNS cache..." 'INFO'
    ipconfig /flushdns | Out-Null
    Add-DetailedResult -Step 'DnsFlush' -Status 'Info' -Message 'DNS cache flushed.'
}


function Convert-RpcSafeString {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value) -replace "[`r`n|]+", ' '
}

function Get-RpcRootCauseAssessment {
    param(
        [Parameter(Mandatory)][bool]$LocalProblem,
        [Parameter(Mandatory)][bool]$RemoteProblem,
        [Parameter(Mandatory)][int]$ServiceProblemCount,
        [Parameter(Mandatory)][bool]$LocalEndpointMapperReachable,
        [Parameter(Mandatory)][bool]$LocalEventLogRead,
        [Parameter(Mandatory)][bool]$LocalWmiQuery,
        [Parameter(Mandatory)][int]$RemoteProblemCount
    )

    if ($LocalProblem) {
        if (-not $LocalEndpointMapperReachable) {
            return 'LOCAL_CLIENT_OR_LOCAL_RPC_SERVICE_PROBLEM: Local RPC endpoint mapper port 135 is not reachable on 127.0.0.1. Check RpcSs/RpcEptMapper/DcomLaunch services and local firewall/security software.'
        }
        if ($ServiceProblemCount -gt 0) {
            return "LOCAL_CLIENT_OR_LOCAL_SERVICE_PROBLEM: One or more required local services are missing/stopped/unhealthy. ServiceProblemCount=$ServiceProblemCount."
        }
        if (-not $LocalEventLogRead -and $LocalWmiQuery) {
            return 'LOCAL_EVENTLOG_ACCESS_PROBLEM: Local RPC/WMI basics are healthy, but local Event Log access failed. Check Windows Event Log service, event log corruption, permissions, or local security policy.'
        }
        if (-not $LocalWmiQuery) {
            return 'LOCAL_WMI_OR_CIM_PROBLEM: Local WMI/CIM query failed. Check Winmgmt service and WMI repository health.'
        }
        return 'LOCAL_CLIENT_SIDE_PROBLEM: Local RPC-related validation failed. Review RPC_LOCAL_* markers above.'
    }

    if ($RemoteProblem) {
        return "REMOTE_ENDPOINT_OR_NETWORK_PROBLEM: Local RPC checks passed, but one or more target systems did not answer DNS/RPC port 135. RemoteProblemCount=$RemoteProblemCount. Check DNS, routing/VLANs, remote firewall, remote RPC services, and domain controller reachability."
    }

    return 'NO_RPC_PROBLEM_DETECTED_BY_BASIC_CHECKS: Local RPC services, local endpoint mapper, Event Log access, WMI/CIM, and tested remote endpoint mapper targets passed. If another section still reports RPC unavailable, that specific provider/query may be failing rather than core RPC.'
}

function Test-TcpPortSafe {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][int]$Port,
        [int]$TimeoutMilliseconds = 2500
    )

    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $connected = $async.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)
        if (-not $connected) {
            return [PSCustomObject]@{ ComputerName = $ComputerName; Port = $Port; Reachable = $false; Error = 'Timeout' }
        }
        $client.EndConnect($async)
        return [PSCustomObject]@{ ComputerName = $ComputerName; Port = $Port; Reachable = $true; Error = '' }
    }
    catch {
        return [PSCustomObject]@{ ComputerName = $ComputerName; Port = $Port; Reachable = $false; Error = $_.Exception.Message }
    }
    finally {
        if ($client) { $client.Close() }
    }
}

function Invoke-RpcDiagnostics {
    Write-Log 'Running RPC diagnostics to separate local client/service issues from remote RPC endpoint issues...' 'INFO'

    $serviceNames = @('RpcSs', 'RpcEptMapper', 'DcomLaunch', 'EventLog', 'Winmgmt')
    $serviceProblems = New-Object System.Collections.Generic.List[string]
    $serviceResults = New-Object System.Collections.Generic.List[object]

    foreach ($name in $serviceNames) {
        try {
            $svc = Get-CimInstance Win32_Service -Filter "Name='$name'" -ErrorAction Stop
            if ($null -eq $svc) {
                $serviceProblems.Add("$name missing") | Out-Null
                Write-Log "RPC_LOCAL_SERVICE|Name=$name|Exists=False|State=Missing|StartMode=Unknown|Status=Problem" 'WARN'
                continue
            }

            $status = if ($svc.State -eq 'Running') { 'OK' } else { 'Problem' }
            if ($status -ne 'OK') { $serviceProblems.Add("$name state $($svc.State)") | Out-Null }

            Write-Log "RPC_LOCAL_SERVICE|Name=$name|Exists=True|State=$(Convert-RpcSafeString $svc.State)|StartMode=$(Convert-RpcSafeString $svc.StartMode)|ProcessId=$($svc.ProcessId)|Status=$status" ($(if ($status -eq 'OK') { 'INFO' } else { 'WARN' }))
            $serviceResults.Add([PSCustomObject]@{ Name=$name; State=$svc.State; StartMode=$svc.StartMode; ProcessId=$svc.ProcessId; Status=$status }) | Out-Null
        }
        catch {
            $serviceProblems.Add("$name query failed") | Out-Null
            Write-Log "RPC_LOCAL_SERVICE|Name=$name|Exists=Unknown|State=Unknown|StartMode=Unknown|Status=Problem|Error=$(Convert-RpcSafeString $_.Exception.Message)" 'WARN'
        }
    }

    $localEndpoint = Test-TcpPortSafe -ComputerName '127.0.0.1' -Port 135
    Write-Log "RPC_ENDPOINT_MAPPER|Target=127.0.0.1|Port=135|Reachable=$($localEndpoint.Reachable)|Error=$(Convert-RpcSafeString $localEndpoint.Error)" ($(if ($localEndpoint.Reachable) { 'INFO' } else { 'WARN' }))

    $eventLogOk = $false
    $eventLogError = ''
    try {
        $null = Get-WinEvent -ListLog System -ErrorAction Stop
        $null = Get-WinEvent -FilterHashtable @{ LogName='System'; StartTime=(Get-Date).AddMinutes(-5) } -MaxEvents 1 -ErrorAction SilentlyContinue
        $eventLogOk = $true
    }
    catch {
        $eventLogError = $_.Exception.Message
    }
    Write-Log "RPC_LOCAL_EVENTLOG_READ|Log=System|Success=$eventLogOk|Error=$(Convert-RpcSafeString $eventLogError)" ($(if ($eventLogOk) { 'INFO' } else { 'WARN' }))

    $cimOk = $false
    $cimError = ''
    try {
        $null = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop | Select-Object -First 1
        $cimOk = $true
    }
    catch {
        $cimError = $_.Exception.Message
    }
    Write-Log "RPC_LOCAL_WMI_QUERY|Class=Win32_OperatingSystem|Success=$cimOk|Error=$(Convert-RpcSafeString $cimError)" ($(if ($cimOk) { 'INFO' } else { 'WARN' }))

    try {
        if (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
            Get-NetFirewallProfile -ErrorAction SilentlyContinue | ForEach-Object {
                Write-Log "RPC_FIREWALL_PROFILE|Name=$(Convert-RpcSafeString $_.Name)|Enabled=$($_.Enabled)|DefaultInboundAction=$(Convert-RpcSafeString $_.DefaultInboundAction)|DefaultOutboundAction=$(Convert-RpcSafeString $_.DefaultOutboundAction)" 'INFO'
            }
        }
    }
    catch {
        Write-Log "RPC_FIREWALL_PROFILE|Status=UnableToQuery|Error=$(Convert-RpcSafeString $_.Exception.Message)" 'WARN'
    }

    $targets = New-Object System.Collections.Generic.List[string]
    if ($env:LOGONSERVER) {
        $targets.Add(($env:LOGONSERVER -replace '^\\+', '')) | Out-Null
    }

    try {
        $dnsServers = @(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | ForEach-Object { $_.ServerAddresses } | Where-Object { $_ } | Select-Object -Unique)
        foreach ($dns in $dnsServers) { $targets.Add([string]$dns) | Out-Null }
    }
    catch { }

    $remoteProblems = New-Object System.Collections.Generic.List[object]
    foreach ($target in ($targets | Where-Object { $_ -and $_ -notmatch '^(127\.|::1|localhost$)' } | Select-Object -Unique -First 8)) {
        $dnsOk = $true
        $dnsError = ''
        try { $null = [System.Net.Dns]::GetHostAddresses($target) }
        catch { $dnsOk = $false; $dnsError = $_.Exception.Message }

        $port135 = Test-TcpPortSafe -ComputerName $target -Port 135
        Write-Log "RPC_TARGET_PORT_TEST|Target=$(Convert-RpcSafeString $target)|DnsResolved=$dnsOk|Port=135|Reachable=$($port135.Reachable)|Error=$(Convert-RpcSafeString $(if (-not $dnsOk) { $dnsError } else { $port135.Error }))" ($(if ($dnsOk -and $port135.Reachable) { 'INFO' } else { 'WARN' }))

        if (-not $dnsOk -or -not $port135.Reachable) {
            $remoteProblems.Add([PSCustomObject]@{ Target=$target; DnsResolved=$dnsOk; Port135Reachable=$port135.Reachable; Error=$(if (-not $dnsOk) { $dnsError } else { $port135.Error }) }) | Out-Null
        }
    }

    $localProblem = ($serviceProblems.Count -gt 0 -or -not $localEndpoint.Reachable -or -not $eventLogOk -or -not $cimOk)
    $remoteProblem = (-not $localProblem -and $remoteProblems.Count -gt 0)

    $script:Summary.RpcClientSideIssueDetected = [bool]$localProblem
    $script:Summary.RpcRemoteSideIssueDetected = [bool]$remoteProblem

    $rpcAssessment = Get-RpcRootCauseAssessment -LocalProblem ([bool]$localProblem) -RemoteProblem ([bool]$remoteProblem) -ServiceProblemCount ($serviceProblems.Count) -LocalEndpointMapperReachable ([bool]$localEndpoint.Reachable) -LocalEventLogRead ([bool]$eventLogOk) -LocalWmiQuery ([bool]$cimOk) -RemoteProblemCount ($remoteProblems.Count)

    if ($localProblem) {
        $script:Summary.RpcDiagnosticsStatus = 'LocalClientOrLocalServiceIssueSuspected'
        Write-Log "RPC_CLIENT_SIDE_SUSPECTED|ServiceProblems=$($serviceProblems.Count)|LocalEndpointMapperReachable=$($localEndpoint.Reachable)|LocalEventLogRead=$eventLogOk|LocalWmiQuery=$cimOk" 'WARN'
        Add-Note 'RPC diagnostics suggest a local client/service-side issue. Review RPC_LOCAL_SERVICE, RPC_ENDPOINT_MAPPER, RPC_LOCAL_EVENTLOG_READ, and RPC_LOCAL_WMI_QUERY markers.'
    }
    elseif ($remoteProblem) {
        $script:Summary.RpcDiagnosticsStatus = 'RemoteEndpointOrNetworkIssueSuspected'
        Write-Log "RPC_REMOTE_SIDE_OR_NETWORK_SUSPECTED|RemoteProblemCount=$($remoteProblems.Count)|LocalEndpointMapperReachable=$($localEndpoint.Reachable)|LocalEventLogRead=$eventLogOk|LocalWmiQuery=$cimOk" 'WARN'
        Add-Note 'RPC diagnostics suggest local RPC is healthy, but one or more domain/DNS/logon targets did not answer RPC endpoint mapper port 135.'
    }
    else {
        $script:Summary.RpcDiagnosticsStatus = 'NoRpcIssueDetectedByBasicChecks'
        Write-Log "RPC_DIAGNOSTICS_OK|LocalEndpointMapperReachable=$($localEndpoint.Reachable)|LocalEventLogRead=$eventLogOk|LocalWmiQuery=$cimOk|RemoteProblemCount=$($remoteProblems.Count)" 'OK'
    }

    $rpcAssessmentLevel = if ($localProblem -or $remoteProblem) { 'WARN' } else { 'OK' }
    Write-Log "RPC_ROOT_CAUSE_ASSESSMENT|$rpcAssessment" $rpcAssessmentLevel
    Write-Log "RPC_DIAGNOSTICS_SUMMARY|Status=$($script:Summary.RpcDiagnosticsStatus)|ClientSideIssue=$($script:Summary.RpcClientSideIssueDetected)|RemoteOrNetworkIssue=$($script:Summary.RpcRemoteSideIssueDetected)|ServiceProblems=$($serviceProblems.Count)|RemoteProblems=$($remoteProblems.Count)|Assessment=$(Convert-RpcSafeString $rpcAssessment)" ($(if ($localProblem -or $remoteProblem) { 'WARN' } else { 'INFO' }))
    Add-Note "RPC root cause assessment: $rpcAssessment"

    Add-DetailedResult -Step 'RpcDiagnostics' -Status 'Info' -Message 'RPC diagnostics completed.' -Data @{
        Status = $script:Summary.RpcDiagnosticsStatus
        ClientSideIssueDetected = $script:Summary.RpcClientSideIssueDetected
        RemoteSideIssueDetected = $script:Summary.RpcRemoteSideIssueDetected
        ServiceResultsJson = (($serviceResults | ConvertTo-Json -Compress) -replace "'", "''")
        RemoteProblemsJson = (($remoteProblems | ConvertTo-Json -Compress) -replace "'", "''")
        LocalEndpointMapperReachable = $localEndpoint.Reachable
        LocalEventLogRead = $eventLogOk
        LocalWmiQuery = $cimOk
        RootCauseAssessment = $rpcAssessment
    }
}



function Convert-ExplorerSafeString {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value) -replace '[\r\n\|]+', ' '
}

function Get-ExplorerEventSummary {
    param([int]$DaysBack = 14)

    Write-Log "Collecting Explorer hang/crash events from the last $DaysBack day(s)..." 'INFO'
    $startTime = (Get-Date).AddDays(-1 * [math]::Abs($DaysBack))
    $events = @()

    try {
        $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $startTime } -ErrorAction Stop | Where-Object {
            ($_.ProviderName -match 'Application Error|Windows Error Reporting|Application Hang') -and
            ($_.Message -match '(?i)explorer\.exe|File Explorer')
        })
    }
    catch {
        Warn-Step -Name 'ExplorerEventCollection' -Reason "Could not read Application log for Explorer events: $($_.Exception.Message)"
        return @()
    }

    $hangs = @($events | Where-Object { $_.ProviderName -match 'Application Hang' -or $_.Message -match '(?i)AppHang|not responding|stopped interacting' })
    $crashes = @($events | Where-Object { $_.ProviderName -match 'Application Error|Windows Error Reporting' -and $_.Message -match '(?i)faulting application|fault bucket|crash|stopped working' })

    $script:Summary.ExplorerHangCount = @($hangs).Count
    $script:Summary.ExplorerCrashCount = @($crashes).Count

    Write-Log "EXPLORER_HEALTH_EVENTS|Days=$DaysBack|Hangs=$($script:Summary.ExplorerHangCount)|Crashes=$($script:Summary.ExplorerCrashCount)|Total=$(@($events).Count)" 'INFO'

    foreach ($event in ($events | Select-Object -First 25)) {
        $kind = if ($event.ProviderName -match 'Application Hang' -or $event.Message -match '(?i)AppHang|not responding|stopped interacting') { 'HANG' } else { 'CRASH' }
        $faultingModule = ''
        if ($event.Message -match '(?im)Faulting module name:\s*([^,\r\n]+)') { $faultingModule = $Matches[1].Trim() }
        $exceptionCode = ''
        if ($event.Message -match '(?im)Exception code:\s*([^\r\n]+)') { $exceptionCode = $Matches[1].Trim() }
        Write-Log "EXPLORER_${kind}_DETECTED|Time=$($event.TimeCreated)|Provider=$(Convert-ExplorerSafeString $event.ProviderName)|EventId=$($event.Id)|FaultingModule=$(Convert-ExplorerSafeString $faultingModule)|ExceptionCode=$(Convert-ExplorerSafeString $exceptionCode)" 'WARN'
    }

    Add-DetailedResult -Step 'ExplorerEventCollection' -Status 'Info' -Message 'Collected Explorer hang/crash event summary.' -Data @{
        DaysBack = $DaysBack
        HangCount = $script:Summary.ExplorerHangCount
        CrashCount = $script:Summary.ExplorerCrashCount
        TotalExplorerEvents = @($events).Count
    }

    return $events
}

function Get-NonMicrosoftShellExtensionInventory {
    Write-Log 'Collecting non-Microsoft shell extension and context-menu handler inventory...' 'INFO'

    $registryRoots = @(
        'Registry::HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Directory\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Drive\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Folder\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\AllFileSystemObjects\shellex\ContextMenuHandlers',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved'
    )

    $items = New-Object System.Collections.Generic.List[object]

    foreach ($root in $registryRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }

        if ($root -match '\\Approved$') {
            try {
                $props = Get-ItemProperty -LiteralPath $root -ErrorAction Stop
                foreach ($prop in $props.PSObject.Properties | Where-Object { $_.Name -match '^\{' }) {
                    $items.Add([PSCustomObject]@{ RegistryPath = $root; Name = $prop.Value; Clsid = $prop.Name; Source = 'ApprovedShellExtension' }) | Out-Null
                }
            }
            catch {
                Write-Log "Could not enumerate $root : $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue | ForEach-Object {
                $defaultValue = ''
                try { $defaultValue = (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop).'(default)' } catch { }
                if ([string]::IsNullOrWhiteSpace($defaultValue)) { $defaultValue = $_.GetValue('') }
                $items.Add([PSCustomObject]@{ RegistryPath = $_.PSPath; Name = $_.PSChildName; Clsid = [string]$defaultValue; Source = 'ContextMenuHandler' }) | Out-Null
            }
        }
    }

    $nonMicrosoft = @($items | Where-Object {
        $text = "$(Convert-ExplorerSafeString $_.Name) $(Convert-ExplorerSafeString $_.RegistryPath)"
        $text -notmatch '(?i)Microsoft|Windows|Offline Files|Sharing|WorkFolders|FileSyncEx|ModernSharing|EPP|Defender'
    } | Sort-Object Source, Name, Clsid -Unique)

    $script:Summary.NonMicrosoftShellExtensionCount = @($nonMicrosoft).Count
    Write-Log "EXPLORER_SHELL_EXTENSION_SUMMARY|NonMicrosoftCount=$($script:Summary.NonMicrosoftShellExtensionCount)|TotalInventoryCount=$($items.Count)" 'INFO'

    foreach ($entry in ($nonMicrosoft | Select-Object -First 75)) {
        Write-Log "NON_MICROSOFT_SHELL_EXTENSION|Source=$(Convert-ExplorerSafeString $entry.Source)|Name=$(Convert-ExplorerSafeString $entry.Name)|Clsid=$(Convert-ExplorerSafeString $entry.Clsid)|RegistryPath=$(Convert-ExplorerSafeString $entry.RegistryPath)" 'WARN'
    }

    Add-DetailedResult -Step 'ExplorerShellExtensionInventory' -Status 'Info' -Message 'Collected shell extension inventory.' -Data @{
        NonMicrosoftCount = $script:Summary.NonMicrosoftShellExtensionCount
        TotalInventoryCount = $items.Count
        NonMicrosoftJson = (($nonMicrosoft | Select-Object -First 100) | ConvertTo-Json -Compress)
    }

    return $nonMicrosoft
}

function Test-ExplorerNetworkPaths {
    Write-Log 'Checking mapped drives, SMB mappings, and disconnected network paths that can hang Explorer...' 'INFO'

    $problems = New-Object System.Collections.Generic.List[object]

    try {
        $drives = @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.DisplayRoot -match '^\\' })
        foreach ($drive in $drives) {
            $ok = $false
            try { $ok = Test-Path -LiteralPath $drive.DisplayRoot -ErrorAction Stop } catch { $ok = $false }
            if (-not $ok) {
                $problems.Add([PSCustomObject]@{ Type='PSDrive'; Name=$drive.Name; Path=$drive.DisplayRoot; Status='Unreachable' }) | Out-Null
            }
        }
    }
    catch { Write-Log "PSDrive network check failed: $($_.Exception.Message)" 'WARN' }

    try {
        if (Get-Command Get-SmbMapping -ErrorAction SilentlyContinue) {
            $mappings = @(Get-SmbMapping -ErrorAction SilentlyContinue)
            foreach ($map in $mappings) {
                $status = [string]$map.Status
                if ($status -and $status -notmatch '(?i)^OK$|Connected') {
                    $problems.Add([PSCustomObject]@{ Type='SmbMapping'; Name=$map.LocalPath; Path=$map.RemotePath; Status=$status }) | Out-Null
                }
            }
        }
    }
    catch { Write-Log "SMB mapping check failed: $($_.Exception.Message)" 'WARN' }

    try {
        $printers = @(Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue | Where-Object { $_.Network -eq $true -and $_.WorkOffline -eq $true })
        foreach ($printer in $printers) {
            $problems.Add([PSCustomObject]@{ Type='NetworkPrinter'; Name=$printer.Name; Path=$printer.PortName; Status='Offline' }) | Out-Null
        }
    }
    catch { Write-Log "Network printer check failed: $($_.Exception.Message)" 'WARN' }

    $script:Summary.BrokenNetworkPathCount = $problems.Count
    Write-Log "EXPLORER_NETWORK_PATH_SUMMARY|BrokenOrOfflineCount=$($script:Summary.BrokenNetworkPathCount)" 'INFO'

    foreach ($item in $problems) {
        Write-Log "EXPLORER_BROKEN_NETWORK_PATH|Type=$(Convert-ExplorerSafeString $item.Type)|Name=$(Convert-ExplorerSafeString $item.Name)|Path=$(Convert-ExplorerSafeString $item.Path)|Status=$(Convert-ExplorerSafeString $item.Status)" 'WARN'
    }

    Add-DetailedResult -Step 'ExplorerNetworkPathCheck' -Status 'Info' -Message 'Checked mapped drives, SMB mappings, and offline network printers.' -Data @{
        BrokenOrOfflineCount = $script:Summary.BrokenNetworkPathCount
        ProblemsJson = ($problems | ConvertTo-Json -Compress)
    }

    return $problems
}

function Invoke-ExplorerCacheCleanup {
    Write-Log 'Cleaning Explorer thumbnail, icon, and Quick Access caches...' 'WARN'

    $deleted = New-Object System.Collections.Generic.List[string]
    $targets = @(
        @{ Path = "$env:LOCALAPPDATA\IconCache.db"; Pattern = $null },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Pattern = 'thumbcache*.db' },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Pattern = 'iconcache*.db' },
        @{ Path = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"; Pattern = '*' },
        @{ Path = "$env:APPDATA\Microsoft\Windows\Recent\CustomDestinations"; Pattern = '*' }
    )

    foreach ($target in $targets) {
        if (-not (Test-Path -LiteralPath $target.Path)) { continue }
        if ($null -eq $target.Pattern) {
            try {
                Remove-Item -LiteralPath $target.Path -Force -ErrorAction Stop
                $deleted.Add($target.Path) | Out-Null
            }
            catch { Write-Log "Could not delete Explorer cache file $($target.Path): $($_.Exception.Message)" 'WARN' }
        }
        else {
            Get-ChildItem -LiteralPath $target.Path -Filter $target.Pattern -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                    $deleted.Add($_.FullName) | Out-Null
                }
                catch { Write-Log "Could not delete Explorer cache item $($_.FullName): $($_.Exception.Message)" 'WARN' }
            }
        }
    }

    Write-Log "EXPLORER_CACHE_CLEANUP|DeletedCount=$($deleted.Count)" 'INFO'
    Add-RepairAttempt 'Explorer thumbnail/icon/Quick Access cache cleanup'
    Add-DetailedResult -Step 'ExplorerCacheCleanup' -Status 'Info' -Message 'Explorer cache cleanup completed.' -Data @{
        DeletedCount = $deleted.Count
        DeletedItems = ($deleted -join '; ')
    }
}

function Invoke-SearchServiceRestart {
    if (-not $AllowSearchServiceRestart) {
        Write-Log 'Windows Search service restart skipped because AllowSearchServiceRestart is disabled.' 'INFO'
        return
    }

    try {
        $svc = Get-Service -Name WSearch -ErrorAction SilentlyContinue
        if ($null -eq $svc) {
            Write-Log 'Windows Search service was not found on this system.' 'INFO'
            return
        }

        if ($svc.Status -eq 'Running') {
            Write-Log 'Restarting Windows Search service to clear possible Explorer/indexer hangs...' 'WARN'
            Restart-Service -Name WSearch -Force -ErrorAction Stop
        }
        else {
            Write-Log "Starting Windows Search service. Current status: $($svc.Status)" 'WARN'
            Start-Service -Name WSearch -ErrorAction Stop
        }

        Add-RepairAttempt 'Windows Search service restart/start'
        Add-DetailedResult -Step 'SearchServiceRestart' -Status 'Info' -Message 'Windows Search service restart/start completed.'
        Write-Log 'EXPLORER_SEARCH_SERVICE_RESTARTED|Service=WSearch' 'INFO'
    }
    catch {
        Warn-Step -Name 'SearchServiceRestart' -Reason $_.Exception.Message
    }
}

function Get-ExplorerRelatedDriverInventory {
    Write-Log 'Collecting GPU and storage controller driver inventory for Explorer hang correlation...' 'INFO'
    $video = @()
    $storage = @()

    try { $video = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue) } catch { }
    try { $storage = @(Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue | Where-Object { $_.DeviceClass -match 'DISPLAY|SCSIAdapter|HDC|System' -and $_.DeviceName -match '(?i)graphics|display|intel|amd|nvidia|storage|nvme|controller|chipset|vmd|raid|sata' }) } catch { }

    foreach ($gpu in $video) {
        Write-Log "EXPLORER_GPU_DRIVER|Name=$(Convert-ExplorerSafeString $gpu.Name)|DriverVersion=$(Convert-ExplorerSafeString $gpu.DriverVersion)|DriverDate=$(Convert-ExplorerSafeString $gpu.DriverDate)|Status=$(Convert-ExplorerSafeString $gpu.Status)" 'INFO'
    }

    foreach ($driver in ($storage | Select-Object -First 50)) {
        Write-Log "EXPLORER_RELATED_DRIVER|Device=$(Convert-ExplorerSafeString $driver.DeviceName)|Class=$(Convert-ExplorerSafeString $driver.DeviceClass)|Provider=$(Convert-ExplorerSafeString $driver.DriverProviderName)|Version=$(Convert-ExplorerSafeString $driver.DriverVersion)|Date=$(Convert-ExplorerSafeString $driver.DriverDate)" 'INFO'
    }

    Add-DetailedResult -Step 'ExplorerRelatedDriverInventory' -Status 'Info' -Message 'Collected GPU/storage driver inventory.' -Data @{
        VideoControllerCount = @($video).Count
        RelatedDriverCount = @($storage).Count
    }
}

function Invoke-ExplorerHealthDiagnostics {
    Write-Log 'Starting Explorer health diagnostics and safe repairs...' 'INFO'

    $events = @(Get-ExplorerEventSummary -DaysBack 14)
    $shellExtensions = @(Get-NonMicrosoftShellExtensionInventory)
    $networkProblems = @(Test-ExplorerNetworkPaths)
    Get-ExplorerRelatedDriverInventory

    if ($script:Summary.ExplorerHangCount -ge 3 -or $script:Summary.ExplorerCrashCount -ge 3 -or @($networkProblems).Count -gt 0) {
        $script:Summary.ExplorerInstabilityDetected = $true
        Write-Log "EXPLORER_INSTABILITY_DETECTED|Hangs=$($script:Summary.ExplorerHangCount)|Crashes=$($script:Summary.ExplorerCrashCount)|BrokenNetworkPaths=$($script:Summary.BrokenNetworkPathCount)|NonMicrosoftShellExtensions=$($script:Summary.NonMicrosoftShellExtensionCount)" 'WARN'
        Add-Note "Explorer instability detected. Review EXPLORER_* log markers, non-Microsoft shell extensions, broken network paths, GPU/storage drivers, and user profile-specific cache/index issues."
    }
    else {
        Write-Log 'No repeated Explorer instability pattern detected from recent Application log events.' 'OK'
    }

    if ($AllowExplorerHealthRepair) {
        Invoke-ExplorerCacheCleanup
        Invoke-SearchServiceRestart
    }
    else {
        Write-Log 'Explorer cache cleanup and Windows Search restart skipped because AllowExplorerHealthRepair is disabled.' 'INFO'
    }

    Add-DetailedResult -Step 'ExplorerHealthDiagnostics' -Status 'Info' -Message 'Explorer health diagnostics completed.' -Data @{
        EventsAnalyzed = @($events).Count
        NonMicrosoftShellExtensionCount = @($shellExtensions).Count
        BrokenNetworkPathCount = @($networkProblems).Count
        RepairsAllowed = $AllowExplorerHealthRepair
    }
}

function Invoke-IconCacheRebuild {
    Write-Log "Rebuilding icon and thumbnail caches..." 'WARN'

    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    $explorerCachePath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
    $deletedFiles = New-Object System.Collections.Generic.List[string]

    $singleFileTargets = @(
        "$env:LOCALAPPDATA\IconCache.db"
    )

    foreach ($path in $singleFileTargets) {
        if (Test-Path -LiteralPath $path) {
            try {
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                $deletedFiles.Add((Split-Path -Leaf $path)) | Out-Null
            }
            catch {
                Write-Log "Failed to delete cache file $path : $($_.Exception.Message)" 'WARN'
            }
        }
    }

    if (Test-Path -LiteralPath $explorerCachePath) {
        $patterns = @(
            'iconcache*',
            'thumbcache_*.db',
            'thumbcache_idx.db'
        )

        foreach ($pattern in $patterns) {
            Get-ChildItem -LiteralPath $explorerCachePath -Filter $pattern -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                    $deletedFiles.Add($_.Name) | Out-Null
                }
                catch {
                    Write-Log "Failed to delete cache file $($_.FullName): $($_.Exception.Message)" 'WARN'
                }
            }
        }
    }

    Start-Process explorer.exe
    Add-RepairAttempt 'Icon and thumbnail cache rebuild'
    Add-DetailedResult -Step 'IconCacheRebuild' -Status 'Info' -Message 'Icon and thumbnail cache rebuild completed.' -Data @{
        DeletedFiles = ($deletedFiles -join '; ')
    }
}


function Set-RegistryDWORDValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
}

function Disable-CopilotForLoadedUsers {
    $targetSids = New-Object System.Collections.Generic.List[string]
    $targetSids.Add('HKEY_CURRENT_USER') | Out-Null

    Get-ChildItem Registry::HKEY_USERS -ErrorAction SilentlyContinue |
        Where-Object {
            $_.PSChildName -match '^S-1-5-21-' -and
            $_.PSChildName -notmatch '_Classes$'
        } |
        ForEach-Object {
            $targetSids.Add("HKEY_USERS\\$($_.PSChildName)") | Out-Null
        }

    foreach ($root in $targetSids | Select-Object -Unique) {
        $policyPath = "Registry::$root\Software\Policies\Microsoft\Windows\WindowsCopilot"
        $explorerPath = "Registry::$root\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        Set-RegistryDWORDValue -Path $policyPath -Name 'TurnOffWindowsCopilot' -Value 1
        Set-RegistryDWORDValue -Path $explorerPath -Name 'ShowCopilotButton' -Value 0

        Add-DetailedResult -Step 'CopilotDisableRegistry' -Status 'Info' -Message 'Applied Copilot disable settings for loaded profile.' -Data @{
            Root = $root
            PolicyPath = $policyPath
            ExplorerPath = $explorerPath
        }
    }
}

function Disable-CopilotForDefaultUser {
    $defaultHiveName = 'HKU\DefaultTempCopilot'
    $defaultHivePsPath = 'Registry::HKEY_USERS\DefaultTempCopilot'
    $defaultUserNtUserDat = 'C:\Users\Default\NTUSER.DAT'

    if (-not (Test-Path -LiteralPath $defaultUserNtUserDat)) {
        Write-Log 'Default User NTUSER.DAT not found; future new users were not updated for Copilot disable.' 'WARN'
        return
    }

    $hiveLoaded = $false
    try {
        $loadResult = & reg.exe load $defaultHiveName $defaultUserNtUserDat
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to load Default User hive: $($loadResult -join ' ')"
        }

        $hiveLoaded = $true
        Start-Sleep -Milliseconds 750

        $policyPath = "$defaultHivePsPath\Software\Policies\Microsoft\Windows\WindowsCopilot"
        $explorerPath = "$defaultHivePsPath\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        Set-RegistryDWORDValue -Path $policyPath -Name 'TurnOffWindowsCopilot' -Value 1
        Set-RegistryDWORDValue -Path $explorerPath -Name 'ShowCopilotButton' -Value 0

        Add-DetailedResult -Step 'CopilotDisableDefaultUser' -Status 'Info' -Message 'Applied Copilot disable settings for Default User profile.' -Data @{
            PolicyPath = $policyPath
            ExplorerPath = $explorerPath
        }
    }
    finally {
        if ($hiveLoaded) {
            Start-Sleep -Milliseconds 750
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 750
            & reg.exe unload $defaultHiveName | Out-Null
        }
    }
}

function Invoke-CopilotDisableAndRemoval {
    Write-Log 'Disabling Microsoft Copilot for current, loaded, and future user profiles...' 'WARN'

    Set-RegistryDWORDValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' -Name 'TurnOffWindowsCopilot' -Value 1
    Set-RegistryDWORDValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer' -Name 'HideCopilotButton' -Value 1

    Disable-CopilotForLoadedUsers
    Disable-CopilotForDefaultUser

    $removedPackages = New-Object System.Collections.Generic.List[string]
    $packagePatterns = @(
        'Microsoft.Windows.Copilot',
        '*Copilot*'
    )

    foreach ($pattern in $packagePatterns) {
        $packages = @(Get-AppxPackage -AllUsers -Name $pattern -ErrorAction SilentlyContinue)
        foreach ($pkg in $packages) {
            if ($removedPackages -contains $pkg.PackageFullName) {
                continue
            }

            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
                $removedPackages.Add($pkg.PackageFullName) | Out-Null
                Write-Log "Removed Copilot Appx package: $($pkg.Name) [$($pkg.PackageFullName)]" 'OK'
            }
            catch {
                Write-Log "Failed to remove Copilot Appx package $($pkg.PackageFullName): $($_.Exception.Message)" 'WARN'
            }
        }

        $provisionedPackages = @(Get-AppxProvisionedPackage -Online | Where-Object {
            $_.DisplayName -like $pattern -or $_.PackageName -like $pattern
        })

        foreach ($prov in $provisionedPackages) {
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
                $removedPackages.Add($prov.PackageName) | Out-Null
                Write-Log "Removed provisioned Copilot package: $($prov.DisplayName) [$($prov.PackageName)]" 'OK'
            }
            catch {
                Write-Log "Failed to remove provisioned Copilot package $($prov.PackageName): $($_.Exception.Message)" 'WARN'
            }
        }
    }

    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer.exe

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Microsoft Copilot disable and removal'
    Add-DetailedResult -Step 'CopilotDisableAndRemoval' -Status 'Info' -Message 'Microsoft Copilot disable and removal routine completed.' -Data @{
        RemovedPackages = ($removedPackages | Select-Object -Unique) -join '; '
    }
}

function Invoke-FirewallReset {
    Write-Log "Resetting Windows Firewall to defaults..." 'WARN'
    netsh advfirewall reset | Out-Null
    Add-RepairAttempt 'Firewall reset'
    Add-DetailedResult -Step 'FirewallReset' -Status 'Info' -Message 'Firewall reset completed.'
}

function Get-ServiceStateSafe {
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop
        return [PSCustomObject]@{
            Name        = $svc.Name
            DisplayName = $svc.DisplayName
            Status      = [string]$svc.Status
            Exists      = $true
            Error       = $null
        }
    }
    catch {
        return [PSCustomObject]@{
            Name        = $Name
            DisplayName = $null
            Status      = 'NotFound'
            Exists      = $false
            Error       = $_.Exception.Message
        }
    }
}

function Wait-ServiceStateSafe {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('Running','Stopped')][string]$DesiredStatus,
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        $state = Get-ServiceStateSafe -Name $Name
        if (-not $state.Exists) {
            return $state
        }

        if ($state.Status -eq $DesiredStatus) {
            return $state
        }

        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $deadline)

    return Get-ServiceStateSafe -Name $Name
}

function Get-ServiceProcessIdSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $cimService = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction Stop
        if ($null -ne $cimService -and $cimService.ProcessId -gt 0) {
            return [int]$cimService.ProcessId
        }
    }
    catch {
        Write-Log "Unable to read process ID for service $Name`: $($_.Exception.Message)" 'WARN'
    }

    return 0
}

function Stop-ServiceWithValidation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$TimeoutSeconds = 30,
        [switch]$KillProcessOnTimeout
    )

    $before = Get-ServiceStateSafe -Name $Name
    $killedProcess = $false
    $processId = 0
    $killMessage = $null

    if (-not $before.Exists) {
        Write-Log "Service validation: $Name was not found. Skipping stop." 'WARN'
        return [PSCustomObject]@{
            Name          = $Name
            BeforeStatus  = $before.Status
            AfterStatus   = $before.Status
            Success       = $true
            ProcessId     = 0
            ProcessKilled = $false
            Message       = 'Service not found; skipped'
        }
    }

    Write-Log "Service validation: $Name current state is $($before.Status)." 'INFO'

    if ($before.Status -eq 'Stopped') {
        Write-Log "Service validation: $Name is already stopped." 'OK'
        return [PSCustomObject]@{
            Name          = $Name
            BeforeStatus  = $before.Status
            AfterStatus   = 'Stopped'
            Success       = $true
            ProcessId     = 0
            ProcessKilled = $false
            Message       = 'Already stopped'
        }
    }

    $processId = Get-ServiceProcessIdSafe -Name $Name

    # Stop-Service can itself block indefinitely while a service remains in StopPending.
    # Send the stop request through sc.exe in a separate process so the request has its
    # own hard timeout before the service-state timeout and process-kill fallback begin.
    $stopRequestTimeoutSeconds = [math]::Min([math]::Max(5, $TimeoutSeconds), 15)
    $scProcess = $null

    try {
        Write-Log "Requesting stop for service $Name. Stop-request timeout: $stopRequestTimeoutSeconds second(s); service-state timeout: $TimeoutSeconds second(s)..." 'INFO'

        $scProcess = Start-Process -FilePath "$env:SystemRoot\System32\sc.exe" `
            -ArgumentList @('stop', $Name) `
            -WindowStyle Hidden `
            -PassThru `
            -ErrorAction Stop

        if (-not $scProcess.WaitForExit($stopRequestTimeoutSeconds * 1000)) {
            Write-Log "The stop request for service $Name exceeded $stopRequestTimeoutSeconds second(s). Terminating the hung sc.exe request process and continuing with validation." 'WARN'
            try {
                $scProcess.Kill()
                $scProcess.WaitForExit(5000) | Out-Null
            }
            catch {
                Write-Log "Unable to terminate the hung sc.exe request process for $Name`: $($_.Exception.Message)" 'WARN'
            }
        }
        elseif ($scProcess.ExitCode -notin @(0, 1062)) {
            Write-Log "sc.exe returned exit code $($scProcess.ExitCode) while requesting that $Name stop. Service-state validation will continue." 'WARN'
        }
    }
    catch {
        Write-Log "The timed stop request reported an issue for $Name`: $($_.Exception.Message). Service-state validation will continue." 'WARN'
    }
    finally {
        if ($null -ne $scProcess) {
            $scProcess.Dispose()
        }
    }

    $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Stopped' -TimeoutSeconds $TimeoutSeconds

    if ($after.Status -ne 'Stopped' -and $KillProcessOnTimeout) {
        if ($processId -le 0) {
            $processId = Get-ServiceProcessIdSafe -Name $Name
        }

        if ($processId -gt 0) {
            try {
                Write-Log "Service $Name did not stop within $TimeoutSeconds second(s). Killing backing process PID $processId..." 'WARN'
                Stop-Process -Id $processId -Force -ErrorAction Stop
                $killedProcess = $true
                Start-Sleep -Seconds 2
                $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Stopped' -TimeoutSeconds 10
                $killMessage = "Killed backing process PID $processId after timeout"
            }
            catch {
                $killMessage = "Failed to kill backing process PID $processId`: $($_.Exception.Message)"
                Write-Log $killMessage 'ERROR'
            }
        }
        else {
            $killMessage = 'No valid backing process ID was available to kill after timeout'
            Write-Log "Service $Name did not stop within $TimeoutSeconds second(s), but no valid backing process ID was found." 'WARN'
        }
    }

    $success = ($after.Status -eq 'Stopped')

    if ($success) {
        if ($killedProcess) {
            Write-Log "Service validation: $Name stopped after killing PID $processId." 'OK'
        }
        else {
            Write-Log "Service validation: $Name stopped successfully." 'OK'
        }
    }
    else {
        Write-Log "Service validation: $Name did not stop. Current state: $($after.Status)." 'WARN'
    }

    $message = if ($success -and $killedProcess) {
        $killMessage
    }
    elseif ($success) {
        'Stopped successfully'
    }
    elseif ($killMessage) {
        "Expected Stopped but found $($after.Status). $killMessage"
    }
    else {
        "Expected Stopped but found $($after.Status)"
    }

    return [PSCustomObject]@{
        Name          = $Name
        BeforeStatus  = $before.Status
        AfterStatus   = $after.Status
        Success       = $success
        ProcessId     = $processId
        ProcessKilled = $killedProcess
        Message       = $message
    }
}

function Start-ServiceWithValidation {
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$TimeoutSeconds = 30
    )

    $before = Get-ServiceStateSafe -Name $Name

    if (-not $before.Exists) {
        Write-Log "Service validation: $Name was not found. Skipping start." 'WARN'
        return [PSCustomObject]@{
            Name         = $Name
            BeforeStatus = $before.Status
            AfterStatus  = $before.Status
            Success      = $true
            Message      = 'Service not found; skipped'
        }
    }

    if ($before.Status -eq 'Running') {
        Write-Log "Service validation: $Name is already running." 'OK'
        return [PSCustomObject]@{
            Name         = $Name
            BeforeStatus = $before.Status
            AfterStatus  = 'Running'
            Success      = $true
            Message      = 'Already running'
        }
    }

    try {
        Write-Log "Starting service $Name..." 'INFO'
        Start-Service -Name $Name -ErrorAction Stop
    }
    catch {
        Write-Log "Start-Service reported an issue for $Name`: $($_.Exception.Message)" 'WARN'
    }

    $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Running' -TimeoutSeconds $TimeoutSeconds
    $success = ($after.Status -eq 'Running')

    if ($success) {
        Write-Log "Service validation: $Name started successfully." 'OK'
    }
    else {
        Write-Log "Service validation: $Name did not start. Current state: $($after.Status)." 'WARN'
    }

    return [PSCustomObject]@{
        Name         = $Name
        BeforeStatus = $before.Status
        AfterStatus  = $after.Status
        Success      = $success
        Message      = if ($success) { 'Started successfully' } else { "Expected Running but found $($after.Status)" }
    }
}


function Remove-PathWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Description,
        [int]$MaxAttempts = 5,
        [int]$InitialDelaySeconds = 2,
        [switch]$ReleaseWindowsUpdateLocks
    )

    $result = [ordered]@{
        Path            = $Path
        Description     = $Description
        ExistsBefore    = $false
        Deleted         = $false
        Attempts        = 0
        ItemCount       = 0
        FileCount       = 0
        FolderCount     = 0
        SizeBytesBefore = [int64]0
        SizeMBBefore    = [double]0
        SizeGBBefore    = [double]0
        SizeBytesAfter  = [int64]0
        SizeMBAfter     = [double]0
        SizeGBAfter     = [double]0
        SpaceFreed      = [int64]0
        EstimatedFreedBytes = [int64]0
        EstimatedFreedMB = [double]0
        EstimatedFreedGB = [double]0
        Message         = $null
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $result.Message = 'Path does not exist; nothing to delete'
        Write-Log "$Description does not exist at $Path. Nothing to delete." 'INFO'
        return [PSCustomObject]$result
    }

    $result.ExistsBefore = $true

    $beforeInfo = Get-FolderSizeInfo -Path $Path
    $result.ItemCount = $beforeInfo.ItemCount
    $result.FileCount = $beforeInfo.FileCount
    $result.FolderCount = $beforeInfo.FolderCount
    $result.SizeBytesBefore = [int64]$beforeInfo.SizeBytes
    $result.SizeMBBefore = [double]$beforeInfo.SizeMB
    $result.SizeGBBefore = [double]$beforeInfo.SizeGB

    Write-Log "Preparing to delete $Description at $Path. No backup will be created." 'INFO'
    Write-Log "Size before deletion for $Description`: $($result.SizeGBBefore) GB ($($result.SizeMBBefore) MB), items: $($result.ItemCount), files: $($result.FileCount), folders: $($result.FolderCount)." 'INFO'

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $result.Attempts = $attempt

        try {
            if (-not (Test-Path -LiteralPath $Path)) {
                $result.Deleted = $true
                $result.Message = 'Path already gone during retry validation'
                Write-Log "$Description no longer exists at $Path." 'OK'
                break
            }

            if ($ReleaseWindowsUpdateLocks -and $attempt -gt 1) {
                Write-Log "Attempt $attempt is releasing possible Windows Update locks before retrying $Description deletion." 'WARN'
                Stop-WindowsUpdateLockingProcesses | Out-Null
            }

            Write-Log "Deletion attempt $attempt of $MaxAttempts for $Description..." 'INFO'
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop

            Start-Sleep -Seconds 1
            if (-not (Test-Path -LiteralPath $Path)) {
                $result.Deleted = $true
                $result.Message = 'Deleted successfully'
                Write-Log "Deleted $Description successfully on attempt $attempt." 'OK'
                break
            }

            throw "$Description still exists after Remove-Item completed."
        }
        catch {
            $result.Message = $_.Exception.Message
            Write-Log "Deletion attempt $attempt failed for $Description`: $($result.Message)" 'WARN'

            if ($ReleaseWindowsUpdateLocks) {
                Stop-WindowsUpdateLockingProcesses | Out-Null
            }

            if ($attempt -lt $MaxAttempts) {
                $delay = $InitialDelaySeconds * $attempt
                Write-Log "Waiting $delay second(s), then retrying $Description deletion. Files may still be locked by services or Windows Update processes." 'INFO'
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                Start-Sleep -Seconds $delay
            }
        }
    }

    if (Test-Path -LiteralPath $Path) {
        $afterInfo = Get-FolderSizeInfo -Path $Path
        $result.SizeBytesAfter = [int64]$afterInfo.SizeBytes
        $result.SizeMBAfter = [double]$afterInfo.SizeMB
        $result.SizeGBAfter = [double]$afterInfo.SizeGB
    }

    $freedBytes = [int64]([math]::Max(0, ([int64]$result.SizeBytesBefore - [int64]$result.SizeBytesAfter)))
    $result.SpaceFreed = $freedBytes
    $result.EstimatedFreedBytes = $freedBytes
    $result.EstimatedFreedMB = [math]::Round(([double]$freedBytes / 1MB), 2)
    $result.EstimatedFreedGB = [math]::Round(([double]$freedBytes / 1GB), 3)

    if ($result.Deleted) {
        Write-Log "Estimated space freed by deleting $Description`: $($result.EstimatedFreedGB) GB ($($result.EstimatedFreedMB) MB)." 'OK'
    }
    else {
        Write-Log "Failed to delete $Description after $MaxAttempts attempt(s). Last error: $($result.Message). Estimated remaining size: $($result.SizeGBAfter) GB." 'ERROR'
    }

    return [PSCustomObject]$result
}

function Invoke-WindowsUpdateComponentReset {
    Write-Log "Resetting Windows Update components..." 'WARN'
    Write-Log "SoftwareDistribution will be deleted directly. No .bak, .bak1, or timestamped backup folder will be created." 'INFO'

    $services = @('wuauserv','bits','cryptsvc','msiserver','usosvc','DoSvc','WaaSMedicSvc')
    $stopResults = New-Object System.Collections.Generic.List[object]
    $startResults = New-Object System.Collections.Generic.List[object]
    $deleteResults = New-Object System.Collections.Generic.List[object]
    $lockProcessResults = New-Object System.Collections.Generic.List[object]
    $mainCleanupResult = $null
    $backupCleanupResult = $null

    foreach ($svc in $services) {
        $stopResults.Add((Stop-ServiceWithValidation -Name $svc -TimeoutSeconds 45 -KillProcessOnTimeout)) | Out-Null
    }

    $criticalServicesStillRunning = @($stopResults | Where-Object {
        $_.Name -in @('wuauserv','bits','cryptsvc') -and $_.Success -eq $false
    })

    if ($criticalServicesStillRunning.Count -gt 0) {
        Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more Windows Update services did not stop cleanly: $($criticalServicesStillRunning.Name -join ', ')"
    }

    Start-Sleep -Seconds 2

    foreach ($lockResult in @(Stop-WindowsUpdateLockingProcesses)) { $lockProcessResults.Add($lockResult) | Out-Null }

    $mainCleanupResult = Remove-SoftwareDistributionMainFolder
    $backupCleanupResult = Remove-SoftwareDistributionBakFolders

    $paths = @(
        @{ Path = "$env:WINDIR\System32\catroot2"; Description = 'Windows Update Catroot2 folder' }
    )

    foreach ($target in $paths) {
        $deleteResults.Add((Remove-PathWithRetry -Path $target.Path -Description $target.Description -MaxAttempts 5 -InitialDelaySeconds 2 -ReleaseWindowsUpdateLocks)) | Out-Null
    }

    foreach ($svc in $services) {
        $startResults.Add((Start-ServiceWithValidation -Name $svc -TimeoutSeconds 45)) | Out-Null
    }

    $failedDeletes = @($deleteResults | Where-Object { $_.ExistsBefore -eq $true -and $_.Deleted -eq $false })
    $failedStarts = @($startResults | Where-Object { $_.Success -eq $false })

    if ($failedDeletes.Count -gt 0) {
        Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more update folders could not be deleted: $($failedDeletes.Description -join ', ')"
    }

    if ($failedStarts.Count -gt 0) {
        Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more update services did not restart cleanly: $($failedStarts.Name -join ', ')"
    }

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Windows Update component reset with robocopy mirror SoftwareDistribution cleanup, SoftwareDistribution.bak* cleanup, folder size logging, lock release, and service validation'
    Add-DetailedResult -Step 'WindowsUpdateComponentReset' -Status 'Info' -Message 'Windows Update components reset. SoftwareDistribution and SoftwareDistribution.bak* folders were cleaned using robocopy mirror cleanup; estimated GB freed was logged.' -Data @{
        StoppedServicesJson = ($stopResults | ConvertTo-Json -Compress)
        LockProcessesJson  = ($lockProcessResults | ConvertTo-Json -Compress)
        MainSoftwareDistributionCleanupJson = ($mainCleanupResult | ConvertTo-Json -Compress)
        BackupCleanupJson  = ($backupCleanupResult | ConvertTo-Json -Compress)
        DeletedPathsJson   = ($deleteResults | ConvertTo-Json -Compress)
        StartedServicesJson = ($startResults | ConvertTo-Json -Compress)
    }
}

function Invoke-ScheduledTaskHealthCheck {
    Write-Log "Checking scheduled task health..." 'INFO'

    $badTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
        $_.State -eq 'Unknown' -or $_.TaskPath -eq $null
    }

    $taskNames = @()

    if ($badTasks) {
        foreach ($task in $badTasks) {
            $name = "$($task.TaskPath)$($task.TaskName)"
            $taskNames += $name
            Warn-Step -Name 'ScheduledTaskCheck' -Reason "Task may need review: $name"
        }
    }

    Add-DetailedResult -Step 'ScheduledTaskHealthCheck' -Status 'Info' -Message 'Scheduled task health check completed.' -Data @{
        SuspectTaskCount = $taskNames.Count
        SuspectTasks     = ($taskNames -join '; ')
    }
}

function Invoke-EventLogSummary {
    Write-Log "Collecting recent system health events..." 'INFO'

    try {
        $recentUnexpected = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Id      = 41, 6008
            StartTime = (Get-Date).AddDays(-7)
        } -ErrorAction Stop

        $count = @($recentUnexpected).Count

        Add-DetailedResult -Step 'EventLogSummary' -Status 'Info' -Message 'Collected recent unexpected shutdown events.' -Data @{
            UnexpectedShutdownCount = $count
        }

        if ($recentUnexpected) {
            Warn-Step -Name 'EventLogSummary' -Reason "Recent unexpected shutdown events found: $count"
        }
    }
    catch {
        $eventLogError = $_.Exception.Message
        Warn-Step -Name 'EventLogSummary' -Reason $eventLogError

        if ($eventLogError -match 'RPC server is unavailable|0x800706BA|1722') {
            $rpcStatus = if ($script:Summary.RpcDiagnosticsStatus) { $script:Summary.RpcDiagnosticsStatus } else { 'RpcDiagnosticsNotRunOrNoStatusAvailable' }
            $clientSide = if ($null -ne $script:Summary.RpcClientSideIssueDetected) { $script:Summary.RpcClientSideIssueDetected } else { $false }
            $remoteSide = if ($null -ne $script:Summary.RpcRemoteSideIssueDetected) { $script:Summary.RpcRemoteSideIssueDetected } else { $false }

            $classification = if ($clientSide) {
                'LOCAL_CLIENT_OR_LOCAL_SERVICE_PROBLEM'
            }
            elseif ($remoteSide) {
                'REMOTE_ENDPOINT_OR_NETWORK_PROBLEM'
            }
            elseif ($rpcStatus -eq 'NoRpcIssueDetectedByBasicChecks') {
                'EVENTLOG_PROVIDER_OR_QUERY_SPECIFIC_PROBLEM'
            }
            else {
                'UNCLASSIFIED_RPC_FAILURE_REVIEW_RPC_DIAGNOSTICS'
            }

            Write-Log "RPC_EVENTLOGSUMMARY_CLASSIFICATION|Classification=$classification|RpcDiagnosticsStatus=$rpcStatus|ClientSideIssue=$clientSide|RemoteOrNetworkIssue=$remoteSide|OriginalError=$(Convert-RpcSafeString $eventLogError)" 'WARN'
            Add-Note "EventLogSummary hit RPC unavailable. Classification: $classification. RPC diagnostics status: $rpcStatus."
        }
    }
}

function Invoke-EventLogClear {
    Write-Log "Clearing classic event logs..." 'WARN'
    wevtutil el | ForEach-Object {
        try {
            wevtutil cl $_ 2>$null
        }
        catch {
        }
    }
    Add-RepairAttempt 'Event log clear'
    Add-DetailedResult -Step 'EventLogClear' -Status 'Info' -Message 'Event log clear attempted.'
}

function Invoke-DiskSpaceCheck {
    $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
    if ($drive) {
        $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
        $sizeGB = [math]::Round($drive.Size / 1GB, 2)
        Write-Log "System drive free space: $freeGB GB of $sizeGB GB" 'INFO'

        Add-DetailedResult -Step 'DiskSpaceCheck' -Status 'Info' -Message 'Disk space checked.' -Data @{
            FreeGB  = $freeGB
            TotalGB = $sizeGB
        }

        if ($freeGB -lt 10) {
            Warn-Step -Name 'DiskSpaceCheck' -Reason "Low free space on system drive: $freeGB GB"
        }
    }
}

function Invoke-DefenderStatusCheck {
    try {
        $status = Get-MpComputerStatus -ErrorAction Stop
        Write-Log "Defender Antivirus Enabled: $($status.AntivirusEnabled)" 'INFO'
        Write-Log "Defender RealTime Protection Enabled: $($status.RealTimeProtectionEnabled)" 'INFO'

        Add-DetailedResult -Step 'DefenderStatusCheck' -Status 'Info' -Message 'Defender status checked.' -Data @{
            AntivirusEnabled          = $status.AntivirusEnabled
            RealTimeProtectionEnabled = $status.RealTimeProtectionEnabled
        }
    }
    catch {
        Warn-Step -Name 'DefenderStatusCheck' -Reason $_.Exception.Message
    }
}

function Invoke-RebootIfNeeded {
    param(
        [int]$DelaySeconds = 60
    )

    $comment = 'Restarting after automated system repair operations.'

    $args = @(
        '/r'
        '/t', $DelaySeconds.ToString()
        '/d', 'p:2:17'
        '/c', "`"$comment`""
        '/f'
    )

    Write-Log "Issuing reboot command: shutdown.exe $($args -join ' ')" 'WARN'
    & "$env:SystemRoot\System32\shutdown.exe" @args

    if ($LASTEXITCODE -ne 0) {
        throw "shutdown.exe returned exit code $LASTEXITCODE"
    }

    $script:RebootInitiated = $true
    Add-DetailedResult -Step 'AutoReboot' -Status 'Info' -Message 'Automatic reboot command issued.' -Data @{
        DelaySeconds = $DelaySeconds
    }
}

function Show-Summary {
    $script:Summary.EndTime = Get-Date

    Write-Log "---------------- Summary ----------------" 'INFO'
    Write-Log "Computer Name: $($script:Summary.ComputerName)" 'INFO'
    Write-Log "Start Time: $($script:Summary.StartTime)" 'INFO'
    Write-Log "End Time:   $($script:Summary.EndTime)" 'INFO'
    Write-Log "YAML Log:   $($script:YamlLogPath)" 'INFO'
    Write-Log "Succeeded:  $($script:Summary.StepsSucceeded)" 'INFO'
    Write-Log "Failed:     $($script:Summary.StepsFailed)" 'INFO'
    Write-Log "Warnings:   $($script:Summary.Warnings)" 'INFO'
    Write-Log "Pending Reboot Detected: $($script:Summary.PendingRebootDetected)" 'INFO'
    Write-Log "Reboot Required: $($script:Summary.RebootRequired)" 'INFO'
    Write-Log "Disk Corruption Suspected: $($script:Summary.DiskCorruptionSuspected)" 'INFO'
    Write-Log "DISM Corruption Detected: $($script:Summary.DismCorruptionDetected)" 'INFO'
    Write-Log "SFC Integrity Violations: $($script:Summary.SfcIntegrityViolations)" 'INFO'
    Write-Log "SFC Repair Attempted: $($script:Summary.SfcRepairAttempted)" 'INFO'
    Write-Log "SFC Repair Succeeded: $($script:Summary.SfcRepairSucceeded)" 'INFO'
    Write-Log "CBS Corrupt File Count: $($script:Summary.CbsCorruptFileCount)" 'INFO'
    Write-Log "CBS Corruption Extract: $($script:Summary.CbsCorruptionExtractPath)" 'INFO'
    Write-Log "CBS Driver Corruption Detected: $($script:Summary.CbsDriverCorruptionDetected)" 'INFO'
    Write-Log "CBS Driver Corruption Remaining: $($script:Summary.CbsDriverCorruptionRemaining)" 'INFO'
    Write-Log "CBS Corrupt Driver Files: $((($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ','))" 'INFO'
    Write-Log "HPIA Driver Repair Attempted: $($script:Summary.HpDriverRepairAttempted)" 'INFO'
    Write-Log "HPIA Driver Repair Succeeded: $($script:Summary.HpDriverRepairSucceeded)" 'INFO'
    Write-Log "HPIA Path: $($script:Summary.HpImageAssistantPath)" 'INFO'
    Write-Log "WMI Repository Inconsistent: $($script:Summary.WmiRepositoryInconsistent)" 'INFO'
    Write-Log "Storage Health Warnings: $($script:Summary.StorageHealthWarnings)" 'INFO'
    Write-Log "Storage Failure Predicted: $($script:Summary.StorageFailurePredicted)" 'INFO'
    Write-Log "Storage Wear Over 50 Percent Detected: $($script:Summary.StorageWearOver50Detected)" 'INFO'
    Write-Log "Explorer Hang Count: $($script:Summary.ExplorerHangCount)" 'INFO'
    Write-Log "Explorer Crash Count: $($script:Summary.ExplorerCrashCount)" 'INFO'
    Write-Log "Explorer Instability Detected: $($script:Summary.ExplorerInstabilityDetected)" 'INFO'
    Write-Log "Non-Microsoft Shell Extension Count: $($script:Summary.NonMicrosoftShellExtensionCount)" 'INFO'
    Write-Log "Broken Network Path Count: $($script:Summary.BrokenNetworkPathCount)" 'INFO'
    Write-Log "RPC Client-Side Issue Detected: $($script:Summary.RpcClientSideIssueDetected)" 'INFO'
    Write-Log "RPC Remote/Network Issue Detected: $($script:Summary.RpcRemoteSideIssueDetected)" 'INFO'
    Write-Log "RPC Diagnostics Status: $($script:Summary.RpcDiagnosticsStatus)" 'INFO'
    Write-Log "RPC Classification Hint: Review RPC_ROOT_CAUSE_ASSESSMENT and RPC_EVENTLOGSUMMARY_CLASSIFICATION markers above for local-vs-remote determination." 'INFO'

    if ($script:Summary.RepairsAttempted.Count -gt 0) {
        Write-Log "Repairs Attempted:" 'INFO'
        foreach ($repair in $script:Summary.RepairsAttempted) {
            Write-Log " - $repair" 'INFO'
        }
    }

    if ($script:Summary.Notes.Count -gt 0) {
        Write-Log "Notes:" 'INFO'
        foreach ($note in $script:Summary.Notes) {
            Write-Log " - $note" 'INFO'
        }
    }
}



function Invoke-RollbackRetentionCleanup {
    [CmdletBinding()]
    param(
        [string]$RollbackRoot = 'C:\Scripts\Rollback',
        [ValidateRange(1,365)]
        [int]$RetentionDays = 7
    )

    Write-Log "Checking updater rollback folders in $RollbackRoot for items older than $RetentionDays day(s)..." 'INFO'

    if (-not (Test-Path -LiteralPath $RollbackRoot -PathType Container)) {
        Write-Log "Rollback directory does not exist: $RollbackRoot" 'INFO'
        Add-DetailedResult -Step 'RollbackRetentionCleanup' -Status 'Info' -Message 'Rollback directory does not exist.' -Data @{
            RollbackRoot  = $RollbackRoot
            RetentionDays = $RetentionDays
            DeletedCount  = 0
            FailedCount   = 0
        }
        return
    }

    $cutoff = (Get-Date).AddDays(-$RetentionDays)

    $oldFolders = @(
        Get-ChildItem -LiteralPath $RollbackRoot -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        Sort-Object LastWriteTime
    )

    $deleted = @()
    $failed  = @()

    if ($oldFolders.Count -eq 0) {
        Write-Log "No rollback folders older than $RetentionDays day(s) were found." 'INFO'
    }

    foreach ($folder in $oldFolders) {
        try {
            $folderPath = $folder.FullName

            # Safety: only direct children of C:\Scripts\Rollback may be removed.
            if ($folder.Parent.FullName.TrimEnd('\') -ine $RollbackRoot.TrimEnd('\')) {
                throw "Safety check blocked rollback cleanup outside the expected root: $folderPath"
            }

            Write-Log "Deleting rollback folder older than $RetentionDays day(s): $folderPath | LastWriteTime=$($folder.LastWriteTime)" 'INFO'
            Remove-Item -LiteralPath $folderPath -Recurse -Force -ErrorAction Stop

            if (Test-Path -LiteralPath $folderPath) {
                throw 'Folder still exists after deletion attempt.'
            }

            $deleted += $folderPath
            Write-Log "Deleted expired rollback folder: $folderPath" 'OK'
        }
        catch {
            $message = "Failed to delete rollback folder $($folder.FullName): $($_.Exception.Message)"
            $failed += $message
            Write-Log $message 'WARN'
        }
    }

    Add-DetailedResult -Step 'RollbackRetentionCleanup' -Status $(if ($failed.Count -gt 0) { 'Warning' } else { 'Info' }) -Message 'Rollback retention cleanup completed.' -Data @{
        RollbackRoot   = $RollbackRoot
        RetentionDays  = $RetentionDays
        Cutoff         = $cutoff
        Candidates     = $oldFolders.Count
        DeletedCount   = $deleted.Count
        FailedCount    = $failed.Count
        DeletedFolders = ($deleted -join '; ')
        Errors         = ($failed -join '; ')
    }

    if ($failed.Count -eq 0) {
        Write-Log "Rollback retention cleanup completed successfully. Deleted: $($deleted.Count)." 'OK'
    }
}

function Remove-WindowsOldFolder {
    [CmdletBinding()]
    param(
        [string]$Path = 'C:\Windows.old'
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "$Path does not exist. No Windows.old cleanup required." 'INFO'
        Add-DetailedResult -Step 'WindowsOldCleanup' -Status 'Info' -Message 'Windows.old does not exist.' -Data @{
            Path    = $Path
            Deleted = $false
            Existed = $false
        }
        return
    }

    Write-Log "Windows.old detected. Preparing forced cleanup of $Path..." 'INFO'

    $sizeBefore = Get-FolderSizeBytesSafe -Path $Path
    $errors = @()

    try {
        # Windows.old can contain TrustedInstaller-owned content. Take ownership
        # and grant SYSTEM full control before using the native directory remover.
        $takeownOutput = & "$env:SystemRoot\System32\takeown.exe" /F $Path /A /R /D Y 2>&1
        if ($LASTEXITCODE -ne 0) {
            $errors += "takeown exit code $LASTEXITCODE"
            Write-Log "takeown returned exit code $LASTEXITCODE while preparing Windows.old cleanup." 'WARN'
        }

        $icaclsOutput = & "$env:SystemRoot\System32\icacls.exe" $Path /grant '*S-1-5-18:(OI)(CI)F' /T /C /Q 2>&1
        if ($LASTEXITCODE -ne 0) {
            $errors += "icacls exit code $LASTEXITCODE"
            Write-Log "icacls returned exit code $LASTEXITCODE while preparing Windows.old cleanup." 'WARN'
        }

        & "$env:SystemRoot\System32\attrib.exe" -R -S -H "$Path\*" /S /D 2>&1 | Out-Null

        # rd.exe is significantly more reliable than Remove-Item for very large
        # Windows.old trees and long legacy paths.
        $cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
        & $cmd /d /c "rd /s /q `"$Path`"" 2>&1 | Out-Null

        if (Test-Path -LiteralPath $Path) {
            # Final PowerShell fallback.
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        }

        $deleted = -not (Test-Path -LiteralPath $Path)

        if ($deleted) {
            Write-Log "Windows.old was deleted successfully: $Path" 'OK'
            Add-DetailedResult -Step 'WindowsOldCleanup' -Status 'Info' -Message 'Windows.old deleted successfully.' -Data @{
                Path             = $Path
                Existed          = $true
                Deleted          = $true
                SizeBeforeBytes  = $sizeBefore
                PreparationNotes = ($errors -join '; ')
            }
        }
        else {
            $message = "Windows.old still exists after forced cleanup attempts: $Path"
            Write-Log $message 'WARN'
            Add-DetailedResult -Step 'WindowsOldCleanup' -Status 'Warning' -Message $message -Data @{
                Path             = $Path
                Existed          = $true
                Deleted          = $false
                SizeBeforeBytes  = $sizeBefore
                PreparationNotes = ($errors -join '; ')
            }
        }
    }
    catch {
        $message = "Windows.old cleanup failed: $($_.Exception.Message)"
        Write-Log $message 'WARN'
        Add-DetailedResult -Step 'WindowsOldCleanup' -Status 'Warning' -Message $message -Data @{
            Path            = $Path
            Existed         = $true
            Deleted         = $false
            SizeBeforeBytes = $sizeBefore
        }
    }
}


function Invoke-LogArchiveRetention {
    [CmdletBinding()]
    param(
        [string]$LogDirectory = 'C:\Logs',
        [string]$ComputerName = $env:COMPUTERNAME
    )

    Write-Log "Starting Sunday-based log archive and retention processing in $LogDirectory" 'INFO'

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        Write-Log "Log directory does not exist: $LogDirectory" 'WARN'
        Add-DetailedResult -Step 'LogArchiveRetention' -Status 'Warning' -Message "Log directory not found: $LogDirectory"
        return
    }

    $now = Get-Date
    $thisSunday = $now.Date.AddDays(-[int]$now.DayOfWeek)
    $previousSunday = $thisSunday.AddDays(-7)
    $twoSundaysAgo = $thisSunday.AddDays(-14)

    Write-Log "This Sunday: $thisSunday" 'INFO'
    Write-Log "Previous Sunday: $previousSunday" 'INFO'
    Write-Log "Two Sundays Ago: $twoSundaysAgo" 'INFO'

    $extensions = @('.yaml', '.yml', '.txt')

    $allLooseLogs = Get-ChildItem -LiteralPath $LogDirectory -File -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $extensions -contains $_.Extension.ToLowerInvariant() -and
            $_.FullName -ne $script:YamlLogPath -and $_.FullName -ne $script:PublishedTextLogPath
        }

    $logsToArchive = $allLooseLogs | Where-Object {
        $_.CreationTime -ge $previousSunday -and $_.CreationTime -lt $thisSunday
    } | Sort-Object CreationTime, Name

    $archiveDateText = $previousSunday.ToString('yyyy-MM-dd')
    $zipPath = Join-Path $LogDirectory ("{0}-logs-{1}.zip" -f $ComputerName, $archiveDateText)

    $archiveSummary = [ordered]@{
        ThisSunday                 = $thisSunday
        PreviousSunday             = $previousSunday
        TwoSundaysAgo              = $twoSundaysAgo
        LooseLogsFound             = @($allLooseLogs).Count
        LogsSelectedForArchive     = @($logsToArchive).Count
        ArchiveCreated             = $false
        ArchivePath                = $null
        DeletedOriginalFiles       = @()
        DeletedOldLooseLogs        = @()
        DeletedExpiredZipFiles     = @()
        Errors                     = @()
    }

    if (@($logsToArchive).Count -gt 0) {
        Write-Log "Preparing archive for previous Sunday week: $zipPath" 'INFO'

        try {
            if (Test-Path -LiteralPath $zipPath) {
                Write-Log "Existing archive found for that Sunday. Removing and recreating: $zipPath" 'WARN'
                Remove-Item -LiteralPath $zipPath -Force -ErrorAction Stop
            }

            Compress-Archive -Path ($logsToArchive.FullName) -DestinationPath $zipPath -CompressionLevel Optimal -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $zipPath)) {
                throw 'ZIP file was not created.'
            }

            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
            $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)

            try {
                $zipEntries = @($zip.Entries)
                if ($zipEntries.Count -lt 1) {
                    throw 'ZIP file was created but contains no entries.'
                }

                if ($zipEntries.Count -lt @($logsToArchive).Count) {
                    throw "ZIP file entry count ($($zipEntries.Count)) is less than expected source file count ($(@($logsToArchive).Count))."
                }
            }
            finally {
                $zip.Dispose()
            }

            $archiveSummary.ArchiveCreated = $true
            $archiveSummary.ArchivePath = $zipPath
            Write-Log "Archive created successfully: $zipPath" 'OK'

            foreach ($file in $logsToArchive) {
                try {
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    $archiveSummary.DeletedOriginalFiles += $file.FullName
                    Write-Log "Deleted archived source log: $($file.FullName)" 'OK'
                }
                catch {
                    $msg = "Failed to delete archived source file $($file.FullName): $($_.Exception.Message)"
                    $archiveSummary.Errors += $msg
                    Write-Log $msg 'WARN'
                }
            }
        }
        catch {
            $msg = "Archive creation/validation failed: $($_.Exception.Message)"
            $archiveSummary.Errors += $msg
            Write-Log $msg 'ERROR'
        }
    }
    else {
        Write-Log 'No loose log files were found for the previous Sunday-to-Saturday period.' 'INFO'
    }

    $remainingLooseLogs = Get-ChildItem -LiteralPath $LogDirectory -File -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $extensions -contains $_.Extension.ToLowerInvariant() -and
            $_.FullName -ne $script:YamlLogPath -and $_.FullName -ne $script:PublishedTextLogPath
        }

    $oldLooseLogsToDelete = $remainingLooseLogs | Where-Object {
        $_.CreationTime -lt $twoSundaysAgo
    }

    foreach ($file in $oldLooseLogsToDelete) {
        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $archiveSummary.DeletedOldLooseLogs += $file.FullName
            Write-Log "Deleted loose log older than two Sundays: $($file.FullName)" 'OK'
        }
        catch {
            $msg = "Failed to delete old loose log $($file.FullName): $($_.Exception.Message)"
            $archiveSummary.Errors += $msg
            Write-Log $msg 'WARN'
        }
    }

    $zipFilesToDelete = Get-ChildItem -LiteralPath $LogDirectory -File -Filter '*.zip' -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like "$ComputerName-logs-*.zip" -and
            $_.CreationTime -lt $twoSundaysAgo
        }

    foreach ($zipFile in $zipFilesToDelete) {
        try {
            Remove-Item -LiteralPath $zipFile.FullName -Force -ErrorAction Stop
            $archiveSummary.DeletedExpiredZipFiles += $zipFile.FullName
            Write-Log "Deleted ZIP archive older than two Sundays: $($zipFile.FullName)" 'OK'
        }
        catch {
            $msg = "Failed to delete expired ZIP $($zipFile.FullName): $($_.Exception.Message)"
            $archiveSummary.Errors += $msg
            Write-Log $msg 'WARN'
        }
    }

    Add-DetailedResult -Step 'LogArchiveRetention' -Status 'Info' -Message 'Sunday-based log archive and retention processing completed.' -Data @{
        ThisSunday                  = $archiveSummary.ThisSunday
        PreviousSunday              = $archiveSummary.PreviousSunday
        TwoSundaysAgo               = $archiveSummary.TwoSundaysAgo
        LooseLogsFound              = $archiveSummary.LooseLogsFound
        LogsSelectedForArchive      = $archiveSummary.LogsSelectedForArchive
        ArchiveCreated              = $archiveSummary.ArchiveCreated
        ArchivePath                 = $archiveSummary.ArchivePath
        DeletedOriginalFilesCount   = @($archiveSummary.DeletedOriginalFiles).Count
        DeletedOldLooseLogsCount    = @($archiveSummary.DeletedOldLooseLogs).Count
        DeletedExpiredZipFilesCount = @($archiveSummary.DeletedExpiredZipFiles).Count
        ErrorsCount                 = @($archiveSummary.Errors).Count
        DeletedOriginalFiles        = ($archiveSummary.DeletedOriginalFiles -join '; ')
        DeletedOldLooseLogs         = ($archiveSummary.DeletedOldLooseLogs -join '; ')
        DeletedExpiredZipFiles      = ($archiveSummary.DeletedExpiredZipFiles -join '; ')
        Errors                      = ($archiveSummary.Errors -join '; ')
    }

    if (@($archiveSummary.Errors).Count -gt 0) {
        Warn-Step -Name 'LogArchiveRetention' -Reason ("Completed with errors: " + ($archiveSummary.Errors -join ' | '))
    }
    else {
        Write-Log 'Sunday-based log archive and retention processing completed successfully.' 'OK'
    }
}


function Get-PendingRebootReasons {
    param([AllowNull()]$State)

    $reasons = New-Object System.Collections.Generic.List[string]
    if ($null -eq $State) { return @() }

    $map = [ordered]@{
        CBServicing_RebootPending         = 'ComponentBasedServicing'
        WindowsUpdate_RebootRequired      = 'WindowsUpdate'
        SessionManager_PendingFileRename  = 'PendingFileRenameOperations'
        SessionManager_PendingFileRename2 = 'PendingFileRenameOperations2'
        UpdateExeVolatile                 = 'UpdateExeVolatile'
        PackagesPending                   = 'PackagesPending'
        WUAU_RebootRequired_COM           = 'MicrosoftUpdateSystemInfo'
    }

    foreach ($entry in $map.GetEnumerator()) {
        try {
            if ([bool]$State.($entry.Key)) { $reasons.Add([string]$entry.Value) | Out-Null }
        }
        catch {}
    }
    return @($reasons)
}

function Get-SystemDriveTelemetry {
    try {
        $drive = Get-DiskSpaceInfo -Path ($env:SystemDrive + '\')
        if ($null -eq $drive) { return $null }
        return [PSCustomObject]@{
            FreeGB  = [math]::Round(([double]$drive.FreeSpace / 1GB), 2)
            UsedGB  = [math]::Round(([double]$drive.UsedSpace / 1GB), 2)
            TotalGB = [math]::Round(([double]$drive.TotalSize / 1GB), 2)
            FreePercent = if ([double]$drive.TotalSize -gt 0) { [math]::Round((100 * [double]$drive.FreeSpace / [double]$drive.TotalSize), 2) } else { $null }
        }
    }
    catch { return $null }
}

function Get-WindowsBuildTelemetry {
    try {
        $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        return [PSCustomObject]@{
            ProductName    = [string]$cv.ProductName
            EditionId      = [string]$cv.EditionID
            DisplayVersion = [string]$cv.DisplayVersion
            ReleaseId      = [string]$cv.ReleaseId
            CurrentBuild   = [string]$cv.CurrentBuild
            UBR            = if ($null -ne $cv.UBR) { [int]$cv.UBR } else { $null }
            Build          = if ($null -ne $cv.UBR) { '{0}.{1}' -f $cv.CurrentBuild, $cv.UBR } else { [string]$cv.CurrentBuild }
        }
    }
    catch { return $null }
}

function Get-StepTelemetry {
    param([Parameter(Mandatory)][string]$StepPattern)

    $matches = @($script:DetailedResults | Where-Object { $_.Step -like $StepPattern })
    return @($matches | ForEach-Object {
        [PSCustomObject]@{
            Timestamp = $_.Timestamp
            Step      = $_.Step
            Status    = $_.Status
            Message   = $_.Message
            Data      = $_.Data
        }
    })
}


function Get-StepSummaryTelemetry {
    param([Parameter(Mandatory)][string]$StepPattern)

    [object[]]$matches = @(
        $script:DetailedResults |
        Where-Object { $_.Step -like $StepPattern }
    )

    return @(
        foreach ($item in $matches) {
            $timestampIso = $null
            try {
                if ($item.Timestamp -is [datetime]) {
                    $timestampIso = ([datetime]$item.Timestamp).ToUniversalTime().ToString('o')
                }
                elseif ($null -ne $item.Timestamp) {
                    $timestampIso = [string]$item.Timestamp
                }
            }
            catch { $timestampIso = $null }

            $dataJson = $null
            if ($null -ne $item.Data) {
                try { $dataJson = $item.Data | ConvertTo-Json -Depth 6 -Compress }
                catch {
                    try { $dataJson = [string]$item.Data } catch { $dataJson = $null }
                }
            }

            [pscustomobject]@{
                Timestamp = $timestampIso
                Step      = [string]$item.Step
                Status    = [string]$item.Status
                Message   = [string]$item.Message
                DataJson  = $dataJson
            }
        }
    )
}

function New-SystemRepairTelemetryEvent {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [AllowNull()][string]$FailureMessage
    )

    $end = Get-Date
    $domain = $null
    try { $domain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()).Name } catch { $domain = $env:USERDNSDOMAIN }

    [string[]]$repairs = @($script:Summary.RepairsAttempted | Select-Object -Unique | ForEach-Object { [string]$_ })
    [string[]]$notes = @($script:Summary.Notes | ForEach-Object { [string]$_ })

    [object[]]$allStepSummaries = @(Get-StepSummaryTelemetry -StepPattern '*')
    [object[]]$failedSteps = @($allStepSummaries | Where-Object { $_.Status -eq 'Failed' })
    [object[]]$warningSteps = @($allStepSummaries | Where-Object { $_.Status -eq 'Warning' })
    [string[]]$failedStepNames = @($failedSteps | ForEach-Object { $_.Step } | Where-Object { $_ } | Select-Object -Unique)
    [string[]]$warningStepNames = @($warningSteps | ForEach-Object { $_.Step } | Where-Object { $_ } | Select-Object -Unique)

    $freeBefore = if ($null -ne $script:DiskSpaceBefore) { $script:DiskSpaceBefore.FreeGB } else { $null }
    $freeAfter = if ($null -ne $script:DiskSpaceAfter) { $script:DiskSpaceAfter.FreeGB } else { $null }
    $reclaimed = if ($null -ne $freeBefore -and $null -ne $freeAfter) { [math]::Round(([double]$freeAfter - [double]$freeBefore), 2) } else { 0.0 }

    $pendingBefore = if ($null -ne $script:PendingRebootBefore) { [bool]$script:PendingRebootBefore.AnyPendingReboot } else { $false }
    $pendingAfter  = if ($null -ne $script:PendingRebootAfter) { [bool]$script:PendingRebootAfter.AnyPendingReboot } else { [bool]$script:Summary.PendingRebootDetected }

    [string[]]$reasonsBefore = @($script:Summary.PendingRebootReasonsBefore | ForEach-Object { [string]$_ } | Select-Object -Unique)
    [string[]]$reasonsAfter  = @($script:Summary.PendingRebootReasonsAfter | ForEach-Object { [string]$_ } | Select-Object -Unique)
    [string[]]$newRebootReasons = @($reasonsAfter | Where-Object { $_ -notin $reasonsBefore })

    $effectiveRebootRequired = [bool]($script:Summary.RebootRequired -or $pendingAfter)
    $newRebootRequirementDetected = [bool]((-not $pendingBefore -and $pendingAfter) -or $newRebootReasons.Count -gt 0)

    $dismRepairAttempted = [bool](@($repairs | Where-Object { $_ -match 'DISM RestoreHealth' }).Count -gt 0)
    $dismRepairStatus = if (-not $script:Summary.DismCorruptionDetected) {
        'NoCorruptionDetected'
    }
    elseif ($dismRepairAttempted) {
        'RepairAttempted'
    }
    elseif (-not $AutoRepairOnDetection) {
        'DetectedNotRepaired_AutoRepairDisabled'
    }
    else {
        'DetectedNotRepaired'
    }

    [string[]]$healthIssues = @()
    if ($script:Summary.DismCorruptionDetected) { $healthIssues += 'DISM component store corruption detected' }
    if ($script:Summary.SfcIntegrityViolations) { $healthIssues += 'SFC integrity violations detected' }
    if ($script:Summary.DiskCorruptionSuspected) { $healthIssues += 'Disk corruption suspected' }
    if ($script:Summary.WmiRepositoryInconsistent) { $healthIssues += 'WMI repository inconsistent' }
    if ($script:Summary.StorageFailurePredicted) { $healthIssues += 'Storage failure predicted' }
    if ($script:Summary.StorageHealthWarnings -gt 0) { $healthIssues += 'Storage health warning detected' }
    if ($script:Summary.ExplorerInstabilityDetected) { $healthIssues += 'Explorer instability detected' }
    if ($script:Summary.RpcClientSideIssueDetected) { $healthIssues += 'RPC client-side issue detected' }
    if ($script:Summary.RpcRemoteSideIssueDetected) { $healthIssues += 'RPC remote-side issue detected' }
    if ($pendingAfter) { $healthIssues += 'Pending reboot remains after system repair' }

    $storageHealthy = [bool](
        -not $script:Summary.StorageFailurePredicted -and
        $script:Summary.StorageHealthWarnings -eq 0 -and
        -not $script:Summary.StorageWearOver50Detected
    )
    $wmiHealthy = [bool](-not $script:Summary.WmiRepositoryInconsistent)
    $windowsBuild = Get-WindowsBuildTelemetry

    [PSCustomObject][ordered]@{
        '@timestamp' = $end.ToUniversalTime().ToString('o')
        event = [ordered]@{
            kind     = 'event'
            category = @('host','configuration')
            type     = @('info')
            action   = 'system-repair'
            outcome  = if ($ExitCode -eq 0 -or $ExitCode -eq 3010) { 'success' } else { 'failure' }
            duration = [int64](($end - $script:RunStart).TotalMilliseconds * 1000000)
        }
        maintenance = [ordered]@{
            execution = [ordered]@{
                RunId           = $script:RunId
                ComputerName    = $script:ComputerName
                Domain          = $domain
                ScriptName      = $script:ScriptName
                ScriptVersion   = $script:ScriptVersion
                Status          = $Status
                ExitCode        = $ExitCode
                StartTime       = $script:RunStart.ToUniversalTime().ToString('o')
                EndTime         = $end.ToUniversalTime().ToString('o')
                DurationSeconds = [math]::Round(($end - $script:RunStart).TotalSeconds, 2)
                ErrorCount      = [int]$script:Summary.StepsFailed
                WarningCount    = [int]$script:Summary.Warnings
                FailureMessage  = $FailureMessage
                TextLogPath     = if (-not [string]::IsNullOrWhiteSpace($script:PublishedTextLogPath)) {
                    $script:PublishedTextLogPath
                }
                else {
                    $script:TextLogPath
                }
            }
            system_repair = [ordered]@{
                AutoRepairOnDetection = [bool]$AutoRepairOnDetection

                RepairSummary = [ordered]@{
                    RepairsAttemptedCount = [int]$repairs.Count
                    RepairsAttempted      = $repairs
                    StepsSucceeded        = [int]$script:Summary.StepsSucceeded
                    StepsFailed           = [int]$script:Summary.StepsFailed
                    FailedStepCount       = [int]$failedSteps.Count
                    FailedStepNames       = $failedStepNames
                    WarningStepCount      = [int]$warningSteps.Count
                    WarningStepNames      = $warningStepNames
                }

                RebootSummary = [ordered]@{
                    PendingBefore                = $pendingBefore
                    PendingAfter                 = $pendingAfter
                    PreExistingPendingReboot     = $pendingBefore
                    EffectiveRebootRequired      = $effectiveRebootRequired
                    NewRebootRequirementDetected = $newRebootRequirementDetected
                    ReasonsBefore                = $reasonsBefore
                    ReasonsAfter                 = $reasonsAfter
                    NewReasonsAfter              = $newRebootReasons
                    RebootInitiated              = [bool]$script:RebootInitiated
                }

                HealthSummary = [ordered]@{
                    HealthIssueCount            = [int]$healthIssues.Count
                    HealthIssues                = $healthIssues
                    DismCorruptionDetected      = [bool]$script:Summary.DismCorruptionDetected
                    DismRepairAttempted         = $dismRepairAttempted
                    DismRepairStatus            = $dismRepairStatus
                    SfcIntegrityViolations      = [bool]$script:Summary.SfcIntegrityViolations
                    SfcRepairAttempted          = [bool]$script:Summary.SfcRepairAttempted
                    SfcRepairSucceeded          = [bool]$script:Summary.SfcRepairSucceeded
                    DiskCorruptionSuspected     = [bool]$script:Summary.DiskCorruptionSuspected
                    WmiHealthy                  = $wmiHealthy
                    StorageHealthy              = $storageHealthy
                    ExplorerInstabilityDetected = [bool]$script:Summary.ExplorerInstabilityDetected
                    RpcClientSideIssueDetected  = [bool]$script:Summary.RpcClientSideIssueDetected
                    RpcRemoteSideIssueDetected  = [bool]$script:Summary.RpcRemoteSideIssueDetected
                }

                StepsSucceeded        = [int]$script:Summary.StepsSucceeded
                StepsFailed           = [int]$script:Summary.StepsFailed
                PendingRebootBefore   = $pendingBefore
                PendingRebootReasonsBefore = $reasonsBefore
                PendingRebootAfter    = $pendingAfter
                PendingRebootReasonsAfter = $reasonsAfter
                RebootRequired        = $effectiveRebootRequired
                RebootInitiated       = [bool]$script:RebootInitiated
                DiskCorruptionSuspected = [bool]$script:Summary.DiskCorruptionSuspected
                RepairVolumeSpotFixSucceeded = [bool]$script:Summary.RepairVolumeSpotFixSucceeded
                DismCorruptionDetected = [bool]$script:Summary.DismCorruptionDetected
                SfcIntegrityViolations = [bool]$script:Summary.SfcIntegrityViolations
                SfcRepairAttempted     = [bool]$script:Summary.SfcRepairAttempted
                SfcRepairSucceeded     = [bool]$script:Summary.SfcRepairSucceeded
                CbsCorruptFileCount    = [int]$script:Summary.CbsCorruptFileCount
                CbsDriverCorruptionDetected = [bool]$script:Summary.CbsDriverCorruptionDetected
                CbsDriverCorruptionRemaining = [bool]$script:Summary.CbsDriverCorruptionRemaining
                CbsCorruptDriverFileNames = @($script:Summary.CbsCorruptDriverFiles | ForEach-Object { [string]$_ } | Select-Object -Unique)
                HpDriverRepairAttempted = [bool]$script:Summary.HpDriverRepairAttempted
                HpDriverRepairSucceeded = [bool]$script:Summary.HpDriverRepairSucceeded
                WmiRepositoryInconsistent = [bool]$script:Summary.WmiRepositoryInconsistent
                StorageHealthWarnings = [int]$script:Summary.StorageHealthWarnings
                StorageFailurePredictionDetected = [bool]$script:Summary.StorageFailurePredicted
                StorageWearOver50Detected = [bool]$script:Summary.StorageWearOver50Detected
                ExplorerHangCount = [int]$script:Summary.ExplorerHangCount
                ExplorerCrashCount = [int]$script:Summary.ExplorerCrashCount
                ExplorerInstabilityDetected = [bool]$script:Summary.ExplorerInstabilityDetected
                NonMicrosoftShellExtensionCount = [int]$script:Summary.NonMicrosoftShellExtensionCount
                BrokenNetworkPathCount = [int]$script:Summary.BrokenNetworkPathCount
                RpcClientSideIssueDetected = [bool]$script:Summary.RpcClientSideIssueDetected
                RpcRemoteSideIssueDetected = [bool]$script:Summary.RpcRemoteSideIssueDetected
                RpcDiagnosticsStatus = [string]$script:Summary.RpcDiagnosticsStatus
                DiskSpaceBefore = $script:DiskSpaceBefore
                DiskSpaceAfter  = $script:DiskSpaceAfter
                SpaceReclaimedGB = $reclaimed
                WindowsBuildInfo = [ordered]@{
                    ProductName    = [string]$windowsBuild.ProductName
                    EditionId      = [string]$windowsBuild.EditionId
                    DisplayVersion = [string]$windowsBuild.DisplayVersion
                    CurrentBuild   = [string]$windowsBuild.CurrentBuild
                    UBR            = $windowsBuild.UBR
                    FullBuild      = [string]$windowsBuild.Build
                }
                DismStepResults = @(Get-StepSummaryTelemetry -StepPattern 'DISM*')
                SfcStepResults = @(Get-StepSummaryTelemetry -StepPattern 'SFC*')
                RepairVolumeStepResults = @(Get-StepSummaryTelemetry -StepPattern 'RepairVolume*')
                StorageStepResults = @(Get-StepSummaryTelemetry -StepPattern 'Storage*')
                WmiStepResults = @(Get-StepSummaryTelemetry -StepPattern 'Wmi*')
                WindowsUpdateResetStepResults = @(Get-StepSummaryTelemetry -StepPattern 'WindowsUpdate*')
                ExplorerStepResults = @(Get-StepSummaryTelemetry -StepPattern 'Explorer*')
                RpcStepResults = @(Get-StepSummaryTelemetry -StepPattern 'Rpc*')
                DefenderStepResults = @(Get-StepSummaryTelemetry -StepPattern 'Defender*')
                Notes = $notes
            }
        }
    }
}

function Write-SystemRepairTelemetry {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [AllowNull()][string]$FailureMessage
    )

    if ($script:TelemetryWritten) { return }

    try {
        Ensure-LogDirectory
        $event = New-SystemRepairTelemetryEvent -Status $Status -ExitCode $ExitCode -FailureMessage $FailureMessage

        try {
            $sr = $event.maintenance.system_repair
            Write-Log ("Elastic system-repair summary: HealthIssues={0}; RepairsAttempted={1}; FailedSteps={2}; PendingRebootAfter={3}; RebootRequired={4}; SpaceReclaimedGB={5}" -f `
                $sr.HealthSummary.HealthIssueCount,
                $sr.RepairSummary.RepairsAttemptedCount,
                $sr.RepairSummary.FailedStepCount,
                $sr.RebootSummary.PendingAfter,
                $sr.RebootSummary.EffectiveRebootRequired,
                $sr.SpaceReclaimedGB) 'INFO'
        }
        catch { }
        $compact = $event | ConvertTo-Json -Depth 12 -Compress
        $pretty = $event | ConvertTo-Json -Depth 12
        Write-MaintenanceTelemetryLine -Path $script:TelemetryNdjsonPath -JsonLine $compact
        [System.IO.File]::WriteAllText($script:LatestTelemetryPath, $pretty + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
        $script:TelemetryWritten = $true
        Write-Log "Telemetry written to $($script:TelemetryNdjsonPath) and $($script:LatestTelemetryPath)" 'OK'
    }
    catch {
        Write-Log "Telemetry write failed: $($_.Exception.Message)" 'ERROR'
    }
}

# Initialize the active runtime log in the staging directory before any
# execution-path logging occurs. The staging path is outside all TempCleanup targets.
# Elastic only receives it after publication.
try {
    Initialize-SystemRepairLogging
}
catch {
    Write-Warning "Unable to initialize staged Script 08 logging: $($_.Exception.Message)"
    # Last-resort fallback also stays under ProgramData so Script 08 cleanup
    # cannot remove its own active log.
    $fallbackRoot = 'C:\ProgramData\Compton\Maintenance-Logs\Fallback'
    if (-not (Test-Path -LiteralPath $fallbackRoot)) {
        New-Item -Path $fallbackRoot -ItemType Directory -Force | Out-Null
    }
    $script:TextLogPath = Join-Path $fallbackRoot ("{0}-08_System_Repair-Fallback-{1}.log" -f $script:ComputerName, $script:TimestampForFile)
    $script:PublishedTextLogPath = $null
}

if (-not (Test-IsAdministrator)) {
    $script:Summary.StepsFailed++
    $script:FailureMessage = 'This script must be run as Administrator.'
    $script:FinalStatus = 'Failed'
    $script:FinalExitCode = 1
    try { Ensure-LogDirectory } catch {}
    try { Write-Log $script:FailureMessage 'ERROR' } catch {}
    Write-Error $script:FailureMessage
    Write-SystemRepairTelemetry -Status $script:FinalStatus -ExitCode $script:FinalExitCode -FailureMessage $script:FailureMessage
    try {
        Write-Log ("Completed {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
            $script:ScriptName,
            $script:FinalStatus,
            $script:FinalExitCode,
            $script:Summary.Warnings,
            $script:Summary.StepsFailed) 'ERROR'
    }
    catch {}
    Publish-SystemRepairLog
    exit $script:FinalExitCode
}

try {
    Ensure-LogDirectory
    $script:DiskSpaceBefore = Get-SystemDriveTelemetry
    Write-Log 'Initializing automated system health and repair script...' 'INFO'
    Write-Log "Starting $($script:ScriptName) version $($script:ScriptVersion)." 'INFO'
    Write-Log "Active staged text log path: $($script:TextLogPath)" 'INFO'
    Write-Log "Dedicated staging root: $($script:DedicatedStagingRoot)" 'INFO'
    Write-Log "Completed text log publish path: $($script:PublishedTextLogPath)" 'INFO'
    if ($AutoRepairOnDetection) {
        Write-Log 'Detection-first mode is enabled. Approved repairs will run only when corresponding problems are detected.' 'INFO'
    }
    else {
        Write-Log 'Detection-only mode is enabled. Automatic repair actions are disabled unless their independent opt-in switches are enabled.' 'INFO'
    }
    Write-Log "Detailed YAML log will be written to $($script:YamlLogPath)" 'INFO'

    Invoke-Safely -Name 'Action1AgentValidation' -ScriptBlock { Ensure-Action1Agent } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DiskSpaceCheck' -ScriptBlock { Invoke-DiskSpaceCheck } -WarnOnly | Out-Null
    Invoke-Safely -Name 'StorageHealthCheck' -ScriptBlock { Invoke-StorageHealthCheck } -WarnOnly | Out-Null

    Invoke-Safely -Name 'PendingRebootCheck' -ScriptBlock {
        $script:PendingRebootBefore = Get-PendingRebootState
        $script:Summary.PendingRebootDetected = [bool]$script:PendingRebootBefore.AnyPendingReboot
        foreach ($reason in @(Get-PendingRebootReasons -State $script:PendingRebootBefore)) {
            $script:Summary.PendingRebootReasonsBefore.Add($reason) | Out-Null
        }
        $script:PendingRebootBefore | Format-List | Out-String | ForEach-Object {
            $_.TrimEnd() -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
        }
        if ($script:PendingRebootBefore.AnyPendingReboot) { Warn-Step -Name 'PendingRebootCheck' -Reason 'A pending reboot was detected before maintenance began.' }
    } -WarnOnly | Out-Null

    Invoke-Safely -Name 'RepairVolumeScan' -ScriptBlock { Invoke-RepairVolumeScan } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DISMDetection' -ScriptBlock { Invoke-DismDetection } -WarnOnly | Out-Null
    Invoke-Safely -Name 'SFCDetection' -ScriptBlock { Invoke-SfcDetection } -WarnOnly | Out-Null
    Invoke-Safely -Name 'WmiRepositoryCheck' -ScriptBlock { Invoke-WmiCheck } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DnsFlush' -ScriptBlock { Invoke-DnsFlushOnly } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DiskUsageInventory' -ScriptBlock { Invoke-DiskUsageInventory -TopDirectories 20 } -WarnOnly | Out-Null
    Invoke-Safely -Name 'TempCleanup' -ScriptBlock { Invoke-TempCleanup } -WarnOnly | Out-Null
    Invoke-Safely -Name 'ScheduledTaskHealthCheck' -ScriptBlock { Invoke-ScheduledTaskHealthCheck } -WarnOnly | Out-Null
    Invoke-Safely -Name 'RpcDiagnostics' -ScriptBlock { Invoke-RpcDiagnostics } -WarnOnly | Out-Null
    Invoke-Safely -Name 'EventLogSummary' -ScriptBlock { Invoke-EventLogSummary } -WarnOnly | Out-Null
    Invoke-Safely -Name 'ExplorerHealthDiagnostics' -ScriptBlock { Invoke-ExplorerHealthDiagnostics } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DefenderStatusCheck' -ScriptBlock { Invoke-DefenderStatusCheck } -WarnOnly | Out-Null

    if ($AutoRepairOnDetection) {
        if ($script:Summary.DismCorruptionDetected) { Invoke-Safely -Name 'DISMRepair' -ScriptBlock { Invoke-DismRepair } | Out-Null }
        if ($script:Summary.SfcIntegrityViolations) { Invoke-Safely -Name 'SfcCbsRepairWorkflow' -ScriptBlock { Invoke-SystemFileRepairWorkflow } -WarnOnly | Out-Null }
        if ($script:Summary.CbsDriverCorruptionDetected -and $script:Summary.SfcIntegrityViolations) { Invoke-Safely -Name 'Tier3HpiaDriverRemediationFromCbs' -ScriptBlock { Invoke-HpiaDriverOnlyRepairFromCbs } -WarnOnly | Out-Null }
        if ($script:Summary.WmiRepositoryInconsistent -and $AllowWmiRepair) { Invoke-Safely -Name 'WMIRepair' -ScriptBlock { Invoke-WmiRepair } -WarnOnly | Out-Null }

        if ($script:Summary.DiskCorruptionSuspected) {
            $script:RepairVolumeSpotFixResult = $false
            Invoke-Safely -Name 'RepairVolumeSpotFix' -ScriptBlock { $script:RepairVolumeSpotFixResult = Invoke-RepairVolumeSpotFix } -WarnOnly | Out-Null
            if ([bool]$script:RepairVolumeSpotFixResult) {
                $script:RepairVolumePostSpotFixScan = $null
                Invoke-Safely -Name 'RepairVolumePostSpotFixScan' -ScriptBlock { $script:RepairVolumePostSpotFixScan = Invoke-RepairVolumeScan -StepName 'RepairVolumePostSpotFixScan' -UpdateSummary:$true } -WarnOnly | Out-Null
                if ($null -ne $script:RepairVolumePostSpotFixScan -and -not [bool]$script:RepairVolumePostSpotFixScan.CorruptionDetected) { Write-Log 'Repair-Volume post-SpotFix verification scan is clean. Offline disk repair will not be scheduled.' 'OK' }
            }
            if ($script:Summary.DiskCorruptionSuspected -and $AllowOfflineDiskRepair) { Invoke-Safely -Name 'OfflineDiskRepair' -ScriptBlock { Invoke-RepairVolumeOfflineFix } -WarnOnly | Out-Null }
            elseif ($script:Summary.DiskCorruptionSuspected) {
                Write-Log 'REPAIR_VOLUME_OFFLINE_RECOMMENDED|Reason=Repair-Volume scan still indicates possible file system corruption after SpotFix or SpotFix did not complete. Re-run script with -AllowOfflineDiskRepair to allow Repair-Volume -OfflineScanAndFix.' 'WARN'
                Add-DetailedResult -Step 'OfflineDiskRepair' -Status 'Skipped' -Message 'Offline disk repair skipped because AllowOfflineDiskRepair is disabled, but post-repair scan still indicates possible corruption.'
            }
        }
    }

    if ($AllowNetworkReset) { Invoke-Safely -Name 'NetworkReset' -ScriptBlock { Invoke-NetworkReset } -WarnOnly | Out-Null }
    if ($AllowIconCacheRebuild) { Invoke-Safely -Name 'IconCacheRebuild' -ScriptBlock { Invoke-IconCacheRebuild } -WarnOnly | Out-Null }
    if ($AllowCopilotRemoval) { Invoke-Safely -Name 'CopilotDisableAndRemoval' -ScriptBlock { Invoke-CopilotDisableAndRemoval } -WarnOnly | Out-Null }
    if ($AllowFirewallReset) { Invoke-Safely -Name 'FirewallReset' -ScriptBlock { Invoke-FirewallReset } -WarnOnly | Out-Null }
    if ($AllowWindowsUpdateReset) { Invoke-Safely -Name 'WindowsUpdateComponentReset' -ScriptBlock { Invoke-WindowsUpdateComponentReset } -WarnOnly | Out-Null }
    if ($ClearEventLogs) { Invoke-Safely -Name 'EventLogClear' -ScriptBlock { Invoke-EventLogClear } -WarnOnly | Out-Null }

    Invoke-Safely -Name 'RollbackRetentionCleanup' -ScriptBlock { Invoke-RollbackRetentionCleanup -RollbackRoot 'C:\Scripts\Rollback' -RetentionDays 7 } -WarnOnly | Out-Null
    Invoke-Safely -Name 'WindowsOldCleanup' -ScriptBlock { Remove-WindowsOldFolder -Path 'C:\Windows.old' } -WarnOnly | Out-Null
    Invoke-Safely -Name 'LogArchiveRetention' -ScriptBlock { Invoke-LogArchiveRetention -LogDirectory $LogDirectory } -WarnOnly | Out-Null

    $script:DiskSpaceAfter = Get-SystemDriveTelemetry
    Invoke-Safely -Name 'PendingRebootFinalCheck' -ScriptBlock {
        $script:PendingRebootAfter = Get-PendingRebootState
        foreach ($reason in @(Get-PendingRebootReasons -State $script:PendingRebootAfter)) { $script:Summary.PendingRebootReasonsAfter.Add($reason) | Out-Null }
        if ($script:PendingRebootAfter.AnyPendingReboot) { $script:Summary.PendingRebootDetected = $true }
    } -WarnOnly | Out-Null

    Show-Summary
    Write-YamlLog

    if ($script:Summary.StepsFailed -gt 0) {
        $script:FinalStatus = 'Failed'
        $script:FinalExitCode = 2
    }
    elseif ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected) {
        $script:FinalStatus = 'SuccessRebootRequired'
        $script:FinalExitCode = 3010
    }
    elseif ($script:Summary.Warnings -gt 0) {
        $script:FinalStatus = 'SuccessWithWarnings'
        $script:FinalExitCode = 0
    }
    else {
        $script:FinalStatus = 'Success'
        $script:FinalExitCode = 0
    }

    if ($AutoRebootIfNeeded -and ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected)) {
        Invoke-RebootIfNeeded -DelaySeconds $AutoRebootDelaySeconds
        $script:FinalStatus = 'SuccessRebootInitiated'
        $script:FinalExitCode = 3010
        Write-YamlLog
    }
}
catch {
    $script:FailureMessage = $_.Exception.Message
    $script:FinalStatus = 'Failed'
    $script:FinalExitCode = 2
    try { Fail-Step -Name 'UnhandledException' -Reason $script:FailureMessage } catch {}
    try { Show-Summary } catch {}
    try { Write-YamlLog } catch {}
}
finally {
    if ($null -eq $script:DiskSpaceAfter) { $script:DiskSpaceAfter = Get-SystemDriveTelemetry }

    # Telemetry is generated before text-log publication so all telemetry-related
    # status messages are still written to the staging log.
    Write-SystemRepairTelemetry -Status $script:FinalStatus -ExitCode $script:FinalExitCode -FailureMessage $script:FailureMessage

    # This must be the final append to the active text log. Once published into
    # C:\Logs, Elastic may open the immutable completed file immediately.
    try {
        Write-Log ("Completed {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
            $script:ScriptName,
            $script:FinalStatus,
            $script:FinalExitCode,
            $script:Summary.Warnings,
            $script:Summary.StepsFailed) $(if ($script:FinalExitCode -eq 0 -or $script:FinalExitCode -eq 3010) { 'SUCCESS' } else { 'ERROR' })
    }
    catch {}

    Publish-SystemRepairLog
}

exit $script:FinalExitCode
﻿# =====================================================================
# ScriptName: 08_System_Repair.ps1
# ScriptVersion: 4.3.0
# LastUpdated: 2026-07-27
# Changes: v4.3.0 adds Elastic-compatible execution and repair telemetry, before/after health snapshots, detailed reboot reasons, repair-result extraction, safe finalization on all exit paths, and corrects status messaging and exit handling.
# Previous: v4.2 fixes Windows PowerShell 5.1 generic-list Count warnings in RPC and Explorer diagnostics.
# Previous: v4.1 prevents service-stop hangs by placing the stop request under its own timeout; if a service remains running after the service-state timeout, the script force-terminates its backing process when requested.
# Previous: v4.0 changes safety defaults to prevent automatic repairs, disables HP driver remediation by default, and removes Windows.old/System.sav cleanup targets while investigating boot/SSD issues.
# Previous: v3.9 adds Tier 3 HP Image Assistant driver-only remediation that runs only when CBS/SFC driver corruption is detected after DISM/SFC repair.
# Previous: v3.8 adds SFC/CBS corruption extraction and a DISM+SFC repair workflow with grep-friendly CBS/SFC markers.
# Previous: v3.7 adds explicit RPC root-cause classification output and ties EventLogSummary RPC errors to the RPC diagnostics assessment.
# Previous: v3.6 adds RPC diagnostics to separate local client/service failures from remote endpoint/network RPC failures, with grep-friendly RPC_* markers.
# Previous: v3.5 adds Explorer hang/crash diagnostics, shell extension inventory, network path checks, Explorer cache cleanup, Search service restart, and grep-friendly Explorer instability markers.
# Previous: v3.4 adds grep-friendly SSD/NVMe health logging, flags wear greater than 50%, improves Repair-Volume scan parsing, and verifies the volume again after SpotFix before considering offline repair.
# Previous: v3.3 deletes C:\HP_Bios_Config during temp cleanup and attempts Repair-Volume -SpotFix first when disk corruption is detected to avoid scheduling a reboot when possible.
# Previous: v3.2 adds native SSD/NVMe health checks, SMART failure prediction reporting, reliability counter logging, and optional smartctl reporting when available.
# Previous: v3.1 removes the storage optimization block and related SSD-detection helper.
# Previous: v3.0 rebuilds SoftwareDistribution.bak* cleanup with resumable state tracking, capped parallel deletion, timeout-safe jobs, and orphan job cleanup for Task Scheduler windows.
# Previous: v2.9 adds capped parallel SoftwareDistribution.bak* deletion, a cleanup time budget, robocopy mirror cleanup per folder, and safer Task Scheduler overlap protection.
# Previous: v2.8 adds guarded SoftwareDistribution deletion with timeout-protected robocopy/rmdir fallbacks and immediate YAML log writes before cleanup steps.
# Previous: v2.7 adds Action1 Agent validation and silent MSI installation from \\filesvr\Labscripts when C:\Windows\Action1\action1_agent.exe is missing.
# Previous: v2.6 replaces SoftwareDistribution cleanup with robocopy empty-folder mirror cleanup for the active folder and SoftwareDistribution.bak* folders.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$AutoRepairOnDetection = $false,
    [switch]$AllowWmiRepair = $true,
    [switch]$AllowNetworkReset = $false,
    [switch]$AllowWindowsUpdateReset = $false,
    [switch]$AllowOfflineDiskRepair = $false,
    [switch]$AllowFirewallReset = $false,
    [switch]$AllowIconCacheRebuild = $false,
    [switch]$AllowExplorerHealthRepair = $true,
    [switch]$AllowSearchServiceRestart = $true,
    [switch]$AllowHpDriverRepairFromCbs = $false,
    [string]$HpImageAssistantSourcePath = '\\filesvr\Labscripts\HPImageAssistant',
    [string]$HpImageAssistantLocalPath = 'C:\ProgramData\SystemRepair\HPImageAssistant',
    [switch]$AllowCopilotRemoval = $false,
    [switch]$AggressiveCleanup = $false,
    [switch]$ClearEventLogs = $false,
    [switch]$AutoRebootIfNeeded = $false,
    [int]$AutoRebootDelaySeconds = 60,
    [int]$MaxParallelSoftwareDistributionDeletes = 2,
    [int]$SoftwareDistributionCleanupTimeLimitMinutes = 40,
    [int]$SoftwareDistributionFolderJobTimeoutMinutes = 15,
    [string]$SoftwareDistributionCleanupStatePath = 'C:\ProgramData\SystemRepair\SoftwareDistributionCleanupState.json',
    [string]$LogDirectory = 'C:\Logs'
)

$ErrorActionPreference = 'Stop'

$script:ScriptName = '08_System_Repair.ps1'
$script:ScriptVersion = '4.3.0'
$script:RunId = [guid]::NewGuid().Guid
$script:TelemetryNdjsonPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestTelemetryPath = Join-Path $LogDirectory '08_System_Repair.latest.json'
$script:TelemetryWritten = $false
$script:FinalExitCode = 0
$script:FinalStatus = 'Running'
$script:FailureMessage = $null
$script:PendingRebootBefore = $null
$script:PendingRebootAfter = $null
$script:DiskSpaceBefore = $null
$script:DiskSpaceAfter = $null
$script:RebootInitiated = $false

$script:RunStart = Get-Date
$script:ComputerName = $env:COMPUTERNAME
$script:TimestampForFile = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
$script:BaseFileName = "{0}-SystemRepair-{1}" -f $script:ComputerName, $script:TimestampForFile
$script:YamlLogPath = Join-Path $LogDirectory ($script:BaseFileName + '.yaml')

$script:Summary = [ordered]@{
    ComputerName                 = $script:ComputerName
    StartTime                    = $script:RunStart
    EndTime                      = $null
    StepsSucceeded               = 0
    StepsFailed                  = 0
    Warnings                     = 0
    RebootRequired               = $false
    PendingRebootDetected        = $false
    PendingRebootReasonsBefore   = New-Object System.Collections.Generic.List[string]
    PendingRebootReasonsAfter    = New-Object System.Collections.Generic.List[string]
    DiskCorruptionSuspected      = $false
    DismCorruptionDetected       = $false
    SfcIntegrityViolations       = $false
    SfcRepairAttempted           = $false
    SfcRepairSucceeded           = $false
    CbsCorruptFileCount          = 0
    CbsCorruptionExtractPath     = $null
    CbsDriverCorruptionDetected  = $false
    CbsDriverCorruptionRemaining = $false
    CbsCorruptDriverFiles        = New-Object System.Collections.Generic.List[string]
    HpDriverRepairAttempted      = $false
    HpDriverRepairSucceeded      = $false
    HpImageAssistantPath         = $null
    WmiRepositoryInconsistent    = $false
    StorageHealthWarnings        = 0
    StorageFailurePredicted      = $false
    StorageWearOver50Detected     = $false
    ExplorerHangCount             = 0
    ExplorerCrashCount            = 0
    ExplorerInstabilityDetected   = $false
    NonMicrosoftShellExtensionCount = 0
    BrokenNetworkPathCount        = 0
    RpcClientSideIssueDetected    = $false
    RpcRemoteSideIssueDetected    = $false
    RpcDiagnosticsStatus          = 'NotRun'
    RepairVolumeSpotFixSucceeded = $false
    RepairsAttempted             = New-Object System.Collections.Generic.List[string]
    Notes                        = New-Object System.Collections.Generic.List[string]
}

$script:DetailedResults = New-Object System.Collections.Generic.List[object]

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Ensure-LogDirectory {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }
}

function ConvertTo-YamlScalar {
    param(
        [Parameter(Mandatory)][AllowNull()]$Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }

    if ($Value -is [datetime]) {
        return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'"
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [string]$Value
    }

    $text = [string]$Value
    $text = $text -replace "'", "''"
    return "'" + $text + "'"
}

function Add-DetailedResult {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [AllowNull()]$Data = $null
    )

    $script:DetailedResults.Add([PSCustomObject]@{
        Timestamp = Get-Date
        Step      = $Step
        Status    = $Status
        Message   = $Message
        Data      = $Data
    }) | Out-Null
}

function Write-YamlLog {
    try {
        Ensure-LogDirectory

        $lines = New-Object System.Collections.Generic.List[string]

        $lines.Add('run:') | Out-Null
        $lines.Add("  computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("  start_time: $(ConvertTo-YamlScalar $script:Summary.StartTime)") | Out-Null
        $lines.Add("  end_time: $(ConvertTo-YamlScalar $script:Summary.EndTime)") | Out-Null
        $lines.Add("  yaml_log_path: $(ConvertTo-YamlScalar $script:YamlLogPath)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('settings:') | Out-Null
        $lines.Add("  auto_repair_on_detection: $(ConvertTo-YamlScalar $AutoRepairOnDetection)") | Out-Null
        $lines.Add("  allow_wmi_repair: $(ConvertTo-YamlScalar $AllowWmiRepair)") | Out-Null
        $lines.Add("  allow_network_reset: $(ConvertTo-YamlScalar $AllowNetworkReset)") | Out-Null
        $lines.Add("  allow_windows_update_reset: $(ConvertTo-YamlScalar $AllowWindowsUpdateReset)") | Out-Null
        $lines.Add("  allow_offline_disk_repair: $(ConvertTo-YamlScalar $AllowOfflineDiskRepair)") | Out-Null
        $lines.Add("  allow_firewall_reset: $(ConvertTo-YamlScalar $AllowFirewallReset)") | Out-Null
        $lines.Add("  allow_icon_cache_rebuild: $(ConvertTo-YamlScalar $AllowIconCacheRebuild)") | Out-Null
        $lines.Add("  allow_explorer_health_repair: $(ConvertTo-YamlScalar $AllowExplorerHealthRepair)") | Out-Null
        $lines.Add("  allow_search_service_restart: $(ConvertTo-YamlScalar $AllowSearchServiceRestart)") | Out-Null
        $lines.Add("  allow_hp_driver_repair_from_cbs: $(ConvertTo-YamlScalar $AllowHpDriverRepairFromCbs)") | Out-Null
        $lines.Add("  hp_image_assistant_source_path: $(ConvertTo-YamlScalar $HpImageAssistantSourcePath)") | Out-Null
        $lines.Add("  hp_image_assistant_local_path: $(ConvertTo-YamlScalar $HpImageAssistantLocalPath)") | Out-Null
        $lines.Add("  allow_copilot_removal: $(ConvertTo-YamlScalar $AllowCopilotRemoval)") | Out-Null
        $lines.Add("  aggressive_cleanup: $(ConvertTo-YamlScalar $AggressiveCleanup)") | Out-Null
        $lines.Add("  clear_event_logs: $(ConvertTo-YamlScalar $ClearEventLogs)") | Out-Null
        $lines.Add("  auto_reboot_if_needed: $(ConvertTo-YamlScalar $AutoRebootIfNeeded)") | Out-Null
        $lines.Add("  auto_reboot_delay_seconds: $(ConvertTo-YamlScalar $AutoRebootDelaySeconds)") | Out-Null
        $lines.Add("  max_parallel_softwaredistribution_deletes: $(ConvertTo-YamlScalar $MaxParallelSoftwareDistributionDeletes)") | Out-Null
        $lines.Add("  softwaredistribution_cleanup_time_limit_minutes: $(ConvertTo-YamlScalar $SoftwareDistributionCleanupTimeLimitMinutes)") | Out-Null
        $lines.Add("  softwaredistribution_folder_job_timeout_minutes: $(ConvertTo-YamlScalar $SoftwareDistributionFolderJobTimeoutMinutes)") | Out-Null
        $lines.Add("  softwaredistribution_cleanup_state_path: $(ConvertTo-YamlScalar $SoftwareDistributionCleanupStatePath)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('summary:') | Out-Null
        $lines.Add("  steps_succeeded: $(ConvertTo-YamlScalar $script:Summary.StepsSucceeded)") | Out-Null
        $lines.Add("  steps_failed: $(ConvertTo-YamlScalar $script:Summary.StepsFailed)") | Out-Null
        $lines.Add("  warnings: $(ConvertTo-YamlScalar $script:Summary.Warnings)") | Out-Null
        $lines.Add("  reboot_required: $(ConvertTo-YamlScalar $script:Summary.RebootRequired)") | Out-Null
        $lines.Add("  pending_reboot_detected: $(ConvertTo-YamlScalar $script:Summary.PendingRebootDetected)") | Out-Null
        $lines.Add("  disk_corruption_suspected: $(ConvertTo-YamlScalar $script:Summary.DiskCorruptionSuspected)") | Out-Null
        $lines.Add("  dism_corruption_detected: $(ConvertTo-YamlScalar $script:Summary.DismCorruptionDetected)") | Out-Null
        $lines.Add("  sfc_integrity_violations: $(ConvertTo-YamlScalar $script:Summary.SfcIntegrityViolations)") | Out-Null
        $lines.Add("  sfc_repair_attempted: $(ConvertTo-YamlScalar $script:Summary.SfcRepairAttempted)") | Out-Null
        $lines.Add("  sfc_repair_succeeded: $(ConvertTo-YamlScalar $script:Summary.SfcRepairSucceeded)") | Out-Null
        $lines.Add("  cbs_corrupt_file_count: $(ConvertTo-YamlScalar $script:Summary.CbsCorruptFileCount)") | Out-Null
        $lines.Add("  cbs_corruption_extract_path: $(ConvertTo-YamlScalar $script:Summary.CbsCorruptionExtractPath)") | Out-Null
        $lines.Add("  cbs_driver_corruption_detected: $(ConvertTo-YamlScalar $script:Summary.CbsDriverCorruptionDetected)") | Out-Null
        $lines.Add("  cbs_driver_corruption_remaining: $(ConvertTo-YamlScalar $script:Summary.CbsDriverCorruptionRemaining)") | Out-Null
        $lines.Add("  cbs_corrupt_driver_files: $(ConvertTo-YamlScalar (($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ','))") | Out-Null
        $lines.Add("  hp_driver_repair_attempted: $(ConvertTo-YamlScalar $script:Summary.HpDriverRepairAttempted)") | Out-Null
        $lines.Add("  hp_driver_repair_succeeded: $(ConvertTo-YamlScalar $script:Summary.HpDriverRepairSucceeded)") | Out-Null
        $lines.Add("  hp_image_assistant_path: $(ConvertTo-YamlScalar $script:Summary.HpImageAssistantPath)") | Out-Null
        $lines.Add("  wmi_repository_inconsistent: $(ConvertTo-YamlScalar $script:Summary.WmiRepositoryInconsistent)") | Out-Null
        $lines.Add("  storage_health_warnings: $(ConvertTo-YamlScalar $script:Summary.StorageHealthWarnings)") | Out-Null
        $lines.Add("  storage_failure_predicted: $(ConvertTo-YamlScalar $script:Summary.StorageFailurePredicted)") | Out-Null
        $lines.Add("  storage_wear_over_50_detected: $(ConvertTo-YamlScalar $script:Summary.StorageWearOver50Detected)") | Out-Null
        $lines.Add("  explorer_hang_count: $(ConvertTo-YamlScalar $script:Summary.ExplorerHangCount)") | Out-Null
        $lines.Add("  explorer_crash_count: $(ConvertTo-YamlScalar $script:Summary.ExplorerCrashCount)") | Out-Null
        $lines.Add("  explorer_instability_detected: $(ConvertTo-YamlScalar $script:Summary.ExplorerInstabilityDetected)") | Out-Null
        $lines.Add("  non_microsoft_shell_extension_count: $(ConvertTo-YamlScalar $script:Summary.NonMicrosoftShellExtensionCount)") | Out-Null
        $lines.Add("  broken_network_path_count: $(ConvertTo-YamlScalar $script:Summary.BrokenNetworkPathCount)") | Out-Null
        $lines.Add("  rpc_client_side_issue_detected: $(ConvertTo-YamlScalar $script:Summary.RpcClientSideIssueDetected)") | Out-Null
        $lines.Add("  rpc_remote_side_issue_detected: $(ConvertTo-YamlScalar $script:Summary.RpcRemoteSideIssueDetected)") | Out-Null
        $lines.Add("  rpc_diagnostics_status: $(ConvertTo-YamlScalar $script:Summary.RpcDiagnosticsStatus)") | Out-Null
        $lines.Add("  repair_volume_spotfix_succeeded: $(ConvertTo-YamlScalar $script:Summary.RepairVolumeSpotFixSucceeded)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('repairs_attempted:') | Out-Null
        if ($script:Summary.RepairsAttempted.Count -gt 0) {
            foreach ($repair in $script:Summary.RepairsAttempted) {
                $lines.Add("  - $(ConvertTo-YamlScalar $repair)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }
        $lines.Add('') | Out-Null

        $lines.Add('notes:') | Out-Null
        if ($script:Summary.Notes.Count -gt 0) {
            foreach ($note in $script:Summary.Notes) {
                $lines.Add("  - $(ConvertTo-YamlScalar $note)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }
        $lines.Add('') | Out-Null

        $lines.Add('detailed_results:') | Out-Null
        if ($script:DetailedResults.Count -gt 0) {
            foreach ($entry in $script:DetailedResults) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    timestamp: $(ConvertTo-YamlScalar $entry.Timestamp)") | Out-Null
                $lines.Add("    step: $(ConvertTo-YamlScalar $entry.Step)") | Out-Null
                $lines.Add("    status: $(ConvertTo-YamlScalar $entry.Status)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $entry.Message)") | Out-Null

                if ($null -eq $entry.Data) {
                    $lines.Add('    data: null') | Out-Null
                }
                elseif ($entry.Data -is [System.Collections.IDictionary]) {
                    $lines.Add('    data:') | Out-Null
                    foreach ($key in $entry.Data.Keys) {
                        $lines.Add("      $key`: $(ConvertTo-YamlScalar $entry.Data[$key])") | Out-Null
                    }
                }
                else {
                    $lines.Add("    data: $(ConvertTo-YamlScalar $entry.Data)") | Out-Null
                }
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        Set-Content -Path $script:YamlLogPath -Value $lines -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }
}

function Add-Note {
    param([string]$Message)
    $script:Summary.Notes.Add($Message) | Out-Null
}

function Add-RepairAttempt {
    param([string]$Message)
    $script:Summary.RepairsAttempted.Add($Message) | Out-Null
}

function Complete-Step {
    param([string]$Name)
    $script:Summary.StepsSucceeded++
    Write-Log "$Name completed." 'OK'
    Add-DetailedResult -Step $Name -Status 'Succeeded' -Message "$Name completed successfully."
}

function Fail-Step {
    param(
        [string]$Name,
        [string]$Reason
    )
    $script:Summary.StepsFailed++
    Write-Log "$Name failed: $Reason" 'ERROR'
    Add-Note "$Name failed: $Reason"
    Add-DetailedResult -Step $Name -Status 'Failed' -Message $Reason
}

function Warn-Step {
    param(
        [string]$Name,
        [string]$Reason
    )
    $script:Summary.Warnings++
    Write-Log "$Name warning: $Reason" 'WARN'
    Add-Note "$Name warning: $Reason"
    Add-DetailedResult -Step $Name -Status 'Warning' -Message $Reason
}

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}


function Set-ClassicContextMenuForHive {
    param(
        [Parameter(Mandatory)][string]$RootKey
    )

    $basePath = Join-Path $RootKey 'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
    $inprocPath = Join-Path $basePath 'InprocServer32'

    if (-not (Test-Path -LiteralPath $basePath)) {
        New-Item -Path $basePath -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $inprocPath)) {
        New-Item -Path $inprocPath -Force | Out-Null
    }

    New-ItemProperty -Path $inprocPath -Name '(default)' -Value '' -PropertyType String -Force | Out-Null
    Write-Log "Classic context menu registry value set for hive: $RootKey" 'OK'
    Add-DetailedResult -Step 'ClassicContextMenuRegistry' -Status 'Info' -Message "Classic context menu registry value set." -Data @{
        RootKey = $RootKey
        RegistryPath = $inprocPath
    }
}

function Enable-ClassicContextMenuAllUsers {
    Write-Log 'Applying classic Windows 10-style context menu for all users...' 'INFO'

    Set-ClassicContextMenuForHive -RootKey 'Registry::HKEY_CURRENT_USER'

    $userSids = Get-ChildItem Registry::HKEY_USERS |
        Where-Object {
            $_.PSChildName -match '^S-1-5-21-' -and
            $_.PSChildName -notmatch '_Classes$'
        } |
        Select-Object -ExpandProperty PSChildName

    foreach ($sid in $userSids) {
        Set-ClassicContextMenuForHive -RootKey "Registry::HKEY_USERS\$sid"
    }

    $defaultHiveName = 'HKU\DefaultTemp'
    $defaultHivePsPath = 'Registry::HKEY_USERS\DefaultTemp'
    $defaultUserNtUserDat = 'C:\Users\Default\NTUSER.DAT'

    if (Test-Path -LiteralPath $defaultUserNtUserDat) {
        $hiveLoaded = $false

        try {
            $loadResult = & reg.exe load $defaultHiveName $defaultUserNtUserDat
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to load Default User hive: $($loadResult -join ' ')"
            }

            $hiveLoaded = $true
            Start-Sleep -Milliseconds 750

            Set-ClassicContextMenuForHive -RootKey $defaultHivePsPath

            Start-Sleep -Milliseconds 750
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 750
        }
        catch {
            Write-Log "Failed to update Default User profile: $($_.Exception.Message)" 'ERROR'
            Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Failed' -Message $_.Exception.Message
        }
        finally {
            if ($hiveLoaded) {
                $unloaded = $false

                foreach ($attempt in 1..5) {
                    $unloadResult = & reg.exe unload $defaultHiveName
                    if ($LASTEXITCODE -eq 0) {
                        $unloaded = $true
                        Write-Log 'Applied classic context menu to Default User profile.' 'OK'
                        Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Succeeded' -Message 'Applied classic context menu to Default User profile.'
                        break
                    }

                    Start-Sleep -Seconds 1
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }

                if (-not $unloaded) {
                    Write-Log 'Classic context menu was written to Default User profile, but unloading the hive failed. A reboot may be required before the hive is released.' 'WARN'
                    Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Warning' -Message 'Classic context menu was written to Default User profile, but unloading the hive failed. A reboot may be required before the hive is released.'
                }
            }
        }
    }
    else {
        Write-Log 'Default User NTUSER.DAT not found; future new users were not updated.' 'WARN'
        Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Warning' -Message 'Default User NTUSER.DAT not found; future new users were not updated.'
    }

    Write-Log 'Classic context menu registry changes applied. Users may need to sign out and back in.' 'INFO'
    Add-DetailedResult -Step 'ClassicContextMenuAllUsers' -Status 'Info' -Message 'Classic context menu registry changes applied for current, loaded, and default user profiles.'
}

function Invoke-Safely {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [switch]$WarnOnly
    )

    try {
        & $ScriptBlock
        Complete-Step -Name $Name
        return $true
    }
    catch {
        if ($WarnOnly) {
            Warn-Step -Name $Name -Reason $_.Exception.Message
        }
        else {
            Fail-Step -Name $Name -Reason $_.Exception.Message
        }
        return $false
    }
}

function Ensure-Action1Agent {
    [CmdletBinding()]
    param(
        [string]$AgentPath = 'C:\Windows\Action1\action1_agent.exe',
        [string]$InstallerPath = '\\filesvr\Labscripts\action1_agent(Compton_College).msi'
    )

    Write-Log 'Checking for Action1 Agent...' 'INFO'

    if (Test-Path -LiteralPath $AgentPath) {
        Write-Log 'Action1 Agent already exists. Skipping installation.' 'OK'
        Add-DetailedResult -Step 'Action1Agent' -Status 'Skipped' -Message 'Action1 Agent already installed.' -Data @{
            AgentPath = $AgentPath
        }
        return
    }

    Write-Log "Action1 Agent not found at $AgentPath. Preparing silent install..." 'WARN'

    if (-not (Test-Path -LiteralPath $InstallerPath)) {
        throw "Action1 Agent installer not found: $InstallerPath"
    }

    $arguments = "/i `"$InstallerPath`" /quiet /qn"
    Write-Log "Running Action1 Agent installer from $InstallerPath" 'INFO'

    $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
    $exitCode = [int]$process.ExitCode

    Add-DetailedResult -Step 'Action1AgentInstall' -Status 'Info' -Message 'Action1 Agent installer finished.' -Data @{
        InstallerPath = $InstallerPath
        AgentPath     = $AgentPath
        ExitCode      = $exitCode
    }

    if ($exitCode -eq 3010 -or $exitCode -eq 1641) {
        $script:Summary.RebootRequired = $true
        Write-Log "Action1 Agent installer returned exit code $exitCode. Reboot required." 'WARN'
    }
    elseif ($exitCode -ne 0) {
        throw "Action1 Agent MSI install failed with exit code $exitCode"
    }

    Start-Sleep -Seconds 3

    if (Test-Path -LiteralPath $AgentPath) {
        Write-Log 'Action1 Agent installed successfully.' 'OK'
        Add-DetailedResult -Step 'Action1Agent' -Status 'Succeeded' -Message 'Action1 Agent detected after installation.' -Data @{
            AgentPath = $AgentPath
        }
    }
    else {
        throw 'Action1 Agent installer completed, but C:\Windows\Action1\action1_agent.exe was not detected afterward.'
    }
}

function Get-PendingRebootState {
    $result = [ordered]@{
        CBServicing_RebootPending         = $false
        WindowsUpdate_RebootRequired      = $false
        SessionManager_PendingFileRename  = $false
        SessionManager_PendingFileRename2 = $false
        UpdateExeVolatile                 = $false
        PackagesPending                   = $false
        WUAU_RebootRequired_COM           = $false
        AnyPendingReboot                  = $false
    }

    $cbsRebootPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    $wuRebootRequired = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    $packagesPending  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    $sessionMgr       = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    $updateExe        = 'HKLM:\SOFTWARE\Microsoft\Updates'

    try { $result.CBServicing_RebootPending = Test-Path -LiteralPath $cbsRebootPending } catch {}
    try { $result.WindowsUpdate_RebootRequired = Test-Path -LiteralPath $wuRebootRequired } catch {}
    try { $result.PackagesPending = Test-Path -LiteralPath $packagesPending } catch {}

    try {
        $pendingRename = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations' -ErrorAction Stop).PendingFileRenameOperations
        if ($null -ne $pendingRename -and $pendingRename.Count -gt 0) {
            $result.SessionManager_PendingFileRename = $true
        }
    }
    catch {}

    try {
        $pendingRename2 = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations2' -ErrorAction Stop).PendingFileRenameOperations2
        if ($null -ne $pendingRename2 -and $pendingRename2.Count -gt 0) {
            $result.SessionManager_PendingFileRename2 = $true
        }
    }
    catch {}

    try {
        $uev = (Get-ItemProperty -Path $updateExe -Name 'UpdateExeVolatile' -ErrorAction Stop).UpdateExeVolatile
        if ($null -ne $uev -and [int]$uev -ne 0) {
            $result.UpdateExeVolatile = $true
        }
    }
    catch {}

    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        if ($sysInfo.RebootRequired) {
            $result.WUAU_RebootRequired_COM = $true
        }
    }
    catch {}

    if ($result.Values -contains $true) {
        $result.AnyPendingReboot = $true
    }

    Add-DetailedResult -Step 'PendingRebootCheckData' -Status 'Info' -Message 'Collected pending reboot state.' -Data $result
    [PSCustomObject]$result
}


function Invoke-DismCommand {
    param([string[]]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\dism.exe"
    $psi.Arguments = $Arguments -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
    }

    Add-DetailedResult -Step 'DISM' -Status 'Info' -Message ("Executed DISM: " + ($Arguments -join ' ')) -Data @{
        ExitCode = $proc.ExitCode
    }

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Invoke-SfcCommand {
    param([string]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\sfc.exe"
    $psi.Arguments = $Arguments
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
    }

    Add-DetailedResult -Step 'SFC' -Status 'Info' -Message ("Executed SFC: " + $Arguments) -Data @{
        ExitCode = $proc.ExitCode
    }

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}



function Get-StorageSmartFailurePrediction {
    [CmdletBinding()]
    param()

    try {
        Get-CimInstance -Namespace 'root\wmi' -ClassName 'MSStorageDriver_FailurePredictStatus' -ErrorAction Stop |
            Select-Object InstanceName, PredictFailure, Reason
    }
    catch {
        Write-Log "SMART failure prediction status is not available through root\wmi on this system: $($_.Exception.Message)" 'WARN'
        Add-DetailedResult -Step 'StorageSmartFailurePrediction' -Status 'Warning' -Message 'SMART failure prediction status is not available through root\wmi.' -Data @{
            Error = $_.Exception.Message
        }
        @()
    }
}

function Get-StorageReliabilitySnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$PhysicalDisk
    )

    try {
        $counter = $PhysicalDisk | Get-StorageReliabilityCounter -ErrorAction Stop
        if ($null -eq $counter) {
            return $null
        }

        [PSCustomObject]@{
            Temperature          = $counter.Temperature
            Wear                 = $counter.Wear
            PowerOnHours         = $counter.PowerOnHours
            ReadErrorsTotal      = $counter.ReadErrorsTotal
            WriteErrorsTotal     = $counter.WriteErrorsTotal
            MediaErrors          = $counter.MediaErrors
            LoadUnloadCycleCount = $counter.LoadUnloadCycleCount
        }
    }
    catch {
        Write-Log "Reliability counters are not available for $($PhysicalDisk.FriendlyName): $($_.Exception.Message)" 'WARN'
        [PSCustomObject]@{
            Temperature          = $null
            Wear                 = $null
            PowerOnHours         = $null
            ReadErrorsTotal      = $null
            WriteErrorsTotal     = $null
            MediaErrors          = $null
            LoadUnloadCycleCount = $null
            Error                = $_.Exception.Message
        }
    }
}

function Invoke-SmartCtlOptionalReport {
    [CmdletBinding()]
    param()

    $smartCtl = Get-Command -Name 'smartctl.exe' -ErrorAction SilentlyContinue
    if (-not $smartCtl) {
        Add-DetailedResult -Step 'SmartCtlOptionalReport' -Status 'Skipped' -Message 'smartctl.exe is not installed or not in PATH. Native Windows storage checks were used.'
        return
    }

    Write-Log "smartctl.exe detected at $($smartCtl.Source). Capturing optional SMART summary." 'INFO'

    try {
        $scanOutput = & $smartCtl.Source --scan 2>&1
        Add-DetailedResult -Step 'SmartCtlScan' -Status 'Info' -Message 'smartctl scan output captured.' -Data @{
            SmartCtlPath = $smartCtl.Source
            ScanOutput   = ($scanOutput -join ' | ')
        }

        foreach ($line in @($scanOutput)) {
            if ($line -match '^(?<device>\S+)\s') {
                $device = $Matches.device
                $healthOutput = & $smartCtl.Source -H $device 2>&1
                Add-DetailedResult -Step 'SmartCtlHealth' -Status 'Info' -Message "smartctl health output captured for $device." -Data @{
                    Device = $device
                    Output = ($healthOutput -join ' | ')
                }
            }
        }
    }
    catch {
        Write-Log "smartctl optional report failed: $($_.Exception.Message)" 'WARN'
        Add-DetailedResult -Step 'SmartCtlOptionalReport' -Status 'Warning' -Message 'smartctl optional report failed.' -Data @{
            Error = $_.Exception.Message
        }
    }
}

function Invoke-StorageHealthCheck {
    [CmdletBinding()]
    param()

    Write-Log 'Checking SSD/NVMe/M.2 health with native Windows storage APIs...' 'INFO'

    $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop)
    if ($physicalDisks.Count -eq 0) {
        throw 'No physical disks were returned by Get-PhysicalDisk.'
    }

    $diskSummaries = New-Object System.Collections.Generic.List[object]
    $warningMessages = New-Object System.Collections.Generic.List[string]

    foreach ($disk in $physicalDisks) {
        $reliability = Get-StorageReliabilitySnapshot -PhysicalDisk $disk
        $mediaType = [string]$disk.MediaType
        $busType = if ($disk.PSObject.Properties.Name -contains 'BusType') { [string]$disk.BusType } else { 'Unknown' }
        $isSolidState = ($mediaType -match 'SSD|SCM') -or ($busType -match 'NVMe')

        $summary = [ordered]@{
            FriendlyName          = [string]$disk.FriendlyName
            SerialNumber          = [string]$disk.SerialNumber
            MediaType             = $mediaType
            BusType               = $busType
            HealthStatus          = [string]$disk.HealthStatus
            OperationalStatus     = (@($disk.OperationalStatus) -join ', ')
            SizeGB                = [math]::Round(([double]$disk.Size / 1GB), 2)
            IsSolidStateOrNVMe    = [bool]$isSolidState
            Temperature           = $reliability.Temperature
            Wear                  = $reliability.Wear
            PowerOnHours          = $reliability.PowerOnHours
            ReadErrorsTotal       = $reliability.ReadErrorsTotal
            WriteErrorsTotal      = $reliability.WriteErrorsTotal
            MediaErrors           = $reliability.MediaErrors
            LoadUnloadCycleCount  = $reliability.LoadUnloadCycleCount
        }

        $diskSummaries.Add([PSCustomObject]$summary) | Out-Null

        Write-Log ("Drive: {0} | Media: {1} | Bus: {2} | Health: {3} | Operational: {4} | Temp: {5}C | Wear: {6}% | PowerOnHours: {7}" -f $summary.FriendlyName, $summary.MediaType, $summary.BusType, $summary.HealthStatus, $summary.OperationalStatus, $summary.Temperature, $summary.Wear, $summary.PowerOnHours) 'INFO'
        Write-Log ("STORAGE_HEALTH_DATA|Computer={0}|Drive={1}|Serial={2}|Media={3}|Bus={4}|Health={5}|Operational={6}|SizeGB={7}|TemperatureC={8}|WearPercent={9}|PowerOnHours={10}|MediaErrors={11}|ReadErrorsTotal={12}|WriteErrorsTotal={13}" -f $script:ComputerName, $summary.FriendlyName, $summary.SerialNumber, $summary.MediaType, $summary.BusType, $summary.HealthStatus, $summary.OperationalStatus, $summary.SizeGB, $summary.Temperature, $summary.Wear, $summary.PowerOnHours, $summary.MediaErrors, $summary.ReadErrorsTotal, $summary.WriteErrorsTotal) 'INFO'

        if ($summary.HealthStatus -and $summary.HealthStatus -notin @('Healthy','Unknown')) {
            $warningMessages.Add("$($summary.FriendlyName) reports HealthStatus $($summary.HealthStatus).") | Out-Null
        }

        if ($summary.OperationalStatus -and $summary.OperationalStatus -notmatch 'OK|Unknown') {
            $warningMessages.Add("$($summary.FriendlyName) reports OperationalStatus $($summary.OperationalStatus).") | Out-Null
        }

        if ($null -ne $summary.Temperature -and [int]$summary.Temperature -ge 70) {
            $warningMessages.Add("$($summary.FriendlyName) temperature is $($summary.Temperature)C.") | Out-Null
        }

        if ($null -ne $summary.Wear -and [int]$summary.Wear -gt 50) {
            $script:Summary.StorageWearOver50Detected = $true
            Write-Log ("STORAGE_WEAR_GT_50|Computer={0}|Drive={1}|Serial={2}|Media={3}|Bus={4}|WearPercent={5}|Health={6}|Operational={7}" -f $script:ComputerName, $summary.FriendlyName, $summary.SerialNumber, $summary.MediaType, $summary.BusType, $summary.Wear, $summary.HealthStatus, $summary.OperationalStatus) 'WARN'
            $warningMessages.Add("$($summary.FriendlyName) wear is $($summary.Wear) percent, which is greater than the 50 percent review threshold. Review vendor interpretation and plan replacement if appropriate.") | Out-Null
        }
        elseif ($null -ne $summary.Wear -and [int]$summary.Wear -ge 80) {
            $warningMessages.Add("$($summary.FriendlyName) wear is $($summary.Wear) percent. Review vendor interpretation and plan replacement if appropriate.") | Out-Null
        }

        if (($null -ne $summary.MediaErrors -and [int64]$summary.MediaErrors -gt 0) -or
            ($null -ne $summary.ReadErrorsTotal -and [int64]$summary.ReadErrorsTotal -gt 0) -or
            ($null -ne $summary.WriteErrorsTotal -and [int64]$summary.WriteErrorsTotal -gt 0)) {
            $warningMessages.Add("$($summary.FriendlyName) reports storage reliability errors. MediaErrors=$($summary.MediaErrors), ReadErrors=$($summary.ReadErrorsTotal), WriteErrors=$($summary.WriteErrorsTotal).") | Out-Null
        }
    }

    $smartPredictions = @(Get-StorageSmartFailurePrediction)
    foreach ($prediction in $smartPredictions) {
        Write-Log ("SMART PredictFailure: {0} | Reason: {1} | Instance: {2}" -f $prediction.PredictFailure, $prediction.Reason, $prediction.InstanceName) 'INFO'
        if ([bool]$prediction.PredictFailure) {
            $script:Summary.StorageFailurePredicted = $true
            $warningMessages.Add("SMART predicts a possible storage failure for instance $($prediction.InstanceName). Reason=$($prediction.Reason).") | Out-Null
        }
    }

    if ($warningMessages.Count -gt 0) {
        $script:Summary.StorageHealthWarnings += $warningMessages.Count
        $script:Summary.Warnings += $warningMessages.Count
        foreach ($warning in $warningMessages) {
            Write-Log $warning 'WARN'
            Add-Note "Storage health warning: $warning"
        }
    }
    else {
        Write-Log 'No SSD/NVMe health warnings detected by native Windows storage checks.' 'OK'
    }

    Add-DetailedResult -Step 'StorageHealthCheck' -Status 'Info' -Message 'SSD/NVMe/M.2 health check completed.' -Data @{
        DiskSummaryJson       = (($diskSummaries | ConvertTo-Json -Depth 5 -Compress) -replace "'", "''")
        SmartPredictionJson   = (($smartPredictions | ConvertTo-Json -Depth 5 -Compress) -replace "'", "''")
        WarningCount          = $warningMessages.Count
        Warnings              = ($warningMessages -join ' | ')
    }

    Invoke-SmartCtlOptionalReport
}

function Get-DiskSpaceInfo {
    param([string]$Path)

    try {
        $driveRoot = Split-Path -Path $Path -Qualifier
        if ([string]::IsNullOrWhiteSpace($driveRoot)) {
            $driveRoot = $env:SystemDrive + '\'
        }

        $drive = [System.IO.DriveInfo]::new($driveRoot)
        return @{
            FreeSpace  = [int64]$drive.AvailableFreeSpace
            TotalSize  = [int64]$drive.TotalSize
            UsedSpace  = [int64]($drive.TotalSize - $drive.AvailableFreeSpace)
        }
    }
    catch {
        return $null
    }
}

function Test-SafeCleanupPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    $normalized = $Path.TrimEnd('\')

    $blockedPaths = @(
        'C:\Windows\System32',
        'C:\Windows\SysWOW64',
        'C:\Program Files',
        'C:\Program Files (x86)',
        'C:\Windows\explorer.exe',
        'C:\Windows\System32\drivers'
    )

    foreach ($blocked in $blockedPaths) {
        if ($normalized -ieq $blocked -or $normalized -like ($blocked + '\*')) {
            return $false
        }
    }

    $allowedPatterns = @(
        'C:\Windows\Temp*',
        'C:\Temp*',
        'C:\SWSetup*',
        'C:\Lab Update Scripts*',
        'C:\ProgramData\Win11UpgradeStage*',
        'C:\Windows\SoftwareDistribution.bak*',
        'C:\SoftwareDistribution.bak*',
        'C:\Windows\SoftwareDistribution\Download*',
        'C:\Windows\Prefetch*',
        'C:\Windows\Logs\CBS*',
        'C:\ProgramData\Microsoft\Windows\WER\ReportQueue*',
        "$env:TEMP*",
        "$env:LOCALAPPDATA\Temp*",
        "$env:LOCALAPPDATA\Microsoft\Windows\INetCache*",
        "$env:LOCALAPPDATA\Microsoft\Windows\WebCache*",
        "$env:LOCALAPPDATA\CrashDumps*",
        "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache*",
        "$env:LOCALAPPDATA\D3DSCache*",
        "$env:LOCALAPPDATA\NVIDIA\DXCache*",
        "$env:LOCALAPPDATA\NVIDIA\GLCache*"
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($pattern in $allowedPatterns) {
        if ($normalized -like $pattern) {
            return $true
        }
    }

    return $false
}

function Remove-FolderContents {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Description,
        [switch]$ContentsOnly
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{
            Success    = $true
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = 'Path does not exist'
        }
    }

    if (-not (Test-SafeCleanupPath -Path $Path)) {
        return @{
            Success    = $false
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = 'Path blocked for security'
        }
    }

    try {
        $items = if ($ContentsOnly) {
            @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)
        }
        else {
            @(Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue)
        }

        $itemCount = $items.Count
        [int64]$sizeBefore = 0

        if ($itemCount -gt 0) {
            $files = if ($ContentsOnly) {
                $items | Where-Object { -not $_.PSIsContainer }
            }
            else {
                $items
            }

            if ($files.Count -gt 0) {
                $sizeSum = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -ne $sizeSum) {
                    $sizeBefore = [int64]$sizeSum
                }
            }
        }

        if ($itemCount -eq 0) {
            return @{
                Success    = $true
                SpaceFreed = [int64]0
                ItemCount  = 0
                Message    = 'Folder is empty'
            }
        }

        if ($ContentsOnly) {
            Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                }
                catch {
                    Write-Log "Could not remove $($_.FullName): $($_.Exception.Message)" 'WARN'
                }
            }
        }
        else {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        }

        return @{
            Success    = $true
            SpaceFreed = $sizeBefore
            ItemCount  = $itemCount
            Message    = 'Successfully cleaned'
        }
    }
    catch {
        return @{
            Success    = $false
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = $_.Exception.Message
        }
    }
}




function Get-FolderSizeInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $info = [ordered]@{
        Path       = $Path
        Exists     = $false
        ItemCount  = 0
        FileCount  = 0
        FolderCount = 0
        SizeBytes  = [int64]0
        SizeMB     = [double]0
        SizeGB     = [double]0
        Message    = $null
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $info.Message = 'Path does not exist'
        return [PSCustomObject]$info
    }

    $info.Exists = $true

    try {
        $items = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue)
        $files = @($items | Where-Object { -not $_.PSIsContainer })
        $folders = @($items | Where-Object { $_.PSIsContainer })
        $sizeBytes = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if ($null -eq $sizeBytes) { $sizeBytes = 0 }

        $info.ItemCount = $items.Count
        $info.FileCount = $files.Count
        $info.FolderCount = $folders.Count
        $info.SizeBytes = [int64]$sizeBytes
        $info.SizeMB = [math]::Round(([double]$info.SizeBytes / 1MB), 2)
        $info.SizeGB = [math]::Round(([double]$info.SizeBytes / 1GB), 3)
        $info.Message = 'Size calculated successfully'
    }
    catch {
        $info.Message = $_.Exception.Message
        Write-Log "Could not calculate folder size for $Path`: $($info.Message)" 'WARN'
    }

    return [PSCustomObject]$info
}

function Stop-WindowsUpdateLockingProcesses {
    [CmdletBinding()]
    param()

    Write-Log 'Checking for Windows Update processes that may lock SoftwareDistribution...' 'INFO'

    $processNames = @(
        'MoUsoCoreWorker',
        'TiWorker',
        'TrustedInstaller',
        'UsoClient',
        'MusNotification',
        'MusNotificationUx',
        'SIHClient'
    )

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($name in $processNames) {
        $processes = @(Get-Process -Name $name -ErrorAction SilentlyContinue)

        if ($processes.Count -eq 0) {
            $results.Add([PSCustomObject]@{
                Name    = $name
                Action  = 'NotRunning'
                Success = $true
                Message = 'Process not running'
            }) | Out-Null
            continue
        }

        foreach ($proc in $processes) {
            try {
                Write-Log "Stopping possible Windows Update lock process: $($proc.ProcessName) PID $($proc.Id)" 'WARN'
                Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                Start-Sleep -Milliseconds 500

                $stillRunning = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
                if ($null -eq $stillRunning) {
                    Write-Log "Stopped process $($proc.ProcessName) PID $($proc.Id)." 'OK'
                    $results.Add([PSCustomObject]@{
                        Name    = $proc.ProcessName
                        ProcessId = $proc.Id
                        Action  = 'Stopped'
                        Success = $true
                        Message = 'Stopped successfully'
                    }) | Out-Null
                }
                else {
                    Write-Log "Process $($proc.ProcessName) PID $($proc.Id) is still running after stop attempt." 'WARN'
                    $results.Add([PSCustomObject]@{
                        Name    = $proc.ProcessName
                        ProcessId = $proc.Id
                        Action  = 'StopAttempted'
                        Success = $false
                        Message = 'Still running after Stop-Process'
                    }) | Out-Null
                }
            }
            catch {
                Write-Log "Could not stop $($proc.ProcessName) PID $($proc.Id): $($_.Exception.Message)" 'WARN'
                $results.Add([PSCustomObject]@{
                    Name    = $proc.ProcessName
                    ProcessId = $proc.Id
                    Action  = 'FailedToStop'
                    Success = $false
                    Message = $_.Exception.Message
                }) | Out-Null
            }
        }
    }

    return @($results)
}

function Get-FolderSizeBytesSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
            return [int64]0
        }

        $sum = (Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer } |
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum

        if ($null -eq $sum) { return [int64]0 }
        return [int64]$sum
    }
    catch {
        Write-Log "Unable to calculate size for ${Path}: $($_.Exception.Message)" 'WARN'
        return [int64]0
    }
}

function Stop-SoftwareDistributionBackupLockingServices {
    [CmdletBinding()]
    param()

    $services = @(
        'wuauserv',
        'bits',
        'cryptsvc',
        'dosvc',
        'UsoSvc',
        'WaaSMedicSvc',
        'TrustedInstaller',
        'msiserver'
    )

    foreach ($svcName in $services) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($null -eq $svc) {
                Write-Log "Service not found while releasing SoftwareDistribution locks: $svcName" 'INFO'
                continue
            }

            if ($svc.Status -ne 'Stopped') {
                Write-Log "Stopping service to release SoftwareDistribution locks: $svcName ($($svc.Status))" 'INFO'
                [void](Stop-ServiceWithValidation -Name $svcName -TimeoutSeconds 30 -KillProcessOnTimeout)
                Start-Sleep -Seconds 2
            }

            $svc.Refresh()
            Write-Log "Service state after stop attempt: $svcName = $($svc.Status)" 'INFO'
        }
        catch {
            Write-Log "Could not stop service ${svcName}: $($_.Exception.Message)" 'WARN'
        }
    }
}

function Stop-SoftwareDistributionBackupLockingProcesses {
    [CmdletBinding()]
    param()

    $processNames = @(
        'TiWorker',
        'TrustedInstaller',
        'MoUsoCoreWorker',
        'UsoClient',
        'wuauclt',
        'bitsadmin',
        'msiexec',
        'MusNotification',
        'MusNotificationUx',
        'SIHClient'
    )

    foreach ($procName in $processNames) {
        try {
            $procs = @(Get-Process -Name $procName -ErrorAction SilentlyContinue)
            foreach ($proc in $procs) {
                Write-Log "Stopping process to release SoftwareDistribution locks: $($proc.ProcessName) PID $($proc.Id)" 'WARN'
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log "Could not stop process ${procName}: $($_.Exception.Message)" 'WARN'
        }
    }
}

function Start-SoftwareDistributionBackupUpdateServices {
    [CmdletBinding()]
    param()

    $services = @('cryptsvc', 'bits', 'wuauserv', 'dosvc', 'UsoSvc')

    foreach ($svcName in $services) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($null -eq $svc) { continue }

            if ($svc.Status -ne 'Running') {
                Write-Log "Restarting update-related service after SoftwareDistribution backup cleanup: $svcName" 'INFO'
                Start-Service -Name $svcName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
            }

            $svc.Refresh()
            Write-Log "Service state after restart attempt: $svcName = $($svc.Status)" 'INFO'
        }
        catch {
            Write-Log "Could not restart service ${svcName}: $($_.Exception.Message)" 'WARN'
        }
    }
}


function New-EmptyMirrorFolderForCleanup {
    [CmdletBinding()]
    param()

    $emptyPath = Join-Path $env:TEMP ('EmptyMirror_{0}' -f ([guid]::NewGuid().ToString('N')))
    New-Item -Path $emptyPath -ItemType Directory -Force | Out-Null
    return $emptyPath
}


function Invoke-NativeProcessWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$ArgumentList,
        [int]$TimeoutSeconds = 60,
        [string]$Description = 'Native process'
    )

    $outputPath = Join-Path $env:TEMP ('NativeProcessOutput_{0}.log' -f ([guid]::NewGuid().ToString('N')))
    $errorPath = Join-Path $env:TEMP ('NativeProcessError_{0}.log' -f ([guid]::NewGuid().ToString('N')))

    $result = [ordered]@{
        FilePath       = $FilePath
        Arguments      = ($ArgumentList -join ' ')
        TimedOut       = $false
        ExitCode       = $null
        Output         = @()
        ErrorOutput    = @()
        Message        = ''
    }

    try {
        Write-Log "$Description starting with timeout ${TimeoutSeconds}s: $FilePath $($ArgumentList -join ' ')" 'INFO'
        Add-DetailedResult -Step 'NativeProcessTimeoutGuard' -Status 'Started' -Message "$Description started." -Data @{
            FilePath = $FilePath
            Arguments = ($ArgumentList -join ' ')
            TimeoutSeconds = $TimeoutSeconds
        }
        Write-YamlLog

        $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -WindowStyle Hidden -RedirectStandardOutput $outputPath -RedirectStandardError $errorPath -PassThru -ErrorAction Stop

        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            $result.TimedOut = $true
            $result.Message = "$Description timed out after ${TimeoutSeconds}s and was terminated."
            Write-Log $result.Message 'WARN'

            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Log "Could not terminate timed-out process $($proc.Id): $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            $result.ExitCode = $proc.ExitCode
            $result.Message = "$Description completed with exit code $($proc.ExitCode)."
            Write-Log $result.Message 'INFO'
        }
    }
    catch {
        $result.Message = "$Description failed to start or monitor: $($_.Exception.Message)"
        Write-Log $result.Message 'WARN'
    }
    finally {
        try {
            if (Test-Path -LiteralPath $outputPath) {
                $result.Output = @(Get-Content -LiteralPath $outputPath -ErrorAction SilentlyContinue)
                Remove-Item -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }

        try {
            if (Test-Path -LiteralPath $errorPath) {
                $result.ErrorOutput = @(Get-Content -LiteralPath $errorPath -ErrorAction SilentlyContinue)
                Remove-Item -LiteralPath $errorPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }

        Add-DetailedResult -Step 'NativeProcessTimeoutGuard' -Status $(if ($result.TimedOut) { 'TimedOut' } elseif ($null -ne $result.ExitCode -and $result.ExitCode -le 7) { 'Completed' } else { 'Warning' }) -Message $result.Message -Data @{
            FilePath = $FilePath
            Arguments = ($ArgumentList -join ' ')
            ExitCode = $result.ExitCode
            TimedOut = $result.TimedOut
        }
        Write-YamlLog
    }

    return [PSCustomObject]$result
}

function Invoke-RobocopyMirrorDeleteFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Description,
        [Parameter(Mandatory = $true)][string]$EmptyMirrorPath,
        [switch]$RecreateAfterDelete,
        [int]$RobocopyTimeoutSeconds = 120,
        [int]$RemoveTimeoutSeconds = 60
    )

    $result = [ordered]@{
        Path             = $Path
        Description      = $Description
        ExistsBefore     = $false
        Deleted          = $false
        Recreated        = $false
        SizeBytesBefore  = [int64]0
        SizeMBBefore     = [double]0
        SizeGBBefore     = [double]0
        ItemCountBefore  = 0
        SpaceFreedBytes  = [int64]0
        SpaceFreedMB     = [double]0
        SpaceFreedGB     = [double]0
        RobocopyExitCode = $null
        RobocopyTimedOut = $false
        RmdirTimedOut    = $false
        Message          = ''
    }

    try {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            $result.Message = 'Path was blank.'
            return [PSCustomObject]$result
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            $result.Message = 'Path did not exist.'
            if ($RecreateAfterDelete) {
                New-Item -Path $Path -ItemType Directory -Force | Out-Null
                $result.Recreated = $true
                $result.Message = 'Path did not exist and was recreated.'
            }
            return [PSCustomObject]$result
        }

        $result.ExistsBefore = $true
        $sizeBytes = Get-FolderSizeBytesSafe -Path $Path
        $result.SizeBytesBefore = [int64]$sizeBytes
        $result.SizeMBBefore = [math]::Round(([double]$sizeBytes / 1MB), 2)
        $result.SizeGBBefore = [math]::Round(([double]$sizeBytes / 1GB), 3)
        $result.ItemCountBefore = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue).Count

        Write-Log "Preparing to clean $Description`: $Path | Size before cleanup: $($result.SizeGBBefore) GB ($($result.SizeMBBefore) MB) | Items: $($result.ItemCountBefore)" 'WARN'
        Add-DetailedResult -Step 'FolderCleanup' -Status 'Started' -Message "Preparing to clean $Description." -Data @{
            Path = $Path
            Description = $Description
            SizeMBBefore = $result.SizeMBBefore
            ItemCountBefore = $result.ItemCountBefore
        }
        Write-YamlLog

        try {
            $attribResult = Invoke-NativeProcessWithTimeout -FilePath 'cmd.exe' -ArgumentList @('/d','/c',"attrib -r -s -h `"$Path`" /s /d") -TimeoutSeconds 45 -Description "Attribute clear for $Description"
            if ($attribResult.TimedOut) {
                Write-Log "Attribute clear timed out for $Path. Continuing to deletion fallback steps." 'WARN'
            }
        }
        catch {
            Write-Log "Could not clear attributes on ${Path}: $($_.Exception.Message)" 'WARN'
        }

        if ($result.ItemCountBefore -eq 0) {
            Write-Log "Folder appears empty; attempting fast empty-directory delete first: $Path" 'INFO'
            try {
                [System.IO.Directory]::Delete($Path, $false)
            }
            catch {
                Write-Log "Fast empty-directory delete did not complete for ${Path}: $($_.Exception.Message). Continuing with guarded robocopy/rmdir cleanup." 'WARN'
            }
        }

        if (Test-Path -LiteralPath $Path) {
            $robocopyArgs = @(
                $EmptyMirrorPath,
                $Path,
                '/MIR',
                '/R:0',
                '/W:0',
                '/XJ',
                '/NP',
                '/NFL',
                '/NDL'
            )

            $robocopyResult = Invoke-NativeProcessWithTimeout -FilePath 'robocopy.exe' -ArgumentList $robocopyArgs -TimeoutSeconds $RobocopyTimeoutSeconds -Description "Robocopy mirror cleanup for $Description"
            $result.RobocopyExitCode = $robocopyResult.ExitCode
            $result.RobocopyTimedOut = [bool]$robocopyResult.TimedOut

            foreach ($line in @($robocopyResult.Output + $robocopyResult.ErrorOutput)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
                    Write-Log "robocopy: $line" 'INFO'
                }
            }

            if ($result.RobocopyTimedOut) {
                Write-Log "Robocopy timed out for $Path. Continuing to guarded rmdir fallback." 'WARN'
            }
            elseif ($null -ne $result.RobocopyExitCode -and $result.RobocopyExitCode -gt 7) {
                Write-Log "Robocopy returned failure exit code $($result.RobocopyExitCode) for $Path. Continuing to guarded rmdir fallback." 'WARN'
            }
        }

        if (Test-Path -LiteralPath $Path) {
            Write-Log "Trying guarded cmd.exe rmdir fallback for: $Path" 'WARN'
            $rmdirResult = Invoke-NativeProcessWithTimeout -FilePath 'cmd.exe' -ArgumentList @('/d','/c',"rmdir /s /q `"$Path`"") -TimeoutSeconds $RemoveTimeoutSeconds -Description "rmdir fallback for $Description"
            $result.RmdirTimedOut = [bool]$rmdirResult.TimedOut
        }

        if (Test-Path -LiteralPath $Path) {
            Write-Log "Trying final PowerShell Remove-Item fallback for: $Path" 'WARN'
            try {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            }
            catch {
                Write-Log "Final Remove-Item fallback reported an issue for ${Path}: $($_.Exception.Message)" 'WARN'
            }
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            $result.Deleted = $true
            $result.SpaceFreedBytes = [int64]$sizeBytes
            $result.SpaceFreedMB = [math]::Round(([double]$sizeBytes / 1MB), 2)
            $result.SpaceFreedGB = [math]::Round(([double]$sizeBytes / 1GB), 3)
            $result.Message = 'Deleted with guarded cleanup logic.'
            Write-Log "Deleted $Description`: $Path | Estimated freed: $($result.SpaceFreedGB) GB ($($result.SpaceFreedMB) MB)" 'OK'
        }
        else {
            $result.Message = 'Folder still exists after guarded robocopy, guarded rmdir, and Remove-Item fallback. Continuing script instead of hanging.'
            Write-Log "FAILED to delete $Description but continuing script: $Path" 'ERROR'
        }

        Add-DetailedResult -Step 'FolderCleanup' -Status $(if ($result.Deleted) { 'Success' } else { 'Warning' }) -Message $result.Message -Data @{
            Path = $Path
            Description = $Description
            Deleted = $result.Deleted
            RobocopyExitCode = $result.RobocopyExitCode
            RobocopyTimedOut = $result.RobocopyTimedOut
            RmdirTimedOut = $result.RmdirTimedOut
            SpaceFreedMB = $result.SpaceFreedMB
        }
        Write-YamlLog

        if ($RecreateAfterDelete) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
            $result.Recreated = $true
            Write-Log "Recreated required folder: $Path" 'OK'
            Write-YamlLog
        }
    }
    catch {
        $result.Message = $_.Exception.Message
        Write-Log "Guarded folder cleanup failed for ${Path}: $($_.Exception.Message). Continuing script." 'ERROR'
        Add-DetailedResult -Step 'FolderCleanup' -Status 'Error' -Message $result.Message -Data @{
            Path = $Path
            Description = $Description
        }
        Write-YamlLog
    }

    return [PSCustomObject]$result
}

function Remove-SoftwareDistributionMainFolder {
    [CmdletBinding()]
    param()

    Write-Log 'Cleaning active C:\Windows\SoftwareDistribution folder with robocopy mirror method.' 'INFO'

    $emptyPath = $null
    try {
        $emptyPath = New-EmptyMirrorFolderForCleanup
        return Invoke-RobocopyMirrorDeleteFolder -Path (Join-Path $env:WINDIR 'SoftwareDistribution') -Description 'Windows Update SoftwareDistribution folder' -EmptyMirrorPath $emptyPath -RecreateAfterDelete
    }
    finally {
        if ($emptyPath -and (Test-Path -LiteralPath $emptyPath)) {
            Remove-Item -LiteralPath $emptyPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Remove-SoftwareDistributionBakFolders {
    [CmdletBinding()]
    param(
        [int]$MaxParallelDeletes = $MaxParallelSoftwareDistributionDeletes,
        [int]$CleanupTimeLimitMinutes = $SoftwareDistributionCleanupTimeLimitMinutes,
        [int]$FolderJobTimeoutMinutes = $SoftwareDistributionFolderJobTimeoutMinutes,
        [string]$StatePath = $SoftwareDistributionCleanupStatePath
    )

    Write-Log 'Cleaning SoftwareDistribution.bak* folders with v3.0 resumable capped parallel cleanup.' 'INFO'

    if ($MaxParallelDeletes -lt 1) { $MaxParallelDeletes = 1 }
    if ($MaxParallelDeletes -gt 4) {
        Write-Log "MaxParallelDeletes was set to $MaxParallelDeletes. Capping at 4 to avoid disk thrashing." 'WARN'
        $MaxParallelDeletes = 4
    }
    if ($CleanupTimeLimitMinutes -lt 5) {
        Write-Log "CleanupTimeLimitMinutes was set to $CleanupTimeLimitMinutes. Raising to 5 minutes." 'WARN'
        $CleanupTimeLimitMinutes = 5
    }
    if ($FolderJobTimeoutMinutes -lt 3) {
        Write-Log "FolderJobTimeoutMinutes was set to $FolderJobTimeoutMinutes. Raising to 3 minutes." 'WARN'
        $FolderJobTimeoutMinutes = 3
    }

    $cleanupStart = Get-Date
    $deadline = $cleanupStart.AddMinutes($CleanupTimeLimitMinutes)
    $jobTimeoutSeconds = [int]($FolderJobTimeoutMinutes * 60)
    $robocopyTimeoutSeconds = [Math]::Max(120, [int]($jobTimeoutSeconds - 60))
    $rmdirTimeoutSeconds = 120

    $stateDirectory = Split-Path -Path $StatePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($stateDirectory) -and -not (Test-Path -LiteralPath $stateDirectory)) {
        New-Item -Path $stateDirectory -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    }

    function New-SoftwareDistributionCleanupState {
        return [PSCustomObject]@{
            Version = 3
            ComputerName = $env:COMPUTERNAME
            LastUpdated = (Get-Date).ToString('o')
            Pending = @()
            Completed = @()
            Failed = @()
            Deferred = @()
        }
    }

    function Read-SoftwareDistributionCleanupState {
        param([string]$Path)
        try {
            if (Test-Path -LiteralPath $Path) {
                $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $loaded = $raw | ConvertFrom-Json -ErrorAction Stop
                    if ($null -ne $loaded) { return $loaded }
                }
            }
        }
        catch {
            Write-Log "Could not read SoftwareDistribution cleanup state file. A new state file will be created. Error: $($_.Exception.Message)" 'WARN'
        }
        return New-SoftwareDistributionCleanupState
    }

    function Save-SoftwareDistributionCleanupState {
        param(
            [Parameter(Mandatory = $true)]$State,
            [Parameter(Mandatory = $true)][string]$Path
        )
        try {
            $State.LastUpdated = (Get-Date).ToString('o')
            $State | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8 -Force
        }
        catch {
            Write-Log "Could not write SoftwareDistribution cleanup state file: $($_.Exception.Message)" 'WARN'
        }
    }

    function Get-StateArray {
        param($Value)
        if ($null -eq $Value) { return @() }
        return @($Value)
    }

    function Add-StatePathUnique {
        param(
            [Parameter(Mandatory = $true)]$State,
            [Parameter(Mandatory = $true)][string]$PropertyName,
            [Parameter(Mandatory = $true)][string]$Path
        )
        $existing = @(Get-StateArray $State.$PropertyName | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($existing -notcontains $Path) { $existing += $Path }
        $State.$PropertyName = @($existing | Sort-Object -Unique)
    }

    function Remove-StatePath {
        param(
            [Parameter(Mandatory = $true)]$State,
            [Parameter(Mandatory = $true)][string]$PropertyName,
            [Parameter(Mandatory = $true)][string]$Path
        )
        $State.$PropertyName = @(Get-StateArray $State.$PropertyName | Where-Object { [string]$_ -ne $Path })
    }

    $state = Read-SoftwareDistributionCleanupState -Path $StatePath
    $basePaths = @('C:\Windows', 'C:\')
    $discoveredTargets = @()

    foreach ($basePath in $basePaths) {
        try {
            if (Test-Path -LiteralPath $basePath) {
                $discoveredTargets += @(
                    Get-ChildItem -LiteralPath $basePath -Directory -Force -Filter 'SoftwareDistribution.bak*' -ErrorAction Stop |
                        Where-Object { $_.Name -match '^SoftwareDistribution\.bak.*$' } |
                        Select-Object -ExpandProperty FullName
                )
            }
        }
        catch {
            Write-Log "Could not scan $basePath for SoftwareDistribution.bak* folders: $($_.Exception.Message)" 'WARN'
        }
    }

    $pendingFromState = @(Get-StateArray $state.Pending)
    $deferredFromState = @(Get-StateArray $state.Deferred)
    $candidateTargets = @($discoveredTargets + $pendingFromState + $deferredFromState |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        Sort-Object -Unique)

    $targets = @()
    foreach ($candidate in $candidateTargets) {
        if (Test-Path -LiteralPath $candidate) {
            $targets += $candidate
        }
        else {
            Remove-StatePath -State $state -PropertyName 'Pending' -Path $candidate
            Remove-StatePath -State $state -PropertyName 'Deferred' -Path $candidate
            Add-StatePathUnique -State $state -PropertyName 'Completed' -Path $candidate
        }
    }
    $targets = @($targets | Sort-Object -Unique)

    $state.Pending = @($targets)
    $state.Deferred = @()
    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath

    if ($targets.Count -eq 0) {
        Write-Log 'No SoftwareDistribution.bak* folders were found or all tracked folders were already removed.' 'INFO'
        Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status 'Success' -Message 'No SoftwareDistribution.bak* folders found.' -Data @{ StatePath = $StatePath }
        Write-YamlLog
        return [PSCustomObject]@{
            Success         = $true
            SpaceFreed      = [int64]0
            SpaceFreedBytes = [int64]0
            SpaceFreedMB    = [double]0
            SpaceFreedGB    = [double]0
            ItemCount       = 0
            DeletedCount    = 0
            FailedCount     = 0
            TimedOutCount   = 0
            DeferredCount   = 0
            Message         = 'No SoftwareDistribution.bak* folders found'
        }
    }

    Write-Log "Found/tracked $($targets.Count) SoftwareDistribution.bak* folder(s). Max parallel deletes: $MaxParallelDeletes. Cleanup budget: $CleanupTimeLimitMinutes minute(s). Per-folder job timeout: $FolderJobTimeoutMinutes minute(s)." 'WARN'
    Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status 'Started' -Message 'Starting v3.0 resumable capped parallel SoftwareDistribution.bak* cleanup.' -Data @{
        TargetCount = $targets.Count
        MaxParallelDeletes = $MaxParallelDeletes
        CleanupTimeLimitMinutes = $CleanupTimeLimitMinutes
        FolderJobTimeoutMinutes = $FolderJobTimeoutMinutes
        StatePath = $StatePath
        Deadline = $deadline
    }
    Write-YamlLog

    [int64]$totalFreedBytes = 0
    [int]$deletedCount = 0
    [int]$failedCount = 0
    [int]$timedOutCount = 0
    [int]$deferredCount = 0
    [int]$totalItems = 0
    $failureMessages = @()
    $script:SDv3TotalFreedBytes = [int64]0
    $script:SDv3DeletedCount = 0
    $script:SDv3FailedCount = 0
    $script:SDv3TimedOutCount = 0
    $script:SDv3DeferredCount = 0
    $script:SDv3TotalItems = 0
    $script:SDv3FailureMessages = @()
    $jobs = @()
    $jobPathById = @{}
    $jobStartedById = @{}
    $queue = New-Object System.Collections.Queue
    foreach ($target in $targets) { [void]$queue.Enqueue($target) }

    Stop-SoftwareDistributionBackupLockingServices
    Stop-SoftwareDistributionBackupLockingProcesses

    $jobScript = {
        param(
            [string]$TargetPath,
            [int]$RobocopyTimeoutSeconds,
            [int]$RmdirTimeoutSeconds
        )

        $result = [ordered]@{
            Path = $TargetPath
            Deleted = $false
            TimedOut = $false
            RobocopyTimedOut = $false
            RmdirTimedOut = $false
            RobocopyExitCode = $null
            SizeBytesBefore = [int64]0
            SizeMBBefore = [double]0
            SizeGBBefore = [double]0
            ItemCountBefore = 0
            Message = ''
            Started = (Get-Date).ToString('o')
            Finished = $null
        }

        function Invoke-NativeProcessWithTimeoutForSdCleanup {
            param(
                [Parameter(Mandatory = $true)][string]$FilePath,
                [Parameter(Mandatory = $true)][string[]]$ArgumentList,
                [Parameter(Mandatory = $true)][int]$TimeoutSeconds
            )

            $p = $null
            $outputFile = Join-Path $env:TEMP ('sdcleanup_out_' + [guid]::NewGuid().ToString('N') + '.txt')
            $errorFile = Join-Path $env:TEMP ('sdcleanup_err_' + [guid]::NewGuid().ToString('N') + '.txt')
            $nativeResult = [ordered]@{ ExitCode = $null; TimedOut = $false; ErrorMessage = '' }

            try {
                $p = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -NoNewWindow -PassThru -RedirectStandardOutput $outputFile -RedirectStandardError $errorFile
                if (-not $p.WaitForExit($TimeoutSeconds * 1000)) {
                    $nativeResult.TimedOut = $true
                    try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { }
                }
                else {
                    $nativeResult.ExitCode = $p.ExitCode
                }
            }
            catch {
                $nativeResult.ErrorMessage = $_.Exception.Message
            }
            finally {
                Remove-Item -LiteralPath $outputFile -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $errorFile -Force -ErrorAction SilentlyContinue
            }

            return [PSCustomObject]$nativeResult
        }

        try {
            if ([string]::IsNullOrWhiteSpace($TargetPath)) {
                $result.Message = 'Target path was blank.'
                return [PSCustomObject]$result
            }

            if (-not (Test-Path -LiteralPath $TargetPath)) {
                $result.Deleted = $true
                $result.Message = 'Target folder was already gone.'
                return [PSCustomObject]$result
            }

            try {
                $files = @(Get-ChildItem -LiteralPath $TargetPath -Force -File -ErrorAction SilentlyContinue)
                foreach ($file in $files) { $result.SizeBytesBefore += [int64]$file.Length; $result.ItemCountBefore++ }
                $result.SizeMBBefore = [math]::Round(([double]$result.SizeBytesBefore / 1MB), 2)
                $result.SizeGBBefore = [math]::Round(([double]$result.SizeBytesBefore / 1GB), 3)
            }
            catch { }

            $emptyMirror = Join-Path $env:TEMP ('EmptyDeleteSource_' + [guid]::NewGuid().ToString('N'))
            New-Item -Path $emptyMirror -ItemType Directory -Force | Out-Null

            try {
                $roboArgs = @($emptyMirror, $TargetPath, '/MIR', '/R:0', '/W:0', '/XJ', '/NP', '/NFL', '/NDL', '/NJH', '/NJS', '/NC', '/NS')
                $robo = Invoke-NativeProcessWithTimeoutForSdCleanup -FilePath 'robocopy.exe' -ArgumentList $roboArgs -TimeoutSeconds $RobocopyTimeoutSeconds
                $result.RobocopyExitCode = $robo.ExitCode
                $result.RobocopyTimedOut = [bool]$robo.TimedOut
                if ($robo.TimedOut) { $result.TimedOut = $true }

                if (Test-Path -LiteralPath $TargetPath) {
                    $rm = Invoke-NativeProcessWithTimeoutForSdCleanup -FilePath 'cmd.exe' -ArgumentList @('/d','/c', ('rmdir /s /q "{0}"' -f $TargetPath)) -TimeoutSeconds $RmdirTimeoutSeconds
                    $result.RmdirTimedOut = [bool]$rm.TimedOut
                    if ($rm.TimedOut) { $result.TimedOut = $true }
                }

                if (Test-Path -LiteralPath $TargetPath) {
                    try { [System.IO.Directory]::Delete($TargetPath, $true) } catch { }
                }
            }
            finally {
                Remove-Item -LiteralPath $emptyMirror -Recurse -Force -ErrorAction SilentlyContinue
            }

            if (-not (Test-Path -LiteralPath $TargetPath)) {
                $result.Deleted = $true
                $result.Message = 'Deleted by v3.0 resumable parallel cleanup job.'
            }
            elseif ($result.TimedOut) {
                $result.Message = 'Timed out before the folder could be fully deleted. It will be retried next run.'
            }
            else {
                $result.Message = 'Folder still exists after cleanup attempt. It will be retried next run.'
            }
        }
        catch {
            $result.Message = $_.Exception.Message
        }
        finally {
            $result.Finished = (Get-Date).ToString('o')
        }

        return [PSCustomObject]$result
    }

    function Receive-FinishedSdCleanupJobs {
        param([switch]$Force)

        $finishedJobs = @()
        foreach ($job in @($jobs)) {
            if ($Force -or $job.State -ne 'Running') { $finishedJobs += $job }
        }

        foreach ($job in $finishedJobs) {
            $targetPath = ''
            if ($jobPathById.ContainsKey($job.Id)) { $targetPath = [string]$jobPathById[$job.Id] }
            try {
                if ($job.State -eq 'Running' -and $Force) {
                    Stop-Job -Job $job -ErrorAction SilentlyContinue
                    $script:SDv3TimedOutCount++
                    $script:SDv3DeferredCount++
                    if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $targetPath
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                        Write-Log "Stopped SoftwareDistribution cleanup job due to cleanup window ending: $targetPath" 'WARN'
                    }
                    continue
                }

                $jobResults = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
                if ($jobResults.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($targetPath)) {
                    $script:SDv3FailedCount++
                    $script:SDv3FailureMessages += "No result returned for $targetPath"
                    Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                    Write-Log "SoftwareDistribution cleanup job returned no result: $targetPath" 'WARN'
                    continue
                }

                foreach ($jobResult in $jobResults) {
                    if ($null -eq $jobResult) { continue }
                    $path = [string]$jobResult.Path
                    if ([string]::IsNullOrWhiteSpace($path)) { $path = $targetPath }

                    if ($jobResult.Deleted) {
                        $script:SDv3DeletedCount++
                        $script:SDv3TotalFreedBytes += [int64]$jobResult.SizeBytesBefore
                        $script:SDv3TotalItems += [int]$jobResult.ItemCountBefore
                        Remove-StatePath -State $state -PropertyName 'Pending' -Path $path
                        Remove-StatePath -State $state -PropertyName 'Deferred' -Path $path
                        Add-StatePathUnique -State $state -PropertyName 'Completed' -Path $path
                        Write-Log "Deleted SoftwareDistribution backup folder: $path" 'OK'
                    }
                    else {
                        if ($jobResult.TimedOut -or $jobResult.RobocopyTimedOut -or $jobResult.RmdirTimedOut) { $script:SDv3TimedOutCount++ }
                        $script:SDv3FailedCount++
                        $script:SDv3FailureMessages += "Failed to delete ${path}: $($jobResult.Message)"
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $path
                        Write-Log "SoftwareDistribution backup folder remains and will be retried: $path | $($jobResult.Message)" 'WARN'
                    }

                    $statusText = 'Warning'
                    if ($jobResult.Deleted) { $statusText = 'Success' }
                    Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status $statusText -Message $jobResult.Message -Data @{
                        Path = $path
                        Deleted = $jobResult.Deleted
                        TimedOut = $jobResult.TimedOut
                        RobocopyExitCode = $jobResult.RobocopyExitCode
                        RobocopyTimedOut = $jobResult.RobocopyTimedOut
                        RmdirTimedOut = $jobResult.RmdirTimedOut
                        StatePath = $StatePath
                    }
                    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                    Write-YamlLog
                }
            }
            catch {
                $script:SDv3FailedCount++
                $script:SDv3FailureMessages += "Failed to receive delete job result for ${targetPath}: $($_.Exception.Message)"
                if (-not [string]::IsNullOrWhiteSpace($targetPath)) { Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath }
                Write-Log "Failed to receive SoftwareDistribution delete job result: $($_.Exception.Message)" 'WARN'
            }
            finally {
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                $jobPathById.Remove($job.Id)
                $jobStartedById.Remove($job.Id)
            }
        }

        $remainingJobs = @()
        foreach ($job in @($jobs)) {
            if ($finishedJobs -notcontains $job) { $remainingJobs += $job }
        }
        $script:__SdCleanupRemainingJobs = $remainingJobs
    }

    try {
        while ($true) {
            Receive-FinishedSdCleanupJobs
            $jobs = @($script:__SdCleanupRemainingJobs)

            $now = Get-Date
            foreach ($job in @($jobs)) {
                if ($jobStartedById.ContainsKey($job.Id)) {
                    $ageSeconds = ($now - [datetime]$jobStartedById[$job.Id]).TotalSeconds
                    if ($ageSeconds -gt $jobTimeoutSeconds) {
                        $targetPath = [string]$jobPathById[$job.Id]
                        Stop-Job -Job $job -ErrorAction SilentlyContinue
                        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                        $jobs = @($jobs | Where-Object { $_.Id -ne $job.Id })
                        $jobPathById.Remove($job.Id)
                        $jobStartedById.Remove($job.Id)
                        $script:SDv3TimedOutCount++
                        $script:SDv3DeferredCount++
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $targetPath
                        Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                        Write-Log "SoftwareDistribution delete job exceeded $FolderJobTimeoutMinutes minute timeout and was stopped: $targetPath" 'WARN'
                    }
                }
            }

            if ((Get-Date) -ge $deadline) {
                while ($queue.Count -gt 0) {
                    $deferredPath = [string]$queue.Dequeue()
                    if (Test-Path -LiteralPath $deferredPath) {
                        $script:SDv3DeferredCount++
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $deferredPath
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $deferredPath
                    }
                }

                foreach ($job in @($jobs)) {
                    $runningPath = ''
                    if ($jobPathById.ContainsKey($job.Id)) { $runningPath = [string]$jobPathById[$job.Id] }
                    if (-not [string]::IsNullOrWhiteSpace($runningPath)) {
                        $script:SDv3DeferredCount++
                        Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $runningPath
                        Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $runningPath
                    }
                    Stop-Job -Job $job -ErrorAction SilentlyContinue
                    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                }
                $jobs = @()
                Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                Write-Log "SoftwareDistribution cleanup time budget reached. Deferred folders will resume next run. Deferred count this run: $($script:SDv3DeferredCount)" 'WARN'
                break
            }

            $runningCount = @($jobs | Where-Object { $_.State -eq 'Running' }).Count
            while ($queue.Count -gt 0 -and $runningCount -lt $MaxParallelDeletes) {
                $target = [string]$queue.Dequeue()
                if (-not (Test-Path -LiteralPath $target)) {
                    Remove-StatePath -State $state -PropertyName 'Pending' -Path $target
                    Add-StatePathUnique -State $state -PropertyName 'Completed' -Path $target
                    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                    Write-Log "SoftwareDistribution.bak* folder no longer exists: $target" 'INFO'
                    continue
                }

                Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $target
                Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
                Write-Log "Starting v3.0 timeout-safe delete job for: $target" 'INFO'
                Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status 'Queued' -Message 'Started timeout-safe resumable delete job.' -Data @{ Path = $target; StatePath = $StatePath }
                Write-YamlLog

                $jobName = 'SDCleanup_' + ([guid]::NewGuid().ToString('N'))
                $newJob = Start-Job -Name $jobName -ScriptBlock $jobScript -ArgumentList $target, $robocopyTimeoutSeconds, $rmdirTimeoutSeconds
                $jobs += $newJob
                $jobPathById[$newJob.Id] = $target
                $jobStartedById[$newJob.Id] = Get-Date
                $runningCount = @($jobs | Where-Object { $_.State -eq 'Running' }).Count
            }

            $runningCount = @($jobs | Where-Object { $_.State -eq 'Running' }).Count
            if ($queue.Count -eq 0 -and $runningCount -eq 0) { break }
            Start-Sleep -Seconds 2
        }

        Receive-FinishedSdCleanupJobs -Force
        $jobs = @($script:__SdCleanupRemainingJobs)
    }
    finally {
        foreach ($job in @($jobs)) {
            try {
                $targetPath = ''
                if ($jobPathById.ContainsKey($job.Id)) { $targetPath = [string]$jobPathById[$job.Id] }
                if ($job.State -eq 'Running') { Stop-Job -Job $job -ErrorAction SilentlyContinue }
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
                    Add-StatePathUnique -State $state -PropertyName 'Pending' -Path $targetPath
                    Add-StatePathUnique -State $state -PropertyName 'Deferred' -Path $targetPath
                }
            }
            catch { }
        }
        Remove-Variable -Name __SdCleanupRemainingJobs -Scope Script -ErrorAction SilentlyContinue
        Save-SoftwareDistributionCleanupState -State $state -Path $StatePath
        Start-SoftwareDistributionBackupUpdateServices
    }

    $remainingPending = @(Get-StateArray $state.Pending | Where-Object { Test-Path -LiteralPath ([string]$_) })
    $state.Pending = @($remainingPending | Sort-Object -Unique)
    Save-SoftwareDistributionCleanupState -State $state -Path $StatePath

    $totalFreedBytes = [int64]$script:SDv3TotalFreedBytes
    $deletedCount = [int]$script:SDv3DeletedCount
    $failedCount = [int]$script:SDv3FailedCount
    $timedOutCount = [int]$script:SDv3TimedOutCount
    $deferredCount = [int]$script:SDv3DeferredCount
    $totalItems = [int]$script:SDv3TotalItems
    $failureMessages = @($script:SDv3FailureMessages)
    $freedMB = [math]::Round(([double]$totalFreedBytes / 1MB), 2)
    $freedGB = [math]::Round(([double]$totalFreedBytes / 1GB), 3)
    if ($state.Pending.Count -gt 0 -and $deferredCount -eq 0) { $deferredCount = $state.Pending.Count }
    $success = ($failedCount -eq 0 -and $state.Pending.Count -eq 0)

    if ($success) {
        Write-Log "SoftwareDistribution.bak* v3.0 parallel cleanup completed. Deleted folders: $deletedCount. Remaining pending: 0." 'OK'
    }
    else {
        Write-Log "SoftwareDistribution.bak* v3.0 cleanup completed with retryable items. Deleted: $deletedCount. Failed attempts: $failedCount. Timed out: $timedOutCount. Pending for next run: $($state.Pending.Count)." 'WARN'
    }

    $finalStatus = 'Warning'
    if ($success) { $finalStatus = 'Success' }
    Add-DetailedResult -Step 'SoftwareDistributionParallelCleanup' -Status $finalStatus -Message 'SoftwareDistribution.bak* v3.0 resumable parallel cleanup finished.' -Data @{
        DeletedCount = $deletedCount
        FailedCount = $failedCount
        TimedOutCount = $timedOutCount
        DeferredCount = $deferredCount
        PendingCount = $state.Pending.Count
        SpaceFreedMB = $freedMB
        SpaceFreedGB = $freedGB
        StatePath = $StatePath
    }
    Write-YamlLog

    $message = 'SoftwareDistribution.bak* v3.0 parallel cleanup completed'
    if (-not $success) {
        $message = (($failureMessages + @("PendingCount=$($state.Pending.Count)", "TimedOutCount=$timedOutCount", "StatePath=$StatePath")) -join '; ')
    }

    return [PSCustomObject]@{
        Success         = $success
        SpaceFreed      = $totalFreedBytes
        SpaceFreedBytes = $totalFreedBytes
        SpaceFreedMB    = $freedMB
        SpaceFreedGB    = $freedGB
        ItemCount       = $totalItems
        DeletedCount    = $deletedCount
        FailedCount     = $failedCount
        TimedOutCount   = $timedOutCount
        DeferredCount   = $deferredCount
        PendingCount    = $state.Pending.Count
        StatePath       = $StatePath
        Message         = $message
    }
}

function Invoke-WindowsCleanup {
    param([int]$TimeoutSec = 300)

    try {
        $cleanmgrPath = Join-Path $env:SystemRoot 'System32\cleanmgr.exe'
        if (-not (Test-Path -LiteralPath $cleanmgrPath)) {
            throw 'Windows Disk Cleanup utility not found.'
        }

        Write-Log "Starting Windows Disk Cleanup (timeout: ${TimeoutSec}s)..." 'INFO'

        $job = Start-Job -ScriptBlock {
            Start-Process -FilePath "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList '/SAGERUN:1','/VERYLOWDISK' -NoNewWindow -Wait -PassThru
        }

        $result = Wait-Job -Job $job -Timeout $TimeoutSec

        if ($null -eq $result -or $job.State -eq 'Running') {
            Stop-Job -Job $job -Force | Out-Null
            Remove-Job -Job $job -Force | Out-Null
            throw "Disk Cleanup timed out after $TimeoutSec seconds."
        }

        $proc = Receive-Job -Job $job
        Remove-Job -Job $job -Force | Out-Null

        return @{
            Success  = $true
            ExitCode = $proc.ExitCode
        }
    }
    catch {
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Invoke-TempCleanup {
    [int64]$totalSpaceFreed = 0
    $cleanupResults = New-Object System.Collections.Generic.List[object]
    $initialSpace = Get-DiskSpaceInfo -Path $env:SystemDrive

    Write-Log 'Cleaning temporary files and caches...' 'INFO'

    $windowsCleanup = Invoke-WindowsCleanup -TimeoutSec 300
    if ($windowsCleanup.Success) {
        Write-Log 'Windows Disk Cleanup completed successfully.' 'OK'
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'cleanmgr.exe'
            Description = 'Windows Disk Cleanup'
            ItemCount   = 0
            SpaceFreed  = [int64]0
            Status      = 'Success'
            Message     = "Exit code $($windowsCleanup.ExitCode)"
        }) | Out-Null
    }
    else {
        Warn-Step -Name 'TempCleanup' -Reason "Windows Disk Cleanup failed: $($windowsCleanup.Error)"
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'cleanmgr.exe'
            Description = 'Windows Disk Cleanup'
            ItemCount   = 0
            SpaceFreed  = [int64]0
            Status      = 'Warning'
            Message     = $windowsCleanup.Error
        }) | Out-Null
    }

    $cleanupTargets = New-Object System.Collections.Generic.List[object]

    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Lab Update Scripts'; Description = 'Lab Update Scripts'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\ProgramData\Win11UpgradeStage'; Description = 'Windows 11 Upgrade Staging'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\SWSetup'; Description = 'HP Software Setup'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\HP_Bios_Config'; Description = 'HP BIOS Config Staging'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Temp'; Description = 'System Temp'; ContentsOnly = $false })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Temp'; Description = 'Windows Temp'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = $env:TEMP; Description = 'User Temp'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Temp"; Description = 'Local Temp'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\SoftwareDistribution\Download'; Description = 'Windows Update Cache'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Prefetch'; Description = 'Windows Prefetch'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Logs\CBS'; Description = 'CBS Logs'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"; Description = 'Internet Cache'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache"; Description = 'Web Cache'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\ProgramData\Microsoft\Windows\WER\ReportQueue'; Description = 'Error Report Queue'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\CrashDumps"; Description = 'Crash Dumps'; ContentsOnly = $true })
    [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"; Description = 'Delivery Optimization Cache'; ContentsOnly = $true })

    if ($AggressiveCleanup) {
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\D3DSCache"; Description = 'Direct3D Shader Cache'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\NVIDIA\DXCache"; Description = 'NVIDIA DX Cache'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\NVIDIA\GLCache"; Description = 'NVIDIA GL Cache'; ContentsOnly = $true })
    }

    foreach ($target in $cleanupTargets) {
        if ($null -eq $target -or [string]::IsNullOrWhiteSpace($target.Path)) {
            Write-Log 'Skipping cleanup target because the path is blank.' 'WARN'
            continue
        }

        Write-Log "Cleaning $($target.Description) at $($target.Path)" 'INFO'
        $result = Remove-FolderContents -Path $target.Path -Description $target.Description -ContentsOnly:([bool]$target.ContentsOnly)

        if ($result.Success) {
            if ($result.SpaceFreed -gt 0) {
                $totalSpaceFreed += [int64]$result.SpaceFreed
                $sizeText = if ($result.SpaceFreed -ge 1GB) {
                    '{0} GB' -f [math]::Round($result.SpaceFreed / 1GB, 2)
                }
                else {
                    '{0} MB' -f [math]::Round($result.SpaceFreed / 1MB, 1)
                }
                Write-Log "Cleaned $($target.Description): $sizeText freed across $($result.ItemCount) item(s)." 'OK'
            }
            else {
                Write-Log "$($target.Description): $($result.Message)" 'INFO'
            }

            $cleanupResults.Add([PSCustomObject]@{
                Path        = $target.Path
                Description = $target.Description
                ItemCount   = $result.ItemCount
                SpaceFreed  = [int64]$result.SpaceFreed
                Status      = 'Success'
                Message     = $result.Message
            }) | Out-Null
        }
        else {
            Warn-Step -Name 'TempCleanup' -Reason "$($target.Description) failed: $($result.Message)"
            $cleanupResults.Add([PSCustomObject]@{
                Path        = $target.Path
                Description = $target.Description
                ItemCount   = 0
                SpaceFreed  = [int64]0
                Status      = 'Failed'
                Message     = $result.Message
            }) | Out-Null
        }
    }

    Write-Log 'Cleaning SoftwareDistribution backup folders at C:\Windows and C:\ root backup variants' 'INFO'
    $sdBackupResult = Remove-SoftwareDistributionBakFolders
    if ($sdBackupResult.Success) {
        if ($sdBackupResult.SpaceFreed -gt 0) {
            $totalSpaceFreed += [int64]$sdBackupResult.SpaceFreed
            Write-Log "Cleaned SoftwareDistribution Backup Folders: $($sdBackupResult.SpaceFreedGB) GB ($($sdBackupResult.SpaceFreedMB) MB) freed across $($sdBackupResult.ItemCount) item(s)." 'OK'
        }
        else {
            Write-Log "SoftwareDistribution Backup Folders: $($sdBackupResult.Message)" 'INFO'
        }

        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'C:\Windows and C:\ SoftwareDistribution backup variants'
            Description = 'SoftwareDistribution Backup Folders'
            ItemCount   = $sdBackupResult.ItemCount
            SpaceFreed  = [int64]$sdBackupResult.SpaceFreed
            Status      = 'Success'
            Message     = $sdBackupResult.Message
        }) | Out-Null
    }
    else {
        Warn-Step -Name 'TempCleanup' -Reason "SoftwareDistribution Backup Folders failed: $($sdBackupResult.Message)"
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'C:\Windows and C:\ SoftwareDistribution backup variants'
            Description = 'SoftwareDistribution Backup Folders'
            ItemCount   = $sdBackupResult.ItemCount
            SpaceFreed  = [int64]0
            Status      = 'Warning'
            Message     = $sdBackupResult.Message
        }) | Out-Null
    }

    $finalSpace = Get-DiskSpaceInfo -Path $env:SystemDrive
    $actualFreed = [int64]0
    if ($initialSpace -and $finalSpace) {
        $actualFreed = [int64]($finalSpace.FreeSpace - $initialSpace.FreeSpace)
    }

    Add-DetailedResult -Step 'TempCleanup' -Status 'Info' -Message 'Enhanced temporary file cleanup completed.' -Data @{
        EstimatedSpaceFreedMB = [math]::Round($totalSpaceFreed / 1MB, 2)
        ActualSpaceFreedMB    = [math]::Round($actualFreed / 1MB, 2)
        TargetsProcessed      = $cleanupResults.Count
        ResultsJson           = ($cleanupResults | ForEach-Object {
            [ordered]@{
                Path         = $_.Path
                Description  = $_.Description
                ItemCount    = $_.ItemCount
                SpaceFreedMB = [math]::Round(([double]$_.SpaceFreed) / 1MB, 2)
                Status       = $_.Status
                Message      = $_.Message
            }
        } | ConvertTo-Json -Compress)
    }
}

function Test-RepairVolumeScanIndicatesCorruption {
    [CmdletBinding()]
    param(
        [AllowNull()]$ScanResult
    )

    if ($null -eq $ScanResult) {
        return $false
    }

    $resultText = ($ScanResult | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($resultText)) {
        return $false
    }

    # Repair-Volume -Scan commonly returns status strings such as NoErrorsFound, ScanNeeded,
    # SpotFixNeeded, or FullRepairNeeded. Do not treat NoErrorsFound as a match simply
    # because it contains the word "Errors".
    if ($resultText -match '(?i)NoErrorsFound|No errors found|No problems found|No corruption detected') {
        return $false
    }

    if ($resultText -match '(?i)SpotFixNeeded|FullRepairNeeded|ScanNeeded|corrupt|corruption|repair needed|needs repair|errors found|file system errors') {
        return $true
    }

    return $false
}

function Invoke-RepairVolumeScan {
    [CmdletBinding()]
    param(
        [string]$StepName = 'RepairVolumeScan',
        [switch]$UpdateSummary = $true
    )

    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Running Repair-Volume scan on $($env:SystemDrive)" 'INFO'
    $result = Repair-Volume -DriveLetter $systemDrive -Scan -ErrorAction Stop

    $resultText = $null
    $corruptionDetected = $false
    if ($null -ne $result) {
        $resultText = ($result | Out-String).Trim()
        if (-not [string]::IsNullOrWhiteSpace($resultText)) {
            Write-Log "Repair-Volume scan result for $($env:SystemDrive): $resultText" 'INFO'
        }
        $corruptionDetected = Test-RepairVolumeScanIndicatesCorruption -ScanResult $result
    }

    if ($corruptionDetected) {
        if ($UpdateSummary) {
            $script:Summary.DiskCorruptionSuspected = $true
        }
        Warn-Step -Name $StepName -Reason 'Disk scan output indicates scan, spot fix, full repair, errors, or corruption may still exist.'
    }
    else {
        if ($UpdateSummary) {
            $script:Summary.DiskCorruptionSuspected = $false
        }
        Write-Log "Repair-Volume scan did not report actionable file system corruption on $($env:SystemDrive)." 'OK'
    }

    Add-DetailedResult -Step $StepName -Status 'Info' -Message 'Repair-Volume scan completed.' -Data @{
        Output             = $resultText
        CorruptionDetected = $corruptionDetected
    }

    return [PSCustomObject]@{
        CorruptionDetected = $corruptionDetected
        Output             = $resultText
    }
}

function Invoke-RepairVolumeSpotFix {
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Attempting online spot fix on $($env:SystemDrive) with Repair-Volume -SpotFix." 'WARN'

    try {
        Repair-Volume -DriveLetter $systemDrive -SpotFix -ErrorAction Stop | Out-Null
        $script:Summary.RepairVolumeSpotFixSucceeded = $true
        Add-RepairAttempt 'Repair-Volume -SpotFix'
        Add-DetailedResult -Step 'RepairVolumeSpotFix' -Status 'Success' -Message 'Repair-Volume -SpotFix completed successfully.'
        Write-Log 'Repair-Volume spot fix completed successfully without scheduling an offline repair.' 'OK'
        return $true
    }
    catch {
        $script:Summary.RepairVolumeSpotFixSucceeded = $false
        Warn-Step -Name 'RepairVolumeSpotFix' -Reason "Repair-Volume -SpotFix could not complete online: $($_.Exception.Message)"
        Add-DetailedResult -Step 'RepairVolumeSpotFix' -Status 'Warning' -Message 'Repair-Volume -SpotFix could not complete online.' -Data @{
            Error = $_.Exception.Message
        }
        return $false
    }
}

function Invoke-RepairVolumeOfflineFix {
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Running offline disk repair on $($env:SystemDrive)" 'WARN'
    Repair-Volume -DriveLetter $systemDrive -OfflineScanAndFix -ErrorAction Stop | Out-Null
    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Repair-Volume -OfflineScanAndFix'
    Add-DetailedResult -Step 'OfflineDiskRepair' -Status 'Info' -Message 'Offline disk repair was started.'
}

function Invoke-DismDetection {
    Write-Log "Running DISM CheckHealth..." 'INFO'
    $check = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/CheckHealth')

    Write-Log "Running DISM ScanHealth..." 'INFO'
    $scan = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/ScanHealth')

    $combined = (($check.StdOut, $scan.StdOut, $check.StdErr, $scan.StdErr) -join "`n")

    if ($check.ExitCode -ne 0 -or $scan.ExitCode -ne 0) {
        $script:Summary.DismCorruptionDetected = $true
        Warn-Step -Name 'DISMDetection' -Reason 'DISM detection returned a non-zero exit code.'
        return
    }

    # Avoid false positives from phrases like "No component store corruption detected."
    if ($combined -match '(?i)No component store corruption detected|No component store corruption was detected|The component store is repairable\s*:\s*No') {
        $script:Summary.DismCorruptionDetected = $false
        Write-Log 'DISM did not detect component store corruption.' 'OK'
        return
    }

    if ($combined -match '(?i)The component store is repairable|component store is repairable|repairable\s*:\s*Yes|corruption detected|component store corruption detected') {
        $script:Summary.DismCorruptionDetected = $true
        Warn-Step -Name 'DISMDetection' -Reason 'DISM detected component store corruption.'
    }
    else {
        $script:Summary.DismCorruptionDetected = $false
        Write-Log 'DISM detection completed without confirmed corruption.' 'OK'
    }
}

function Invoke-DismRepair {
    Write-Log "Running DISM RestoreHealth..." 'WARN'
    $repair = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/RestoreHealth')
    if ($repair.ExitCode -ne 0) {
        throw "DISM RestoreHealth exited with code $($repair.ExitCode)"
    }

    Write-Log "Running DISM StartComponentCleanup..." 'INFO'
    $cleanup = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/StartComponentCleanup')
    if ($cleanup.ExitCode -ne 0) {
        throw "DISM StartComponentCleanup exited with code $($cleanup.ExitCode)"
    }

    Add-RepairAttempt 'DISM RestoreHealth + StartComponentCleanup'
}

function Invoke-SfcDetection {
    Write-Log "Running SFC verify-only scan..." 'INFO'
    $result = Invoke-SfcCommand -Arguments '/verifyonly'

    $combined = (($result.StdOut, $result.StdErr) -join "`n")

    if ($combined -match '(?i)Windows Resource Protection found integrity violations|found integrity violations') {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason 'SFC detected integrity violations.'
        return
    }

    if ($combined -match '(?i)Windows Resource Protection did not find any integrity violations|did not find any integrity violations') {
        $script:Summary.SfcIntegrityViolations = $false
        Write-Log 'SFC did not detect integrity violations.' 'OK'
        return
    }

    if ($combined -match '(?i)Windows Resource Protection found corrupt files and successfully repaired them|Windows Resource Protection found corrupt files but was unable to fix some of them') {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason 'SFC reported corrupt files.'
        return
    }

    if ($result.ExitCode -notin 0,1) {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason "SFC verify returned unusual exit code $($result.ExitCode)."
    }
    else {
        Write-Log 'SFC detection completed without confirmed integrity violations.' 'OK'
    }
}

function Invoke-SfcRepair {
    Write-Log "Running SFC /SCANNOW..." 'WARN'
    $result = Invoke-SfcCommand -Arguments '/scannow'

    if ($result.ExitCode -notin 0,1) {
        throw "SFC /SCANNOW exited with code $($result.ExitCode)"
    }

    Add-RepairAttempt 'SFC /SCANNOW'
}

function Export-CbsCorruptionDetails {
    param(
        [int]$TailLines = 20000
    )

    $cbsPath = Join-Path $env:SystemRoot 'Logs\CBS\CBS.log'
    $extractPath = Join-Path $LogDirectory ($script:BaseFileName + '-CBS-SFC-CorruptionExtract.txt')

    if (-not (Test-Path -LiteralPath $cbsPath)) {
        Write-Log "CBS_CORRUPTION_EXTRACT|Status=MissingCBSLog|Path=$cbsPath" 'WARN'
        Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Warning' -Message 'CBS.log was not found.' -Data @{ Path = $cbsPath }
        return @()
    }

    try {
        $lines = @(Get-Content -LiteralPath $cbsPath -Tail $TailLines -ErrorAction Stop)
        $matches = @($lines | Where-Object {
            $_ -match '\[SR\]' -or
            $_ -match 'Corrupt file:' -or
            $_ -match 'Cannot repair member file' -or
            $_ -match 'Repairing corrupted file' -or
            $_ -match 'Repaired file' -or
            $_ -match 'could not reproject corrupted file' -or
            $_ -match 'Hash mismatch' -or
            $_ -match 'CSI.*corrupt'
        })

        if ($matches.Count -gt 0) {
            $header = @(
                "# CBS/SFC corruption extract",
                "# Computer: $($script:ComputerName)",
                "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
                "# Source: $cbsPath",
                "# TailLinesAnalyzed: $TailLines",
                ""
            )
            Ensure-LogDirectory
            $header + $matches | Set-Content -LiteralPath $extractPath -Encoding UTF8 -Force
            $script:Summary.CbsCorruptionExtractPath = $extractPath

            $corruptFileLines = @($matches | Where-Object { $_ -match 'Corrupt file:|Cannot repair member file|Hash mismatch|could not reproject corrupted file' })
            $script:Summary.CbsCorruptFileCount = $corruptFileLines.Count

            Write-Log "CBS_CORRUPTION_EXTRACT|Status=Created|Path=$extractPath|MatchingLines=$($matches.Count)|CorruptFileLines=$($corruptFileLines.Count)" 'WARN'

            foreach ($line in ($corruptFileLines | Select-Object -First 25)) {
                $safeLine = ([string]$line) -replace '\s+', ' '
                Write-Log "CBS_CORRUPTION_DETAIL|$safeLine" 'WARN'
            }

            $driverCorruptionLines = @($corruptFileLines | Where-Object { $_ -match '\\System32\\drivers\\|\.sys' })
            if ($driverCorruptionLines.Count -gt 0) {
                $script:Summary.CbsDriverCorruptionDetected = $true
                foreach ($driverLine in $driverCorruptionLines) {
                    $driverMatches = [regex]::Matches([string]$driverLine, '(?i)[A-Za-z0-9_.-]+\.sys')
                    foreach ($driverMatch in $driverMatches) {
                        $driverName = $driverMatch.Value
                        if (-not [string]::IsNullOrWhiteSpace($driverName) -and -not $script:Summary.CbsCorruptDriverFiles.Contains($driverName)) {
                            $script:Summary.CbsCorruptDriverFiles.Add($driverName) | Out-Null
                        }
                    }
                }
                $driverList = (($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ',')
                Write-Log "CBS_DRIVER_CORRUPTION_DETECTED|Drivers=$driverList|Action=DISM_SFC_FIRST_THEN_HPIA_DRIVER_ONLY_IF_STILL_CORRUPT" 'WARN'
                Add-Note 'CBS corruption references driver files. If DISM/SFC does not clear it, HP Image Assistant driver-only remediation can be run.'
            }

            Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Warning' -Message 'CBS/SFC corruption details were extracted.' -Data @{
                SourcePath = $cbsPath
                ExtractPath = $extractPath
                MatchingLines = $matches.Count
                CorruptFileLines = $corruptFileLines.Count
            }
        }
        else {
            Write-Log "CBS_CORRUPTION_EXTRACT|Status=NoMatchingCorruptionLines|Path=$cbsPath|TailLines=$TailLines" 'INFO'
            Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Info' -Message 'No CBS/SFC corruption detail lines found in recent CBS.log tail.' -Data @{
                SourcePath = $cbsPath
                TailLines = $TailLines
            }
        }

        return $matches
    }
    catch {
        Write-Log "CBS_CORRUPTION_EXTRACT|Status=Failed|Error=$($_.Exception.Message)" 'WARN'
        Add-DetailedResult -Step 'CBSCorruptionExtract' -Status 'Warning' -Message 'Failed to extract CBS/SFC corruption details.' -Data @{ Error = $_.Exception.Message }
        return @()
    }
}

function Invoke-SystemFileRepairWorkflow {
    Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Status=Starting|Actions=ExtractCBS,DISMRestoreHealth,SFCScannow,VerifyOnly' 'WARN'
    $script:Summary.SfcRepairAttempted = $true

    Export-CbsCorruptionDetails | Out-Null

    try {
        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=DISMRestoreHealth|Status=Starting' 'WARN'
        $dismRestore = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/RestoreHealth')
        if ($dismRestore.ExitCode -ne 0) {
            throw "DISM RestoreHealth exited with code $($dismRestore.ExitCode)"
        }
        Add-RepairAttempt 'DISM RestoreHealth for SFC/CBS corruption'
        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=DISMRestoreHealth|Status=Completed' 'OK'

        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=SFCScannow|Status=Starting' 'WARN'
        $sfcRepair = Invoke-SfcCommand -Arguments '/scannow'
        if ($sfcRepair.ExitCode -notin 0,1) {
            throw "SFC /SCANNOW exited with code $($sfcRepair.ExitCode)"
        }
        Add-RepairAttempt 'SFC /SCANNOW for CBS corruption'
        Write-Log "SFC_CBS_REPAIR_WORKFLOW|Step=SFCScannow|Status=Completed|ExitCode=$($sfcRepair.ExitCode)" 'OK'

        Write-Log 'SFC_CBS_REPAIR_WORKFLOW|Step=VerifyOnlyAfterRepair|Status=Starting' 'INFO'
        $verify = Invoke-SfcCommand -Arguments '/verifyonly'
        $verifyText = (($verify.StdOut, $verify.StdErr) -join "`n")

        if ($verifyText -match '(?i)did not find any integrity violations') {
            $script:Summary.SfcIntegrityViolations = $false
            $script:Summary.SfcRepairSucceeded = $true
            $script:Summary.CbsDriverCorruptionRemaining = $false
            Write-Log 'SFC_REPAIR_RESULT|Status=CleanAfterRepair|IntegrityViolations=False' 'OK'
        }
        elseif ($verifyText -match '(?i)found integrity violations|found corrupt files|unable to fix') {
            $script:Summary.SfcIntegrityViolations = $true
            $script:Summary.SfcRepairSucceeded = $false
            Write-Log 'SFC_REPAIR_RESULT|Status=StillDetectedAfterRepair|IntegrityViolations=True|NextStep=If CBS driver corruption remains, Tier3 HPIA driver-only remediation may run' 'WARN'
            Export-CbsCorruptionDetails | Out-Null
            $script:Summary.CbsDriverCorruptionRemaining = [bool]$script:Summary.CbsDriverCorruptionDetected
        }
        else {
            $script:Summary.SfcRepairSucceeded = $false
            Write-Log "SFC_REPAIR_RESULT|Status=UnknownAfterRepair|ExitCode=$($verify.ExitCode)" 'WARN'
        }

        Add-DetailedResult -Step 'SfcCbsRepairWorkflow' -Status 'Info' -Message 'SFC/CBS repair workflow completed.' -Data @{
            SfcRepairSucceeded = $script:Summary.SfcRepairSucceeded
            SfcIntegrityViolations = $script:Summary.SfcIntegrityViolations
            CbsCorruptFileCount = $script:Summary.CbsCorruptFileCount
            CbsCorruptionExtractPath = $script:Summary.CbsCorruptionExtractPath
        }
    }
    catch {
        $script:Summary.SfcRepairSucceeded = $false
        Write-Log "SFC_REPAIR_RESULT|Status=Failed|Error=$($_.Exception.Message)" 'ERROR'
        Add-DetailedResult -Step 'SfcCbsRepairWorkflow' -Status 'Failed' -Message $_.Exception.Message -Data @{
            CbsCorruptionExtractPath = $script:Summary.CbsCorruptionExtractPath
        }
        throw
    }
}

function Test-IsHpSystem {
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $manufacturer = [string]$cs.Manufacturer
        if ($manufacturer -match '(?i)\bHP\b|Hewlett-Packard|Hewlett Packard') {
            Write-Log "HP_PLATFORM_DETECTED|Manufacturer=$(Convert-ExplorerSafeString $manufacturer)|Model=$(Convert-ExplorerSafeString $cs.Model)" 'INFO'
            return $true
        }

        Write-Log "HP_PLATFORM_DETECTED|Status=False|Manufacturer=$(Convert-ExplorerSafeString $manufacturer)|Model=$(Convert-ExplorerSafeString $cs.Model)" 'INFO'
        return $false
    }
    catch {
        Write-Log "HP_PLATFORM_DETECTED|Status=Unknown|Error=$($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Find-HpImageAssistantExecutable {
    [CmdletBinding()]
    param()

    $candidatePaths = @(
        'C:\Program Files\HP\HP Image Assistant\HPImageAssistant.exe',
        'C:\Program Files (x86)\HP\HP Image Assistant\HPImageAssistant.exe',
        (Join-Path $HpImageAssistantLocalPath 'HPImageAssistant.exe')
    )

    foreach ($candidate in $candidatePaths) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    try {
        $localFound = Get-ChildItem -LiteralPath $HpImageAssistantLocalPath -Filter 'HPImageAssistant.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($localFound) { return $localFound.FullName }
    }
    catch { }

    return $null
}

function Ensure-HpImageAssistantAvailable {
    [CmdletBinding()]
    param()

    $existing = Find-HpImageAssistantExecutable
    if ($existing) {
        $script:Summary.HpImageAssistantPath = $existing
        Write-Log "HPIA_AVAILABLE|Status=AlreadyPresent|Path=$existing" 'OK'
        return $existing
    }

    Write-Log "HPIA_AVAILABLE|Status=NotFound|Action=StageFromShare|Source=$HpImageAssistantSourcePath|Destination=$HpImageAssistantLocalPath" 'WARN'

    if (-not (Test-Path -LiteralPath $HpImageAssistantSourcePath)) {
        throw "HP Image Assistant source path was not found: $HpImageAssistantSourcePath"
    }

    if (-not (Test-Path -LiteralPath $HpImageAssistantLocalPath)) {
        New-Item -Path $HpImageAssistantLocalPath -ItemType Directory -Force | Out-Null
    }

    $sourceExe = $null
    try {
        $sourceExe = Get-ChildItem -LiteralPath $HpImageAssistantSourcePath -Filter 'HPImageAssistant.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    catch { }

    if ($sourceExe) {
        Write-Log "HPIA_STAGE|Mode=CopyExtractedFolder|Source=$($sourceExe.DirectoryName)|Destination=$HpImageAssistantLocalPath" 'INFO'
        Copy-Item -LiteralPath (Join-Path $sourceExe.DirectoryName '*') -Destination $HpImageAssistantLocalPath -Recurse -Force -ErrorAction Stop
        $stagedExe = Find-HpImageAssistantExecutable
        if ($stagedExe) {
            $script:Summary.HpImageAssistantPath = $stagedExe
            Write-Log "HPIA_AVAILABLE|Status=Staged|Path=$stagedExe" 'OK'
            return $stagedExe
        }
    }

    $installer = $null
    try {
        $installer = Get-ChildItem -LiteralPath $HpImageAssistantSourcePath -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '(?i)hp.*image.*assistant|hpia|sp\d+' -and $_.Extension -match '(?i)\.exe|\.msi' } |
            Sort-Object Length -Descending |
            Select-Object -First 1
    }
    catch { }

    if (-not $installer) {
        throw "Could not find HPImageAssistant.exe or an HPIA installer under $HpImageAssistantSourcePath"
    }

    Write-Log "HPIA_STAGE|Mode=Installer|Installer=$($installer.FullName)|Destination=$HpImageAssistantLocalPath" 'INFO'

    if ($installer.Extension -ieq '.msi') {
        $args = "/i `"$($installer.FullName)`" /qn /norestart"
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -notin 0,3010,1641) {
            throw "HPIA MSI installer failed with exit code $($proc.ExitCode)"
        }
    }
    else {
        $extractArgs = "/s /e /f `"$HpImageAssistantLocalPath`""
        $proc = Start-Process -FilePath $installer.FullName -ArgumentList $extractArgs -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -notin 0,1168,3010,1641) {
            throw "HPIA installer/extractor failed with exit code $($proc.ExitCode)"
        }
    }

    $afterInstall = Find-HpImageAssistantExecutable
    if (-not $afterInstall) {
        throw "HPImageAssistant.exe was not found after staging/installing HPIA."
    }

    $script:Summary.HpImageAssistantPath = $afterInstall
    Write-Log "HPIA_AVAILABLE|Status=InstalledOrExtracted|Path=$afterInstall" 'OK'
    return $afterInstall
}

function Invoke-HpiaDriverOnlyRepairFromCbs {
    [CmdletBinding()]
    param()

    Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Evaluating|Trigger=CBS_DRIVER_CORRUPTION_REMAINING' 'INFO'

    if (-not $AllowHpDriverRepairFromCbs) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=AllowHpDriverRepairFromCbsDisabled' 'INFO'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'HP driver remediation is disabled by parameter.'
        return
    }

    if (-not $script:Summary.CbsDriverCorruptionDetected) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=NoCBSDriverCorruptionDetected' 'INFO'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'No CBS driver corruption marker was detected.'
        return
    }

    if ($script:Summary.SfcRepairSucceeded -and -not $script:Summary.SfcIntegrityViolations) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=DISMAndSfcClearedIntegrityViolations' 'OK'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'DISM/SFC cleared the issue; HPIA driver remediation not needed.'
        return
    }

    if (-not (Test-IsHpSystem)) {
        Write-Log 'TIER3_DRIVER_REMEDIATION|Status=Skipped|Reason=NotAnHPSystem' 'WARN'
        Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Skipped' -Message 'System manufacturer is not HP/Hewlett-Packard.'
        return
    }

    $script:Summary.HpDriverRepairAttempted = $true
    $driverList = (($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ',')
    if ([string]::IsNullOrWhiteSpace($driverList)) { $driverList = 'UnknownDriverFilesFromCBS' }

    Write-Log "HPIA_DRIVER_REMEDIATION|Status=Starting|Trigger=CBS_DRIVER_CORRUPTION|Drivers=$driverList|Category=DriversOnly|Excluded=BIOS,Firmware" 'WARN'

    $hpiaExe = Ensure-HpImageAssistantAvailable
    $hpiaRoot = Join-Path 'C:\ProgramData\SystemRepair' 'HPIA-CBS-DriverRepair'
    $hpiaReportFolder = Join-Path $hpiaRoot 'Reports'
    $hpiaDownloadFolder = Join-Path $hpiaRoot 'SoftPaqs'
    New-Item -Path $hpiaReportFolder -ItemType Directory -Force | Out-Null
    New-Item -Path $hpiaDownloadFolder -ItemType Directory -Force | Out-Null

    $args = @(
        '/Operation:Analyze',
        '/Action:Install',
        '/Selection:All',
        '/Category:Drivers',
        '/Silent',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`""
    )

    Write-Log "HPIA_DRIVER_REMEDIATION|Command=$hpiaExe $($args -join ' ')" 'INFO'

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $hpiaExe
    $psi.Arguments = $args -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 80 | ForEach-Object { Write-Log "HPIA_STDOUT|$($_.Trim())" 'INFO' }
    }
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 80 | ForEach-Object { Write-Log "HPIA_STDERR|$($_.Trim())" 'WARN' }
    }

    $reportFiles = @()
    try { $reportFiles = @(Get-ChildItem -LiteralPath $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName) } catch { }

    if ($proc.ExitCode -in 0,256,3010,3020,4096) {
        if ($proc.ExitCode -in 3010,3020) { $script:Summary.RebootRequired = $true }
        $script:Summary.HpDriverRepairSucceeded = $true
        Add-RepairAttempt 'HP Image Assistant driver-only remediation for CBS driver corruption'
        Write-Log "HPIA_DRIVER_REMEDIATION|Status=Completed|ExitCode=$($proc.ExitCode)|ReportFolder=$hpiaReportFolder|SoftpaqDownloadFolder=$hpiaDownloadFolder|RebootRequired=$($script:Summary.RebootRequired)" 'OK'
    }
    else {
        $script:Summary.HpDriverRepairSucceeded = $false
        Write-Log "HPIA_DRIVER_REMEDIATION|Status=Failed|ExitCode=$($proc.ExitCode)|ReportFolder=$hpiaReportFolder|SoftpaqDownloadFolder=$hpiaDownloadFolder" 'ERROR'
    }

    if ($script:Summary.HpDriverRepairSucceeded) {
        Write-Log 'HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=Starting' 'INFO'
        try {
            $postHpiaVerify = Invoke-SfcCommand -Arguments '/verifyonly'
            $postHpiaVerifyText = (($postHpiaVerify.StdOut, $postHpiaVerify.StdErr) -join "`n")
            if ($postHpiaVerifyText -match '(?i)did not find any integrity violations') {
                $script:Summary.SfcIntegrityViolations = $false
                $script:Summary.CbsDriverCorruptionRemaining = $false
                Write-Log 'HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=CleanAfterDriverRepair|IntegrityViolations=False' 'OK'
            }
            elseif ($postHpiaVerifyText -match '(?i)found integrity violations|found corrupt files|unable to fix') {
                $script:Summary.SfcIntegrityViolations = $true
                $script:Summary.CbsDriverCorruptionRemaining = $true
                Write-Log 'HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=StillCorruptAfterDriverRepair|IntegrityViolations=True|NextStep=Review CBS extract and consider in-place repair install' 'WARN'
                Export-CbsCorruptionDetails | Out-Null
            }
            else {
                Write-Log "HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=Unknown|ExitCode=$($postHpiaVerify.ExitCode)" 'WARN'
            }
        }
        catch {
            Write-Log "HPIA_DRIVER_REMEDIATION|Step=PostHpiaSfcVerify|Status=Failed|Error=$($_.Exception.Message)" 'WARN'
        }
    }

    Add-DetailedResult -Step 'HpiaDriverRepairFromCbs' -Status 'Info' -Message 'HPIA driver-only remediation attempted because CBS driver corruption was detected.' -Data @{
        ExitCode = $proc.ExitCode
        HpiaPath = $hpiaExe
        ReportFolder = $hpiaReportFolder
        SoftpaqDownloadFolder = $hpiaDownloadFolder
        DriverFilesFromCbs = $driverList
        ReportFiles = ($reportFiles -join ';')
        PostRepairSfcIntegrityViolations = $script:Summary.SfcIntegrityViolations
    }

    if (-not $script:Summary.HpDriverRepairSucceeded) {
        throw "HPIA driver-only remediation failed with exit code $($proc.ExitCode)"
    }
}

function Invoke-WmiCheck {
    Write-Log "Checking WMI repository consistency..." 'INFO'
    $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /verifyrepository 2>&1
    $text = ($output | Out-String).Trim()

    if ($text) {
        $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    Add-DetailedResult -Step 'WmiRepositoryCheck' -Status 'Info' -Message 'WMI verify completed.' -Data @{
        Output = $text
    }

    if ($text -match 'inconsistent') {
        $script:Summary.WmiRepositoryInconsistent = $true
        Warn-Step -Name 'WmiRepositoryCheck' -Reason 'WMI repository reported as inconsistent.'
    }
}

function Invoke-WmiRepair {
    Write-Log "Repairing WMI repository..." 'WARN'
    $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /salvagerepository 2>&1
    $text = ($output | Out-String).Trim()

    if ($text) {
        $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    Add-RepairAttempt 'winmgmt /salvagerepository'
    Add-DetailedResult -Step 'WmiRepair' -Status 'Info' -Message 'WMI salvage completed.' -Data @{
        Output = $text
    }
}

function Invoke-NetworkReset {
    Write-Log "Flushing DNS cache..." 'INFO'
    ipconfig /flushdns | Out-Null

    Write-Log "Resetting Winsock..." 'WARN'
    netsh winsock reset | Out-Null

    Write-Log "Resetting TCP/IP stack..." 'WARN'
    netsh int ip reset | Out-Null

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Winsock/TCPIP reset'
    Add-DetailedResult -Step 'NetworkReset' -Status 'Info' -Message 'Network reset completed.'
}

function Invoke-DnsFlushOnly {
    Write-Log "Flushing DNS cache..." 'INFO'
    ipconfig /flushdns | Out-Null
    Add-DetailedResult -Step 'DnsFlush' -Status 'Info' -Message 'DNS cache flushed.'
}


function Convert-RpcSafeString {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value) -replace "[`r`n|]+", ' '
}

function Get-RpcRootCauseAssessment {
    param(
        [Parameter(Mandatory)][bool]$LocalProblem,
        [Parameter(Mandatory)][bool]$RemoteProblem,
        [Parameter(Mandatory)][int]$ServiceProblemCount,
        [Parameter(Mandatory)][bool]$LocalEndpointMapperReachable,
        [Parameter(Mandatory)][bool]$LocalEventLogRead,
        [Parameter(Mandatory)][bool]$LocalWmiQuery,
        [Parameter(Mandatory)][int]$RemoteProblemCount
    )

    if ($LocalProblem) {
        if (-not $LocalEndpointMapperReachable) {
            return 'LOCAL_CLIENT_OR_LOCAL_RPC_SERVICE_PROBLEM: Local RPC endpoint mapper port 135 is not reachable on 127.0.0.1. Check RpcSs/RpcEptMapper/DcomLaunch services and local firewall/security software.'
        }
        if ($ServiceProblemCount -gt 0) {
            return "LOCAL_CLIENT_OR_LOCAL_SERVICE_PROBLEM: One or more required local services are missing/stopped/unhealthy. ServiceProblemCount=$ServiceProblemCount."
        }
        if (-not $LocalEventLogRead -and $LocalWmiQuery) {
            return 'LOCAL_EVENTLOG_ACCESS_PROBLEM: Local RPC/WMI basics are healthy, but local Event Log access failed. Check Windows Event Log service, event log corruption, permissions, or local security policy.'
        }
        if (-not $LocalWmiQuery) {
            return 'LOCAL_WMI_OR_CIM_PROBLEM: Local WMI/CIM query failed. Check Winmgmt service and WMI repository health.'
        }
        return 'LOCAL_CLIENT_SIDE_PROBLEM: Local RPC-related validation failed. Review RPC_LOCAL_* markers above.'
    }

    if ($RemoteProblem) {
        return "REMOTE_ENDPOINT_OR_NETWORK_PROBLEM: Local RPC checks passed, but one or more target systems did not answer DNS/RPC port 135. RemoteProblemCount=$RemoteProblemCount. Check DNS, routing/VLANs, remote firewall, remote RPC services, and domain controller reachability."
    }

    return 'NO_RPC_PROBLEM_DETECTED_BY_BASIC_CHECKS: Local RPC services, local endpoint mapper, Event Log access, WMI/CIM, and tested remote endpoint mapper targets passed. If another section still reports RPC unavailable, that specific provider/query may be failing rather than core RPC.'
}

function Test-TcpPortSafe {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][int]$Port,
        [int]$TimeoutMilliseconds = 2500
    )

    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $connected = $async.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)
        if (-not $connected) {
            return [PSCustomObject]@{ ComputerName = $ComputerName; Port = $Port; Reachable = $false; Error = 'Timeout' }
        }
        $client.EndConnect($async)
        return [PSCustomObject]@{ ComputerName = $ComputerName; Port = $Port; Reachable = $true; Error = '' }
    }
    catch {
        return [PSCustomObject]@{ ComputerName = $ComputerName; Port = $Port; Reachable = $false; Error = $_.Exception.Message }
    }
    finally {
        if ($client) { $client.Close() }
    }
}

function Invoke-RpcDiagnostics {
    Write-Log 'Running RPC diagnostics to separate local client/service issues from remote RPC endpoint issues...' 'INFO'

    $serviceNames = @('RpcSs', 'RpcEptMapper', 'DcomLaunch', 'EventLog', 'Winmgmt')
    $serviceProblems = New-Object System.Collections.Generic.List[string]
    $serviceResults = New-Object System.Collections.Generic.List[object]

    foreach ($name in $serviceNames) {
        try {
            $svc = Get-CimInstance Win32_Service -Filter "Name='$name'" -ErrorAction Stop
            if ($null -eq $svc) {
                $serviceProblems.Add("$name missing") | Out-Null
                Write-Log "RPC_LOCAL_SERVICE|Name=$name|Exists=False|State=Missing|StartMode=Unknown|Status=Problem" 'WARN'
                continue
            }

            $status = if ($svc.State -eq 'Running') { 'OK' } else { 'Problem' }
            if ($status -ne 'OK') { $serviceProblems.Add("$name state $($svc.State)") | Out-Null }

            Write-Log "RPC_LOCAL_SERVICE|Name=$name|Exists=True|State=$(Convert-RpcSafeString $svc.State)|StartMode=$(Convert-RpcSafeString $svc.StartMode)|ProcessId=$($svc.ProcessId)|Status=$status" ($(if ($status -eq 'OK') { 'INFO' } else { 'WARN' }))
            $serviceResults.Add([PSCustomObject]@{ Name=$name; State=$svc.State; StartMode=$svc.StartMode; ProcessId=$svc.ProcessId; Status=$status }) | Out-Null
        }
        catch {
            $serviceProblems.Add("$name query failed") | Out-Null
            Write-Log "RPC_LOCAL_SERVICE|Name=$name|Exists=Unknown|State=Unknown|StartMode=Unknown|Status=Problem|Error=$(Convert-RpcSafeString $_.Exception.Message)" 'WARN'
        }
    }

    $localEndpoint = Test-TcpPortSafe -ComputerName '127.0.0.1' -Port 135
    Write-Log "RPC_ENDPOINT_MAPPER|Target=127.0.0.1|Port=135|Reachable=$($localEndpoint.Reachable)|Error=$(Convert-RpcSafeString $localEndpoint.Error)" ($(if ($localEndpoint.Reachable) { 'INFO' } else { 'WARN' }))

    $eventLogOk = $false
    $eventLogError = ''
    try {
        $null = Get-WinEvent -ListLog System -ErrorAction Stop
        $null = Get-WinEvent -FilterHashtable @{ LogName='System'; StartTime=(Get-Date).AddMinutes(-5) } -MaxEvents 1 -ErrorAction SilentlyContinue
        $eventLogOk = $true
    }
    catch {
        $eventLogError = $_.Exception.Message
    }
    Write-Log "RPC_LOCAL_EVENTLOG_READ|Log=System|Success=$eventLogOk|Error=$(Convert-RpcSafeString $eventLogError)" ($(if ($eventLogOk) { 'INFO' } else { 'WARN' }))

    $cimOk = $false
    $cimError = ''
    try {
        $null = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop | Select-Object -First 1
        $cimOk = $true
    }
    catch {
        $cimError = $_.Exception.Message
    }
    Write-Log "RPC_LOCAL_WMI_QUERY|Class=Win32_OperatingSystem|Success=$cimOk|Error=$(Convert-RpcSafeString $cimError)" ($(if ($cimOk) { 'INFO' } else { 'WARN' }))

    try {
        if (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
            Get-NetFirewallProfile -ErrorAction SilentlyContinue | ForEach-Object {
                Write-Log "RPC_FIREWALL_PROFILE|Name=$(Convert-RpcSafeString $_.Name)|Enabled=$($_.Enabled)|DefaultInboundAction=$(Convert-RpcSafeString $_.DefaultInboundAction)|DefaultOutboundAction=$(Convert-RpcSafeString $_.DefaultOutboundAction)" 'INFO'
            }
        }
    }
    catch {
        Write-Log "RPC_FIREWALL_PROFILE|Status=UnableToQuery|Error=$(Convert-RpcSafeString $_.Exception.Message)" 'WARN'
    }

    $targets = New-Object System.Collections.Generic.List[string]
    if ($env:LOGONSERVER) {
        $targets.Add(($env:LOGONSERVER -replace '^\\+', '')) | Out-Null
    }

    try {
        $dnsServers = @(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | ForEach-Object { $_.ServerAddresses } | Where-Object { $_ } | Select-Object -Unique)
        foreach ($dns in $dnsServers) { $targets.Add([string]$dns) | Out-Null }
    }
    catch { }

    $remoteProblems = New-Object System.Collections.Generic.List[object]
    foreach ($target in ($targets | Where-Object { $_ -and $_ -notmatch '^(127\.|::1|localhost$)' } | Select-Object -Unique -First 8)) {
        $dnsOk = $true
        $dnsError = ''
        try { $null = [System.Net.Dns]::GetHostAddresses($target) }
        catch { $dnsOk = $false; $dnsError = $_.Exception.Message }

        $port135 = Test-TcpPortSafe -ComputerName $target -Port 135
        Write-Log "RPC_TARGET_PORT_TEST|Target=$(Convert-RpcSafeString $target)|DnsResolved=$dnsOk|Port=135|Reachable=$($port135.Reachable)|Error=$(Convert-RpcSafeString $(if (-not $dnsOk) { $dnsError } else { $port135.Error }))" ($(if ($dnsOk -and $port135.Reachable) { 'INFO' } else { 'WARN' }))

        if (-not $dnsOk -or -not $port135.Reachable) {
            $remoteProblems.Add([PSCustomObject]@{ Target=$target; DnsResolved=$dnsOk; Port135Reachable=$port135.Reachable; Error=$(if (-not $dnsOk) { $dnsError } else { $port135.Error }) }) | Out-Null
        }
    }

    $localProblem = ($serviceProblems.Count -gt 0 -or -not $localEndpoint.Reachable -or -not $eventLogOk -or -not $cimOk)
    $remoteProblem = (-not $localProblem -and $remoteProblems.Count -gt 0)

    $script:Summary.RpcClientSideIssueDetected = [bool]$localProblem
    $script:Summary.RpcRemoteSideIssueDetected = [bool]$remoteProblem

    $rpcAssessment = Get-RpcRootCauseAssessment -LocalProblem ([bool]$localProblem) -RemoteProblem ([bool]$remoteProblem) -ServiceProblemCount ($serviceProblems.Count) -LocalEndpointMapperReachable ([bool]$localEndpoint.Reachable) -LocalEventLogRead ([bool]$eventLogOk) -LocalWmiQuery ([bool]$cimOk) -RemoteProblemCount ($remoteProblems.Count)

    if ($localProblem) {
        $script:Summary.RpcDiagnosticsStatus = 'LocalClientOrLocalServiceIssueSuspected'
        Write-Log "RPC_CLIENT_SIDE_SUSPECTED|ServiceProblems=$($serviceProblems.Count)|LocalEndpointMapperReachable=$($localEndpoint.Reachable)|LocalEventLogRead=$eventLogOk|LocalWmiQuery=$cimOk" 'WARN'
        Add-Note 'RPC diagnostics suggest a local client/service-side issue. Review RPC_LOCAL_SERVICE, RPC_ENDPOINT_MAPPER, RPC_LOCAL_EVENTLOG_READ, and RPC_LOCAL_WMI_QUERY markers.'
    }
    elseif ($remoteProblem) {
        $script:Summary.RpcDiagnosticsStatus = 'RemoteEndpointOrNetworkIssueSuspected'
        Write-Log "RPC_REMOTE_SIDE_OR_NETWORK_SUSPECTED|RemoteProblemCount=$($remoteProblems.Count)|LocalEndpointMapperReachable=$($localEndpoint.Reachable)|LocalEventLogRead=$eventLogOk|LocalWmiQuery=$cimOk" 'WARN'
        Add-Note 'RPC diagnostics suggest local RPC is healthy, but one or more domain/DNS/logon targets did not answer RPC endpoint mapper port 135.'
    }
    else {
        $script:Summary.RpcDiagnosticsStatus = 'NoRpcIssueDetectedByBasicChecks'
        Write-Log "RPC_DIAGNOSTICS_OK|LocalEndpointMapperReachable=$($localEndpoint.Reachable)|LocalEventLogRead=$eventLogOk|LocalWmiQuery=$cimOk|RemoteProblemCount=$($remoteProblems.Count)" 'OK'
    }

    $rpcAssessmentLevel = if ($localProblem -or $remoteProblem) { 'WARN' } else { 'OK' }
    Write-Log "RPC_ROOT_CAUSE_ASSESSMENT|$rpcAssessment" $rpcAssessmentLevel
    Write-Log "RPC_DIAGNOSTICS_SUMMARY|Status=$($script:Summary.RpcDiagnosticsStatus)|ClientSideIssue=$($script:Summary.RpcClientSideIssueDetected)|RemoteOrNetworkIssue=$($script:Summary.RpcRemoteSideIssueDetected)|ServiceProblems=$($serviceProblems.Count)|RemoteProblems=$($remoteProblems.Count)|Assessment=$(Convert-RpcSafeString $rpcAssessment)" ($(if ($localProblem -or $remoteProblem) { 'WARN' } else { 'INFO' }))
    Add-Note "RPC root cause assessment: $rpcAssessment"

    Add-DetailedResult -Step 'RpcDiagnostics' -Status 'Info' -Message 'RPC diagnostics completed.' -Data @{
        Status = $script:Summary.RpcDiagnosticsStatus
        ClientSideIssueDetected = $script:Summary.RpcClientSideIssueDetected
        RemoteSideIssueDetected = $script:Summary.RpcRemoteSideIssueDetected
        ServiceResultsJson = (($serviceResults | ConvertTo-Json -Compress) -replace "'", "''")
        RemoteProblemsJson = (($remoteProblems | ConvertTo-Json -Compress) -replace "'", "''")
        LocalEndpointMapperReachable = $localEndpoint.Reachable
        LocalEventLogRead = $eventLogOk
        LocalWmiQuery = $cimOk
        RootCauseAssessment = $rpcAssessment
    }
}



function Convert-ExplorerSafeString {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value) -replace '[\r\n\|]+', ' '
}

function Get-ExplorerEventSummary {
    param([int]$DaysBack = 14)

    Write-Log "Collecting Explorer hang/crash events from the last $DaysBack day(s)..." 'INFO'
    $startTime = (Get-Date).AddDays(-1 * [math]::Abs($DaysBack))
    $events = @()

    try {
        $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $startTime } -ErrorAction Stop | Where-Object {
            ($_.ProviderName -match 'Application Error|Windows Error Reporting|Application Hang') -and
            ($_.Message -match '(?i)explorer\.exe|File Explorer')
        })
    }
    catch {
        Warn-Step -Name 'ExplorerEventCollection' -Reason "Could not read Application log for Explorer events: $($_.Exception.Message)"
        return @()
    }

    $hangs = @($events | Where-Object { $_.ProviderName -match 'Application Hang' -or $_.Message -match '(?i)AppHang|not responding|stopped interacting' })
    $crashes = @($events | Where-Object { $_.ProviderName -match 'Application Error|Windows Error Reporting' -and $_.Message -match '(?i)faulting application|fault bucket|crash|stopped working' })

    $script:Summary.ExplorerHangCount = @($hangs).Count
    $script:Summary.ExplorerCrashCount = @($crashes).Count

    Write-Log "EXPLORER_HEALTH_EVENTS|Days=$DaysBack|Hangs=$($script:Summary.ExplorerHangCount)|Crashes=$($script:Summary.ExplorerCrashCount)|Total=$(@($events).Count)" 'INFO'

    foreach ($event in ($events | Select-Object -First 25)) {
        $kind = if ($event.ProviderName -match 'Application Hang' -or $event.Message -match '(?i)AppHang|not responding|stopped interacting') { 'HANG' } else { 'CRASH' }
        $faultingModule = ''
        if ($event.Message -match '(?im)Faulting module name:\s*([^,\r\n]+)') { $faultingModule = $Matches[1].Trim() }
        $exceptionCode = ''
        if ($event.Message -match '(?im)Exception code:\s*([^\r\n]+)') { $exceptionCode = $Matches[1].Trim() }
        Write-Log "EXPLORER_${kind}_DETECTED|Time=$($event.TimeCreated)|Provider=$(Convert-ExplorerSafeString $event.ProviderName)|EventId=$($event.Id)|FaultingModule=$(Convert-ExplorerSafeString $faultingModule)|ExceptionCode=$(Convert-ExplorerSafeString $exceptionCode)" 'WARN'
    }

    Add-DetailedResult -Step 'ExplorerEventCollection' -Status 'Info' -Message 'Collected Explorer hang/crash event summary.' -Data @{
        DaysBack = $DaysBack
        HangCount = $script:Summary.ExplorerHangCount
        CrashCount = $script:Summary.ExplorerCrashCount
        TotalExplorerEvents = @($events).Count
    }

    return $events
}

function Get-NonMicrosoftShellExtensionInventory {
    Write-Log 'Collecting non-Microsoft shell extension and context-menu handler inventory...' 'INFO'

    $registryRoots = @(
        'Registry::HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Directory\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Drive\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\Folder\shellex\ContextMenuHandlers',
        'Registry::HKEY_CLASSES_ROOT\AllFileSystemObjects\shellex\ContextMenuHandlers',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved'
    )

    $items = New-Object System.Collections.Generic.List[object]

    foreach ($root in $registryRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }

        if ($root -match '\\Approved$') {
            try {
                $props = Get-ItemProperty -LiteralPath $root -ErrorAction Stop
                foreach ($prop in $props.PSObject.Properties | Where-Object { $_.Name -match '^\{' }) {
                    $items.Add([PSCustomObject]@{ RegistryPath = $root; Name = $prop.Value; Clsid = $prop.Name; Source = 'ApprovedShellExtension' }) | Out-Null
                }
            }
            catch {
                Write-Log "Could not enumerate $root : $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue | ForEach-Object {
                $defaultValue = ''
                try { $defaultValue = (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop).'(default)' } catch { }
                if ([string]::IsNullOrWhiteSpace($defaultValue)) { $defaultValue = $_.GetValue('') }
                $items.Add([PSCustomObject]@{ RegistryPath = $_.PSPath; Name = $_.PSChildName; Clsid = [string]$defaultValue; Source = 'ContextMenuHandler' }) | Out-Null
            }
        }
    }

    $nonMicrosoft = @($items | Where-Object {
        $text = "$(Convert-ExplorerSafeString $_.Name) $(Convert-ExplorerSafeString $_.RegistryPath)"
        $text -notmatch '(?i)Microsoft|Windows|Offline Files|Sharing|WorkFolders|FileSyncEx|ModernSharing|EPP|Defender'
    } | Sort-Object Source, Name, Clsid -Unique)

    $script:Summary.NonMicrosoftShellExtensionCount = @($nonMicrosoft).Count
    Write-Log "EXPLORER_SHELL_EXTENSION_SUMMARY|NonMicrosoftCount=$($script:Summary.NonMicrosoftShellExtensionCount)|TotalInventoryCount=$($items.Count)" 'INFO'

    foreach ($entry in ($nonMicrosoft | Select-Object -First 75)) {
        Write-Log "NON_MICROSOFT_SHELL_EXTENSION|Source=$(Convert-ExplorerSafeString $entry.Source)|Name=$(Convert-ExplorerSafeString $entry.Name)|Clsid=$(Convert-ExplorerSafeString $entry.Clsid)|RegistryPath=$(Convert-ExplorerSafeString $entry.RegistryPath)" 'WARN'
    }

    Add-DetailedResult -Step 'ExplorerShellExtensionInventory' -Status 'Info' -Message 'Collected shell extension inventory.' -Data @{
        NonMicrosoftCount = $script:Summary.NonMicrosoftShellExtensionCount
        TotalInventoryCount = $items.Count
        NonMicrosoftJson = (($nonMicrosoft | Select-Object -First 100) | ConvertTo-Json -Compress)
    }

    return $nonMicrosoft
}

function Test-ExplorerNetworkPaths {
    Write-Log 'Checking mapped drives, SMB mappings, and disconnected network paths that can hang Explorer...' 'INFO'

    $problems = New-Object System.Collections.Generic.List[object]

    try {
        $drives = @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.DisplayRoot -match '^\\' })
        foreach ($drive in $drives) {
            $ok = $false
            try { $ok = Test-Path -LiteralPath $drive.DisplayRoot -ErrorAction Stop } catch { $ok = $false }
            if (-not $ok) {
                $problems.Add([PSCustomObject]@{ Type='PSDrive'; Name=$drive.Name; Path=$drive.DisplayRoot; Status='Unreachable' }) | Out-Null
            }
        }
    }
    catch { Write-Log "PSDrive network check failed: $($_.Exception.Message)" 'WARN' }

    try {
        if (Get-Command Get-SmbMapping -ErrorAction SilentlyContinue) {
            $mappings = @(Get-SmbMapping -ErrorAction SilentlyContinue)
            foreach ($map in $mappings) {
                $status = [string]$map.Status
                if ($status -and $status -notmatch '(?i)^OK$|Connected') {
                    $problems.Add([PSCustomObject]@{ Type='SmbMapping'; Name=$map.LocalPath; Path=$map.RemotePath; Status=$status }) | Out-Null
                }
            }
        }
    }
    catch { Write-Log "SMB mapping check failed: $($_.Exception.Message)" 'WARN' }

    try {
        $printers = @(Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue | Where-Object { $_.Network -eq $true -and $_.WorkOffline -eq $true })
        foreach ($printer in $printers) {
            $problems.Add([PSCustomObject]@{ Type='NetworkPrinter'; Name=$printer.Name; Path=$printer.PortName; Status='Offline' }) | Out-Null
        }
    }
    catch { Write-Log "Network printer check failed: $($_.Exception.Message)" 'WARN' }

    $script:Summary.BrokenNetworkPathCount = $problems.Count
    Write-Log "EXPLORER_NETWORK_PATH_SUMMARY|BrokenOrOfflineCount=$($script:Summary.BrokenNetworkPathCount)" 'INFO'

    foreach ($item in $problems) {
        Write-Log "EXPLORER_BROKEN_NETWORK_PATH|Type=$(Convert-ExplorerSafeString $item.Type)|Name=$(Convert-ExplorerSafeString $item.Name)|Path=$(Convert-ExplorerSafeString $item.Path)|Status=$(Convert-ExplorerSafeString $item.Status)" 'WARN'
    }

    Add-DetailedResult -Step 'ExplorerNetworkPathCheck' -Status 'Info' -Message 'Checked mapped drives, SMB mappings, and offline network printers.' -Data @{
        BrokenOrOfflineCount = $script:Summary.BrokenNetworkPathCount
        ProblemsJson = ($problems | ConvertTo-Json -Compress)
    }

    return $problems
}

function Invoke-ExplorerCacheCleanup {
    Write-Log 'Cleaning Explorer thumbnail, icon, and Quick Access caches...' 'WARN'

    $deleted = New-Object System.Collections.Generic.List[string]
    $targets = @(
        @{ Path = "$env:LOCALAPPDATA\IconCache.db"; Pattern = $null },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Pattern = 'thumbcache*.db' },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Pattern = 'iconcache*.db' },
        @{ Path = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"; Pattern = '*' },
        @{ Path = "$env:APPDATA\Microsoft\Windows\Recent\CustomDestinations"; Pattern = '*' }
    )

    foreach ($target in $targets) {
        if (-not (Test-Path -LiteralPath $target.Path)) { continue }
        if ($null -eq $target.Pattern) {
            try {
                Remove-Item -LiteralPath $target.Path -Force -ErrorAction Stop
                $deleted.Add($target.Path) | Out-Null
            }
            catch { Write-Log "Could not delete Explorer cache file $($target.Path): $($_.Exception.Message)" 'WARN' }
        }
        else {
            Get-ChildItem -LiteralPath $target.Path -Filter $target.Pattern -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                    $deleted.Add($_.FullName) | Out-Null
                }
                catch { Write-Log "Could not delete Explorer cache item $($_.FullName): $($_.Exception.Message)" 'WARN' }
            }
        }
    }

    Write-Log "EXPLORER_CACHE_CLEANUP|DeletedCount=$($deleted.Count)" 'INFO'
    Add-RepairAttempt 'Explorer thumbnail/icon/Quick Access cache cleanup'
    Add-DetailedResult -Step 'ExplorerCacheCleanup' -Status 'Info' -Message 'Explorer cache cleanup completed.' -Data @{
        DeletedCount = $deleted.Count
        DeletedItems = ($deleted -join '; ')
    }
}

function Invoke-SearchServiceRestart {
    if (-not $AllowSearchServiceRestart) {
        Write-Log 'Windows Search service restart skipped because AllowSearchServiceRestart is disabled.' 'INFO'
        return
    }

    try {
        $svc = Get-Service -Name WSearch -ErrorAction SilentlyContinue
        if ($null -eq $svc) {
            Write-Log 'Windows Search service was not found on this system.' 'INFO'
            return
        }

        if ($svc.Status -eq 'Running') {
            Write-Log 'Restarting Windows Search service to clear possible Explorer/indexer hangs...' 'WARN'
            Restart-Service -Name WSearch -Force -ErrorAction Stop
        }
        else {
            Write-Log "Starting Windows Search service. Current status: $($svc.Status)" 'WARN'
            Start-Service -Name WSearch -ErrorAction Stop
        }

        Add-RepairAttempt 'Windows Search service restart/start'
        Add-DetailedResult -Step 'SearchServiceRestart' -Status 'Info' -Message 'Windows Search service restart/start completed.'
        Write-Log 'EXPLORER_SEARCH_SERVICE_RESTARTED|Service=WSearch' 'INFO'
    }
    catch {
        Warn-Step -Name 'SearchServiceRestart' -Reason $_.Exception.Message
    }
}

function Get-ExplorerRelatedDriverInventory {
    Write-Log 'Collecting GPU and storage controller driver inventory for Explorer hang correlation...' 'INFO'
    $video = @()
    $storage = @()

    try { $video = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue) } catch { }
    try { $storage = @(Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue | Where-Object { $_.DeviceClass -match 'DISPLAY|SCSIAdapter|HDC|System' -and $_.DeviceName -match '(?i)graphics|display|intel|amd|nvidia|storage|nvme|controller|chipset|vmd|raid|sata' }) } catch { }

    foreach ($gpu in $video) {
        Write-Log "EXPLORER_GPU_DRIVER|Name=$(Convert-ExplorerSafeString $gpu.Name)|DriverVersion=$(Convert-ExplorerSafeString $gpu.DriverVersion)|DriverDate=$(Convert-ExplorerSafeString $gpu.DriverDate)|Status=$(Convert-ExplorerSafeString $gpu.Status)" 'INFO'
    }

    foreach ($driver in ($storage | Select-Object -First 50)) {
        Write-Log "EXPLORER_RELATED_DRIVER|Device=$(Convert-ExplorerSafeString $driver.DeviceName)|Class=$(Convert-ExplorerSafeString $driver.DeviceClass)|Provider=$(Convert-ExplorerSafeString $driver.DriverProviderName)|Version=$(Convert-ExplorerSafeString $driver.DriverVersion)|Date=$(Convert-ExplorerSafeString $driver.DriverDate)" 'INFO'
    }

    Add-DetailedResult -Step 'ExplorerRelatedDriverInventory' -Status 'Info' -Message 'Collected GPU/storage driver inventory.' -Data @{
        VideoControllerCount = @($video).Count
        RelatedDriverCount = @($storage).Count
    }
}

function Invoke-ExplorerHealthDiagnostics {
    Write-Log 'Starting Explorer health diagnostics and safe repairs...' 'INFO'

    $events = @(Get-ExplorerEventSummary -DaysBack 14)
    $shellExtensions = @(Get-NonMicrosoftShellExtensionInventory)
    $networkProblems = @(Test-ExplorerNetworkPaths)
    Get-ExplorerRelatedDriverInventory

    if ($script:Summary.ExplorerHangCount -ge 3 -or $script:Summary.ExplorerCrashCount -ge 3 -or @($networkProblems).Count -gt 0) {
        $script:Summary.ExplorerInstabilityDetected = $true
        Write-Log "EXPLORER_INSTABILITY_DETECTED|Hangs=$($script:Summary.ExplorerHangCount)|Crashes=$($script:Summary.ExplorerCrashCount)|BrokenNetworkPaths=$($script:Summary.BrokenNetworkPathCount)|NonMicrosoftShellExtensions=$($script:Summary.NonMicrosoftShellExtensionCount)" 'WARN'
        Add-Note "Explorer instability detected. Review EXPLORER_* log markers, non-Microsoft shell extensions, broken network paths, GPU/storage drivers, and user profile-specific cache/index issues."
    }
    else {
        Write-Log 'No repeated Explorer instability pattern detected from recent Application log events.' 'OK'
    }

    if ($AllowExplorerHealthRepair) {
        Invoke-ExplorerCacheCleanup
        Invoke-SearchServiceRestart
    }
    else {
        Write-Log 'Explorer cache cleanup and Windows Search restart skipped because AllowExplorerHealthRepair is disabled.' 'INFO'
    }

    Add-DetailedResult -Step 'ExplorerHealthDiagnostics' -Status 'Info' -Message 'Explorer health diagnostics completed.' -Data @{
        EventsAnalyzed = @($events).Count
        NonMicrosoftShellExtensionCount = @($shellExtensions).Count
        BrokenNetworkPathCount = @($networkProblems).Count
        RepairsAllowed = $AllowExplorerHealthRepair
    }
}

function Invoke-IconCacheRebuild {
    Write-Log "Rebuilding icon and thumbnail caches..." 'WARN'

    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    $explorerCachePath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
    $deletedFiles = New-Object System.Collections.Generic.List[string]

    $singleFileTargets = @(
        "$env:LOCALAPPDATA\IconCache.db"
    )

    foreach ($path in $singleFileTargets) {
        if (Test-Path -LiteralPath $path) {
            try {
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                $deletedFiles.Add((Split-Path -Leaf $path)) | Out-Null
            }
            catch {
                Write-Log "Failed to delete cache file $path : $($_.Exception.Message)" 'WARN'
            }
        }
    }

    if (Test-Path -LiteralPath $explorerCachePath) {
        $patterns = @(
            'iconcache*',
            'thumbcache_*.db',
            'thumbcache_idx.db'
        )

        foreach ($pattern in $patterns) {
            Get-ChildItem -LiteralPath $explorerCachePath -Filter $pattern -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                    $deletedFiles.Add($_.Name) | Out-Null
                }
                catch {
                    Write-Log "Failed to delete cache file $($_.FullName): $($_.Exception.Message)" 'WARN'
                }
            }
        }
    }

    Start-Process explorer.exe
    Add-RepairAttempt 'Icon and thumbnail cache rebuild'
    Add-DetailedResult -Step 'IconCacheRebuild' -Status 'Info' -Message 'Icon and thumbnail cache rebuild completed.' -Data @{
        DeletedFiles = ($deletedFiles -join '; ')
    }
}


function Set-RegistryDWORDValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
}

function Disable-CopilotForLoadedUsers {
    $targetSids = New-Object System.Collections.Generic.List[string]
    $targetSids.Add('HKEY_CURRENT_USER') | Out-Null

    Get-ChildItem Registry::HKEY_USERS -ErrorAction SilentlyContinue |
        Where-Object {
            $_.PSChildName -match '^S-1-5-21-' -and
            $_.PSChildName -notmatch '_Classes$'
        } |
        ForEach-Object {
            $targetSids.Add("HKEY_USERS\\$($_.PSChildName)") | Out-Null
        }

    foreach ($root in $targetSids | Select-Object -Unique) {
        $policyPath = "Registry::$root\Software\Policies\Microsoft\Windows\WindowsCopilot"
        $explorerPath = "Registry::$root\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        Set-RegistryDWORDValue -Path $policyPath -Name 'TurnOffWindowsCopilot' -Value 1
        Set-RegistryDWORDValue -Path $explorerPath -Name 'ShowCopilotButton' -Value 0

        Add-DetailedResult -Step 'CopilotDisableRegistry' -Status 'Info' -Message 'Applied Copilot disable settings for loaded profile.' -Data @{
            Root = $root
            PolicyPath = $policyPath
            ExplorerPath = $explorerPath
        }
    }
}

function Disable-CopilotForDefaultUser {
    $defaultHiveName = 'HKU\DefaultTempCopilot'
    $defaultHivePsPath = 'Registry::HKEY_USERS\DefaultTempCopilot'
    $defaultUserNtUserDat = 'C:\Users\Default\NTUSER.DAT'

    if (-not (Test-Path -LiteralPath $defaultUserNtUserDat)) {
        Write-Log 'Default User NTUSER.DAT not found; future new users were not updated for Copilot disable.' 'WARN'
        return
    }

    $hiveLoaded = $false
    try {
        $loadResult = & reg.exe load $defaultHiveName $defaultUserNtUserDat
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to load Default User hive: $($loadResult -join ' ')"
        }

        $hiveLoaded = $true
        Start-Sleep -Milliseconds 750

        $policyPath = "$defaultHivePsPath\Software\Policies\Microsoft\Windows\WindowsCopilot"
        $explorerPath = "$defaultHivePsPath\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        Set-RegistryDWORDValue -Path $policyPath -Name 'TurnOffWindowsCopilot' -Value 1
        Set-RegistryDWORDValue -Path $explorerPath -Name 'ShowCopilotButton' -Value 0

        Add-DetailedResult -Step 'CopilotDisableDefaultUser' -Status 'Info' -Message 'Applied Copilot disable settings for Default User profile.' -Data @{
            PolicyPath = $policyPath
            ExplorerPath = $explorerPath
        }
    }
    finally {
        if ($hiveLoaded) {
            Start-Sleep -Milliseconds 750
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 750
            & reg.exe unload $defaultHiveName | Out-Null
        }
    }
}

function Invoke-CopilotDisableAndRemoval {
    Write-Log 'Disabling Microsoft Copilot for current, loaded, and future user profiles...' 'WARN'

    Set-RegistryDWORDValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' -Name 'TurnOffWindowsCopilot' -Value 1
    Set-RegistryDWORDValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer' -Name 'HideCopilotButton' -Value 1

    Disable-CopilotForLoadedUsers
    Disable-CopilotForDefaultUser

    $removedPackages = New-Object System.Collections.Generic.List[string]
    $packagePatterns = @(
        'Microsoft.Windows.Copilot',
        '*Copilot*'
    )

    foreach ($pattern in $packagePatterns) {
        $packages = @(Get-AppxPackage -AllUsers -Name $pattern -ErrorAction SilentlyContinue)
        foreach ($pkg in $packages) {
            if ($removedPackages -contains $pkg.PackageFullName) {
                continue
            }

            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
                $removedPackages.Add($pkg.PackageFullName) | Out-Null
                Write-Log "Removed Copilot Appx package: $($pkg.Name) [$($pkg.PackageFullName)]" 'OK'
            }
            catch {
                Write-Log "Failed to remove Copilot Appx package $($pkg.PackageFullName): $($_.Exception.Message)" 'WARN'
            }
        }

        $provisionedPackages = @(Get-AppxProvisionedPackage -Online | Where-Object {
            $_.DisplayName -like $pattern -or $_.PackageName -like $pattern
        })

        foreach ($prov in $provisionedPackages) {
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
                $removedPackages.Add($prov.PackageName) | Out-Null
                Write-Log "Removed provisioned Copilot package: $($prov.DisplayName) [$($prov.PackageName)]" 'OK'
            }
            catch {
                Write-Log "Failed to remove provisioned Copilot package $($prov.PackageName): $($_.Exception.Message)" 'WARN'
            }
        }
    }

    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer.exe

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Microsoft Copilot disable and removal'
    Add-DetailedResult -Step 'CopilotDisableAndRemoval' -Status 'Info' -Message 'Microsoft Copilot disable and removal routine completed.' -Data @{
        RemovedPackages = ($removedPackages | Select-Object -Unique) -join '; '
    }
}

function Invoke-FirewallReset {
    Write-Log "Resetting Windows Firewall to defaults..." 'WARN'
    netsh advfirewall reset | Out-Null
    Add-RepairAttempt 'Firewall reset'
    Add-DetailedResult -Step 'FirewallReset' -Status 'Info' -Message 'Firewall reset completed.'
}

function Get-ServiceStateSafe {
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop
        return [PSCustomObject]@{
            Name        = $svc.Name
            DisplayName = $svc.DisplayName
            Status      = [string]$svc.Status
            Exists      = $true
            Error       = $null
        }
    }
    catch {
        return [PSCustomObject]@{
            Name        = $Name
            DisplayName = $null
            Status      = 'NotFound'
            Exists      = $false
            Error       = $_.Exception.Message
        }
    }
}

function Wait-ServiceStateSafe {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('Running','Stopped')][string]$DesiredStatus,
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        $state = Get-ServiceStateSafe -Name $Name
        if (-not $state.Exists) {
            return $state
        }

        if ($state.Status -eq $DesiredStatus) {
            return $state
        }

        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $deadline)

    return Get-ServiceStateSafe -Name $Name
}

function Get-ServiceProcessIdSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $cimService = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction Stop
        if ($null -ne $cimService -and $cimService.ProcessId -gt 0) {
            return [int]$cimService.ProcessId
        }
    }
    catch {
        Write-Log "Unable to read process ID for service $Name`: $($_.Exception.Message)" 'WARN'
    }

    return 0
}

function Stop-ServiceWithValidation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$TimeoutSeconds = 30,
        [switch]$KillProcessOnTimeout
    )

    $before = Get-ServiceStateSafe -Name $Name
    $killedProcess = $false
    $processId = 0
    $killMessage = $null

    if (-not $before.Exists) {
        Write-Log "Service validation: $Name was not found. Skipping stop." 'WARN'
        return [PSCustomObject]@{
            Name          = $Name
            BeforeStatus  = $before.Status
            AfterStatus   = $before.Status
            Success       = $true
            ProcessId     = 0
            ProcessKilled = $false
            Message       = 'Service not found; skipped'
        }
    }

    Write-Log "Service validation: $Name current state is $($before.Status)." 'INFO'

    if ($before.Status -eq 'Stopped') {
        Write-Log "Service validation: $Name is already stopped." 'OK'
        return [PSCustomObject]@{
            Name          = $Name
            BeforeStatus  = $before.Status
            AfterStatus   = 'Stopped'
            Success       = $true
            ProcessId     = 0
            ProcessKilled = $false
            Message       = 'Already stopped'
        }
    }

    $processId = Get-ServiceProcessIdSafe -Name $Name

    # Stop-Service can itself block indefinitely while a service remains in StopPending.
    # Send the stop request through sc.exe in a separate process so the request has its
    # own hard timeout before the service-state timeout and process-kill fallback begin.
    $stopRequestTimeoutSeconds = [math]::Min([math]::Max(5, $TimeoutSeconds), 15)
    $scProcess = $null

    try {
        Write-Log "Requesting stop for service $Name. Stop-request timeout: $stopRequestTimeoutSeconds second(s); service-state timeout: $TimeoutSeconds second(s)..." 'INFO'

        $scProcess = Start-Process -FilePath "$env:SystemRoot\System32\sc.exe" `
            -ArgumentList @('stop', $Name) `
            -WindowStyle Hidden `
            -PassThru `
            -ErrorAction Stop

        if (-not $scProcess.WaitForExit($stopRequestTimeoutSeconds * 1000)) {
            Write-Log "The stop request for service $Name exceeded $stopRequestTimeoutSeconds second(s). Terminating the hung sc.exe request process and continuing with validation." 'WARN'
            try {
                $scProcess.Kill()
                $scProcess.WaitForExit(5000) | Out-Null
            }
            catch {
                Write-Log "Unable to terminate the hung sc.exe request process for $Name`: $($_.Exception.Message)" 'WARN'
            }
        }
        elseif ($scProcess.ExitCode -notin @(0, 1062)) {
            Write-Log "sc.exe returned exit code $($scProcess.ExitCode) while requesting that $Name stop. Service-state validation will continue." 'WARN'
        }
    }
    catch {
        Write-Log "The timed stop request reported an issue for $Name`: $($_.Exception.Message). Service-state validation will continue." 'WARN'
    }
    finally {
        if ($null -ne $scProcess) {
            $scProcess.Dispose()
        }
    }

    $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Stopped' -TimeoutSeconds $TimeoutSeconds

    if ($after.Status -ne 'Stopped' -and $KillProcessOnTimeout) {
        if ($processId -le 0) {
            $processId = Get-ServiceProcessIdSafe -Name $Name
        }

        if ($processId -gt 0) {
            try {
                Write-Log "Service $Name did not stop within $TimeoutSeconds second(s). Killing backing process PID $processId..." 'WARN'
                Stop-Process -Id $processId -Force -ErrorAction Stop
                $killedProcess = $true
                Start-Sleep -Seconds 2
                $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Stopped' -TimeoutSeconds 10
                $killMessage = "Killed backing process PID $processId after timeout"
            }
            catch {
                $killMessage = "Failed to kill backing process PID $processId`: $($_.Exception.Message)"
                Write-Log $killMessage 'ERROR'
            }
        }
        else {
            $killMessage = 'No valid backing process ID was available to kill after timeout'
            Write-Log "Service $Name did not stop within $TimeoutSeconds second(s), but no valid backing process ID was found." 'WARN'
        }
    }

    $success = ($after.Status -eq 'Stopped')

    if ($success) {
        if ($killedProcess) {
            Write-Log "Service validation: $Name stopped after killing PID $processId." 'OK'
        }
        else {
            Write-Log "Service validation: $Name stopped successfully." 'OK'
        }
    }
    else {
        Write-Log "Service validation: $Name did not stop. Current state: $($after.Status)." 'WARN'
    }

    $message = if ($success -and $killedProcess) {
        $killMessage
    }
    elseif ($success) {
        'Stopped successfully'
    }
    elseif ($killMessage) {
        "Expected Stopped but found $($after.Status). $killMessage"
    }
    else {
        "Expected Stopped but found $($after.Status)"
    }

    return [PSCustomObject]@{
        Name          = $Name
        BeforeStatus  = $before.Status
        AfterStatus   = $after.Status
        Success       = $success
        ProcessId     = $processId
        ProcessKilled = $killedProcess
        Message       = $message
    }
}

function Start-ServiceWithValidation {
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$TimeoutSeconds = 30
    )

    $before = Get-ServiceStateSafe -Name $Name

    if (-not $before.Exists) {
        Write-Log "Service validation: $Name was not found. Skipping start." 'WARN'
        return [PSCustomObject]@{
            Name         = $Name
            BeforeStatus = $before.Status
            AfterStatus  = $before.Status
            Success      = $true
            Message      = 'Service not found; skipped'
        }
    }

    if ($before.Status -eq 'Running') {
        Write-Log "Service validation: $Name is already running." 'OK'
        return [PSCustomObject]@{
            Name         = $Name
            BeforeStatus = $before.Status
            AfterStatus  = 'Running'
            Success      = $true
            Message      = 'Already running'
        }
    }

    try {
        Write-Log "Starting service $Name..." 'INFO'
        Start-Service -Name $Name -ErrorAction Stop
    }
    catch {
        Write-Log "Start-Service reported an issue for $Name`: $($_.Exception.Message)" 'WARN'
    }

    $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Running' -TimeoutSeconds $TimeoutSeconds
    $success = ($after.Status -eq 'Running')

    if ($success) {
        Write-Log "Service validation: $Name started successfully." 'OK'
    }
    else {
        Write-Log "Service validation: $Name did not start. Current state: $($after.Status)." 'WARN'
    }

    return [PSCustomObject]@{
        Name         = $Name
        BeforeStatus = $before.Status
        AfterStatus  = $after.Status
        Success      = $success
        Message      = if ($success) { 'Started successfully' } else { "Expected Running but found $($after.Status)" }
    }
}


function Remove-PathWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Description,
        [int]$MaxAttempts = 5,
        [int]$InitialDelaySeconds = 2,
        [switch]$ReleaseWindowsUpdateLocks
    )

    $result = [ordered]@{
        Path            = $Path
        Description     = $Description
        ExistsBefore    = $false
        Deleted         = $false
        Attempts        = 0
        ItemCount       = 0
        FileCount       = 0
        FolderCount     = 0
        SizeBytesBefore = [int64]0
        SizeMBBefore    = [double]0
        SizeGBBefore    = [double]0
        SizeBytesAfter  = [int64]0
        SizeMBAfter     = [double]0
        SizeGBAfter     = [double]0
        SpaceFreed      = [int64]0
        EstimatedFreedBytes = [int64]0
        EstimatedFreedMB = [double]0
        EstimatedFreedGB = [double]0
        Message         = $null
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $result.Message = 'Path does not exist; nothing to delete'
        Write-Log "$Description does not exist at $Path. Nothing to delete." 'INFO'
        return [PSCustomObject]$result
    }

    $result.ExistsBefore = $true

    $beforeInfo = Get-FolderSizeInfo -Path $Path
    $result.ItemCount = $beforeInfo.ItemCount
    $result.FileCount = $beforeInfo.FileCount
    $result.FolderCount = $beforeInfo.FolderCount
    $result.SizeBytesBefore = [int64]$beforeInfo.SizeBytes
    $result.SizeMBBefore = [double]$beforeInfo.SizeMB
    $result.SizeGBBefore = [double]$beforeInfo.SizeGB

    Write-Log "Preparing to delete $Description at $Path. No backup will be created." 'INFO'
    Write-Log "Size before deletion for $Description`: $($result.SizeGBBefore) GB ($($result.SizeMBBefore) MB), items: $($result.ItemCount), files: $($result.FileCount), folders: $($result.FolderCount)." 'INFO'

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $result.Attempts = $attempt

        try {
            if (-not (Test-Path -LiteralPath $Path)) {
                $result.Deleted = $true
                $result.Message = 'Path already gone during retry validation'
                Write-Log "$Description no longer exists at $Path." 'OK'
                break
            }

            if ($ReleaseWindowsUpdateLocks -and $attempt -gt 1) {
                Write-Log "Attempt $attempt is releasing possible Windows Update locks before retrying $Description deletion." 'WARN'
                Stop-WindowsUpdateLockingProcesses | Out-Null
            }

            Write-Log "Deletion attempt $attempt of $MaxAttempts for $Description..." 'INFO'
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop

            Start-Sleep -Seconds 1
            if (-not (Test-Path -LiteralPath $Path)) {
                $result.Deleted = $true
                $result.Message = 'Deleted successfully'
                Write-Log "Deleted $Description successfully on attempt $attempt." 'OK'
                break
            }

            throw "$Description still exists after Remove-Item completed."
        }
        catch {
            $result.Message = $_.Exception.Message
            Write-Log "Deletion attempt $attempt failed for $Description`: $($result.Message)" 'WARN'

            if ($ReleaseWindowsUpdateLocks) {
                Stop-WindowsUpdateLockingProcesses | Out-Null
            }

            if ($attempt -lt $MaxAttempts) {
                $delay = $InitialDelaySeconds * $attempt
                Write-Log "Waiting $delay second(s), then retrying $Description deletion. Files may still be locked by services or Windows Update processes." 'INFO'
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                Start-Sleep -Seconds $delay
            }
        }
    }

    if (Test-Path -LiteralPath $Path) {
        $afterInfo = Get-FolderSizeInfo -Path $Path
        $result.SizeBytesAfter = [int64]$afterInfo.SizeBytes
        $result.SizeMBAfter = [double]$afterInfo.SizeMB
        $result.SizeGBAfter = [double]$afterInfo.SizeGB
    }

    $freedBytes = [int64]([math]::Max(0, ([int64]$result.SizeBytesBefore - [int64]$result.SizeBytesAfter)))
    $result.SpaceFreed = $freedBytes
    $result.EstimatedFreedBytes = $freedBytes
    $result.EstimatedFreedMB = [math]::Round(([double]$freedBytes / 1MB), 2)
    $result.EstimatedFreedGB = [math]::Round(([double]$freedBytes / 1GB), 3)

    if ($result.Deleted) {
        Write-Log "Estimated space freed by deleting $Description`: $($result.EstimatedFreedGB) GB ($($result.EstimatedFreedMB) MB)." 'OK'
    }
    else {
        Write-Log "Failed to delete $Description after $MaxAttempts attempt(s). Last error: $($result.Message). Estimated remaining size: $($result.SizeGBAfter) GB." 'ERROR'
    }

    return [PSCustomObject]$result
}

function Invoke-WindowsUpdateComponentReset {
    Write-Log "Resetting Windows Update components..." 'WARN'
    Write-Log "SoftwareDistribution will be deleted directly. No .bak, .bak1, or timestamped backup folder will be created." 'INFO'

    $services = @('wuauserv','bits','cryptsvc','msiserver','usosvc','DoSvc','WaaSMedicSvc')
    $stopResults = New-Object System.Collections.Generic.List[object]
    $startResults = New-Object System.Collections.Generic.List[object]
    $deleteResults = New-Object System.Collections.Generic.List[object]
    $lockProcessResults = New-Object System.Collections.Generic.List[object]
    $mainCleanupResult = $null
    $backupCleanupResult = $null

    foreach ($svc in $services) {
        $stopResults.Add((Stop-ServiceWithValidation -Name $svc -TimeoutSeconds 45 -KillProcessOnTimeout)) | Out-Null
    }

    $criticalServicesStillRunning = @($stopResults | Where-Object {
        $_.Name -in @('wuauserv','bits','cryptsvc') -and $_.Success -eq $false
    })

    if ($criticalServicesStillRunning.Count -gt 0) {
        Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more Windows Update services did not stop cleanly: $($criticalServicesStillRunning.Name -join ', ')"
    }

    Start-Sleep -Seconds 2

    foreach ($lockResult in @(Stop-WindowsUpdateLockingProcesses)) { $lockProcessResults.Add($lockResult) | Out-Null }

    $mainCleanupResult = Remove-SoftwareDistributionMainFolder
    $backupCleanupResult = Remove-SoftwareDistributionBakFolders

    $paths = @(
        @{ Path = "$env:WINDIR\System32\catroot2"; Description = 'Windows Update Catroot2 folder' }
    )

    foreach ($target in $paths) {
        $deleteResults.Add((Remove-PathWithRetry -Path $target.Path -Description $target.Description -MaxAttempts 5 -InitialDelaySeconds 2 -ReleaseWindowsUpdateLocks)) | Out-Null
    }

    foreach ($svc in $services) {
        $startResults.Add((Start-ServiceWithValidation -Name $svc -TimeoutSeconds 45)) | Out-Null
    }

    $failedDeletes = @($deleteResults | Where-Object { $_.ExistsBefore -eq $true -and $_.Deleted -eq $false })
    $failedStarts = @($startResults | Where-Object { $_.Success -eq $false })

    if ($failedDeletes.Count -gt 0) {
        Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more update folders could not be deleted: $($failedDeletes.Description -join ', ')"
    }

    if ($failedStarts.Count -gt 0) {
        Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more update services did not restart cleanly: $($failedStarts.Name -join ', ')"
    }

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Windows Update component reset with robocopy mirror SoftwareDistribution cleanup, SoftwareDistribution.bak* cleanup, folder size logging, lock release, and service validation'
    Add-DetailedResult -Step 'WindowsUpdateComponentReset' -Status 'Info' -Message 'Windows Update components reset. SoftwareDistribution and SoftwareDistribution.bak* folders were cleaned using robocopy mirror cleanup; estimated GB freed was logged.' -Data @{
        StoppedServicesJson = ($stopResults | ConvertTo-Json -Compress)
        LockProcessesJson  = ($lockProcessResults | ConvertTo-Json -Compress)
        MainSoftwareDistributionCleanupJson = ($mainCleanupResult | ConvertTo-Json -Compress)
        BackupCleanupJson  = ($backupCleanupResult | ConvertTo-Json -Compress)
        DeletedPathsJson   = ($deleteResults | ConvertTo-Json -Compress)
        StartedServicesJson = ($startResults | ConvertTo-Json -Compress)
    }
}

function Invoke-ScheduledTaskHealthCheck {
    Write-Log "Checking scheduled task health..." 'INFO'

    $badTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
        $_.State -eq 'Unknown' -or $_.TaskPath -eq $null
    }

    $taskNames = @()

    if ($badTasks) {
        foreach ($task in $badTasks) {
            $name = "$($task.TaskPath)$($task.TaskName)"
            $taskNames += $name
            Warn-Step -Name 'ScheduledTaskCheck' -Reason "Task may need review: $name"
        }
    }

    Add-DetailedResult -Step 'ScheduledTaskHealthCheck' -Status 'Info' -Message 'Scheduled task health check completed.' -Data @{
        SuspectTaskCount = $taskNames.Count
        SuspectTasks     = ($taskNames -join '; ')
    }
}

function Invoke-EventLogSummary {
    Write-Log "Collecting recent system health events..." 'INFO'

    try {
        $recentUnexpected = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Id      = 41, 6008
            StartTime = (Get-Date).AddDays(-7)
        } -ErrorAction Stop

        $count = @($recentUnexpected).Count

        Add-DetailedResult -Step 'EventLogSummary' -Status 'Info' -Message 'Collected recent unexpected shutdown events.' -Data @{
            UnexpectedShutdownCount = $count
        }

        if ($recentUnexpected) {
            Warn-Step -Name 'EventLogSummary' -Reason "Recent unexpected shutdown events found: $count"
        }
    }
    catch {
        $eventLogError = $_.Exception.Message
        Warn-Step -Name 'EventLogSummary' -Reason $eventLogError

        if ($eventLogError -match 'RPC server is unavailable|0x800706BA|1722') {
            $rpcStatus = if ($script:Summary.RpcDiagnosticsStatus) { $script:Summary.RpcDiagnosticsStatus } else { 'RpcDiagnosticsNotRunOrNoStatusAvailable' }
            $clientSide = if ($null -ne $script:Summary.RpcClientSideIssueDetected) { $script:Summary.RpcClientSideIssueDetected } else { $false }
            $remoteSide = if ($null -ne $script:Summary.RpcRemoteSideIssueDetected) { $script:Summary.RpcRemoteSideIssueDetected } else { $false }

            $classification = if ($clientSide) {
                'LOCAL_CLIENT_OR_LOCAL_SERVICE_PROBLEM'
            }
            elseif ($remoteSide) {
                'REMOTE_ENDPOINT_OR_NETWORK_PROBLEM'
            }
            elseif ($rpcStatus -eq 'NoRpcIssueDetectedByBasicChecks') {
                'EVENTLOG_PROVIDER_OR_QUERY_SPECIFIC_PROBLEM'
            }
            else {
                'UNCLASSIFIED_RPC_FAILURE_REVIEW_RPC_DIAGNOSTICS'
            }

            Write-Log "RPC_EVENTLOGSUMMARY_CLASSIFICATION|Classification=$classification|RpcDiagnosticsStatus=$rpcStatus|ClientSideIssue=$clientSide|RemoteOrNetworkIssue=$remoteSide|OriginalError=$(Convert-RpcSafeString $eventLogError)" 'WARN'
            Add-Note "EventLogSummary hit RPC unavailable. Classification: $classification. RPC diagnostics status: $rpcStatus."
        }
    }
}

function Invoke-EventLogClear {
    Write-Log "Clearing classic event logs..." 'WARN'
    wevtutil el | ForEach-Object {
        try {
            wevtutil cl $_ 2>$null
        }
        catch {
        }
    }
    Add-RepairAttempt 'Event log clear'
    Add-DetailedResult -Step 'EventLogClear' -Status 'Info' -Message 'Event log clear attempted.'
}

function Invoke-DiskSpaceCheck {
    $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
    if ($drive) {
        $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
        $sizeGB = [math]::Round($drive.Size / 1GB, 2)
        Write-Log "System drive free space: $freeGB GB of $sizeGB GB" 'INFO'

        Add-DetailedResult -Step 'DiskSpaceCheck' -Status 'Info' -Message 'Disk space checked.' -Data @{
            FreeGB  = $freeGB
            TotalGB = $sizeGB
        }

        if ($freeGB -lt 10) {
            Warn-Step -Name 'DiskSpaceCheck' -Reason "Low free space on system drive: $freeGB GB"
        }
    }
}

function Invoke-DefenderStatusCheck {
    try {
        $status = Get-MpComputerStatus -ErrorAction Stop
        Write-Log "Defender Antivirus Enabled: $($status.AntivirusEnabled)" 'INFO'
        Write-Log "Defender RealTime Protection Enabled: $($status.RealTimeProtectionEnabled)" 'INFO'

        Add-DetailedResult -Step 'DefenderStatusCheck' -Status 'Info' -Message 'Defender status checked.' -Data @{
            AntivirusEnabled          = $status.AntivirusEnabled
            RealTimeProtectionEnabled = $status.RealTimeProtectionEnabled
        }
    }
    catch {
        Warn-Step -Name 'DefenderStatusCheck' -Reason $_.Exception.Message
    }
}

function Invoke-RebootIfNeeded {
    param(
        [int]$DelaySeconds = 60
    )

    $comment = 'Restarting after automated system repair operations.'

    $args = @(
        '/r'
        '/t', $DelaySeconds.ToString()
        '/d', 'p:2:17'
        '/c', "`"$comment`""
        '/f'
    )

    Write-Log "Issuing reboot command: shutdown.exe $($args -join ' ')" 'WARN'
    & "$env:SystemRoot\System32\shutdown.exe" @args

    if ($LASTEXITCODE -ne 0) {
        throw "shutdown.exe returned exit code $LASTEXITCODE"
    }

    $script:RebootInitiated = $true
    Add-DetailedResult -Step 'AutoReboot' -Status 'Info' -Message 'Automatic reboot command issued.' -Data @{
        DelaySeconds = $DelaySeconds
    }
}

function Show-Summary {
    $script:Summary.EndTime = Get-Date

    Write-Log "---------------- Summary ----------------" 'INFO'
    Write-Log "Computer Name: $($script:Summary.ComputerName)" 'INFO'
    Write-Log "Start Time: $($script:Summary.StartTime)" 'INFO'
    Write-Log "End Time:   $($script:Summary.EndTime)" 'INFO'
    Write-Log "YAML Log:   $($script:YamlLogPath)" 'INFO'
    Write-Log "Succeeded:  $($script:Summary.StepsSucceeded)" 'INFO'
    Write-Log "Failed:     $($script:Summary.StepsFailed)" 'INFO'
    Write-Log "Warnings:   $($script:Summary.Warnings)" 'INFO'
    Write-Log "Pending Reboot Detected: $($script:Summary.PendingRebootDetected)" 'INFO'
    Write-Log "Reboot Required: $($script:Summary.RebootRequired)" 'INFO'
    Write-Log "Disk Corruption Suspected: $($script:Summary.DiskCorruptionSuspected)" 'INFO'
    Write-Log "DISM Corruption Detected: $($script:Summary.DismCorruptionDetected)" 'INFO'
    Write-Log "SFC Integrity Violations: $($script:Summary.SfcIntegrityViolations)" 'INFO'
    Write-Log "SFC Repair Attempted: $($script:Summary.SfcRepairAttempted)" 'INFO'
    Write-Log "SFC Repair Succeeded: $($script:Summary.SfcRepairSucceeded)" 'INFO'
    Write-Log "CBS Corrupt File Count: $($script:Summary.CbsCorruptFileCount)" 'INFO'
    Write-Log "CBS Corruption Extract: $($script:Summary.CbsCorruptionExtractPath)" 'INFO'
    Write-Log "CBS Driver Corruption Detected: $($script:Summary.CbsDriverCorruptionDetected)" 'INFO'
    Write-Log "CBS Driver Corruption Remaining: $($script:Summary.CbsDriverCorruptionRemaining)" 'INFO'
    Write-Log "CBS Corrupt Driver Files: $((($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique) -join ','))" 'INFO'
    Write-Log "HPIA Driver Repair Attempted: $($script:Summary.HpDriverRepairAttempted)" 'INFO'
    Write-Log "HPIA Driver Repair Succeeded: $($script:Summary.HpDriverRepairSucceeded)" 'INFO'
    Write-Log "HPIA Path: $($script:Summary.HpImageAssistantPath)" 'INFO'
    Write-Log "WMI Repository Inconsistent: $($script:Summary.WmiRepositoryInconsistent)" 'INFO'
    Write-Log "Storage Health Warnings: $($script:Summary.StorageHealthWarnings)" 'INFO'
    Write-Log "Storage Failure Predicted: $($script:Summary.StorageFailurePredicted)" 'INFO'
    Write-Log "Storage Wear Over 50 Percent Detected: $($script:Summary.StorageWearOver50Detected)" 'INFO'
    Write-Log "Explorer Hang Count: $($script:Summary.ExplorerHangCount)" 'INFO'
    Write-Log "Explorer Crash Count: $($script:Summary.ExplorerCrashCount)" 'INFO'
    Write-Log "Explorer Instability Detected: $($script:Summary.ExplorerInstabilityDetected)" 'INFO'
    Write-Log "Non-Microsoft Shell Extension Count: $($script:Summary.NonMicrosoftShellExtensionCount)" 'INFO'
    Write-Log "Broken Network Path Count: $($script:Summary.BrokenNetworkPathCount)" 'INFO'
    Write-Log "RPC Client-Side Issue Detected: $($script:Summary.RpcClientSideIssueDetected)" 'INFO'
    Write-Log "RPC Remote/Network Issue Detected: $($script:Summary.RpcRemoteSideIssueDetected)" 'INFO'
    Write-Log "RPC Diagnostics Status: $($script:Summary.RpcDiagnosticsStatus)" 'INFO'
    Write-Log "RPC Classification Hint: Review RPC_ROOT_CAUSE_ASSESSMENT and RPC_EVENTLOGSUMMARY_CLASSIFICATION markers above for local-vs-remote determination." 'INFO'

    if ($script:Summary.RepairsAttempted.Count -gt 0) {
        Write-Log "Repairs Attempted:" 'INFO'
        foreach ($repair in $script:Summary.RepairsAttempted) {
            Write-Log " - $repair" 'INFO'
        }
    }

    if ($script:Summary.Notes.Count -gt 0) {
        Write-Log "Notes:" 'INFO'
        foreach ($note in $script:Summary.Notes) {
            Write-Log " - $note" 'INFO'
        }
    }
}


function Invoke-LogArchiveRetention {
    [CmdletBinding()]
    param(
        [string]$LogDirectory = 'C:\Logs',
        [string]$ComputerName = $env:COMPUTERNAME
    )

    Write-Log "Starting Sunday-based log archive and retention processing in $LogDirectory" 'INFO'

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        Write-Log "Log directory does not exist: $LogDirectory" 'WARN'
        Add-DetailedResult -Step 'LogArchiveRetention' -Status 'Warning' -Message "Log directory not found: $LogDirectory"
        return
    }

    $now = Get-Date
    $thisSunday = $now.Date.AddDays(-[int]$now.DayOfWeek)
    $previousSunday = $thisSunday.AddDays(-7)
    $twoSundaysAgo = $thisSunday.AddDays(-14)

    Write-Log "This Sunday: $thisSunday" 'INFO'
    Write-Log "Previous Sunday: $previousSunday" 'INFO'
    Write-Log "Two Sundays Ago: $twoSundaysAgo" 'INFO'

    $extensions = @('.yaml', '.yml', '.txt')

    $allLooseLogs = Get-ChildItem -LiteralPath $LogDirectory -File -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $extensions -contains $_.Extension.ToLowerInvariant() -and
            $_.FullName -ne $script:YamlLogPath
        }

    $logsToArchive = $allLooseLogs | Where-Object {
        $_.CreationTime -ge $previousSunday -and $_.CreationTime -lt $thisSunday
    } | Sort-Object CreationTime, Name

    $archiveDateText = $previousSunday.ToString('yyyy-MM-dd')
    $zipPath = Join-Path $LogDirectory ("{0}-logs-{1}.zip" -f $ComputerName, $archiveDateText)

    $archiveSummary = [ordered]@{
        ThisSunday                 = $thisSunday
        PreviousSunday             = $previousSunday
        TwoSundaysAgo              = $twoSundaysAgo
        LooseLogsFound             = @($allLooseLogs).Count
        LogsSelectedForArchive     = @($logsToArchive).Count
        ArchiveCreated             = $false
        ArchivePath                = $null
        DeletedOriginalFiles       = @()
        DeletedOldLooseLogs        = @()
        DeletedExpiredZipFiles     = @()
        Errors                     = @()
    }

    if (@($logsToArchive).Count -gt 0) {
        Write-Log "Preparing archive for previous Sunday week: $zipPath" 'INFO'

        try {
            if (Test-Path -LiteralPath $zipPath) {
                Write-Log "Existing archive found for that Sunday. Removing and recreating: $zipPath" 'WARN'
                Remove-Item -LiteralPath $zipPath -Force -ErrorAction Stop
            }

            Compress-Archive -Path ($logsToArchive.FullName) -DestinationPath $zipPath -CompressionLevel Optimal -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $zipPath)) {
                throw 'ZIP file was not created.'
            }

            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
            $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)

            try {
                $zipEntries = @($zip.Entries)
                if ($zipEntries.Count -lt 1) {
                    throw 'ZIP file was created but contains no entries.'
                }

                if ($zipEntries.Count -lt @($logsToArchive).Count) {
                    throw "ZIP file entry count ($($zipEntries.Count)) is less than expected source file count ($(@($logsToArchive).Count))."
                }
            }
            finally {
                $zip.Dispose()
            }

            $archiveSummary.ArchiveCreated = $true
            $archiveSummary.ArchivePath = $zipPath
            Write-Log "Archive created successfully: $zipPath" 'OK'

            foreach ($file in $logsToArchive) {
                try {
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    $archiveSummary.DeletedOriginalFiles += $file.FullName
                    Write-Log "Deleted archived source log: $($file.FullName)" 'OK'
                }
                catch {
                    $msg = "Failed to delete archived source file $($file.FullName): $($_.Exception.Message)"
                    $archiveSummary.Errors += $msg
                    Write-Log $msg 'WARN'
                }
            }
        }
        catch {
            $msg = "Archive creation/validation failed: $($_.Exception.Message)"
            $archiveSummary.Errors += $msg
            Write-Log $msg 'ERROR'
        }
    }
    else {
        Write-Log 'No loose log files were found for the previous Sunday-to-Saturday period.' 'INFO'
    }

    $remainingLooseLogs = Get-ChildItem -LiteralPath $LogDirectory -File -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $extensions -contains $_.Extension.ToLowerInvariant() -and
            $_.FullName -ne $script:YamlLogPath
        }

    $oldLooseLogsToDelete = $remainingLooseLogs | Where-Object {
        $_.CreationTime -lt $twoSundaysAgo
    }

    foreach ($file in $oldLooseLogsToDelete) {
        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $archiveSummary.DeletedOldLooseLogs += $file.FullName
            Write-Log "Deleted loose log older than two Sundays: $($file.FullName)" 'OK'
        }
        catch {
            $msg = "Failed to delete old loose log $($file.FullName): $($_.Exception.Message)"
            $archiveSummary.Errors += $msg
            Write-Log $msg 'WARN'
        }
    }

    $zipFilesToDelete = Get-ChildItem -LiteralPath $LogDirectory -File -Filter '*.zip' -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like "$ComputerName-logs-*.zip" -and
            $_.CreationTime -lt $twoSundaysAgo
        }

    foreach ($zipFile in $zipFilesToDelete) {
        try {
            Remove-Item -LiteralPath $zipFile.FullName -Force -ErrorAction Stop
            $archiveSummary.DeletedExpiredZipFiles += $zipFile.FullName
            Write-Log "Deleted ZIP archive older than two Sundays: $($zipFile.FullName)" 'OK'
        }
        catch {
            $msg = "Failed to delete expired ZIP $($zipFile.FullName): $($_.Exception.Message)"
            $archiveSummary.Errors += $msg
            Write-Log $msg 'WARN'
        }
    }

    Add-DetailedResult -Step 'LogArchiveRetention' -Status 'Info' -Message 'Sunday-based log archive and retention processing completed.' -Data @{
        ThisSunday                  = $archiveSummary.ThisSunday
        PreviousSunday              = $archiveSummary.PreviousSunday
        TwoSundaysAgo               = $archiveSummary.TwoSundaysAgo
        LooseLogsFound              = $archiveSummary.LooseLogsFound
        LogsSelectedForArchive      = $archiveSummary.LogsSelectedForArchive
        ArchiveCreated              = $archiveSummary.ArchiveCreated
        ArchivePath                 = $archiveSummary.ArchivePath
        DeletedOriginalFilesCount   = @($archiveSummary.DeletedOriginalFiles).Count
        DeletedOldLooseLogsCount    = @($archiveSummary.DeletedOldLooseLogs).Count
        DeletedExpiredZipFilesCount = @($archiveSummary.DeletedExpiredZipFiles).Count
        ErrorsCount                 = @($archiveSummary.Errors).Count
        DeletedOriginalFiles        = ($archiveSummary.DeletedOriginalFiles -join '; ')
        DeletedOldLooseLogs         = ($archiveSummary.DeletedOldLooseLogs -join '; ')
        DeletedExpiredZipFiles      = ($archiveSummary.DeletedExpiredZipFiles -join '; ')
        Errors                      = ($archiveSummary.Errors -join '; ')
    }

    if (@($archiveSummary.Errors).Count -gt 0) {
        Warn-Step -Name 'LogArchiveRetention' -Reason ("Completed with errors: " + ($archiveSummary.Errors -join ' | '))
    }
    else {
        Write-Log 'Sunday-based log archive and retention processing completed successfully.' 'OK'
    }
}


function Get-PendingRebootReasons {
    param([AllowNull()]$State)

    $reasons = New-Object System.Collections.Generic.List[string]
    if ($null -eq $State) { return @() }

    $map = [ordered]@{
        CBServicing_RebootPending         = 'ComponentBasedServicing'
        WindowsUpdate_RebootRequired      = 'WindowsUpdate'
        SessionManager_PendingFileRename  = 'PendingFileRenameOperations'
        SessionManager_PendingFileRename2 = 'PendingFileRenameOperations2'
        UpdateExeVolatile                 = 'UpdateExeVolatile'
        PackagesPending                   = 'PackagesPending'
        WUAU_RebootRequired_COM           = 'MicrosoftUpdateSystemInfo'
    }

    foreach ($entry in $map.GetEnumerator()) {
        try {
            if ([bool]$State.($entry.Key)) { $reasons.Add([string]$entry.Value) | Out-Null }
        }
        catch {}
    }
    return @($reasons)
}

function Get-SystemDriveTelemetry {
    try {
        $drive = Get-DiskSpaceInfo -Path ($env:SystemDrive + '\')
        if ($null -eq $drive) { return $null }
        return [PSCustomObject]@{
            FreeGB  = [math]::Round(([double]$drive.FreeSpace / 1GB), 2)
            UsedGB  = [math]::Round(([double]$drive.UsedSpace / 1GB), 2)
            TotalGB = [math]::Round(([double]$drive.TotalSize / 1GB), 2)
            FreePercent = if ([double]$drive.TotalSize -gt 0) { [math]::Round((100 * [double]$drive.FreeSpace / [double]$drive.TotalSize), 2) } else { $null }
        }
    }
    catch { return $null }
}

function Get-WindowsBuildTelemetry {
    try {
        $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        return [PSCustomObject]@{
            ProductName    = [string]$cv.ProductName
            EditionId      = [string]$cv.EditionID
            DisplayVersion = [string]$cv.DisplayVersion
            ReleaseId      = [string]$cv.ReleaseId
            CurrentBuild   = [string]$cv.CurrentBuild
            UBR            = if ($null -ne $cv.UBR) { [int]$cv.UBR } else { $null }
            Build          = if ($null -ne $cv.UBR) { '{0}.{1}' -f $cv.CurrentBuild, $cv.UBR } else { [string]$cv.CurrentBuild }
        }
    }
    catch { return $null }
}

function Get-StepTelemetry {
    param([Parameter(Mandatory)][string]$StepPattern)

    $matches = @($script:DetailedResults | Where-Object { $_.Step -like $StepPattern })
    return @($matches | ForEach-Object {
        [PSCustomObject]@{
            Timestamp = $_.Timestamp
            Step      = $_.Step
            Status    = $_.Status
            Message   = $_.Message
            Data      = $_.Data
        }
    })
}

function New-SystemRepairTelemetryEvent {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [AllowNull()][string]$FailureMessage
    )

    $end = Get-Date
    $domain = $null
    try { $domain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()).Name } catch { $domain = $env:USERDNSDOMAIN }

    $repairs = @($script:Summary.RepairsAttempted | Select-Object -Unique)
    $notes = @($script:Summary.Notes)
    $failedSteps = @(Get-StepTelemetry -StepPattern '*' | Where-Object { $_.Status -eq 'Failed' })
    $warningSteps = @(Get-StepTelemetry -StepPattern '*' | Where-Object { $_.Status -eq 'Warning' })

    $freeBefore = if ($null -ne $script:DiskSpaceBefore) { $script:DiskSpaceBefore.FreeGB } else { $null }
    $freeAfter = if ($null -ne $script:DiskSpaceAfter) { $script:DiskSpaceAfter.FreeGB } else { $null }
    $reclaimed = if ($null -ne $freeBefore -and $null -ne $freeAfter) { [math]::Round(([double]$freeAfter - [double]$freeBefore), 2) } else { $null }

    [PSCustomObject][ordered]@{
        '@timestamp' = $end.ToUniversalTime().ToString('o')
        event = [ordered]@{
            kind     = 'event'
            category = @('host','configuration')
            type     = @('info')
            action   = 'system-repair'
            outcome  = if ($ExitCode -eq 0 -or $ExitCode -eq 3010) { 'success' } else { 'failure' }
            duration = [int64](($end - $script:RunStart).TotalMilliseconds * 1000000)
        }
        maintenance = [ordered]@{
            execution = [ordered]@{
                RunId           = $script:RunId
                ComputerName    = $script:ComputerName
                Domain          = $domain
                ScriptName      = $script:ScriptName
                ScriptVersion   = $script:ScriptVersion
                Status          = $Status
                ExitCode        = $ExitCode
                StartTime       = $script:RunStart.ToUniversalTime().ToString('o')
                EndTime         = $end.ToUniversalTime().ToString('o')
                DurationSeconds = [math]::Round(($end - $script:RunStart).TotalSeconds, 2)
                ErrorCount      = [int]$script:Summary.StepsFailed
                WarningCount    = [int]$script:Summary.Warnings
                FailureMessage  = $FailureMessage
            }
            system_repair = [ordered]@{
                AutoRepairOnDetection = [bool]$AutoRepairOnDetection
                RepairsAttempted      = $repairs
                StepsSucceeded        = [int]$script:Summary.StepsSucceeded
                StepsFailed           = [int]$script:Summary.StepsFailed
                FailedSteps           = $failedSteps
                WarningSteps          = $warningSteps
                PendingRebootBefore   = if ($null -ne $script:PendingRebootBefore) { [bool]$script:PendingRebootBefore.AnyPendingReboot } else { $false }
                PendingRebootReasonsBefore = @($script:Summary.PendingRebootReasonsBefore)
                PendingRebootAfter    = if ($null -ne $script:PendingRebootAfter) { [bool]$script:PendingRebootAfter.AnyPendingReboot } else { [bool]$script:Summary.PendingRebootDetected }
                PendingRebootReasonsAfter = @($script:Summary.PendingRebootReasonsAfter)
                RebootRequired        = [bool]$script:Summary.RebootRequired
                RebootInitiated       = [bool]$script:RebootInitiated
                DiskCorruptionSuspected = [bool]$script:Summary.DiskCorruptionSuspected
                RepairVolumeSpotFixSucceeded = [bool]$script:Summary.RepairVolumeSpotFixSucceeded
                DismCorruptionDetected = [bool]$script:Summary.DismCorruptionDetected
                SfcIntegrityViolations = [bool]$script:Summary.SfcIntegrityViolations
                SfcRepairAttempted     = [bool]$script:Summary.SfcRepairAttempted
                SfcRepairSucceeded     = [bool]$script:Summary.SfcRepairSucceeded
                CbsCorruptFileCount    = [int]$script:Summary.CbsCorruptFileCount
                CbsDriverCorruptionDetected = [bool]$script:Summary.CbsDriverCorruptionDetected
                CbsDriverCorruptionRemaining = [bool]$script:Summary.CbsDriverCorruptionRemaining
                CbsCorruptDriverFiles = @($script:Summary.CbsCorruptDriverFiles | Select-Object -Unique)
                HpDriverRepairAttempted = [bool]$script:Summary.HpDriverRepairAttempted
                HpDriverRepairSucceeded = [bool]$script:Summary.HpDriverRepairSucceeded
                WmiRepositoryInconsistent = [bool]$script:Summary.WmiRepositoryInconsistent
                StorageHealthWarnings = [int]$script:Summary.StorageHealthWarnings
                StorageFailurePredicted = [bool]$script:Summary.StorageFailurePredicted
                StorageWearOver50Detected = [bool]$script:Summary.StorageWearOver50Detected
                ExplorerHangCount = [int]$script:Summary.ExplorerHangCount
                ExplorerCrashCount = [int]$script:Summary.ExplorerCrashCount
                ExplorerInstabilityDetected = [bool]$script:Summary.ExplorerInstabilityDetected
                NonMicrosoftShellExtensionCount = [int]$script:Summary.NonMicrosoftShellExtensionCount
                BrokenNetworkPathCount = [int]$script:Summary.BrokenNetworkPathCount
                RpcClientSideIssueDetected = [bool]$script:Summary.RpcClientSideIssueDetected
                RpcRemoteSideIssueDetected = [bool]$script:Summary.RpcRemoteSideIssueDetected
                RpcDiagnosticsStatus = [string]$script:Summary.RpcDiagnosticsStatus
                DiskSpaceBefore = $script:DiskSpaceBefore
                DiskSpaceAfter  = $script:DiskSpaceAfter
                SpaceReclaimedGB = $reclaimed
                WindowsBuild = Get-WindowsBuildTelemetry
                DismResults = @(Get-StepTelemetry -StepPattern 'DISM*')
                SfcResults = @(Get-StepTelemetry -StepPattern 'SFC*')
                RepairVolumeResults = @(Get-StepTelemetry -StepPattern 'RepairVolume*')
                StorageResults = @(Get-StepTelemetry -StepPattern 'Storage*')
                WmiResults = @(Get-StepTelemetry -StepPattern 'Wmi*')
                WindowsUpdateResetResults = @(Get-StepTelemetry -StepPattern 'WindowsUpdate*')
                ExplorerResults = @(Get-StepTelemetry -StepPattern 'Explorer*')
                RpcResults = @(Get-StepTelemetry -StepPattern 'Rpc*')
                DefenderResults = @(Get-StepTelemetry -StepPattern 'Defender*')
                Notes = $notes
            }
        }
    }
}

function Write-SystemRepairTelemetry {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [AllowNull()][string]$FailureMessage
    )

    if ($script:TelemetryWritten) { return }

    try {
        Ensure-LogDirectory
        $event = New-SystemRepairTelemetryEvent -Status $Status -ExitCode $ExitCode -FailureMessage $FailureMessage
        $compact = $event | ConvertTo-Json -Depth 12 -Compress
        $pretty = $event | ConvertTo-Json -Depth 12
        Write-MaintenanceTelemetryLine -Path $script:TelemetryNdjsonPath -JsonLine $compact
        [System.IO.File]::WriteAllText($script:LatestTelemetryPath, $pretty + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
        $script:TelemetryWritten = $true
        Write-Log "Telemetry written to $($script:TelemetryNdjsonPath) and $($script:LatestTelemetryPath)" 'OK'
    }
    catch {
        Write-Log "Telemetry write failed: $($_.Exception.Message)" 'ERROR'
    }
}

if (-not (Test-IsAdministrator)) {
    $script:Summary.StepsFailed++
    $script:FailureMessage = 'This script must be run as Administrator.'
    $script:FinalStatus = 'Failed'
    $script:FinalExitCode = 1
    try { Ensure-LogDirectory } catch {}
    Write-Error $script:FailureMessage
    Write-SystemRepairTelemetry -Status $script:FinalStatus -ExitCode $script:FinalExitCode -FailureMessage $script:FailureMessage
    exit $script:FinalExitCode
}

try {
    Ensure-LogDirectory
    $script:DiskSpaceBefore = Get-SystemDriveTelemetry
    Write-Log 'Initializing automated system health and repair script...' 'INFO'
    if ($AutoRepairOnDetection) {
        Write-Log 'Detection-first mode is enabled. Approved repairs will run only when corresponding problems are detected.' 'INFO'
    }
    else {
        Write-Log 'Detection-only mode is enabled. Automatic repair actions are disabled unless their independent opt-in switches are enabled.' 'INFO'
    }
    Write-Log "Detailed YAML log will be written to $($script:YamlLogPath)" 'INFO'

    Invoke-Safely -Name 'Action1AgentValidation' -ScriptBlock { Ensure-Action1Agent } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DiskSpaceCheck' -ScriptBlock { Invoke-DiskSpaceCheck } -WarnOnly | Out-Null
    Invoke-Safely -Name 'StorageHealthCheck' -ScriptBlock { Invoke-StorageHealthCheck } -WarnOnly | Out-Null

    Invoke-Safely -Name 'PendingRebootCheck' -ScriptBlock {
        $script:PendingRebootBefore = Get-PendingRebootState
        $script:Summary.PendingRebootDetected = [bool]$script:PendingRebootBefore.AnyPendingReboot
        foreach ($reason in @(Get-PendingRebootReasons -State $script:PendingRebootBefore)) {
            $script:Summary.PendingRebootReasonsBefore.Add($reason) | Out-Null
        }
        $script:PendingRebootBefore | Format-List | Out-String | ForEach-Object {
            $_.TrimEnd() -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
        }
        if ($script:PendingRebootBefore.AnyPendingReboot) { Warn-Step -Name 'PendingRebootCheck' -Reason 'A pending reboot was detected before maintenance began.' }
    } -WarnOnly | Out-Null

    Invoke-Safely -Name 'RepairVolumeScan' -ScriptBlock { Invoke-RepairVolumeScan } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DISMDetection' -ScriptBlock { Invoke-DismDetection } -WarnOnly | Out-Null
    Invoke-Safely -Name 'SFCDetection' -ScriptBlock { Invoke-SfcDetection } -WarnOnly | Out-Null
    Invoke-Safely -Name 'WmiRepositoryCheck' -ScriptBlock { Invoke-WmiCheck } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DnsFlush' -ScriptBlock { Invoke-DnsFlushOnly } -WarnOnly | Out-Null
    Invoke-Safely -Name 'TempCleanup' -ScriptBlock { Invoke-TempCleanup } -WarnOnly | Out-Null
    Invoke-Safely -Name 'ScheduledTaskHealthCheck' -ScriptBlock { Invoke-ScheduledTaskHealthCheck } -WarnOnly | Out-Null
    Invoke-Safely -Name 'RpcDiagnostics' -ScriptBlock { Invoke-RpcDiagnostics } -WarnOnly | Out-Null
    Invoke-Safely -Name 'EventLogSummary' -ScriptBlock { Invoke-EventLogSummary } -WarnOnly | Out-Null
    Invoke-Safely -Name 'ExplorerHealthDiagnostics' -ScriptBlock { Invoke-ExplorerHealthDiagnostics } -WarnOnly | Out-Null
    Invoke-Safely -Name 'DefenderStatusCheck' -ScriptBlock { Invoke-DefenderStatusCheck } -WarnOnly | Out-Null

    if ($AutoRepairOnDetection) {
        if ($script:Summary.DismCorruptionDetected) { Invoke-Safely -Name 'DISMRepair' -ScriptBlock { Invoke-DismRepair } | Out-Null }
        if ($script:Summary.SfcIntegrityViolations) { Invoke-Safely -Name 'SfcCbsRepairWorkflow' -ScriptBlock { Invoke-SystemFileRepairWorkflow } -WarnOnly | Out-Null }
        if ($script:Summary.CbsDriverCorruptionDetected -and $script:Summary.SfcIntegrityViolations) { Invoke-Safely -Name 'Tier3HpiaDriverRemediationFromCbs' -ScriptBlock { Invoke-HpiaDriverOnlyRepairFromCbs } -WarnOnly | Out-Null }
        if ($script:Summary.WmiRepositoryInconsistent -and $AllowWmiRepair) { Invoke-Safely -Name 'WMIRepair' -ScriptBlock { Invoke-WmiRepair } -WarnOnly | Out-Null }

        if ($script:Summary.DiskCorruptionSuspected) {
            $script:RepairVolumeSpotFixResult = $false
            Invoke-Safely -Name 'RepairVolumeSpotFix' -ScriptBlock { $script:RepairVolumeSpotFixResult = Invoke-RepairVolumeSpotFix } -WarnOnly | Out-Null
            if ([bool]$script:RepairVolumeSpotFixResult) {
                $script:RepairVolumePostSpotFixScan = $null
                Invoke-Safely -Name 'RepairVolumePostSpotFixScan' -ScriptBlock { $script:RepairVolumePostSpotFixScan = Invoke-RepairVolumeScan -StepName 'RepairVolumePostSpotFixScan' -UpdateSummary:$true } -WarnOnly | Out-Null
                if ($null -ne $script:RepairVolumePostSpotFixScan -and -not [bool]$script:RepairVolumePostSpotFixScan.CorruptionDetected) { Write-Log 'Repair-Volume post-SpotFix verification scan is clean. Offline disk repair will not be scheduled.' 'OK' }
            }
            if ($script:Summary.DiskCorruptionSuspected -and $AllowOfflineDiskRepair) { Invoke-Safely -Name 'OfflineDiskRepair' -ScriptBlock { Invoke-RepairVolumeOfflineFix } -WarnOnly | Out-Null }
            elseif ($script:Summary.DiskCorruptionSuspected) {
                Write-Log 'REPAIR_VOLUME_OFFLINE_RECOMMENDED|Reason=Repair-Volume scan still indicates possible file system corruption after SpotFix or SpotFix did not complete. Re-run script with -AllowOfflineDiskRepair to allow Repair-Volume -OfflineScanAndFix.' 'WARN'
                Add-DetailedResult -Step 'OfflineDiskRepair' -Status 'Skipped' -Message 'Offline disk repair skipped because AllowOfflineDiskRepair is disabled, but post-repair scan still indicates possible corruption.'
            }
        }
    }

    if ($AllowNetworkReset) { Invoke-Safely -Name 'NetworkReset' -ScriptBlock { Invoke-NetworkReset } -WarnOnly | Out-Null }
    if ($AllowIconCacheRebuild) { Invoke-Safely -Name 'IconCacheRebuild' -ScriptBlock { Invoke-IconCacheRebuild } -WarnOnly | Out-Null }
    if ($AllowCopilotRemoval) { Invoke-Safely -Name 'CopilotDisableAndRemoval' -ScriptBlock { Invoke-CopilotDisableAndRemoval } -WarnOnly | Out-Null }
    if ($AllowFirewallReset) { Invoke-Safely -Name 'FirewallReset' -ScriptBlock { Invoke-FirewallReset } -WarnOnly | Out-Null }
    if ($AllowWindowsUpdateReset) { Invoke-Safely -Name 'WindowsUpdateComponentReset' -ScriptBlock { Invoke-WindowsUpdateComponentReset } -WarnOnly | Out-Null }
    if ($ClearEventLogs) { Invoke-Safely -Name 'EventLogClear' -ScriptBlock { Invoke-EventLogClear } -WarnOnly | Out-Null }

    Invoke-Safely -Name 'LogArchiveRetention' -ScriptBlock { Invoke-LogArchiveRetention -LogDirectory $LogDirectory } -WarnOnly | Out-Null

    $script:DiskSpaceAfter = Get-SystemDriveTelemetry
    Invoke-Safely -Name 'PendingRebootFinalCheck' -ScriptBlock {
        $script:PendingRebootAfter = Get-PendingRebootState
        foreach ($reason in @(Get-PendingRebootReasons -State $script:PendingRebootAfter)) { $script:Summary.PendingRebootReasonsAfter.Add($reason) | Out-Null }
        if ($script:PendingRebootAfter.AnyPendingReboot) { $script:Summary.PendingRebootDetected = $true }
    } -WarnOnly | Out-Null

    Show-Summary
    Write-YamlLog

    if ($script:Summary.StepsFailed -gt 0) {
        $script:FinalStatus = 'Failed'
        $script:FinalExitCode = 2
    }
    elseif ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected) {
        $script:FinalStatus = 'SuccessRebootRequired'
        $script:FinalExitCode = 3010
    }
    elseif ($script:Summary.Warnings -gt 0) {
        $script:FinalStatus = 'SuccessWithWarnings'
        $script:FinalExitCode = 0
    }
    else {
        $script:FinalStatus = 'Success'
        $script:FinalExitCode = 0
    }

    if ($AutoRebootIfNeeded -and ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected)) {
        Invoke-RebootIfNeeded -DelaySeconds $AutoRebootDelaySeconds
        $script:FinalStatus = 'SuccessRebootInitiated'
        $script:FinalExitCode = 3010
        Write-YamlLog
    }
}
catch {
    $script:FailureMessage = $_.Exception.Message
    $script:FinalStatus = 'Failed'
    $script:FinalExitCode = 2
    try { Fail-Step -Name 'UnhandledException' -Reason $script:FailureMessage } catch {}
    try { Show-Summary } catch {}
    try { Write-YamlLog } catch {}
}
finally {
    if ($null -eq $script:DiskSpaceAfter) { $script:DiskSpaceAfter = Get-SystemDriveTelemetry }
    Write-SystemRepairTelemetry -Status $script:FinalStatus -ExitCode $script:FinalExitCode -FailureMessage $script:FailureMessage
}

exit $script:FinalExitCode
