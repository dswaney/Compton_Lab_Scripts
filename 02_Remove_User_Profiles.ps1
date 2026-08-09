# =====================================================================
# ScriptName: 02_Remove_User_Profiles.ps1
# ScriptVersion: 2.1.0
# LastUpdated: 2026-07-27
# =====================================================================

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [string[]]$ExcludedProfiles = @(
        'Default',
        'Default User',
        'Public',
        'All Users',
        'MISAdmin',
        'CC-Student',
        'dswaney'
    ),

    [string]$UsersRoot = 'C:\Users',

    [switch]$SkipLoadedProfiles = $true,

    [switch]$SkipSpecialProfiles = $true,

    [int]$OlderThanDays = 0,

    [string]$LogDirectory = 'C:\Logs',

    [int]$MaxParallelProfileDeletes = 2,

    [int]$ProfileCleanupTimeLimitMinutes = 40,

    [int]$ProfileFolderJobTimeoutMinutes = 15,

    [string]$ProfileCleanupStatePath = 'C:\ProgramData\Compton\ProfileCleanupState.json',

    # Add additional computer name patterns here as needed.
    # Examples:
    #   'SSB-114*' matches any computer name that begins with SSB-114.
    #   'LAB-205-01' matches one exact computer name.
    [string[]]$EdgeInPrivateComputerNamePatterns = @(
        'SSB-114*',
        'SSB-122*'
    ),

    [string]$EdgeInPrivateUrl = 'https://www.compton.edu'
)

$ErrorActionPreference = 'Stop'

$script:RunStart = Get-Date
$script:ComputerName = $env:COMPUTERNAME
$script:TimestampForFile = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
$script:BaseFileName = "{0}-RemoveUserProfiles-{1}" -f $script:ComputerName, $script:TimestampForFile
$script:YamlLogPath = Join-Path $LogDirectory ($script:BaseFileName + '.yaml')
$script:RunId = [guid]::NewGuid().ToString('N')
$script:ScriptName = '02_Remove_User_Profiles.ps1'
$script:ScriptVersion = '2.1.0'
$script:Domain = $env:USERDNSDOMAIN
$script:TelemetryNdjsonPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestTelemetryPath = Join-Path $LogDirectory '02_Remove_User_Profiles.latest.json'
$script:WarningCount = 0
$script:ErrorCount = 0
$script:Finalized = $false
$script:EdgeInPrivateMatched = $false
$script:EdgeInPrivateConfigured = $false
$script:DiskBefore = $null
$script:DiskAfter = $null

$script:Summary = [ordered]@{
    ComputerName       = $script:ComputerName
    StartTime          = $script:RunStart
    EndTime            = $null
    FoundProfiles      = 0
    ExcludedProfiles   = 0
    SkippedLoaded      = 0
    SkippedSpecial     = 0
    SkippedByAge       = 0
    QueuedProfiles     = 0
    DeletedProfiles    = 0
    OneDriveTasksRemoved = 0
    FailedProfiles     = 0
    TimedOutProfiles   = 0
    DeferredProfiles   = 0
    DiskFreeGBBefore   = $null
    DiskFreeGBAfter    = $null
    SpaceReclaimedGB   = $null
}

$script:DeletedProfileDetails = New-Object System.Collections.ArrayList
$script:SkippedProfileDetails = New-Object System.Collections.ArrayList
$script:FailedProfileDetails  = New-Object System.Collections.ArrayList
$script:DeferredProfileDetails = New-Object System.Collections.ArrayList


# =====================================================================
# Windows 11 UI preference configuration
# =====================================================================
# Maintain the standard lab UI settings in this one table. To change a
# setting, edit its Value. To temporarily omit a setting, set Enabled to
# $false. The application functions below process this table for current,
# loaded, offline, and future user profiles.
$script:Windows11UIPreferences = @(
    # Taskbar and Start menu preferences.
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'TaskbarAl';          Type = 'REG_DWORD'; Value = 0; Description = 'taskbar/start alignment left' },
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'TaskbarDa';          Type = 'REG_DWORD'; Value = 0; Description = 'hide Widgets button' },
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'ShowTaskViewButton'; Type = 'REG_DWORD'; Value = 0; Description = 'hide Task View button' },
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'ShowCopilotButton';  Type = 'REG_DWORD'; Value = 0; Description = 'hide Copilot button' },

    # Display the full taskbar search box instead of only the search icon.
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Search'; Name = 'SearchboxTaskbarMode'; Type = 'REG_DWORD'; Value = 2; Description = 'show the full taskbar search box' },

    # Restore the Windows 10-style full right-click context menu in Windows 11.
    # This setting uses the unnamed/default registry value under InprocServer32.
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32'; Name = $null; Type = 'REG_SZ'; Value = ''; Description = 'enable the classic Windows 10-style context menu' },

    # Search and Copilot preferences.
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\SearchSettings'; Name = 'IsDynamicSearchBoxEnabled'; Type = 'REG_DWORD'; Value = 0; Description = 'disable Search Highlights' },
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Policies\Microsoft\Windows\WindowsCopilot'; Name = 'TurnOffWindowsCopilot'; Type = 'REG_DWORD'; Value = 1; Description = 'disable Windows Copilot for this user' },

    # File Explorer preferences.
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'LaunchTo';    Type = 'REG_DWORD'; Value = 1; Description = 'open File Explorer to This PC' },
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'HideFileExt'; Type = 'REG_DWORD'; Value = 0; Description = 'show file extensions' },
    [pscustomobject]@{ Enabled = $true; SubKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'Hidden';      Type = 'REG_DWORD'; Value = 1; Description = 'show hidden files' }
)

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Ensure-LogDirectory {
    Ensure-Directory -Path $LogDirectory
}


function Get-SystemDriveSpace {
    try {
        $driveName = ([System.IO.Path]::GetPathRoot($UsersRoot)).TrimEnd('\\').TrimEnd(':')
        if ([string]::IsNullOrWhiteSpace($driveName)) { $driveName = $env:SystemDrive.TrimEnd(':') }
        $drive = Get-PSDrive -Name $driveName -ErrorAction Stop
        return [PSCustomObject]@{
            Drive          = "$driveName`:"
            FreeBytes      = [int64]$drive.Free
            UsedBytes      = [int64]$drive.Used
            CapacityBytes  = [int64]($drive.Free + $drive.Used)
            FreeGB         = [math]::Round($drive.Free / 1GB, 2)
            FreePercent    = if (($drive.Free + $drive.Used) -gt 0) { [math]::Round(($drive.Free / ($drive.Free + $drive.Used)) * 100, 2) } else { $null }
        }
    }
    catch {
        Write-Log "Unable to collect system-drive space: $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Write-ExecutionTelemetry {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode
    )

    try {
        Ensure-LogDirectory
        $endTime = if ($script:Summary.EndTime) { [datetime]$script:Summary.EndTime } else { Get-Date }
        $durationSeconds = [math]::Round(($endTime - $script:RunStart).TotalSeconds, 2)

        $event = [ordered]@{
            EventType                  = 'maintenance.execution'
            ComputerName               = $script:ComputerName
            Domain                     = $script:Domain
            ScriptName                 = $script:ScriptName
            ScriptVersion              = $script:ScriptVersion
            RunId                      = $script:RunId
            Status                     = $Status
            ExitCode                   = $ExitCode
            StartTime                  = $script:RunStart.ToString('o')
            EndTime                    = $endTime.ToString('o')
            DurationSeconds            = $durationSeconds
            WarningCount               = $script:WarningCount
            ErrorCount                 = $script:ErrorCount
            Timestamp                  = (Get-Date).ToUniversalTime().ToString('o')
            UsersRoot                  = $UsersRoot
            OlderThanDays              = $OlderThanDays
            FoundProfiles              = $script:Summary.FoundProfiles
            ExcludedProfiles           = $script:Summary.ExcludedProfiles
            SkippedLoaded              = $script:Summary.SkippedLoaded
            SkippedSpecial             = $script:Summary.SkippedSpecial
            SkippedByAge               = $script:Summary.SkippedByAge
            QueuedProfiles             = $script:Summary.QueuedProfiles
            DeletedProfiles            = $script:Summary.DeletedProfiles
            FailedProfiles             = $script:Summary.FailedProfiles
            TimedOutProfiles           = $script:Summary.TimedOutProfiles
            DeferredProfiles           = $script:Summary.DeferredProfiles
            OneDriveTasksRemoved       = $script:Summary.OneDriveTasksRemoved
            DiskFreeGBBefore           = $script:Summary.DiskFreeGBBefore
            DiskFreeGBAfter            = $script:Summary.DiskFreeGBAfter
            SpaceReclaimedGB           = $script:Summary.SpaceReclaimedGB
            EdgeInPrivateMatched       = $script:EdgeInPrivateMatched
            EdgeInPrivateConfigured    = $script:EdgeInPrivateConfigured
            DeletedProfileNames        = @($script:DeletedProfileDetails | ForEach-Object { $_.ProfileName })
            FailedProfileNames         = @($script:FailedProfileDetails | ForEach-Object { $_.ProfileName })
            DeferredProfileNames       = @($script:DeferredProfileDetails | ForEach-Object { $_.ProfileName })
            SkippedProfileCountsByReason = @($script:SkippedProfileDetails | Group-Object Reason | ForEach-Object {
                [ordered]@{ Reason = $_.Name; Count = $_.Count }
            })
        }

        $compact = $event | ConvertTo-Json -Depth 8 -Compress
        Write-MaintenanceTelemetryLine -Path $script:TelemetryNdjsonPath -JsonLine $compact
        $event | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:LatestTelemetryPath -Encoding UTF8 -Force
    }
    catch {
        Write-Warning "Failed to write execution telemetry: $($_.Exception.Message)"
    }
}

function Complete-ScriptExecution {
    param([Parameter(Mandatory)][int]$ExitCode)

    if ($script:Finalized) { exit $ExitCode }
    $script:Finalized = $true
    if (-not $script:Summary.EndTime) { $script:Summary.EndTime = Get-Date }

    $script:DiskAfter = Get-SystemDriveSpace
    if ($script:DiskAfter) {
        $script:Summary.DiskFreeGBAfter = $script:DiskAfter.FreeGB
        if ($script:DiskBefore) {
            $script:Summary.SpaceReclaimedGB = [math]::Round(($script:DiskAfter.FreeBytes - $script:DiskBefore.FreeBytes) / 1GB, 2)
        }
    }

    $status = if ($ExitCode -ne 0 -or $script:Summary.FailedProfiles -gt 0 -or $script:ErrorCount -gt 0) {
        'Failed'
    }
    elseif ($script:Summary.DeferredProfiles -gt 0 -or $script:Summary.TimedOutProfiles -gt 0 -or $script:WarningCount -gt 0) {
        'SuccessWithWarnings'
    }
    else {
        'Success'
    }

    Write-YamlLog
    Write-ExecutionTelemetry -Status $status -ExitCode $ExitCode
    exit $ExitCode
}

function ConvertTo-YamlScalar {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
    if ($Value -is [datetime]) { return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'" }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [string]$Value }

    $text = [string]$Value
    $text = $text -replace "'", "''"
    return "'" + $text + "'"
}

function Write-YamlList {
    param(
        [Parameter(Mandatory)]$Lines,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Items = @(),
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$Properties
    )

    $safeItems = @()
    if ($null -ne $Items) {
        $safeItems = @($Items)
    }

    $Lines.Add("${Name}:") | Out-Null
    if ($safeItems.Count -eq 0) {
        $Lines.Add('  []') | Out-Null
        return
    }

    foreach ($entry in $safeItems) {
        $first = $true
        foreach ($property in $Properties) {
            $value = $entry.$property
            if ($first) {
                $Lines.Add("  - $($property.ToLowerInvariant()): $(ConvertTo-YamlScalar $value)") | Out-Null
                $first = $false
            }
            else {
                $Lines.Add("    $($property.ToLowerInvariant()): $(ConvertTo-YamlScalar $value)") | Out-Null
            }
        }
    }
}

function Write-YamlLog {
    try {
        Ensure-LogDirectory

        $lines = New-Object System.Collections.ArrayList

        $lines.Add('run:') | Out-Null
        $lines.Add("  computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("  start_time: $(ConvertTo-YamlScalar $script:Summary.StartTime)") | Out-Null
        $lines.Add("  end_time: $(ConvertTo-YamlScalar $script:Summary.EndTime)") | Out-Null
        $lines.Add("  yaml_log_path: $(ConvertTo-YamlScalar $script:YamlLogPath)") | Out-Null
        $lines.Add("  run_id: $(ConvertTo-YamlScalar $script:RunId)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('settings:') | Out-Null
        $lines.Add("  users_root: $(ConvertTo-YamlScalar $UsersRoot)") | Out-Null
        $lines.Add("  skip_loaded_profiles: $(ConvertTo-YamlScalar $SkipLoadedProfiles)") | Out-Null
        $lines.Add("  skip_special_profiles: $(ConvertTo-YamlScalar $SkipSpecialProfiles)") | Out-Null
        $lines.Add("  older_than_days: $(ConvertTo-YamlScalar $OlderThanDays)") | Out-Null
        $lines.Add("  max_parallel_profile_deletes: $(ConvertTo-YamlScalar $MaxParallelProfileDeletes)") | Out-Null
        $lines.Add("  profile_cleanup_time_limit_minutes: $(ConvertTo-YamlScalar $ProfileCleanupTimeLimitMinutes)") | Out-Null
        $lines.Add("  profile_folder_job_timeout_minutes: $(ConvertTo-YamlScalar $ProfileFolderJobTimeoutMinutes)") | Out-Null
        $lines.Add("  profile_cleanup_state_path: $(ConvertTo-YamlScalar $ProfileCleanupStatePath)") | Out-Null
        $lines.Add("  edge_inprivate_url: $(ConvertTo-YamlScalar $EdgeInPrivateUrl)") | Out-Null
        $lines.Add('  edge_inprivate_computer_name_patterns:') | Out-Null
        if ($EdgeInPrivateComputerNamePatterns.Count -gt 0) {
            foreach ($pattern in $EdgeInPrivateComputerNamePatterns) { $lines.Add("    - $(ConvertTo-YamlScalar $pattern)") | Out-Null }
        }
        else { $lines.Add('    []') | Out-Null }
        $lines.Add('  excluded_profiles:') | Out-Null
        if ($ExcludedProfiles.Count -gt 0) {
            foreach ($name in $ExcludedProfiles) { $lines.Add("    - $(ConvertTo-YamlScalar $name)") | Out-Null }
        }
        else { $lines.Add('    []') | Out-Null }
        $lines.Add('') | Out-Null

        $lines.Add('summary:') | Out-Null
        foreach ($key in $script:Summary.Keys) {
            $lines.Add("  $($key.ToString().ToLowerInvariant()): $(ConvertTo-YamlScalar ($script:Summary[$key]))") | Out-Null
        }
        $lines.Add('') | Out-Null

        $deletedYamlItems  = @($script:DeletedProfileDetails.ToArray())
        $skippedYamlItems  = @($script:SkippedProfileDetails.ToArray())
        $deferredYamlItems = @($script:DeferredProfileDetails.ToArray())
        $failedYamlItems   = @($script:FailedProfileDetails.ToArray())

        Write-YamlList -Lines $lines -Name 'deleted_profiles' -Items $deletedYamlItems -Properties @('ProfileName','LocalPath','SID','Loaded','Special','CreatedTime','LastUseTime','DaysOnSystem','DeleteSeconds','Message')
        $lines.Add('') | Out-Null
        Write-YamlList -Lines $lines -Name 'skipped_profiles' -Items $skippedYamlItems -Properties @('ProfileName','LocalPath','SID','Reason','CreatedTime','LastUseTime','DaysOnSystem')
        $lines.Add('') | Out-Null
        Write-YamlList -Lines $lines -Name 'deferred_profiles' -Items $deferredYamlItems -Properties @('ProfileName','LocalPath','SID','Reason','CreatedTime','LastUseTime','DaysOnSystem')
        $lines.Add('') | Out-Null
        Write-YamlList -Lines $lines -Name 'failed_profiles' -Items $failedYamlItems -Properties @('ProfileName','LocalPath','SID','Error','CreatedTime','LastUseTime','DaysOnSystem')

        Set-Content -Path $script:YamlLogPath -Value @($lines.ToArray()) -Encoding UTF8 -Force
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

    if ($Level -eq 'WARN') { $script:WarningCount++ }
    elseif ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }
}

function Test-ComputerNamePatternMatch {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) {
            continue
        }

        if ($ComputerName -like $pattern) {
            return $true
        }
    }

    return $false
}

function Set-EdgeInPrivateAutoLaunch {
    param(
        [Parameter(Mandatory)][string]$Url
    )

    $edgePaths = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path ${env:ProgramFiles} 'Microsoft\Edge\Application\msedge.exe')
    )

    $edgePath = $edgePaths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

    if (-not $edgePath) {
        Write-Log 'Microsoft Edge was not found. Edge InPrivate auto-launch was not configured.' 'ERROR'
        return $false
    }

    $runKey = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'
    $runValueName = 'LaunchComptonEdge'
    $command = '"{0}" --inprivate --new-window --start-maximized "{1}"' -f $edgePath, $Url

    try {
        New-Item -Path $runKey -Force | Out-Null

        New-ItemProperty `
            -Path $runKey `
            -Name $runValueName `
            -Value $command `
            -PropertyType String `
            -Force | Out-Null

        Write-Log "Configured Edge InPrivate auto-launch for every user at logon: $Url" 'OK'
        Write-Log "Run key value: $runValueName = $command" 'INFO'
        return $true
    }
    catch {
        Write-Log "Failed to configure Edge InPrivate auto-launch: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { return $false }
}

function Get-ProfileFolderName {
    param([Parameter(Mandatory)][string]$Path)
    try { return (Split-Path -Path $Path -Leaf) } catch { return $null }
}

function Get-ProfileAgeData {
    param(
        [Parameter(Mandatory)][string]$ProfilePath,
        [AllowNull()]$LastUseTime = $null
    )

    $createdTime = $null
    $daysOnSystem = $null
    $normalizedLastUseTime = $null

    try {
        if ($null -ne $LastUseTime -and -not [string]::IsNullOrWhiteSpace([string]$LastUseTime)) {
            if ($LastUseTime -is [datetime]) { $normalizedLastUseTime = $LastUseTime }
            else { $normalizedLastUseTime = [Management.ManagementDateTimeConverter]::ToDateTime([string]$LastUseTime) }
        }
    }
    catch { $normalizedLastUseTime = $null }

    try {
        if (Test-Path -LiteralPath $ProfilePath) {
            $item = Get-Item -LiteralPath $ProfilePath -ErrorAction Stop
            $createdTime = $item.CreationTime
            $daysOnSystem = [math]::Round(((Get-Date) - $createdTime).TotalDays, 2)
        }
    }
    catch { }

    return [PSCustomObject]@{
        CreatedTime  = $createdTime
        LastUseTime  = $normalizedLastUseTime
        DaysOnSystem = $daysOnSystem
    }
}

function New-ProfileCleanupState {
    return [PSCustomObject]@{
        Version       = 2
        ComputerName  = $script:ComputerName
        LastUpdated   = (Get-Date).ToString('o')
        Pending       = @()
        Completed     = @()
        Failed        = @()
        TimedOut      = @()
    }
}

function Read-ProfileCleanupState {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return New-ProfileCleanupState }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return New-ProfileCleanupState }
        $state = $raw | ConvertFrom-Json -ErrorAction Stop
        foreach ($property in @('Pending','Completed','Failed','TimedOut')) {
            if (-not $state.PSObject.Properties.Name.Contains($property) -or $null -eq $state.$property) {
                $state | Add-Member -MemberType NoteProperty -Name $property -Value @() -Force
            }
        }
        return $state
    }
    catch {
        Write-Log "Could not read profile cleanup state. Creating a new state file. Error: $($_.Exception.Message)" 'WARN'
        return New-ProfileCleanupState
    }
}

function Save-ProfileCleanupState {
    param(
        [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $parent = Split-Path -Path $Path -Parent
        if (-not [string]::IsNullOrWhiteSpace($parent)) { Ensure-Directory -Path $parent }
        $State.LastUpdated = (Get-Date).ToString('o')
        $State | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Log "Could not write profile cleanup state file: $($_.Exception.Message)" 'WARN'
    }
}

function Add-StateItemUnique {
    param(
        [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][string]$ListName,
        [Parameter(Mandatory)][string]$Value
    )

    $existing = @($State.$ListName)
    if ($existing -notcontains $Value) {
        $State.$ListName = @($existing + $Value)
    }
}

function Remove-StateItem {
    param(
        [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][string]$ListName,
        [Parameter(Mandatory)][string]$Value
    )

    $State.$ListName = @(@($State.$ListName) | Where-Object { $_ -ne $Value })
}

function Stop-ProfileCleanupJobs {
    param([string]$NamePrefix = 'ProfileDelete_')

    try {
        $jobs = @(Get-Job -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "$NamePrefix*" })
        foreach ($job in $jobs) {
            try {
                if ($job.State -eq 'Running') {
                    Stop-Job -Job $job -Force -ErrorAction SilentlyContinue
                }
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
            catch { }
        }
        if ($jobs.Count -gt 0) {
            Write-Log "Cleaned up $($jobs.Count) existing profile deletion job(s) before starting." 'WARN'
        }
    }
    catch {
        Write-Log "Failed while cleaning existing profile deletion jobs: $($_.Exception.Message)" 'WARN'
    }
}

function Get-RunningJobCount {
    param([Parameter(Mandatory)]$Jobs)
    return @($Jobs | Where-Object { $_.Job.State -eq 'Running' }).Count
}

function Remove-OneDriveTasksForProfile {
    param(
        [Parameter(Mandatory)][string]$Sid,
        [Parameter(Mandatory)][string]$ProfileName
    )

    if ([string]::IsNullOrWhiteSpace($Sid)) {
        Write-Log "Cannot check OneDrive reporting tasks because the SID is blank for profile: $ProfileName" 'WARN'
        return 0
    }

    try {
        Import-Module ScheduledTasks -ErrorAction Stop

        $escapedSid = [regex]::Escape($Sid)

        # OneDrive normally creates tasks named similar to:
        # OneDrive Reporting Task-S-1-5-21-...
        # OneDrive Startup Task-S-1-5-21-...
        # Match only those OneDrive task types that contain this exact profile SID.
        $matchingTasks = @(
            Get-ScheduledTask -ErrorAction SilentlyContinue |
            Where-Object {
                (
                    $_.TaskName -like 'OneDrive Reporting Task-*' -or
                    $_.TaskName -like 'OneDrive Startup Task-*'
                ) -and
                $_.TaskName -match $escapedSid
            }
        )

        if ($matchingTasks.Count -eq 0) {
            Write-Log "No matching OneDrive reporting or startup task was found for $ProfileName ($Sid)." 'OK'
            return 0
        }

        $removedCount = 0

        foreach ($task in $matchingTasks) {
            try {
                Write-Log "Removing OneDrive task for deleted profile: $($task.TaskPath)$($task.TaskName)" 'INFO'

                Unregister-ScheduledTask `
                    -TaskName $task.TaskName `
                    -TaskPath $task.TaskPath `
                    -Confirm:$false `
                    -ErrorAction Stop

                $removedCount++
                Write-Log "Removed OneDrive task: $($task.TaskPath)$($task.TaskName)" 'OK'
            }
            catch {
                Write-Log "Could not remove OneDrive task '$($task.TaskPath)$($task.TaskName)': $($_.Exception.Message)" 'WARN'
            }
        }

        return $removedCount
    }
    catch {
        Write-Log "OneDrive scheduled-task cleanup failed for $ProfileName ($Sid): $($_.Exception.Message)" 'WARN'
        return 0
    }
}

function Remove-ProfileRegistration {
    param(
        [Parameter(Mandatory)][string]$Sid,
        [Parameter(Mandatory)][string]$ProfilePath,
        [Parameter(Mandatory)][string]$ProfileName
    )

    try {
        $escapedSid = $Sid.Replace("'", "''")
        $profileRegistration = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$escapedSid'" -ErrorAction SilentlyContinue
        if ($null -eq $profileRegistration) {
            Write-Log "Profile registration already absent for SID ${Sid}: $ProfileName" 'OK'
            return
        }

        Remove-CimInstance -InputObject $profileRegistration -Confirm:$false -ErrorAction Stop
        Write-Log "Removed profile registration for SID ${Sid}: $ProfileName" 'OK'
    }
    catch {
        if (-not (Test-Path -LiteralPath $ProfilePath)) {
            Write-Log "Profile folder is gone, but profile registration cleanup returned: $($_.Exception.Message)" 'WARN'
        }
        else {
            throw
        }
    }
}

$profileDeleteJobScript = {
    param(
        [Parameter(Mandatory)][string]$TargetPath,
        [Parameter(Mandatory)][string]$EmptyMirrorRoot,
        [int]$RobocopyTimeoutSeconds = 600,
        [int]$RmdirTimeoutSeconds = 120
    )

    $started = Get-Date
    $result = [ordered]@{
        Path            = $TargetPath
        Success         = $false
        TimedOut        = $false
        RobocopyExitCode = $null
        Message         = ''
        Started         = $started.ToString('o')
        Ended           = $null
        DurationSeconds = $null
    }

    function Invoke-ProcessTimeoutLocal {
        param(
            [Parameter(Mandatory)][string]$FilePath,
            [Parameter(Mandatory)][string]$Arguments,
            [int]$TimeoutSeconds = 600
        )

        $proc = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden
        try {
            Wait-Process -Id $proc.Id -Timeout $TimeoutSeconds -ErrorAction Stop
            $proc.Refresh()
            return [PSCustomObject]@{ ExitCode = $proc.ExitCode; TimedOut = $false }
        }
        catch {
            try { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } catch { }
            return [PSCustomObject]@{ ExitCode = $null; TimedOut = $true }
        }
    }

    try {
        if ([string]::IsNullOrWhiteSpace($TargetPath)) { throw 'TargetPath is blank.' }

        if (-not (Test-Path -LiteralPath $TargetPath)) {
            $result.Success = $true
            $result.Message = 'Target path already removed before job started.'
            return [PSCustomObject]$result
        }

        if (-not (Test-Path -LiteralPath $EmptyMirrorRoot)) {
            New-Item -Path $EmptyMirrorRoot -ItemType Directory -Force | Out-Null
        }

        $emptyMirrorPath = Join-Path $EmptyMirrorRoot ('EmptyProfileMirror_' + ([guid]::NewGuid().ToString('N')))
        New-Item -Path $emptyMirrorPath -ItemType Directory -Force | Out-Null

        try {
            # Best-effort ownership/ACL reset helps with stubborn abandoned profiles.
            $takeownArgs = '/f "{0}" /r /d y' -f $TargetPath
            [void](Invoke-ProcessTimeoutLocal -FilePath 'takeown.exe' -Arguments $takeownArgs -TimeoutSeconds 120)

            $icaclsArgs = '"{0}" /grant Administrators:F /t /c /q' -f $TargetPath
            [void](Invoke-ProcessTimeoutLocal -FilePath 'icacls.exe' -Arguments $icaclsArgs -TimeoutSeconds 180)

            $robocopyArgs = '"{0}" "{1}" /MIR /R:0 /W:0 /XJ /COPY:DAT /DCOPY:DAT /NP /NFL /NDL /NJH /NJS /NC /NS' -f $emptyMirrorPath, $TargetPath
            $robo = Invoke-ProcessTimeoutLocal -FilePath 'robocopy.exe' -Arguments $robocopyArgs -TimeoutSeconds $RobocopyTimeoutSeconds
            $result.RobocopyExitCode = $robo.ExitCode

            if ($robo.TimedOut) {
                $result.TimedOut = $true
                $result.Message = "Robocopy timed out after $RobocopyTimeoutSeconds seconds."
                return [PSCustomObject]$result
            }

            if ($robo.ExitCode -gt 7) {
                $result.Message = "Robocopy failed with exit code $($robo.ExitCode)."
                return [PSCustomObject]$result
            }

            if (Test-Path -LiteralPath $TargetPath) {
                $rdArgs = '/c rd /s /q "{0}"' -f $TargetPath
                $rd = Invoke-ProcessTimeoutLocal -FilePath 'cmd.exe' -Arguments $rdArgs -TimeoutSeconds $RmdirTimeoutSeconds
                if ($rd.TimedOut) {
                    $result.TimedOut = $true
                    $result.Message = "cmd rd timed out after $RmdirTimeoutSeconds seconds."
                    return [PSCustomObject]$result
                }
            }

            if (Test-Path -LiteralPath $TargetPath) {
                $result.Message = 'Target path still exists after robocopy mirror and rd fallback.'
                return [PSCustomObject]$result
            }

            $result.Success = $true
            $result.Message = "Profile folder removed successfully. Robocopy exit code: $($result.RobocopyExitCode)."
            return [PSCustomObject]$result
        }
        finally {
            if ($emptyMirrorPath -and (Test-Path -LiteralPath $emptyMirrorPath)) {
                Remove-Item -LiteralPath $emptyMirrorPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        $result.Message = $_.Exception.Message
        return [PSCustomObject]$result
    }
    finally {
        $ended = Get-Date
        $result.Ended = $ended.ToString('o')
        $result.DurationSeconds = [math]::Round(($ended - $started).TotalSeconds, 2)
    }
}

function Invoke-ParallelProfileDeletion {
    param(
        [Parameter(Mandatory)][object[]]$ProfilesToDelete
    )

    if ($MaxParallelProfileDeletes -lt 1) { $MaxParallelProfileDeletes = 1 }
    if ($MaxParallelProfileDeletes -gt 4) {
        Write-Log "MaxParallelProfileDeletes was set to $MaxParallelProfileDeletes. Capping at 4 to avoid disk thrashing." 'WARN'
        $MaxParallelProfileDeletes = 4
    }

    if ($ProfileCleanupTimeLimitMinutes -lt 1) { $ProfileCleanupTimeLimitMinutes = 40 }
    if ($ProfileFolderJobTimeoutMinutes -lt 1) { $ProfileFolderJobTimeoutMinutes = 15 }

    $deadline = (Get-Date).AddMinutes($ProfileCleanupTimeLimitMinutes)
    $emptyMirrorRoot = 'C:\ProgramData\Compton\ProfileDeleteEmptyMirrors'
    Ensure-Directory -Path $emptyMirrorRoot

    $state = Read-ProfileCleanupState -Path $ProfileCleanupStatePath
    Save-ProfileCleanupState -State $state -Path $ProfileCleanupStatePath

    Stop-ProfileCleanupJobs

    $queue = New-Object System.Collections.Queue
    foreach ($item in $ProfilesToDelete) {
        Add-StateItemUnique -State $state -ListName 'Pending' -Value $item.LocalPath
        $queue.Enqueue($item)
    }
    Save-ProfileCleanupState -State $state -Path $ProfileCleanupStatePath

    $running = New-Object System.Collections.ArrayList
    $robocopyTimeoutSeconds = [math]::Max(60, ($ProfileFolderJobTimeoutMinutes * 60) - 60)
    $rmdirTimeoutSeconds = 60

    Write-Log "Starting capped parallel profile deletion. Max parallel: $MaxParallelProfileDeletes. Cleanup budget: $ProfileCleanupTimeLimitMinutes minute(s). Per-folder timeout: $ProfileFolderJobTimeoutMinutes minute(s)." 'INFO'

    function Receive-FinishedProfileJob {
        param(
            [Parameter(Mandatory)]$JobInfo,
            [switch]$TimedOutByController
        )

        $job = $JobInfo.Job
        $item = $JobInfo.Profile
        $jobResult = $null

        try {
            if ($TimedOutByController) {
                Stop-Job -Job $job -Force -ErrorAction SilentlyContinue
                $jobResult = [PSCustomObject]@{
                    Path = $item.LocalPath
                    Success = $false
                    TimedOut = $true
                    RobocopyExitCode = $null
                    Message = "Controller stopped job after $ProfileFolderJobTimeoutMinutes minute timeout."
                    DurationSeconds = $null
                }
            }
            else {
                $received = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
                $jobResult = $received | Select-Object -Last 1
                if ($null -eq $jobResult) {
                    $jobResult = [PSCustomObject]@{
                        Path = $item.LocalPath
                        Success = $false
                        TimedOut = $false
                        RobocopyExitCode = $null
                        Message = 'Delete job returned no result.'
                        DurationSeconds = $null
                    }
                }
            }
        }
        finally {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }

        Remove-StateItem -State $state -ListName 'Pending' -Value $item.LocalPath

        if ($jobResult.Success -eq $true) {
            try {
                Remove-ProfileRegistration -Sid $item.SID -ProfilePath $item.LocalPath -ProfileName $item.ProfileName
                $script:Summary.DeletedProfiles++
                Add-StateItemUnique -State $state -ListName 'Completed' -Value $item.LocalPath
                Write-Log "Successfully deleted profile: $($item.ProfileName) ($($item.LocalPath))" 'OK'

                $removedOneDriveTasks = Remove-OneDriveTasksForProfile `
                    -Sid $item.SID `
                    -ProfileName $item.ProfileName
                $script:Summary.OneDriveTasksRemoved += $removedOneDriveTasks

                $script:DeletedProfileDetails.Add([PSCustomObject]@{
                    ProfileName   = $item.ProfileName
                    LocalPath     = $item.LocalPath
                    SID           = $item.SID
                    Loaded        = $item.Loaded
                    Special       = $item.Special
                    CreatedTime   = $item.CreatedTime
                    LastUseTime   = $item.LastUseTime
                    DaysOnSystem  = $item.DaysOnSystem
                    DeleteSeconds = $jobResult.DurationSeconds
                    Message       = $jobResult.Message
                }) | Out-Null
            }
            catch {
                $script:Summary.FailedProfiles++
                Add-StateItemUnique -State $state -ListName 'Failed' -Value $item.LocalPath
                Write-Log "Profile folder was deleted but registration cleanup failed for $($item.ProfileName): $($_.Exception.Message)" 'ERROR'
                $script:FailedProfileDetails.Add([PSCustomObject]@{
                    ProfileName  = $item.ProfileName
                    LocalPath    = $item.LocalPath
                    SID          = $item.SID
                    Error        = $_.Exception.Message
                    CreatedTime  = $item.CreatedTime
                    LastUseTime  = $item.LastUseTime
                    DaysOnSystem = $item.DaysOnSystem
                }) | Out-Null
            }
        }
        elseif ($jobResult.TimedOut -eq $true) {
            $script:Summary.TimedOutProfiles++
            $script:Summary.FailedProfiles++
            Add-StateItemUnique -State $state -ListName 'TimedOut' -Value $item.LocalPath
            Add-StateItemUnique -State $state -ListName 'Failed' -Value $item.LocalPath
            Write-Log "Profile deletion timed out and will be retried on a future run: $($item.ProfileName) ($($item.LocalPath)) | $($jobResult.Message)" 'WARN'
            $script:FailedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $item.ProfileName
                LocalPath    = $item.LocalPath
                SID          = $item.SID
                Error        = $jobResult.Message
                CreatedTime  = $item.CreatedTime
                LastUseTime  = $item.LastUseTime
                DaysOnSystem = $item.DaysOnSystem
            }) | Out-Null
        }
        else {
            $script:Summary.FailedProfiles++
            Add-StateItemUnique -State $state -ListName 'Failed' -Value $item.LocalPath
            Write-Log "Failed to delete profile folder: $($item.ProfileName) ($($item.LocalPath)) | $($jobResult.Message)" 'ERROR'
            $script:FailedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $item.ProfileName
                LocalPath    = $item.LocalPath
                SID          = $item.SID
                Error        = $jobResult.Message
                CreatedTime  = $item.CreatedTime
                LastUseTime  = $item.LastUseTime
                DaysOnSystem = $item.DaysOnSystem
            }) | Out-Null
        }

        Save-ProfileCleanupState -State $state -Path $ProfileCleanupStatePath
        Write-YamlLog
    }

    while ($queue.Count -gt 0 -or $running.Count -gt 0) {
        $now = Get-Date

        for ($i = $running.Count - 1; $i -ge 0; $i--) {
            $jobInfo = $running[$i]
            $jobAgeMinutes = ((Get-Date) - $jobInfo.Started).TotalMinutes

            if ($jobInfo.Job.State -ne 'Running') {
                Receive-FinishedProfileJob -JobInfo $jobInfo
                $running.RemoveAt($i)
            }
            elseif ($jobAgeMinutes -ge $ProfileFolderJobTimeoutMinutes) {
                Receive-FinishedProfileJob -JobInfo $jobInfo -TimedOutByController
                $running.RemoveAt($i)
            }
        }

        if ((Get-Date) -ge $deadline) {
            Write-Log 'Profile cleanup time budget reached. Stopping active jobs and deferring remaining profiles to the next run.' 'WARN'

            for ($i = $running.Count - 1; $i -ge 0; $i--) {
                $jobInfo = $running[$i]
                Stop-Job -Job $jobInfo.Job -Force -ErrorAction SilentlyContinue
                Remove-Job -Job $jobInfo.Job -Force -ErrorAction SilentlyContinue
                $script:Summary.DeferredProfiles++
                Add-StateItemUnique -State $state -ListName 'Pending' -Value $jobInfo.Profile.LocalPath
                $script:DeferredProfileDetails.Add([PSCustomObject]@{
                    ProfileName  = $jobInfo.Profile.ProfileName
                    LocalPath    = $jobInfo.Profile.LocalPath
                    SID          = $jobInfo.Profile.SID
                    Reason       = 'CleanupWindowEndedActiveJobStopped'
                    CreatedTime  = $jobInfo.Profile.CreatedTime
                    LastUseTime  = $jobInfo.Profile.LastUseTime
                    DaysOnSystem = $jobInfo.Profile.DaysOnSystem
                }) | Out-Null
                $running.RemoveAt($i)
            }

            while ($queue.Count -gt 0) {
                $deferred = $queue.Dequeue()
                $script:Summary.DeferredProfiles++
                Add-StateItemUnique -State $state -ListName 'Pending' -Value $deferred.LocalPath
                $script:DeferredProfileDetails.Add([PSCustomObject]@{
                    ProfileName  = $deferred.ProfileName
                    LocalPath    = $deferred.LocalPath
                    SID          = $deferred.SID
                    Reason       = 'CleanupWindowEndedQueued'
                    CreatedTime  = $deferred.CreatedTime
                    LastUseTime  = $deferred.LastUseTime
                    DaysOnSystem = $deferred.DaysOnSystem
                }) | Out-Null
            }

            Save-ProfileCleanupState -State $state -Path $ProfileCleanupStatePath
            Write-YamlLog
            break
        }

        while ($queue.Count -gt 0 -and (Get-RunningJobCount -Jobs $running) -lt $MaxParallelProfileDeletes) {
            $item = $queue.Dequeue()

            if (-not (Test-Path -LiteralPath $item.LocalPath)) {
                Write-Log "Profile folder already absent before queue start: $($item.ProfileName) ($($item.LocalPath))" 'WARN'
                try {
                    Remove-ProfileRegistration -Sid $item.SID -ProfilePath $item.LocalPath -ProfileName $item.ProfileName
                    $script:Summary.DeletedProfiles++
                    $removedOneDriveTasks = Remove-OneDriveTasksForProfile `
                        -Sid $item.SID `
                        -ProfileName $item.ProfileName
                    $script:Summary.OneDriveTasksRemoved += $removedOneDriveTasks
                    Remove-StateItem -State $state -ListName 'Pending' -Value $item.LocalPath
                    Add-StateItemUnique -State $state -ListName 'Completed' -Value $item.LocalPath
                }
                catch {
                    $script:Summary.FailedProfiles++
                    Add-StateItemUnique -State $state -ListName 'Failed' -Value $item.LocalPath
                    Write-Log "Profile registration cleanup failed for absent folder $($item.ProfileName): $($_.Exception.Message)" 'ERROR'
                }
                Save-ProfileCleanupState -State $state -Path $ProfileCleanupStatePath
                continue
            }

            try {
                Write-Log "Removing Win32_UserProfile registration before filesystem cleanup: $($item.ProfileName) ($($item.SID))" 'INFO'
                Remove-ProfileRegistration -Sid $item.SID -ProfilePath $item.LocalPath -ProfileName $item.ProfileName
            }
            catch {
                Write-Log "Native profile registration cleanup did not complete before folder deletion for $($item.ProfileName): $($_.Exception.Message). Continuing with filesystem cleanup." 'WARN'
            }

            if (-not (Test-Path -LiteralPath $item.LocalPath)) {
                Write-Log "Native profile removal deleted the folder before job start: $($item.ProfileName) ($($item.LocalPath))" 'OK'
                $script:Summary.DeletedProfiles++
                $removedOneDriveTasks = Remove-OneDriveTasksForProfile `
                    -Sid $item.SID `
                    -ProfileName $item.ProfileName
                $script:Summary.OneDriveTasksRemoved += $removedOneDriveTasks
                Remove-StateItem -State $state -ListName 'Pending' -Value $item.LocalPath
                Add-StateItemUnique -State $state -ListName 'Completed' -Value $item.LocalPath
                $script:DeletedProfileDetails.Add([PSCustomObject]@{
                    ProfileName   = $item.ProfileName
                    LocalPath     = $item.LocalPath
                    SID           = $item.SID
                    Loaded        = $item.Loaded
                    Special       = $item.Special
                    CreatedTime   = $item.CreatedTime
                    LastUseTime   = $item.LastUseTime
                    DaysOnSystem  = $item.DaysOnSystem
                    DeleteSeconds = 0
                    Message       = 'Removed by Win32_UserProfile cleanup before robocopy job.'
                }) | Out-Null
                Save-ProfileCleanupState -State $state -Path $ProfileCleanupStatePath
                continue
            }

            $safeJobName = 'ProfileDelete_' + ($item.SID -replace '[^A-Za-z0-9_\-]', '_')
            Write-Log "Starting timeout-safe profile delete job: $($item.ProfileName) ($($item.LocalPath))" 'INFO'
            $job = Start-Job -Name $safeJobName -ScriptBlock $profileDeleteJobScript -ArgumentList $item.LocalPath, $emptyMirrorRoot, $robocopyTimeoutSeconds, $rmdirTimeoutSeconds
            $running.Add([PSCustomObject]@{
                Job     = $job
                Profile = $item
                Started = Get-Date
            }) | Out-Null
        }

        if ($running.Count -gt 0 -or $queue.Count -gt 0) {
            Start-Sleep -Seconds 2
        }
    }

    Stop-ProfileCleanupJobs
    Save-ProfileCleanupState -State $state -Path $ProfileCleanupStatePath

    try {
        Get-ChildItem -LiteralPath $emptyMirrorRoot -Directory -Force -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch { }
}

# =====================================================================
# Windows 11 UI Preferences
# Applies the standard lab UI settings to loaded user profiles, offline
# profiles that remain after cleanup, and the Default User profile.
# =====================================================================

function Invoke-WithLoadedUserHive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$HiveFile,

        [Parameter(Mandatory)]
        [string]$DisplayName,

        [Parameter(Mandatory)]
        [scriptblock]$Action
    )

    if (-not (Test-Path -LiteralPath $HiveFile)) {
        Write-Log "User hive file not found for $DisplayName : $HiveFile" 'WARN'
        return $false
    }

    $tempHiveName = "TempWindows11UI_$([guid]::NewGuid().ToString('N'))"
    $loaded = $false

    try {
        $loadOutput = & reg.exe load "HKU\$tempHiveName" "$HiveFile" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Failed to load hive for $DisplayName : $($loadOutput -join ' ')" 'WARN'
            return $false
        }

        $loaded = $true
        $registryRoot = "Registry::HKEY_USERS\$tempHiveName"
        & $Action $registryRoot $DisplayName | Out-Null
        return $true
    }
    catch {
        Write-Log "Unexpected error while processing hive for $DisplayName : $($_.Exception.Message)" 'WARN'
        return $false
    }
    finally {
        if ($loaded) {
            try {
                [gc]::Collect()
                [gc]::WaitForPendingFinalizers()
                Start-Sleep -Milliseconds 300
                $unloadOutput = & reg.exe unload "HKU\$tempHiveName" 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "Unloaded temporary registry hive for $DisplayName" 'OK'
                }
                else {
                    Write-Log "Failed to unload temporary registry hive for $DisplayName : $($unloadOutput -join ' ')" 'WARN'
                }
            }
            catch {
                Write-Log "Unexpected error unloading hive for $DisplayName : $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Set-Windows11UIPreferencesForHive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RegistryRoot,

        [Parameter(Mandatory)]
        [string]$DisplayName
    )

    # Convert the PowerShell registry path to a native reg.exe path.
    # reg.exe is used here because it is more reliable when writing values
    # to temporarily loaded offline user hives.
    if ($RegistryRoot -eq 'HKCU:') {
        $nativeRegistryRoot = 'HKCU'
    }
    elseif ($RegistryRoot -like 'Registry::HKEY_USERS\*') {
        $nativeRegistryRoot = $RegistryRoot -replace '^Registry::HKEY_USERS', 'HKU'
    }
    else {
        Write-Log "Unsupported registry root for ${DisplayName}: $RegistryRoot" 'WARN'
        return $false
    }

    $enabledSettings = @(
        $script:Windows11UIPreferences |
            Where-Object { $_.Enabled }
    )

    if ($enabledSettings.Count -eq 0) {
        Write-Log "No enabled Windows 11 UI preferences are configured for $DisplayName." 'INFO'
        return $true
    }

    $successCount = 0
    $failureCount = 0

    foreach ($setting in $enabledSettings) {
        $nativeKey = "$nativeRegistryRoot\$($setting.SubKey)"

        try {
            # Use /ve when the setting targets the unnamed/default registry value.
            # Named values continue to use /v with the configured value name.
            if ([string]::IsNullOrEmpty([string]$setting.Name)) {
                # Omitting /d creates the unnamed REG_SZ value with empty data,
                # which is the value required for the classic context menu CLSID.
                if ([string]::IsNullOrEmpty([string]$setting.Value)) {
                    $regOutput = & reg.exe add $nativeKey /ve /t $setting.Type /f 2>&1
                }
                else {
                    $regOutput = & reg.exe add $nativeKey /ve /t $setting.Type /d $setting.Value /f 2>&1
                }
                $registryValueLabel = '(Default)'
            }
            else {
                $regOutput = & reg.exe add $nativeKey /v $setting.Name /t $setting.Type /d $setting.Value /f 2>&1
                $registryValueLabel = [string]$setting.Name
            }

            $regExitCode = $LASTEXITCODE

            if ($regExitCode -eq 0) {
                $successCount++
                Write-Log "Applied $($setting.Description) for $DisplayName." 'OK'
            }
            else {
                $failureCount++
                Write-Log "Failed to apply $($setting.Description) for $DisplayName. Registry value: $registryValueLabel. reg.exe exit code: $regExitCode. Output: $($regOutput -join ' ')" 'WARN'
            }
        }
        catch {
            $failureCount++
            $registryValueLabel = if ([string]::IsNullOrEmpty([string]$setting.Name)) { '(Default)' } else { [string]$setting.Name }
            Write-Log "Failed to apply $($setting.Description) for $DisplayName. Registry value: $registryValueLabel. Error: $($_.Exception.Message)" 'WARN'
        }
    }

    # Microsoft Teams / Chat is intentionally not changed.
    # Do not set TaskbarMn here because Teams is used on campus.

    if ($failureCount -eq 0) {
        Write-Log "Applied all $successCount Windows 11 UI preference(s) for $DisplayName." 'OK'
        return $true
    }

    Write-Log "Windows 11 UI preference result for ${DisplayName}: Applied=$successCount, Failed=$failureCount." 'WARN'
    return ($successCount -gt 0)
}

function Set-Windows11UIPreferencesForAllUsers {
    [CmdletBinding()]
    param()

    $enabledPreferenceCount = @($script:Windows11UIPreferences | Where-Object { $_.Enabled }).Count
    Write-Log "Applying $enabledPreferenceCount configured Windows 11 UI preference(s) for all users..." 'INFO'
    Write-Log "Microsoft Teams / Chat taskbar settings are intentionally left unchanged." 'INFO'

    $processedSids = New-Object 'System.Collections.Generic.HashSet[string]'

    [void](Set-Windows11UIPreferencesForHive -RegistryRoot 'HKCU:' -DisplayName 'current user')

    try {
        $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        [void]$processedSids.Add($currentSid)
    }
    catch {
        Write-Log "Could not determine current user SID: $($_.Exception.Message)" 'WARN'
    }

    try {
        Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction Stop |
            Where-Object {
                $_.PSChildName -match '^S-1-5-21-' -and
                $_.PSChildName -notmatch '_Classes$'
            } |
            ForEach-Object {
                $sid = $_.PSChildName

                if ($processedSids.Contains($sid)) {
                    Write-Log "Loaded profile $sid was already processed; skipping duplicate UI preference application." 'INFO'
                    return
                }

                [void]$processedSids.Add($sid)
                [void](Set-Windows11UIPreferencesForHive -RegistryRoot "Registry::HKEY_USERS\$sid" -DisplayName "loaded profile $sid")
            }
    }
    catch {
        Write-Log "Failed to enumerate loaded user registry hives for Windows 11 UI preferences: $($_.Exception.Message)" 'WARN'
    }

    try {
        $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
            Where-Object {
                -not $_.Special -and
                $_.SID -match '^S-1-5-21-' -and
                $_.LocalPath -and
                (Test-Path -LiteralPath (Join-Path $_.LocalPath 'NTUSER.DAT'))
            }

        foreach ($profile in $profiles) {
            if ($processedSids.Contains($profile.SID) -or (Test-Path -LiteralPath "Registry::HKEY_USERS\$($profile.SID)")) {
                Write-Log "Profile already loaded or processed; skipping offline Windows 11 UI preference load for $($profile.LocalPath)" 'INFO'
                continue
            }

            $ntUserDat = Join-Path $profile.LocalPath 'NTUSER.DAT'
            [void](Invoke-WithLoadedUserHive -HiveFile $ntUserDat -DisplayName "offline profile $($profile.LocalPath)" -Action {
                param($RegistryRoot, $DisplayName)
                Set-Windows11UIPreferencesForHive -RegistryRoot $RegistryRoot -DisplayName $DisplayName
            })
        }
    }
    catch {
        Write-Log "Failed to process offline user profiles for Windows 11 UI preferences: $($_.Exception.Message)" 'WARN'
    }

    $defaultHive = Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT'
    if (Test-Path -LiteralPath $defaultHive) {
        [void](Invoke-WithLoadedUserHive -HiveFile $defaultHive -DisplayName 'Default User profile for future users' -Action {
            param($RegistryRoot, $DisplayName)
            Set-Windows11UIPreferencesForHive -RegistryRoot $RegistryRoot -DisplayName $DisplayName
        })
    }
    else {
        Write-Log "Default User hive not found at expected path: $defaultHive" 'WARN'
    }

    Write-Log "Completed Windows 11 UI preference application for all available users." 'OK'
}

if (-not (Test-IsAdministrator)) {
    Write-Error 'Please run this script as Administrator.'
    $script:ErrorCount++
    Complete-ScriptExecution -ExitCode 1
}

Ensure-LogDirectory
$script:DiskBefore = Get-SystemDriveSpace
if ($script:DiskBefore) { $script:Summary.DiskFreeGBBefore = $script:DiskBefore.FreeGB }
Write-YamlLog

if (Test-ComputerNamePatternMatch -ComputerName $script:ComputerName -Patterns $EdgeInPrivateComputerNamePatterns) {
    $script:EdgeInPrivateMatched = $true
    Write-Log "Computer name '$($script:ComputerName)' matches Edge InPrivate auto-launch list." 'INFO'
    $script:EdgeInPrivateConfigured = [bool](Set-EdgeInPrivateAutoLaunch -Url $EdgeInPrivateUrl)
}
else {
    Write-Log "Computer name '$($script:ComputerName)' does not match Edge InPrivate auto-launch list. Skipping Edge configuration." 'INFO'
}

Write-Log 'Starting profile cleanup.' 'INFO'
Write-Log "Users root: $UsersRoot" 'INFO'
Write-Log "Excluded profile names: $($ExcludedProfiles -join ', ')" 'INFO'
Write-Log "Skip loaded profiles: $SkipLoadedProfiles" 'INFO'
Write-Log "Skip special profiles: $SkipSpecialProfiles" 'INFO'
Write-Log "OlderThanDays filter: $OlderThanDays" 'INFO'
Write-Log "Max parallel profile deletes: $MaxParallelProfileDeletes" 'INFO'
Write-Log "Profile cleanup time budget: $ProfileCleanupTimeLimitMinutes minute(s)" 'INFO'
Write-Log "Profile folder job timeout: $ProfileFolderJobTimeoutMinutes minute(s)" 'INFO'
Write-Log "Profile cleanup state path: $ProfileCleanupStatePath" 'INFO'
Write-Log "YAML log path: $($script:YamlLogPath)" 'INFO'

try {
    $allUserProfiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.LocalPath) -and
        $_.LocalPath -like "$UsersRoot\*"
    }
}
catch {
    Write-Log "Failed to enumerate Win32_UserProfile instances: $($_.Exception.Message)" 'ERROR'
    $script:Summary.EndTime = Get-Date
    Complete-ScriptExecution -ExitCode 2
}

$script:Summary.FoundProfiles = @($allUserProfiles).Count
Write-Log "Found $($script:Summary.FoundProfiles) profile(s) under $UsersRoot." 'INFO'

$profilesToDelete = New-Object System.Collections.ArrayList

foreach ($profile in $allUserProfiles) {
    $profilePath = $profile.LocalPath
    $profileName = Get-ProfileFolderName -Path $profilePath

    if ([string]::IsNullOrWhiteSpace($profileName)) {
        Write-Log "Skipping profile with invalid path: $profilePath" 'WARN'
        continue
    }

    $ageData = Get-ProfileAgeData -ProfilePath $profilePath -LastUseTime $profile.LastUseTime

    if ($ExcludedProfiles -icontains $profileName) {
        $script:Summary.ExcludedProfiles++
        Write-Log "Skipping excluded profile: $profileName" 'WARN'
        $script:SkippedProfileDetails.Add([PSCustomObject]@{
            ProfileName  = $profileName
            LocalPath    = $profilePath
            SID          = $profile.SID
            Reason       = 'Excluded'
            CreatedTime  = $ageData.CreatedTime
            LastUseTime  = $ageData.LastUseTime
            DaysOnSystem = $ageData.DaysOnSystem
        }) | Out-Null
        continue
    }

    if ($SkipSpecialProfiles -and $profile.Special) {
        $script:Summary.SkippedSpecial++
        Write-Log "Skipping special/system profile: $profileName ($profilePath)" 'WARN'
        $script:SkippedProfileDetails.Add([PSCustomObject]@{
            ProfileName  = $profileName
            LocalPath    = $profilePath
            SID          = $profile.SID
            Reason       = 'SpecialProfile'
            CreatedTime  = $ageData.CreatedTime
            LastUseTime  = $ageData.LastUseTime
            DaysOnSystem = $ageData.DaysOnSystem
        }) | Out-Null
        continue
    }

    if ($SkipLoadedProfiles -and $profile.Loaded) {
        $script:Summary.SkippedLoaded++
        Write-Log "Skipping loaded profile: $profileName ($profilePath)" 'WARN'
        $script:SkippedProfileDetails.Add([PSCustomObject]@{
            ProfileName  = $profileName
            LocalPath    = $profilePath
            SID          = $profile.SID
            Reason       = 'LoadedProfile'
            CreatedTime  = $ageData.CreatedTime
            LastUseTime  = $ageData.LastUseTime
            DaysOnSystem = $ageData.DaysOnSystem
        }) | Out-Null
        continue
    }

    if ($OlderThanDays -gt 0) {
        try {
            $cutoff = (Get-Date).AddDays(-$OlderThanDays)
            $lastUse = $ageData.LastUseTime
            if ($lastUse -and $lastUse -gt $cutoff) {
                $script:Summary.SkippedByAge++
                Write-Log "Skipping recent profile: $profileName (LastUseTime: $lastUse)" 'WARN'
                $script:SkippedProfileDetails.Add([PSCustomObject]@{
                    ProfileName  = $profileName
                    LocalPath    = $profilePath
                    SID          = $profile.SID
                    Reason       = 'TooRecent'
                    CreatedTime  = $ageData.CreatedTime
                    LastUseTime  = $ageData.LastUseTime
                    DaysOnSystem = $ageData.DaysOnSystem
                }) | Out-Null
                continue
            }
        }
        catch {
            Write-Log "Could not evaluate LastUseTime for ${profileName}: $($_.Exception.Message)" 'WARN'
        }
    }

    $targetDescription = "$profileName ($profilePath)"
    if ($PSCmdlet.ShouldProcess($targetDescription, 'Queue profile folder for capped parallel timeout-safe deletion and remove profile registration')) {
        $profilesToDelete.Add([PSCustomObject]@{
            ProfileName  = $profileName
            LocalPath    = $profilePath
            SID          = $profile.SID
            Loaded       = $profile.Loaded
            Special      = $profile.Special
            CreatedTime  = $ageData.CreatedTime
            LastUseTime  = $ageData.LastUseTime
            DaysOnSystem = $ageData.DaysOnSystem
        }) | Out-Null
    }
}

$script:Summary.QueuedProfiles = $profilesToDelete.Count
Write-Log "Queued $($profilesToDelete.Count) profile(s) for capped parallel deletion." 'INFO'
Write-YamlLog

if ($profilesToDelete.Count -gt 0) {
    Invoke-ParallelProfileDeletion -ProfilesToDelete @($profilesToDelete.ToArray())
}
else {
    Write-Log 'No eligible profiles were queued for deletion.' 'OK'
}

# ---------------------------------------------------------------------
# Apply the standard Windows 11 lab UI and classic context-menu settings
# after profile removal. This updates remaining users and the Default User
# profile so future users inherit the same configuration.
# ---------------------------------------------------------------------
Write-Log 'Profile removal stage complete. Applying Windows 11 UI and classic context-menu preferences to remaining and future user profiles.' 'INFO'
Set-Windows11UIPreferencesForAllUsers
Write-YamlLog

$script:Summary.EndTime = Get-Date

Write-Log 'Profile cleanup complete.' 'INFO'
Write-Log "Summary: Found=$($script:Summary.FoundProfiles), Excluded=$($script:Summary.ExcludedProfiles), LoadedSkipped=$($script:Summary.SkippedLoaded), SpecialSkipped=$($script:Summary.SkippedSpecial), AgeSkipped=$($script:Summary.SkippedByAge), Queued=$($script:Summary.QueuedProfiles), Deleted=$($script:Summary.DeletedProfiles), OneDriveTasksRemoved=$($script:Summary.OneDriveTasksRemoved), Failed=$($script:Summary.FailedProfiles), TimedOut=$($script:Summary.TimedOutProfiles), Deferred=$($script:Summary.DeferredProfiles)" 'INFO'

Write-YamlLog

if ($script:Summary.FailedProfiles -gt 0) { Complete-ScriptExecution -ExitCode 2 }
Complete-ScriptExecution -ExitCode 0
