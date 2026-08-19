# =====================================================================
# ScriptName: 07_Force_Reboot_Install_Updates.ps1
# ScriptVersion: 2.0.5
# LastUpdated: 2026-08-18
# ChangeLog: v2.0.5 adds -StartupResume support. Startup-triggered executions exit immediately when no reboot cycle is active.
#            v2.0.4 extends the cycle to a maximum of three reboots, then performs a final no-more-reboots verification pass.
#            Adds structured persistent-reboot cause telemetry including likely source and affected file/change details.
#            v2.0.3 adds Maintenance.Framework v2.4 staged text logging and publishes the completed log before script exit/reboot.
#            Previous telemetry, reboot verification, boot/uptime, update status, and safe reboot-flag handling remain intact.
# =====================================================================

[CmdletBinding()]
param(
    [ValidateRange(0,86400)]
    [int]$RebootDelaySeconds = 30,
    [string]$LogDirectory = 'C:\Logs',
    [string]$StateDirectory = 'C:\ProgramData\MISMaintenance\State',
    [string]$StateFileName = '07_Force_Reboot_Install_Updates_State.json',
    [switch]$AllowUnsafeRebootFlagCleanup,
    [switch]$StartupResume
)

$ErrorActionPreference = 'Stop'

$script:RunStart         = Get-Date
$script:ScriptName       = '07_Force_Reboot_Install_Updates.ps1'
$script:ScriptVersion    = '2.0.5'
$script:RunId            = [guid]::NewGuid().ToString()
$script:Domain           = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { $env:USERDOMAIN }
$script:TelemetryPath    = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestJsonPath   = Join-Path $LogDirectory '07_Force_Reboot_Install_Updates.latest.json'
$script:TextLogPath      = $null
$script:PublishedLogPath = $null
$script:LogSession       = $null
$script:LogPublished     = $false
$script:ExitCode         = 0
$script:BootSnapshot     = $null
$script:UpdateSnapshot   = $null
$script:LoggedOnUsers    = @()
$script:RebootVerified   = $false
$script:PreviousBootTime = $null
$script:StateRecovered   = $false
$script:ComputerName     = $env:COMPUTERNAME
$script:StateFilePath    = Join-Path $StateDirectory $StateFileName
$script:YamlLogPath      = $null
$script:OverallResult    = 'Unknown'
$script:FailureMessage   = $null
$script:CurrentFlags     = @()
$script:RebootCauseDetails = @()
$script:PersistentAfterThirdReboot = $false
$script:FinalVerificationAfterThirdReboot = $false
$script:ClearResults     = @()
$script:ActionHistory    = @()
$script:CurrentStage     = 0
$script:RebootIssued     = $false
$script:RebootReason     = $null
$script:MutexName        = 'Global\MIS_07_Force_Reboot_Install_Updates'
$script:Mutex            = $null
$script:MutexAcquired    = $false

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 07 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

function Ensure-Folder {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Initialize-Paths {
    Ensure-Folder -Path $LogDirectory
    Ensure-Folder -Path $StateDirectory

    $script:TelemetryPath  = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
    $script:LatestJsonPath = Join-Path $LogDirectory '07_Force_Reboot_Install_Updates.latest.json'

    Archive-MaintenanceLogs `
        -ScriptName $script:ScriptName `
        -LogRoot $LogDirectory `
        -AdditionalPatterns @(
            '07_Force_Reboot_Install_Updates.log',
            '*-07_Force_Reboot_Install_Updates-*.log'
        ) | Out-Null

    $script:LogSession = New-MaintenanceStagedLog `
        -ScriptName $script:ScriptName `
        -LogRoot $LogDirectory `
        -StagingRoot $MaintenanceConfig.LogStagingRoot `
        -ComputerName $script:ComputerName `
        -Timestamp $script:RunStart

    $script:TextLogPath = [string]$script:LogSession.WorkingPath
    $script:PublishedLogPath = [string]$script:LogSession.PublishedPath

    $timestamp = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $baseName = "$($script:ComputerName)-ForceRebootInstallUpdates-$timestamp"
    $script:YamlLogPath = Join-Path $LogDirectory ($baseName + '.yaml')
}

function Initialize-SingleInstanceLock {
    try {
        $createdNew = $false
        $script:Mutex = New-Object System.Threading.Mutex($false, $script:MutexName, [ref]$createdNew)

        if (-not $script:Mutex.WaitOne(0)) {
            Write-Status 'Another instance of Script 07 is already running. Exiting to prevent state-file corruption.' 'WARN'
            $script:OverallResult = 'SkippedAlreadyRunning'
            $script:ExitCode = 0
            Write-YamlLog
            Complete-TextLogPublication -Reason 'Skipped because another instance is running'
            exit 0
        }

        $script:MutexAcquired = $true
        Write-Status 'Single-instance lock acquired.' 'INFO'
    }
    catch {
        Write-Status "Failed to acquire single-instance lock: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

function Release-SingleInstanceLock {
    try {
        if ($script:MutexAcquired -and $null -ne $script:Mutex) {
            [void]$script:Mutex.ReleaseMutex()
            $script:MutexAcquired = $false
            Write-Status 'Single-instance lock released.' 'INFO'
        }
    }
    catch {
        Write-Warning "Failed to release single-instance lock: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $script:Mutex) {
            $script:Mutex.Dispose()
            $script:Mutex = $null
        }
    }
}

function Confirm-FileWriteToDisk {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        $stream.Flush($true)
    }
    catch {
        Write-Status "Unable to force-flush file to disk [$Path]: $($_.Exception.Message)" 'WARN'
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Confirm-StateWriteBeforeReboot {
    try {
        Confirm-FileWriteToDisk -Path $script:StateFilePath
        if ($script:YamlLogPath) {
            Confirm-FileWriteToDisk -Path $script:YamlLogPath
        }
        if ($script:TelemetryPath) {
            Confirm-FileWriteToDisk -Path $script:TelemetryPath
        }
        if ($script:LatestJsonPath) {
            Confirm-FileWriteToDisk -Path $script:LatestJsonPath
        }
        if ($script:TextLogPath) {
            Confirm-FileWriteToDisk -Path $script:TextLogPath
        }
        Start-Sleep -Seconds 3
        Write-Status 'State and YAML log writes confirmed before reboot.' 'INFO'
    }
    catch {
        Write-Status "Pre-reboot write confirmation failed: $($_.Exception.Message)" 'WARN'
    }
}

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO','OK','WARN','ERROR','SUCCESS','WARNING')]
        [string]$Level = 'INFO'
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
        $activeTextLogDirectory = Split-Path -Parent $script:TextLogPath
        Ensure-Folder -Path $activeTextLogDirectory
        Add-Content -LiteralPath $script:TextLogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write text log '$($script:TextLogPath)': $($_.Exception.Message)"
    }

    $script:ActionHistory += [PSCustomObject]@{
        Time    = $timestamp
        Level   = $normalizedLevel
        Message = $Message
    }
}

function Complete-TextLogPublication {
    [CmdletBinding()]
    param(
        [string]$Reason = 'Script completion'
    )

    if ($script:LogPublished -or $null -eq $script:LogSession) {
        return
    }

    # This is intentionally the final text-log append. Do not call Write-Status
    # after this point because Elastic may open the published file immediately.
    Write-Status ("Finalizing text log before exit. Reason={0}; Result={1}; ExitCode={2}" -f `
        $Reason,
        $script:OverallResult,
        $script:ExitCode) 'INFO'

    Confirm-FileWriteToDisk -Path $script:TextLogPath
    $publishResult = Publish-MaintenanceLog -LogSession $script:LogSession

    if ($publishResult.Published) {
        $script:LogPublished = $true
        Write-Host ("Published completed script 07 text log for Elastic: {0}" -f $script:PublishedLogPath) -ForegroundColor Green
    }
    else {
        Write-Warning ("Script 07 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
    }
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


function Get-BootSnapshot {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $boot = [datetime]$os.LastBootUpTime
        return [PSCustomObject]@{
            LastBootTime  = $boot.ToString('o')
            UptimeSeconds = [math]::Round(((Get-Date) - $boot).TotalSeconds, 0)
            UptimeDays    = [math]::Round(((Get-Date) - $boot).TotalDays, 2)
        }
    }
    catch {
        Write-Status "Unable to collect boot information: $($_.Exception.Message)" 'WARN'
        return [PSCustomObject]@{ LastBootTime = $null; UptimeSeconds = $null; UptimeDays = $null }
    }
}

function Get-LoggedOnUserSnapshot {
    $results = New-Object System.Collections.Generic.List[object]
    try {
        $explorers = @(Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue)
        foreach ($process in $explorers) {
            try {
                $owner = Invoke-CimMethod -InputObject $process -MethodName GetOwner -ErrorAction Stop
                if ($owner.ReturnValue -eq 0 -and $owner.User) {
                    $name = if ($owner.Domain) { "$($owner.Domain)\$($owner.User)" } else { $owner.User }
                    if (-not ($results | Where-Object { $_.UserName -eq $name })) {
                        $results.Add([PSCustomObject]@{ UserName = $name; SessionProcessId = [int]$process.ProcessId }) | Out-Null
                    }
                }
            }
            catch { }
        }
    }
    catch {
        Write-Status "Unable to enumerate interactive users: $($_.Exception.Message)" 'WARN'
    }
    return @($results | ForEach-Object { $_ })
}

function Get-WindowsUpdateSnapshot {
    $result = [ordered]@{
        WindowsEdition          = $null
        DisplayVersion          = $null
        CurrentBuild            = $null
        UBR                     = $null
        BuildString             = $null
        LatestInstalledHotFix   = $null
        LatestHotFixInstalledOn = $null
        RecentUpdateEventCount  = 0
        RecentUpdateFailures    = @()
    }

    try {
        $cv = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $result.WindowsEdition = $cv.ProductName
        $result.DisplayVersion = $cv.DisplayVersion
        $result.CurrentBuild = $cv.CurrentBuildNumber
        $result.UBR = $cv.UBR
        if ($cv.CurrentBuildNumber) {
            $result.BuildString = if ($null -ne $cv.UBR) { "$($cv.CurrentBuildNumber).$($cv.UBR)" } else { [string]$cv.CurrentBuildNumber }
        }
    }
    catch {
        Write-Status "Unable to read Windows build information: $($_.Exception.Message)" 'WARN'
    }

    try {
        $latest = Get-HotFix -ErrorAction Stop | Where-Object InstalledOn | Sort-Object InstalledOn -Descending | Select-Object -First 1
        if ($latest) {
            $result.LatestInstalledHotFix = $latest.HotFixID
            $result.LatestHotFixInstalledOn = ([datetime]$latest.InstalledOn).ToString('o')
        }
    }
    catch {
        Write-Status "Unable to read latest installed hotfix: $($_.Exception.Message)" 'WARN'
    }

    try {
        $start = (Get-Date).AddDays(-7)
        $events = @(Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-WindowsUpdateClient/Operational'; StartTime=$start } -ErrorAction SilentlyContinue)
        $result.RecentUpdateEventCount = $events.Count
        $failures = @($events | Where-Object { $_.LevelDisplayName -eq 'Error' } | Select-Object -First 20)
        $result.RecentUpdateFailures = @($failures | ForEach-Object {
            [PSCustomObject]@{
                TimeCreated = if ($_.TimeCreated) { $_.TimeCreated.ToString('o') } else { $null }
                Id          = $_.Id
                Message     = $_.Message
            }
        })
    }
    catch {
        Write-Status "Unable to read recent Windows Update events: $($_.Exception.Message)" 'WARN'
    }

    return [PSCustomObject]$result
}

function Get-ExecutionStatus {
    if ($script:OverallResult -eq 'Failed') { return 'Failed' }
    if ($script:RebootIssued) { return 'SuccessRebootInitiated' }
    if ($script:OverallResult -match 'Persistent') { return 'SuccessWithWarnings' }
    if (@($script:ActionHistory | Where-Object Level -eq 'ERROR').Count -gt 0) { return 'Failed' }
    if (@($script:ActionHistory | Where-Object Level -eq 'WARNING').Count -gt 0) { return 'SuccessWithWarnings' }
    return 'Success'
}

function Write-ExecutionTelemetry {
    try {
        $runEnd = Get-Date
        $warnings = @($script:ActionHistory | Where-Object Level -eq 'WARNING').Count
        $errors = @($script:ActionHistory | Where-Object Level -eq 'ERROR').Count
        $status = Get-ExecutionStatus
        $flags = @($script:CurrentFlags)
        $clear = @($script:ClearResults)
        $causeDetails = @($script:RebootCauseDetails)
        $previousUpdateContext = Get-PreviousWindowsUpdateContext

        $event = [ordered]@{
            EventType                   = 'maintenance.execution'
            ComputerName                = $script:ComputerName
            Domain                      = $script:Domain
            ScriptName                  = $script:ScriptName
            ScriptVersion               = $script:ScriptVersion
            RunId                       = $script:RunId
            Status                      = $status
            ExitCode                    = $script:ExitCode
            StartTime                   = $script:RunStart.ToString('o')
            EndTime                     = $runEnd.ToString('o')
            DurationSeconds             = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 2)
            WarningCount                = $warnings
            ErrorCount                  = $errors
            Timestamp                   = $runEnd.ToUniversalTime().ToString('o')
            Stage                       = $script:CurrentStage
            OverallResult               = $script:OverallResult
            FailureMessage              = $script:FailureMessage
            RebootIssued                = $script:RebootIssued
            RebootReason                = $script:RebootReason
            RebootDelaySeconds          = $RebootDelaySeconds
            RebootVerified              = $script:RebootVerified
            PreviousBootTime            = $script:PreviousBootTime
            CurrentBoot                 = $script:BootSnapshot
            LoggedOnUsers               = @($script:LoggedOnUsers)
            LoggedOnUserCount           = @($script:LoggedOnUsers).Count
            PendingReboot                = ($flags.Count -gt 0)
            PendingRebootReasonCount     = $flags.Count
            PendingRebootReasons        = @($flags | Select-Object Name,Type,Path,ValueName,Details)
            RebootCauseDetails           = $causeDetails
            FinalVerificationAfterThirdReboot = [bool]$script:FinalVerificationAfterThirdReboot
            PersistentAfterThirdReboot   = [bool]$script:PersistentAfterThirdReboot
            PreviousWindowsUpdateContext = $previousUpdateContext
            StartupResumeInvocation      = [bool]$StartupResume
            UnsafeFlagCleanupEnabled     = [bool]$AllowUnsafeRebootFlagCleanup
            ClearActionCount             = $clear.Count
            ClearActions                 = $clear
            WindowsUpdate                = $script:UpdateSnapshot
            StateFilePath                = $script:StateFilePath
            StateRecovered               = $script:StateRecovered
            TextLogPath                  = $script:PublishedLogPath
        }

        $jsonCompact = $event | ConvertTo-Json -Depth 14 -Compress
        Write-MaintenanceTelemetryLine -Path $script:TelemetryPath -JsonLine $jsonCompact
        $event | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath $script:LatestJsonPath -Encoding UTF8 -Force
        Confirm-FileWriteToDisk -Path $script:TelemetryPath
        Confirm-FileWriteToDisk -Path $script:LatestJsonPath
    }
    catch {
        Write-Warning "Failed to write execution telemetry: $($_.Exception.Message)"
    }
}

function ConvertTo-YamlScalar {
    param(
        [AllowNull()]$Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [string]$Value
    }

    if ($Value -is [datetime]) {
        return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'"
    }

    $text = [string]$Value
    $text = $text -replace "`r", ' '
    $text = $text -replace "`n", ' '
    $text = $text -replace "'", "''"
    return "'" + $text + "'"
}

function New-FlagRecord {
    param(
        [string]$Name,
        [string]$Type,
        [string]$Path,
        [string]$ValueName,
        [string]$Details
    )

    [PSCustomObject]@{
        Name      = $Name
        Type      = $Type
        Path      = $Path
        ValueName = $ValueName
        Details   = $Details
    }
}

function New-ClearResult {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Action,
        [string]$Status,
        [string]$Message
    )

    [PSCustomObject]@{
        Name    = $Name
        Path    = $Path
        Action  = $Action
        Status  = $Status
        Message = $Message
    }
}

function Test-RegistryKeyExists {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    try {
        Test-Path -LiteralPath $Path
    }
    catch {
        $false
    }
}

function Get-RegistryValueSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
        $item.$Name
    }
    catch {
        $null
    }
}


function Get-RebootCauseDetails {
    param(
        [Parameter(Mandatory)]
        [array]$Flags
    )

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($flag in @($Flags)) {
        $likelySource = 'Unknown'
        $affectedItems = @()
        $explanation = [string]$flag.Details

        switch ([string]$flag.Name) {
            'WindowsUpdateRebootRequired' {
                $likelySource = 'Windows Update'
                $explanation = 'Windows Update explicitly reports that a reboot is required to finish update processing.'
            }
            'WindowsUpdatePostRebootReporting' {
                $likelySource = 'Windows Update'
                $explanation = 'Windows Update has post-reboot reporting work pending.'
            }
            'CBSRebootPending' {
                $likelySource = 'Component Based Servicing'
                $explanation = 'Windows servicing (CBS) reports a component/package change that has not completed across a reboot.'
            }
            'CBSRebootInProgress' {
                $likelySource = 'Component Based Servicing'
                $explanation = 'Windows servicing reports a reboot transaction still in progress.'
            }
            'CBSPackagesPending' {
                $likelySource = 'Component Based Servicing'
                $explanation = 'One or more Windows component packages remain pending.'
            }
            'UpdateExeVolatile' {
                $likelySource = 'Windows Installer / Update installer'
                $explanation = "UpdateExeVolatile remains non-zero. $($flag.Details)"
            }
            { $_ -in @('PendingFileRenameOperations','PendingFileRenameOperations2') } {
                $likelySource = 'Pending file replacement/deletion'

                # PendingFileRenameOperations is typically a sequence of source/destination
                # paths. Preserve the exact paths as evidence and infer the subsystem.
                $rawParts = @(
                    ([string]$flag.Details -split '\s+\|\s+') |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                )

                $affectedItems = @($rawParts)

                $joined = ($rawParts -join ' ')
                if ($joined -match '(?i)\\AppPatch\\|AcPluginDlls') {
                    $likelySource = 'Windows Application Compatibility / AppPatch'
                    $explanation = 'Windows application-compatibility files are queued for rename/deletion at reboot.'
                }
                elseif ($joined -match '(?i)\\DriverStore\\|\\System32\\drivers\\') {
                    $likelySource = 'Driver installation/update'
                    $explanation = 'Driver files are queued for replacement/deletion at reboot.'
                }
                elseif ($joined -match '(?i)\\WinSxS\\|\\servicing\\') {
                    $likelySource = 'Windows Component Servicing'
                    $explanation = 'Windows servicing files are queued for replacement/deletion at reboot.'
                }
                elseif ($joined -match '(?i)\\Program Files\\Elastic\\|Elastic\\Agent') {
                    $likelySource = 'Elastic Agent'
                    $explanation = 'Elastic Agent files are queued for replacement/deletion at reboot.'
                }
                elseif ($joined -match '(?i)\.msi\b|\\Installer\\|\\Temp\\') {
                    $likelySource = 'Application installer'
                    $explanation = 'An application installer left file replacement/deletion operations pending.'
                }
                else {
                    $explanation = 'Windows has file rename/delete operations queued for the next reboot. AffectedItems contains the raw file paths.'
                }
            }
        }

        $results.Add([pscustomobject]@{
            FlagName       = [string]$flag.Name
            FlagType       = [string]$flag.Type
            RegistryPath   = [string]$flag.Path
            RegistryValue  = [string]$flag.ValueName
            LikelySource   = $likelySource
            Explanation    = $explanation
            AffectedItems  = @($affectedItems)
            RawDetails     = [string]$flag.Details
        }) | Out-Null
    }

    return @($results | ForEach-Object { $_ })
}

function Get-PreviousWindowsUpdateContext {
    $path = Join-Path $LogDirectory '06_Weekend_Windows_Updates.latest.json'

    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return $null
    }

    try {
        $obj = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

        return [pscustomobject]@{
            ScriptVersion            = $obj.ScriptVersion
            RunId                    = $obj.RunId
            Status                   = $obj.Status
            EndTime                  = $obj.EndTime
            UpdatesInstalled         = $obj.UpdatesInstalled
            InstalledKBs             = @($obj.InstalledKBs)
            InstalledUpdates         = @($obj.InstalledUpdates | ForEach-Object {
                [pscustomobject]@{
                    Title = $_.Title
                    KB = $_.KB
                    Result = $_.Result
                    Classification = $_.Classification
                }
            })
            PendingRebootAfter       = $obj.PendingRebootAfter
            PendingRebootReasonsAfter = @($obj.PendingRebootReasonsAfter)
        }
    }
    catch {
        Write-Status "Unable to read Script 06 latest update context: $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Get-PendingRebootFlags {
    $flags = @()

    $wuRebootRequired = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    if (Test-RegistryKeyExists -Path $wuRebootRequired) {
        $flags += New-FlagRecord -Name 'WindowsUpdateRebootRequired' -Type 'RegistryKey' -Path $wuRebootRequired -ValueName '' -Details 'Windows Update indicates a reboot is required.'
    }

    $wuPostRebootReporting = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting'
    if (Test-RegistryKeyExists -Path $wuPostRebootReporting) {
        $flags += New-FlagRecord -Name 'WindowsUpdatePostRebootReporting' -Type 'RegistryKey' -Path $wuPostRebootReporting -ValueName '' -Details 'Windows Update post-reboot reporting flag is present.'
    }

    $cbsRebootPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    if (Test-RegistryKeyExists -Path $cbsRebootPending) {
        $flags += New-FlagRecord -Name 'CBSRebootPending' -Type 'RegistryKey' -Path $cbsRebootPending -ValueName '' -Details 'Component Based Servicing reports RebootPending.'
    }

    $cbsRebootInProgress = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'
    if (Test-RegistryKeyExists -Path $cbsRebootInProgress) {
        $flags += New-FlagRecord -Name 'CBSRebootInProgress' -Type 'RegistryKey' -Path $cbsRebootInProgress -ValueName '' -Details 'Component Based Servicing reports RebootInProgress.'
    }

    $cbsPackagesPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    if (Test-RegistryKeyExists -Path $cbsPackagesPending) {
        $flags += New-FlagRecord -Name 'CBSPackagesPending' -Type 'RegistryKey' -Path $cbsPackagesPending -ValueName '' -Details 'Component Based Servicing reports PackagesPending.'
    }

    $sessionManagerPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'

    $pendingFileRenameOperations = Get-RegistryValueSafe -Path $sessionManagerPath -Name 'PendingFileRenameOperations'
    if ($null -ne $pendingFileRenameOperations) {
        $details = if ($pendingFileRenameOperations -is [System.Array]) {
            ($pendingFileRenameOperations | ForEach-Object { [string]$_ }) -join ' | '
        }
        else {
            [string]$pendingFileRenameOperations
        }

        $flags += New-FlagRecord -Name 'PendingFileRenameOperations' -Type 'RegistryValue' -Path $sessionManagerPath -ValueName 'PendingFileRenameOperations' -Details $details
    }

    $pendingFileRenameOperations2 = Get-RegistryValueSafe -Path $sessionManagerPath -Name 'PendingFileRenameOperations2'
    if ($null -ne $pendingFileRenameOperations2) {
        $details = if ($pendingFileRenameOperations2 -is [System.Array]) {
            ($pendingFileRenameOperations2 | ForEach-Object { [string]$_ }) -join ' | '
        }
        else {
            [string]$pendingFileRenameOperations2
        }

        $flags += New-FlagRecord -Name 'PendingFileRenameOperations2' -Type 'RegistryValue' -Path $sessionManagerPath -ValueName 'PendingFileRenameOperations2' -Details $details
    }

    $updatesPath = 'HKLM:\SOFTWARE\Microsoft\Updates'
    $updateExeVolatile = Get-RegistryValueSafe -Path $updatesPath -Name 'UpdateExeVolatile'
    if ($null -ne $updateExeVolatile) {
        try {
            if ([int]$updateExeVolatile -ne 0) {
                $flags += New-FlagRecord -Name 'UpdateExeVolatile' -Type 'RegistryValue' -Path $updatesPath -ValueName 'UpdateExeVolatile' -Details "Value is $updateExeVolatile"
            }
        }
        catch {
            $flags += New-FlagRecord -Name 'UpdateExeVolatile' -Type 'RegistryValue' -Path $updatesPath -ValueName 'UpdateExeVolatile' -Details "Non-integer value detected: $updateExeVolatile"
        }
    }

    $script:CurrentFlags = @($flags)
    $script:RebootCauseDetails = @(Get-RebootCauseDetails -Flags $flags)
    return @($flags)
}

function New-DefaultState {
    [PSCustomObject]@{
        Stage     = 0
        FirstSeen = $null
        LastRun      = $null
        LastBootTime = $null
        LastFlags    = @()
    }
}

function ConvertTo-StateObject {
    param(
        [Parameter(Mandatory)]
        $InputObject
    )

    $stage = 0
    try {
        $stage = [int]($InputObject.Stage)
    }
    catch {
        $stage = 0
    }

    if ($stage -lt 0 -or $stage -gt 3) {
        $stage = 0
    }

    [PSCustomObject]@{
        Stage     = $stage
        FirstSeen = $InputObject.FirstSeen
        LastRun      = $InputObject.LastRun
        LastBootTime = $InputObject.LastBootTime
        LastFlags    = @($InputObject.LastFlags)
    }
}

function Backup-BadStateFile {
    param(
        [string]$Reason
    )

    try {
        if (Test-Path -LiteralPath $script:StateFilePath) {
            $backupPath = "$($script:StateFilePath).bad-$((Get-Date).ToString('yyyyMMdd-HHmmss'))"
            Copy-Item -LiteralPath $script:StateFilePath -Destination $backupPath -Force -ErrorAction Stop
            Write-Status "Backed up unreadable reboot state file to: $backupPath" 'WARN'
        }
    }
    catch {
        Write-Status "Could not back up unreadable reboot state file. $($_.Exception.Message)" 'WARN'
    }

    Write-Status "State file reset reason: $Reason" 'WARN'
}

function Test-RebootCycleActive {
    if (-not (Test-Path -LiteralPath $script:StateFilePath -PathType Leaf)) {
        return $false
    }

    try {
        $state = Get-Content -LiteralPath $script:StateFilePath -Raw -Encoding UTF8 -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop

        $stage = [int]$state.Stage
        return ($stage -ge 1 -and $stage -le 3)
    }
    catch {
        # A corrupt state file must not cause a normal startup task to begin
        # a brand-new reboot cycle. The scheduled Sunday run can handle recovery.
        return $false
    }
}

function Get-State {
    if (-not (Test-Path -LiteralPath $script:StateFilePath)) {
        return New-DefaultState
    }

    try {
        $raw = Get-Content -LiteralPath $script:StateFilePath -Raw -Encoding UTF8 -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($raw)) {
            Backup-BadStateFile -Reason 'State file was empty.'
            return New-DefaultState
        }

        $trimmed = $raw.Trim()

        try {
            $obj = $trimmed | ConvertFrom-Json -ErrorAction Stop
            return (ConvertTo-StateObject -InputObject $obj)
        }
        catch {
            # Recovery path: some systems have produced a state file missing the opening brace,
            # for example: \"Stage\": 1, ... }. Wrap it and retry instead of forcing a false first pass.
            if (($trimmed -notmatch '^\s*\{') -and ($trimmed -match '\"Stage\"\s*:') -and ($trimmed -match '\}\s*$')) {
                try {
                    $repairedJson = '{' + $trimmed
                    $obj = $repairedJson | ConvertFrom-Json -ErrorAction Stop
                    $script:StateRecovered = $true
                    Write-Status 'State file was missing the opening JSON brace. Repaired in memory and continuing staged reboot logic.' 'WARN'
                    Save-State -Stage ([int]$obj.Stage) -FirstSeen $obj.FirstSeen -LastRun $obj.LastRun -LastBootTime $obj.LastBootTime -LastFlags @($obj.LastFlags)
                    return (ConvertTo-StateObject -InputObject $obj)
                }
                catch {
                    Backup-BadStateFile -Reason "Automatic JSON brace repair failed. $($_.Exception.Message)"
                    return New-DefaultState
                }
            }

            Backup-BadStateFile -Reason "State file unreadable. $($_.Exception.Message)"
            return New-DefaultState
        }
    }
    catch {
        Backup-BadStateFile -Reason "State file could not be read. $($_.Exception.Message)"
        return New-DefaultState
    }
}

function Save-State {
    param(
        [int]$Stage,
        $FirstSeen,
        $LastRun,
        $LastBootTime,
        $LastFlags
    )

    Ensure-Folder -Path $StateDirectory

    $state = [PSCustomObject]@{
        Stage     = $Stage
        FirstSeen = if ($null -ne $FirstSeen) { [string]$FirstSeen } else { $null }
        LastRun      = if ($null -ne $LastRun) { [string]$LastRun } else { $null }
        LastBootTime = if ($null -ne $LastBootTime) { [string]$LastBootTime } else { $null }
        LastFlags    = @($LastFlags)
    }

    $json = $state | ConvertTo-Json -Depth 12
    if ([string]::IsNullOrWhiteSpace($json) -or $json.TrimStart()[0] -ne '{') {
        throw 'Generated state JSON did not begin with an opening brace. Refusing to write invalid state file.'
    }

    $unique = [guid]::NewGuid().ToString('N')
    $tempPath = Join-Path $StateDirectory ("$StateFileName.$unique.tmp")

    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $fileStream = [System.IO.File]::Open($tempPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $writer = New-Object System.IO.StreamWriter($fileStream, $utf8NoBom)
            try {
                $writer.Write($json)
                $writer.Flush()
                $fileStream.Flush($true)
            }
            finally {
                $writer.Dispose()
            }
        }
        finally {
            $fileStream.Dispose()
        }

        # Validate the temp file before replacing the active state file.
        $validated = (Get-Content -LiteralPath $tempPath -Raw -Encoding UTF8 -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $validated -or $null -eq $validated.Stage) {
            throw 'Temporary state file validation failed because required Stage property was missing.'
        }

        if (Test-Path -LiteralPath $script:StateFilePath) {
            Copy-Item -LiteralPath $script:StateFilePath -Destination "$($script:StateFilePath).prev" -Force -ErrorAction SilentlyContinue
        }

        Move-Item -LiteralPath $tempPath -Destination $script:StateFilePath -Force -ErrorAction Stop
        Confirm-FileWriteToDisk -Path $script:StateFilePath
        Write-Status "Saved reboot state tracking at stage $Stage." 'INFO'
    }
    catch {
        Write-Status "Failed to save reboot state file safely. Error: $($_.Exception.Message)" 'ERROR'
        try {
            if (Test-Path -LiteralPath $tempPath) {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }
        throw
    }
}

function Reset-State {
    try {
        if (Test-Path -LiteralPath $script:StateFilePath) {
            Remove-Item -LiteralPath $script:StateFilePath -Force -ErrorAction Stop
            Write-Status "Reset reboot state tracking." 'OK'
        }
    }
    catch {
        Write-Status "Failed to remove state file: $($_.Exception.Message)" 'WARN'
    }
}

function Remove-RegistryKeySafe {
    param(
        [string]$Name,
        [string]$Path
    )

    try {
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            $script:ClearResults += New-ClearResult -Name $Name -Path $Path -Action 'RemoveKey' -Status 'Succeeded' -Message 'Registry key removed.'
            Write-Status "Removed key for [$Name]: $Path" 'OK'
        }
        else {
            $script:ClearResults += New-ClearResult -Name $Name -Path $Path -Action 'RemoveKey' -Status 'Skipped' -Message 'Registry key not present.'
            Write-Status "Key already absent for [$Name]: $Path" 'INFO'
        }
    }
    catch {
        $script:ClearResults += New-ClearResult -Name $Name -Path $Path -Action 'RemoveKey' -Status 'Failed' -Message $_.Exception.Message
        Write-Status "Failed removing key for [$Name]: $($_.Exception.Message)" 'ERROR'
    }
}

function Remove-RegistryValueSafe {
    param(
        [string]$Name,
        [string]$Path,
        [string]$ValueName
    )

    try {
        if (Test-Path -LiteralPath $Path) {
            $currentValue = Get-RegistryValueSafe -Path $Path -Name $ValueName
            if ($null -ne $currentValue) {
                Remove-ItemProperty -LiteralPath $Path -Name $ValueName -ErrorAction Stop
                $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Succeeded' -Message 'Registry value removed.'
                Write-Status "Removed value for [$Name]: $Path\$ValueName" 'OK'
            }
            else {
                $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Skipped' -Message 'Registry value not present.'
                Write-Status "Value already absent for [$Name]: $Path\$ValueName" 'INFO'
            }
        }
        else {
            $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Skipped' -Message 'Registry path not present.'
            Write-Status "Path absent for [$Name]: $Path" 'INFO'
        }
    }
    catch {
        $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Failed' -Message $_.Exception.Message
        Write-Status "Failed removing value for [$Name]: $($_.Exception.Message)" 'ERROR'
    }
}

function Set-RegistryValueSafe {
    param(
        [string]$Name,
        [string]$Path,
        [string]$ValueName,
        $Value,
        [string]$Type
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }

        New-ItemProperty -LiteralPath $Path -Name $ValueName -PropertyType $Type -Value $Value -Force | Out-Null
        $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'SetValue' -Status 'Succeeded' -Message "Set to $Value."
        Write-Status "Set [$Name] value: $Path\$ValueName = $Value" 'OK'
    }
    catch {
        $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'SetValue' -Status 'Failed' -Message $_.Exception.Message
        Write-Status "Failed setting value for [$Name]: $($_.Exception.Message)" 'ERROR'
    }
}

function Clear-PendingRebootFlags {
    $script:ClearResults = @()
    Write-Status "Attempting to clear persistent reboot flags..." 'WARN'

    Remove-RegistryKeySafe   -Name 'WindowsUpdateRebootRequired'      -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    Remove-RegistryKeySafe   -Name 'WindowsUpdatePostRebootReporting' -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting'
    Remove-RegistryKeySafe   -Name 'CBSRebootPending'                 -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    Remove-RegistryKeySafe   -Name 'CBSRebootInProgress'              -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'
    Remove-RegistryKeySafe   -Name 'CBSPackagesPending'               -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    Remove-RegistryValueSafe -Name 'PendingFileRenameOperations'      -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ValueName 'PendingFileRenameOperations'
    Remove-RegistryValueSafe -Name 'PendingFileRenameOperations2'     -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ValueName 'PendingFileRenameOperations2'
    Set-RegistryValueSafe    -Name 'UpdateExeVolatile'                -Path 'HKLM:\SOFTWARE\Microsoft\Updates' -ValueName 'UpdateExeVolatile' -Value 0 -Type 'DWord'

    return @($script:ClearResults)
}

function Test-AppPatchPendingRenameFlag {
    param(
        [array]$Flags
    )

    foreach ($flag in @($Flags)) {
        if ($flag.Name -in @('PendingFileRenameOperations','PendingFileRenameOperations2') -and $flag.Details -match '(?i)apppatch\\AcPluginDlls') {
            return $true
        }
    }

    return $false
}

function Stop-UpdateLockingProcessesSafe {
    $lockingProcesses = @(
        'TiWorker',
        'TrustedInstaller',
        'wuauclt',
        'UsoClient'
    )

    foreach ($processName in $lockingProcesses) {
        try {
            $processes = @(Get-Process -Name $processName -ErrorAction SilentlyContinue)
            foreach ($process in $processes) {
                try {
                    Write-Status "Stopping possible update-locking process: $($process.Name) PID $($process.Id)" 'WARN'
                    Stop-Process -Id $process.Id -Force -ErrorAction Stop
                }
                catch {
                    Write-Status "Could not stop process $($process.Name) PID $($process.Id): $($_.Exception.Message)" 'WARN'
                }
            }
        }
        catch {
            Write-Status "Process check failed for [$processName]: $($_.Exception.Message)" 'WARN'
        }
    }
}

function Resolve-PersistentRebootFlags {
    param(
        [array]$InitialFlags
    )

    if (-not $AllowUnsafeRebootFlagCleanup) {
        Write-Status 'Reboot flags remain after three reboot attempts. Registry servicing flags will not be deleted automatically.' 'WARN'
        foreach ($flag in @($InitialFlags)) {
            Write-Status "Persistent flag retained: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'WARN'
        }

        if ($script:PersistentAfterThirdReboot) {
            $script:OverallResult = 'PersistentFlagsAfterThirdReboot'
        }
        else {
            $script:OverallResult = 'PersistentFlagsRemain'
        }
        $script:ExitCode = 2
        Reset-State
        Write-YamlLog
        exit 2
    }

    Write-Status 'Unsafe reboot-flag cleanup was explicitly enabled. Attempting legacy registry cleanup.' 'WARN'

    if (Test-AppPatchPendingRenameFlag -Flags $InitialFlags) {
        Write-Status 'Detected AppPatch/AcPluginDlls PendingFileRenameOperations.' 'WARN'
    }

    Stop-UpdateLockingProcessesSafe
    [void](Clear-PendingRebootFlags)
    Start-Sleep -Seconds 2

    $remainingFlags = @(Get-PendingRebootFlags)
    if ($remainingFlags.Count -gt 0) {
        Write-Status 'Reboot flags are still present after opt-in cleanup.' 'WARN'
        foreach ($flag in $remainingFlags) {
            Write-Status "Remaining flag: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'WARN'
        }
        $script:OverallResult = 'PersistentFlagsRemainAfterCleanup'
        $script:ExitCode = 2
    }
    else {
        Write-Status 'Persistent reboot flags cleared after explicit opt-in cleanup.' 'OK'
        $script:OverallResult = 'PersistentFlagsCleared'
        $script:ExitCode = 0
    }

    Reset-State
    Write-YamlLog
    exit $script:ExitCode
}


function Write-YamlLog {
    try {
        $runEnd = Get-Date
        $duration = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 0)
        $flags = @($script:CurrentFlags)
        $clear = @($script:ClearResults)
        $actions = @($script:ActionHistory)

        $lines = New-Object System.Collections.Generic.List[string]

        $lines.Add("computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("script_name: '07_Force_Reboot_Install_Updates.ps1'") | Out-Null
        $lines.Add("script_version: '$($script:ScriptVersion)'") | Out-Null
        $lines.Add("run_id: $(ConvertTo-YamlScalar $script:RunId)") | Out-Null
        $lines.Add("domain: $(ConvertTo-YamlScalar $script:Domain)") | Out-Null
        $lines.Add("status: $(ConvertTo-YamlScalar (Get-ExecutionStatus))") | Out-Null
        $lines.Add("exit_code: $($script:ExitCode)") | Out-Null
        $lines.Add("run_started: $(ConvertTo-YamlScalar $script:RunStart)") | Out-Null
        $lines.Add("run_finished: $(ConvertTo-YamlScalar $runEnd)") | Out-Null
        $lines.Add("duration_seconds: $duration") | Out-Null
        $lines.Add("stage: $($script:CurrentStage)") | Out-Null
        $lines.Add("reboot_issued: $(ConvertTo-YamlScalar $script:RebootIssued)") | Out-Null
        $lines.Add("reboot_reason: $(ConvertTo-YamlScalar $script:RebootReason)") | Out-Null
        $lines.Add("overall_result: $(ConvertTo-YamlScalar $script:OverallResult)") | Out-Null
        $lines.Add("failure_message: $(ConvertTo-YamlScalar $script:FailureMessage)") | Out-Null
        $lines.Add("final_verification_after_third_reboot: $(ConvertTo-YamlScalar $script:FinalVerificationAfterThirdReboot)") | Out-Null
        $lines.Add("persistent_after_third_reboot: $(ConvertTo-YamlScalar $script:PersistentAfterThirdReboot)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('flags_detected:') | Out-Null
        if ($flags.Count -gt 0) {
            foreach ($flag in $flags) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    name: $(ConvertTo-YamlScalar $flag.Name)") | Out-Null
                $lines.Add("    type: $(ConvertTo-YamlScalar $flag.Type)") | Out-Null
                $lines.Add("    path: $(ConvertTo-YamlScalar $flag.Path)") | Out-Null
                $lines.Add("    value_name: $(ConvertTo-YamlScalar $flag.ValueName)") | Out-Null
                $lines.Add("    details: $(ConvertTo-YamlScalar $flag.Details)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        $lines.Add('') | Out-Null
        $lines.Add('clear_actions:') | Out-Null
        if ($clear.Count -gt 0) {
            foreach ($item in $clear) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    name: $(ConvertTo-YamlScalar $item.Name)") | Out-Null
                $lines.Add("    path: $(ConvertTo-YamlScalar $item.Path)") | Out-Null
                $lines.Add("    action: $(ConvertTo-YamlScalar $item.Action)") | Out-Null
                $lines.Add("    status: $(ConvertTo-YamlScalar $item.Status)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $item.Message)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        $lines.Add('') | Out-Null
        $lines.Add('actions:') | Out-Null
        if ($actions.Count -gt 0) {
            foreach ($action in $actions) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    time: $(ConvertTo-YamlScalar $action.Time)") | Out-Null
                $lines.Add("    level: $(ConvertTo-YamlScalar $action.Level)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $action.Message)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        $yamlTempPath = "$($script:YamlLogPath).tmp"
        Set-Content -LiteralPath $yamlTempPath -Value @($lines | ForEach-Object { $_ }) -Encoding UTF8 -Force
        Move-Item -LiteralPath $yamlTempPath -Destination $script:YamlLogPath -Force
        Confirm-FileWriteToDisk -Path $script:YamlLogPath
        Write-ExecutionTelemetry
    }
    catch {
        Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
    }
}

function Invoke-ForcedReboot {
    param(
        [Parameter(Mandatory)]
        [string]$Reason
    )

    $script:RebootIssued = $true
    $script:RebootReason = $Reason
    $script:ExitCode = 3010

    Write-Status "Issuing forced reboot in $RebootDelaySeconds seconds. Reason: $Reason" 'WARN'
    Write-YamlLog
    Confirm-StateWriteBeforeReboot

    $arguments = @(
        '/r'
        '/f'
        '/t'
        [string]$RebootDelaySeconds
        '/c'
        $Reason
    )

    & shutdown.exe @arguments | Out-Null

    if ($LASTEXITCODE -ne 0) {
        throw "shutdown.exe returned exit code $LASTEXITCODE"
    }

    exit 3010
}

# Main
# The AtStartup task always supplies -StartupResume. It exists only to continue
# an already-active Script 07 reboot sequence. It must never start a new sequence
# during an ordinary workstation boot.
if ($StartupResume -and -not (Test-RebootCycleActive)) {
    Write-Host "Script 07 startup resume: no active reboot cycle. Exiting without action." -ForegroundColor DarkGray
    exit 0
}

Initialize-Paths
Write-Status "Starting $($script:ScriptName) version $($script:ScriptVersion)." 'INFO'
Write-Status "Active staged text log path: $($script:TextLogPath)" 'INFO'
Write-Status "Completed text log publish path: $($script:PublishedLogPath)" 'INFO'
Initialize-SingleInstanceLock

try {
    if (-not (Test-IsAdministrator)) {
        $script:FailureMessage = 'Administrator privileges are required.'
        $script:OverallResult = 'Failed'
        $script:ExitCode = 1
        Write-Status $script:FailureMessage 'ERROR'
        Write-YamlLog
        exit 1
    }

    $script:BootSnapshot = Get-BootSnapshot
    $script:LoggedOnUsers = @(Get-LoggedOnUserSnapshot)
    $script:UpdateSnapshot = Get-WindowsUpdateSnapshot

    Write-Status "Starting reboot flag evaluation..." 'INFO'

    try {
        $state = Get-State
        $script:PreviousBootTime = $state.LastBootTime
        if ($state.LastBootTime -and $script:BootSnapshot.LastBootTime) {
            try {
                $script:RebootVerified = ([datetime]$script:BootSnapshot.LastBootTime -gt [datetime]$state.LastBootTime)
            }
            catch {
                $script:RebootVerified = $false
            }
        }
        $flags = @(Get-PendingRebootFlags)

        $stage = 0
        try {
            $stage = [int]$state.Stage
        }
        catch {
            $stage = 0
        }

        $script:CurrentStage = $stage

        switch ($stage) {
            0 {
                Write-Status "First pass detected. Forcing reboot regardless of flag state." 'WARN'
                Save-State -Stage 1 -FirstSeen ((Get-Date).ToString('o')) -LastRun ((Get-Date).ToString('o')) -LastBootTime $script:BootSnapshot.LastBootTime -LastFlags $flags
                $script:OverallResult = 'ForcedInitialReboot'
                Invoke-ForcedReboot -Reason 'Initial forced reboot for update cycle.'
            }

            1 {
                if ($flags.Count -gt 0) {
                    Write-Status "Second pass: reboot flags still detected. Issuing second reboot." 'WARN'
                    foreach ($flag in $flags) {
                        Write-Status "Flag detected: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'WARN'
                    }

                    Save-State -Stage 2 -FirstSeen $state.FirstSeen -LastRun ((Get-Date).ToString('o')) -LastBootTime $script:BootSnapshot.LastBootTime -LastFlags $flags
                    $script:OverallResult = 'SecondPassRebootIssued'
                    Invoke-ForcedReboot -Reason 'Reboot flags remain after initial reboot. Issuing second reboot.'
                }
                else {
                    Write-Status "Second pass: no reboot flags detected. Resetting state and exiting normally." 'OK'
                    Reset-State
                    $script:OverallResult = 'SecondPassNoFlags'
                    Write-YamlLog
                    exit 0
                }
            }

            2 {
                if ($flags.Count -gt 0) {
                    Write-Status "Third pass: reboot flags still detected after two reboots. Issuing the third and final automatic reboot." 'WARN'
                    foreach ($flag in $flags) {
                        Write-Status "Flag before third reboot: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'WARN'
                    }

                    Save-State -Stage 3 -FirstSeen $state.FirstSeen -LastRun ((Get-Date).ToString('o')) -LastBootTime $script:BootSnapshot.LastBootTime -LastFlags $flags
                    $script:OverallResult = 'ThirdRebootIssued'
                    Invoke-ForcedReboot -Reason 'Reboot flags remain after two reboots. Issuing third and final automatic reboot.'
                }
                else {
                    Write-Status "Third pass: no reboot flags detected after two reboots. Resetting state and exiting normally." 'OK'
                    Reset-State
                    $script:OverallResult = 'ThirdPassNoFlags'
                    Write-YamlLog
                    exit 0
                }
            }

            3 {
                # This is the important safety gate before the workstation can be
                # considered ready to re-freeze. No fourth automatic reboot occurs.
                $script:FinalVerificationAfterThirdReboot = $true

                if ($flags.Count -gt 0) {
                    $script:PersistentAfterThirdReboot = $true
                    Write-Status "FINAL CHECK: reboot flags are still present after the third reboot. System is NOT ready to re-freeze." 'ERROR'

                    foreach ($cause in @($script:RebootCauseDetails)) {
                        Write-Status ("Persistent reboot cause: Flag={0}; LikelySource={1}; Explanation={2}; AffectedItems={3}" -f `
                            $cause.FlagName,
                            $cause.LikelySource,
                            $cause.Explanation,
                            (@($cause.AffectedItems) -join ' | ')) 'ERROR'
                    }

                    $script:OverallResult = 'PersistentFlagsAfterThirdReboot'
                    $script:ExitCode = 2

                    # Preserve the existing opt-in cleanup behavior, but do not
                    # automatically delete servicing state unless explicitly enabled.
                    Resolve-PersistentRebootFlags -InitialFlags $flags
                }
                else {
                    Write-Status "FINAL CHECK: no reboot flags remain after the third reboot. System is safe to proceed to re-freeze." 'OK'
                    $script:PersistentAfterThirdReboot = $false
                    Reset-State
                    $script:OverallResult = 'ThirdRebootVerifiedClean'
                    $script:ExitCode = 0
                    Write-YamlLog
                    exit 0
                }
            }

            default {
                Write-Status "Unexpected state value [$stage]. Resetting state and starting over with forced initial reboot." 'WARN'
                Reset-State
                Save-State -Stage 1 -FirstSeen ((Get-Date).ToString('o')) -LastRun ((Get-Date).ToString('o')) -LastBootTime $script:BootSnapshot.LastBootTime -LastFlags $flags
                $script:CurrentStage = 0
                $script:OverallResult = 'ForcedInitialRebootAfterStateReset'
                Invoke-ForcedReboot -Reason 'State reset occurred. Performing initial forced reboot for update cycle.'
            }
        }
    }
    catch {
        $script:FailureMessage = $_.Exception.Message
        $script:OverallResult = 'Failed'
        $script:ExitCode = 3
        Write-Status "Script failed: $($_.Exception.Message)" 'ERROR'
        Write-YamlLog
        exit 3
    }
}
finally {
    # Release-SingleInstanceLock writes a status entry, so it must happen before
    # the staged log is published into C:\Logs.
    Release-SingleInstanceLock
    Complete-TextLogPublication -Reason 'Script exit or reboot handoff'
}
