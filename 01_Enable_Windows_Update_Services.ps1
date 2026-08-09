# =====================================================================
# ScriptName: 01_Enable_Windows_Update_Services.ps1
# ScriptVersion: 2.6.0
# LastUpdated: 2026-07-27
# Purpose: Restore Windows Update services, tasks, policy settings,
#          Windows 11 UI preferences, and classic right-click context menu behavior for all users;
#          verify required services are running, retry startup failures
#          up to 4 total attempts, and force a reboot if critical
#          services still refuse to start.
# Fix:      Removed StartWhenAvailable from task creation and added direct
#          enforcement to clear missed-run behavior from maintenance tasks.
# =====================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$script:ScriptName       = '01_Enable_Windows_Update_Services.ps1'
$script:ScriptVersion    = '2.6.0'
$script:ExecutionStart   = Get-Date
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:ExitCode         = 0
$script:FinalStatus      = 'Running'
$script:RebootInitiated  = $false
$script:RebootReason     = $null
$script:TaskReconcileStatus = 'NotRun'
$script:ServiceNames     = @('wuauserv','bits','dosvc','UsoSvc','WaaSMedicSvc','CryptSvc','msiserver')
$script:ServiceStateBefore = @()
$script:ServiceStateAfter  = @()
$script:PolicyConflictsBefore = @()
$script:TelemetryDirectory = 'C:\Logs'
$script:NdjsonPath = Join-Path $script:TelemetryDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestJsonPath = Join-Path $script:TelemetryDirectory '01_Enable_Windows_Update_Services.latest.json'


# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }
    switch ($Level) {
        'INFO'  { Write-Host "[$timestamp] [INFO ] $Message" -ForegroundColor Cyan }
        'OK'    { Write-Host "[$timestamp] [ OK  ] $Message" -ForegroundColor Green }
        'WARN'  { Write-Host "[$timestamp] [WARN ] $Message" -ForegroundColor Yellow }
        'ERROR' { Write-Host "[$timestamp] [ERROR] $Message" -ForegroundColor Red }
    }
}


function Get-FileSha256Safe {
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
        }
    }
    catch {
    }

    return $null
}

function Invoke-UpdaterBootstrap {
    [CmdletBinding()]
    param()

    $preferredUpdater = '\\filesvr\Labscripts\00_Update-Scripts-FromShare.ps1'
    $fallbackUpdater  = '\\10.2.3.30\Labscripts\00_Update-Scripts-FromShare.ps1'
    $localUpdater     = 'C:\Scripts\00_Update-Scripts-FromShare.ps1'
    $powershellExe    = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    $remoteUpdater = @($preferredUpdater, $fallbackUpdater) |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1

    if (-not $remoteUpdater) {
        Write-Status 'Updater bootstrap skipped because neither network source is available.' 'WARN'
        return
    }

    try {
        $remoteHash = Get-FileSha256Safe -Path $remoteUpdater
        $localHash = Get-FileSha256Safe -Path $localUpdater
        $updaterChanged = $remoteHash -ne $localHash -or $null -eq $localHash

        if ($updaterChanged) {
            Write-Status "A newer or different 00 updater was found: $remoteUpdater" 'WARN'

            $localDirectory = Split-Path -Path $localUpdater -Parent
            if (-not (Test-Path -LiteralPath $localDirectory -PathType Container)) {
                New-Item -Path $localDirectory -ItemType Directory -Force | Out-Null
            }

            $temporaryUpdater = "$localUpdater.new"
            Copy-Item -LiteralPath $remoteUpdater -Destination $temporaryUpdater -Force
            Move-Item -LiteralPath $temporaryUpdater -Destination $localUpdater -Force

            if ((Get-FileSha256Safe -Path $localUpdater) -ne $remoteHash) {
                throw 'The copied updater failed SHA-256 verification.'
            }

            Write-Status 'The local 00 updater was refreshed successfully.' 'OK'
        }
        else {
            Write-Status 'The local 00 updater is already current.' 'OK'
        }

        # Run 00 now so any newly introduced scripts and task definitions are
        # deployed during this same maintenance window.
        Write-Status 'Running the current 00 updater to synchronize scripts and tasks.' 'INFO'

        $process = Start-Process `
            -FilePath $powershellExe `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$localUpdater`"" `
            -Wait `
            -PassThru `
            -WindowStyle Hidden

        if ($process.ExitCode -eq 0) {
            Write-Status 'Updater synchronization completed successfully.' 'OK'
        }
        else {
            Write-Status "Updater synchronization returned exit code $($process.ExitCode)." 'WARN'
        }
    }
    catch {
        Write-Status "Updater bootstrap failed: $($_.Exception.Message)" 'WARN'
    }
}


function Get-ServiceTelemetrySnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Names)

    $results = foreach ($name in $Names) {
        try {
            $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction Stop
            [pscustomobject]@{
                Name        = $name
                Exists      = $true
                State       = [string]$service.State
                StartMode   = [string]$service.StartMode
                ProcessId   = [int]$service.ProcessId
            }
        }
        catch {
            [pscustomobject]@{
                Name        = $name
                Exists      = $false
                State       = $null
                StartMode   = $null
                ProcessId   = 0
            }
        }
    }

    return @($results)
}

function Get-WindowsUpdatePolicyConflicts {
    [CmdletBinding()]
    param()

    $conflicts = New-Object System.Collections.Generic.List[string]
    $wuPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'

    foreach ($item in @(
        @{ Path = $wuPath; Name = 'WUServer' },
        @{ Path = $wuPath; Name = 'WUStatusServer' },
        @{ Path = $wuPath; Name = 'UpdateServiceUrlAlternate' },
        @{ Path = $wuPath; Name = 'DisableWindowsUpdateAccess' },
        @{ Path = $auPath; Name = 'UseWUServer' }
    )) {
        try {
            $value = (Get-ItemProperty -Path $item.Path -Name $item.Name -ErrorAction Stop).$($item.Name)
            if ($null -ne $value -and $value -notin @(0, '')) {
                $conflicts.Add("$($item.Path)\\$($item.Name)=$value")
            }
        }
        catch { }
    }

    try {
        $noAutoUpdate = (Get-ItemProperty -Path $auPath -Name 'NoAutoUpdate' -ErrorAction Stop).NoAutoUpdate
        if ([int]$noAutoUpdate -ne 0) { $conflicts.Add("$auPath\\NoAutoUpdate=$noAutoUpdate") }
    }
    catch { }

    return @($conflicts)
}

function Write-ExecutionTelemetry {
    [CmdletBinding()]
    param()

    try {
        if (-not (Test-Path -LiteralPath $script:TelemetryDirectory -PathType Container)) {
            New-Item -Path $script:TelemetryDirectory -ItemType Directory -Force | Out-Null
        }

        $end = Get-Date
        $duration = [math]::Round(($end - $script:ExecutionStart).TotalSeconds, 3)
        $afterByName = @{}
        foreach ($service in $script:ServiceStateAfter) { $afterByName[$service.Name] = $service }

        $serviceChanges = foreach ($before in $script:ServiceStateBefore) {
            $after = $afterByName[$before.Name]
            if ($null -ne $after -and ($before.State -ne $after.State -or $before.StartMode -ne $after.StartMode)) {
                [pscustomobject]@{
                    Name            = $before.Name
                    StateBefore     = $before.State
                    StateAfter      = $after.State
                    StartModeBefore = $before.StartMode
                    StartModeAfter  = $after.StartMode
                }
            }
        }

        $required = @('wuauserv','bits','dosvc','UsoSvc')
        $servicesFailed = @($script:ServiceStateAfter | Where-Object { $_.Name -in $required -and $_.State -ne 'Running' } | ForEach-Object Name)
        $servicesRunning = @($script:ServiceStateAfter | Where-Object State -eq 'Running' | ForEach-Object Name)
        $policyConflictsAfter = @(Get-WindowsUpdatePolicyConflicts)

        $domain = $env:USERDNSDOMAIN
        if ([string]::IsNullOrWhiteSpace($domain)) {
            try { $domain = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Domain } catch { $domain = $null }
        }

        $event = [ordered]@{
            EventType                 = 'maintenance.execution'
            ComputerName              = $env:COMPUTERNAME
            Domain                    = $domain
            ScriptName                = $script:ScriptName
            ScriptVersion             = $script:ScriptVersion
            Status                    = $script:FinalStatus
            ExitCode                  = $script:ExitCode
            StartTime                 = $script:ExecutionStart.ToString('o')
            EndTime                   = $end.ToString('o')
            DurationSeconds           = $duration
            ErrorCount                = $script:ErrorCount
            WarningCount              = $script:WarningCount
            Timestamp                 = (Get-Date).ToUniversalTime().ToString('o')
            RebootInitiated           = $script:RebootInitiated
            RebootReason              = $script:RebootReason
            TaskReconciliationStatus  = $script:TaskReconcileStatus
            PolicyConflictDetected    = ($script:PolicyConflictsBefore.Count -gt 0)
            PolicyConflictsBefore     = @($script:PolicyConflictsBefore)
            PolicyConflictsAfter      = $policyConflictsAfter
            ServicesChecked           = $script:ServiceNames.Count
            ServicesRunningAfter      = $servicesRunning
            ServicesFailed            = $servicesFailed
            ServicesRepairedCount     = @($serviceChanges).Count
            ServiceChanges            = @($serviceChanges)
            ServiceStateBefore        = @($script:ServiceStateBefore)
            ServiceStateAfter         = @($script:ServiceStateAfter)
            ApprovedServiceState      = ($servicesFailed.Count -eq 0)
        }

        $jsonCompact = $event | ConvertTo-Json -Depth 8 -Compress
        Write-MaintenanceTelemetryLine -Path $script:NdjsonPath -JsonLine $jsonCompact
        $event | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:LatestJsonPath -Encoding UTF8
    }
    catch {
        Write-Host "Telemetry write failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-ServiceStartRegistry {
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName,

        [Parameter(Mandatory)]
        [int]$StartValue
    )

    $paths = @(
        "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName",
        "HKLM:\SYSTEM\ControlSet001\Services\$ServiceName"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            try {
                Set-ItemProperty -Path $path -Name Start -Value $StartValue -Type DWord -ErrorAction Stop
                Write-Status "Set registry Start=$StartValue for $ServiceName at $path" 'OK'
            }
            catch {
                Write-Status "Failed setting Start for $ServiceName at $path : $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Status "Registry path not found for $ServiceName at $path" 'WARN'
        }
    }
}

function Set-ServiceStartupAndStart {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Automatic','Manual')]
        [string]$StartupType
    )

    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop

        try {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction Stop
            Write-Status "Set startup type for $Name to $StartupType" 'OK'
        }
        catch {
            Write-Status "Set-Service failed for $Name. Trying sc.exe config..." 'WARN'
            $startValue = if ($StartupType -eq 'Automatic') { 'auto' } else { 'demand' }
            & sc.exe config $Name start= $startValue | Out-Null
            Write-Status "Configured startup type for $Name via sc.exe" 'OK'
        }

        try {
            Start-Service -Name $Name -ErrorAction Stop
            Write-Status "Started service: $Name" 'OK'
        }
        catch {
            Write-Status "Could not start service $Name immediately: $($_.Exception.Message)" 'WARN'
        }
    }
    catch {
        Write-Status "Service not found or inaccessible: $Name" 'WARN'
    }
}

function Get-ServiceStateSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        return (Get-Service -Name $Name -ErrorAction Stop).Status
    }
    catch {
        return $null
    }
}

function Wait-ForServiceRunning {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [int]$TimeoutSeconds = 15
    )

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($stopWatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $status = Get-ServiceStateSafe -Name $Name
        if ($status -eq 'Running') {
            return $true
        }

        Start-Sleep -Seconds 2
    }

    return $false
}

function Ensure-ServiceRunningWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Automatic','Manual')]
        [string]$StartupType,

        [int]$MaxAttempts = 4,

        [int]$WaitPerAttemptSeconds = 15
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $currentState = Get-ServiceStateSafe -Name $Name

        if ($currentState -eq 'Running') {
            Write-Status "Service $Name is already running." 'OK'
            return $true
        }

        Write-Status "Attempt $attempt of $MaxAttempts to start service $Name..." 'INFO'

        try {
            Set-ServiceStartupAndStart -Name $Name -StartupType $StartupType
        }
        catch {
            Write-Status "Unexpected error while attempting to start $Name : $($_.Exception.Message)" 'WARN'
        }

        if (Wait-ForServiceRunning -Name $Name -TimeoutSeconds $WaitPerAttemptSeconds) {
            Write-Status "Verified service is running: $Name" 'OK'
            return $true
        }

        $stateAfterWait = Get-ServiceStateSafe -Name $Name
        Write-Status "Service $Name did not reach Running state after attempt $attempt. Current state: $stateAfterWait" 'WARN'

        if ($attempt -lt $MaxAttempts) {
            Start-Sleep -Seconds 5
        }
    }

    Write-Status "Service $Name failed to reach Running state after $MaxAttempts attempts." 'ERROR'
    return $false
}

function Force-RebootNow {
    param(
        [string]$Reason = 'Required Windows Update services failed to start after multiple attempts.'
    )

    $script:RebootInitiated = $true
    $script:RebootReason = $Reason
    $script:FinalStatus = 'FailedRebootScheduled'
    $script:ExitCode = 1
    Write-Status "FORCING REBOOT: $Reason" 'ERROR'

    try {
        shutdown.exe /r /f /t 30 /c "$Reason" | Out-Null
        Write-Status "Forced reboot command issued successfully. System will restart in 30 seconds." 'ERROR'
    }
    catch {
        Write-Status "Failed to issue shutdown.exe reboot command: $($_.Exception.Message)" 'ERROR'
    }

    exit 1
}

function Enable-ScheduledTaskSafe {
    param(
        [Parameter(Mandatory)]
        [string]$TaskPath,

        [Parameter(Mandatory)]
        [string]$TaskName
    )

    try {
        $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
        if ($task.State -eq 'Disabled') {
            Enable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
            Write-Status "Enabled scheduled task: $TaskPath$TaskName" 'OK'
        }
        else {
            Write-Status "Scheduled task already enabled or available: $TaskPath$TaskName" 'INFO'
        }
    }
    catch {
        Write-Status "Scheduled task not found or could not be enabled: $TaskPath$TaskName" 'WARN'
    }
}

function Remove-RegistryValueSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        if (Test-Path $Path) {
            $prop = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $prop) {
                Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                Write-Status "Removed $Path\$Name" 'OK'
            }
            else {
                Write-Status "Registry value not present: $Path\$Name" 'INFO'
            }
        }
        else {
            Write-Status "Registry path not present: $Path" 'INFO'
        }
    }
    catch {
        Write-Status "Failed to remove $Path\$Name : $($_.Exception.Message)" 'WARN'
    }
}

function Set-RegistryDwordSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [int]$Value
    )

    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }

        New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force | Out-Null
        Write-Status "Set $Path\$Name = $Value" 'OK'
    }
    catch {
        Write-Status "Failed to set $Path\$Name : $($_.Exception.Message)" 'ERROR'
    }
}

function Set-ClassicRightClickMenuForHive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RegistryRoot,

        [Parameter(Mandatory)]
        [string]$DisplayName
    )

    $clsid = '{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
    $basePath = "$RegistryRoot\Software\Classes\CLSID\$clsid"
    $subPath  = "$basePath\InprocServer32"

    try {
        if (-not (Test-Path -LiteralPath $basePath)) {
            New-Item -Path $basePath -Force -ErrorAction Stop | Out-Null
            Write-Status "Created classic right-click menu CLSID key for $DisplayName" 'OK'
        }
        else {
            Write-Status "Classic right-click menu CLSID key already exists for $DisplayName" 'INFO'
        }

        if (-not (Test-Path -LiteralPath $subPath)) {
            New-Item -Path $subPath -Force -ErrorAction Stop | Out-Null
            Write-Status "Created classic right-click menu InprocServer32 key for $DisplayName" 'OK'
        }
        else {
            Write-Status "Classic right-click menu InprocServer32 key already exists for $DisplayName" 'INFO'
        }

        Set-Item -Path $subPath -Value '' -ErrorAction Stop
        Write-Status "Enabled Windows 11 classic right-click menu for $DisplayName." 'OK'
        return $true
    }
    catch {
        Write-Status "Failed to enable Windows 11 classic right-click menu for $DisplayName : $($_.Exception.Message)" 'WARN'
        return $false
    }
}

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
        Write-Status "User hive file not found for $DisplayName : $HiveFile" 'WARN'
        return $false
    }

    $tempHiveName = "TempClassicContextMenu_$([guid]::NewGuid().ToString('N'))"
    $loaded = $false

    try {
        $loadOutput = & reg.exe load "HKU\$tempHiveName" "$HiveFile" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Failed to load hive for $DisplayName : $($loadOutput -join ' ')" 'WARN'
            return $false
        }

        $loaded = $true
        $registryRoot = "Registry::HKEY_USERS\$tempHiveName"
        & $Action $registryRoot $DisplayName | Out-Null
        return $true
    }
    catch {
        Write-Status "Unexpected error while processing hive for $DisplayName : $($_.Exception.Message)" 'WARN'
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
                    Write-Status "Unloaded temporary registry hive for $DisplayName" 'OK'
                }
                else {
                    Write-Status "Failed to unload temporary registry hive for $DisplayName : $($unloadOutput -join ' ')" 'WARN'
                }
            }
            catch {
                Write-Status "Unexpected error unloading hive for $DisplayName : $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Enable-ClassicWindows11RightClickMenu {
    [CmdletBinding()]
    param()

    Write-Status "Applying Windows 11 classic right-click menu for all users..." 'INFO'

    $processedSids = New-Object 'System.Collections.Generic.HashSet[string]'

    [void](Set-ClassicRightClickMenuForHive -RegistryRoot 'HKCU:' -DisplayName 'current user')

    try {
        $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        [void]$processedSids.Add($currentSid)
    }
    catch {
        Write-Status "Could not determine current user SID: $($_.Exception.Message)" 'WARN'
    }

    try {
        Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction Stop |
            Where-Object {
                $_.PSChildName -match '^S-1-5-21-' -and
                $_.PSChildName -notmatch '_Classes$'
            } |
            ForEach-Object {
                $sid = $_.PSChildName
                [void]$processedSids.Add($sid)
                [void](Set-ClassicRightClickMenuForHive -RegistryRoot "Registry::HKEY_USERS\$sid" -DisplayName "loaded profile $sid")
            }
    }
    catch {
        Write-Status "Failed to enumerate loaded user registry hives: $($_.Exception.Message)" 'WARN'
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
                Write-Status "Profile already loaded or processed; skipping offline load for $($profile.LocalPath)" 'INFO'
                continue
            }

            $ntUserDat = Join-Path $profile.LocalPath 'NTUSER.DAT'
            [void](Invoke-WithLoadedUserHive -HiveFile $ntUserDat -DisplayName "offline profile $($profile.LocalPath)" -Action {
                param($RegistryRoot, $DisplayName)
                Set-ClassicRightClickMenuForHive -RegistryRoot $RegistryRoot -DisplayName $DisplayName
            })
        }
    }
    catch {
        Write-Status "Failed to process offline user profiles: $($_.Exception.Message)" 'WARN'
    }

    $defaultHive = Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT'
    if (Test-Path -LiteralPath $defaultHive) {
        [void](Invoke-WithLoadedUserHive -HiveFile $defaultHive -DisplayName 'Default User profile for future users' -Action {
            param($RegistryRoot, $DisplayName)
            Set-ClassicRightClickMenuForHive -RegistryRoot $RegistryRoot -DisplayName $DisplayName
        })
    }
    else {
        Write-Status "Default User hive not found at expected path: $defaultHive" 'WARN'
    }

    Write-Status "Completed Windows 11 classic right-click menu application for all available users." 'OK'
}


function Set-Windows11UIPreferencesForHive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RegistryRoot,

        [Parameter(Mandatory)]
        [string]$DisplayName
    )

    $advancedPath      = "$RegistryRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $searchSettingsPath = "$RegistryRoot\Software\Microsoft\Windows\CurrentVersion\SearchSettings"
    $copilotPolicyPath = "$RegistryRoot\Software\Policies\Microsoft\Windows\WindowsCopilot"

    $settings = @(
        # Windows 11 taskbar/start alignment
        @{ Path = $advancedPath;       Name = 'TaskbarAl';                 Value = 0; Description = 'taskbar/start alignment left' },

        # Windows 11 taskbar buttons - intentionally does NOT modify Microsoft Teams/Chat
        @{ Path = $advancedPath;       Name = 'TaskbarDa';                 Value = 0; Description = 'hide Widgets button' },
        @{ Path = $advancedPath;       Name = 'ShowTaskViewButton';        Value = 0; Description = 'hide Task View button' },
        @{ Path = $advancedPath;       Name = 'ShowCopilotButton';         Value = 0; Description = 'hide Copilot button' },

        # Search / Copilot preferences
        @{ Path = $searchSettingsPath; Name = 'IsDynamicSearchBoxEnabled'; Value = 0; Description = 'disable Search Highlights' },
        @{ Path = $copilotPolicyPath;  Name = 'TurnOffWindowsCopilot';     Value = 1; Description = 'disable Windows Copilot for this user' },

        # File Explorer preferences
        @{ Path = $advancedPath;       Name = 'LaunchTo';                  Value = 1; Description = 'open File Explorer to This PC' },
        @{ Path = $advancedPath;       Name = 'HideFileExt';               Value = 0; Description = 'show file extensions' },
        @{ Path = $advancedPath;       Name = 'Hidden';                    Value = 1; Description = 'show hidden files' }
    )

    try {
        foreach ($setting in $settings) {
            if (-not (Test-Path -LiteralPath $setting.Path)) {
                New-Item -Path $setting.Path -Force -ErrorAction Stop | Out-Null
            }

            New-ItemProperty `
                -Path $setting.Path `
                -Name $setting.Name `
                -Value $setting.Value `
                -PropertyType DWord `
                -Force `
                -ErrorAction Stop | Out-Null

            Write-Status "Applied $($setting.Description) for $DisplayName." 'OK'
        }

        # NOTE: Microsoft Teams / Chat is intentionally not changed.
        # Do not set TaskbarMn here because Teams is used on campus.

        return $true
    }
    catch {
        Write-Status "Failed to apply Windows 11 UI preferences for $DisplayName : $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Set-Windows11UIPreferencesForAllUsers {
    [CmdletBinding()]
    param()

    Write-Status "Applying Windows 11 UI preferences for all users..." 'INFO'
    Write-Status "Microsoft Teams / Chat taskbar settings are intentionally left unchanged." 'INFO'

    $processedSids = New-Object 'System.Collections.Generic.HashSet[string]'

    [void](Set-Windows11UIPreferencesForHive -RegistryRoot 'HKCU:' -DisplayName 'current user')

    try {
        $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        [void]$processedSids.Add($currentSid)
    }
    catch {
        Write-Status "Could not determine current user SID: $($_.Exception.Message)" 'WARN'
    }

    try {
        Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction Stop |
            Where-Object {
                $_.PSChildName -match '^S-1-5-21-' -and
                $_.PSChildName -notmatch '_Classes$'
            } |
            ForEach-Object {
                $sid = $_.PSChildName
                [void]$processedSids.Add($sid)
                [void](Set-Windows11UIPreferencesForHive -RegistryRoot "Registry::HKEY_USERS\$sid" -DisplayName "loaded profile $sid")
            }
    }
    catch {
        Write-Status "Failed to enumerate loaded user registry hives for Windows 11 UI preferences: $($_.Exception.Message)" 'WARN'
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
                Write-Status "Profile already loaded or processed; skipping offline Windows 11 UI preference load for $($profile.LocalPath)" 'INFO'
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
        Write-Status "Failed to process offline user profiles for Windows 11 UI preferences: $($_.Exception.Message)" 'WARN'
    }

    $defaultHive = Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT'
    if (Test-Path -LiteralPath $defaultHive) {
        [void](Invoke-WithLoadedUserHive -HiveFile $defaultHive -DisplayName 'Default User profile for future users' -Action {
            param($RegistryRoot, $DisplayName)
            Set-Windows11UIPreferencesForHive -RegistryRoot $RegistryRoot -DisplayName $DisplayName
        })
    }
    else {
        Write-Status "Default User hive not found at expected path: $defaultHive" 'WARN'
    }

    Write-Status "Completed Windows 11 UI preference application for all available users." 'OK'
}


function Get-ScriptHeaderValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$HeaderName
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $pattern = '^\s*#\s*' + [regex]::Escape($HeaderName) + '\s*:\s*(.+?)\s*$'
        $match = Select-String -LiteralPath $Path -Pattern $pattern -CaseSensitive:$false -ErrorAction Stop | Select-Object -First 1
        if ($match -and $match.Matches.Count -gt 0) {
            return $match.Matches[0].Groups[1].Value.Trim()
        }
    }
    catch {
        Write-Status "Unable to read $HeaderName from $Path : $($_.Exception.Message)" 'WARN'
    }

    return $null
}

function Convert-ToVersionObjectSafe {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$VersionText
    )

    if ([string]::IsNullOrWhiteSpace($VersionText)) {
        return [version]'0.0'
    }

    try {
        return [version]$VersionText.Trim()
    }
    catch {
        $cleanVersion = ($VersionText -replace '[^0-9\.]', '').Trim('.')
        if ([string]::IsNullOrWhiteSpace($cleanVersion)) {
            return [version]'0.0'
        }

        try {
            return [version]$cleanVersion
        }
        catch {
            Write-Status "Unable to parse version [$VersionText]. Treating as 0.0." 'WARN'
            return [version]'0.0'
        }
    }
}

function Get-WeeklySundayTriggerSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Triggers
    )

    foreach ($trigger in $Triggers) {
        try {
            $days = [string]$trigger.DaysOfWeek
            $startBoundary = [string]$trigger.StartBoundary
            $startTime = $null

            if (-not [string]::IsNullOrWhiteSpace($startBoundary)) {
                $startTime = ([datetime]$startBoundary).ToString('HH:mm')
            }

            if ($days -match 'Sunday' -and $startTime -eq '01:15') {
                return $true
            }
        }
        catch {
        }
    }

    return $false
}

function Test-CheckForUpdatedScriptsTaskCompliant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Task,

        [Parameter(Mandatory)]
        [string]$ExpectedScriptPath,

        [Parameter(Mandatory)]
        [string]$ExpectedPowerShellExe
    )

    $issues = New-Object System.Collections.Generic.List[string]

    $action = $Task.Actions | Select-Object -First 1
    if ($null -eq $action) {
        $issues.Add('Task has no action.') | Out-Null
    }
    else {
        $actualExecute = [string]$action.Execute
        $actualArguments = [string]$action.Arguments

        if (($actualExecute -ine $ExpectedPowerShellExe) -and ((Split-Path -Leaf $actualExecute) -ine 'powershell.exe')) {
            $issues.Add("Action executable is not Windows PowerShell. Current: $actualExecute") | Out-Null
        }

        if ($actualArguments -notmatch [regex]::Escape($ExpectedScriptPath)) {
            $issues.Add("Action does not point to expected updater script. Current arguments: $actualArguments") | Out-Null
        }

        if ($actualArguments -notmatch '-NoProfile') {
            $issues.Add('Action is missing -NoProfile.') | Out-Null
        }

        if ($actualArguments -notmatch '-ExecutionPolicy\s+Bypass') {
            $issues.Add('Action is missing -ExecutionPolicy Bypass.') | Out-Null
        }
    }

    if ($Task.Principal.UserId -notmatch 'SYSTEM') {
        $issues.Add("Task principal is not SYSTEM. Current: $($Task.Principal.UserId)") | Out-Null
    }

    if ($Task.Principal.RunLevel -ne 'Highest') {
        $issues.Add("Task run level is not Highest. Current: $($Task.Principal.RunLevel)") | Out-Null
    }

    if (-not (Get-WeeklySundayTriggerSummary -Triggers $Task.Triggers)) {
        $issues.Add('Task trigger is not Sunday at 01:15.') | Out-Null
    }

    return $issues
}


function Disable-MissedRunBehaviorForMaintenanceTasks {
    [CmdletBinding()]
    param(
        [string[]]$TaskNames = @(
            '01. Check for Updated Scripts',
            '02. Remove User Profiles',
            '03. Weekend Apps Update',
            '04. Update Edge Silent',
            '05. Weekend HP Drivers Update',
            '06. Weekend Windows Updates',
            '07. Force Reboot Install Updates',
            '08. System Repair',
            '09. Disable Windows Update Services',
            '10. Sync System Time'
        )
    )

    try {
        Import-Module ScheduledTasks -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Status "ScheduledTasks module is unavailable. Cannot clear missed-run behavior: $($_.Exception.Message)" 'WARN'
        return
    }

    foreach ($taskName in $TaskNames) {
        try {
            $tasks = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
        }
        catch {
            Write-Status "Task not found while clearing missed-run behavior: $taskName" 'WARN'
            continue
        }

        foreach ($task in @($tasks)) {
            try {
                if ($task.Settings.StartWhenAvailable -eq $true) {
                    $task.Settings.StartWhenAvailable = $false
                    Set-ScheduledTask -InputObject $task -ErrorAction Stop | Out-Null
                    Write-Status "Disabled missed-run behavior for task: $($task.TaskPath)$($task.TaskName)" 'OK'
                }
                else {
                    Write-Status "Missed-run behavior already disabled for task: $($task.TaskPath)$($task.TaskName)" 'INFO'
                }
            }
            catch {
                Write-Status "Failed to clear missed-run behavior for task $($task.TaskPath)$($task.TaskName): $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Ensure-CheckForUpdatedScriptsTask {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,

        [string]$TaskName = '01. Check for Updated Scripts'
    )

    $windowsPowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    if (-not (Test-Path -LiteralPath $windowsPowerShellExe)) {
        Write-Status "Windows PowerShell executable not found: $windowsPowerShellExe" 'ERROR'
        return $false
    }

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Status "Cannot register task because updater script is missing: $ScriptPath" 'ERROR'
        return $false
    }

    try {
        Import-Module ScheduledTasks -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Status "ScheduledTasks module is unavailable: $($_.Exception.Message)" 'ERROR'
        return $false
    }

    $argumentString = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
    $needsRebuild = $true

    try {
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        $issues = Test-CheckForUpdatedScriptsTaskCompliant -Task $existingTask -ExpectedScriptPath $ScriptPath -ExpectedPowerShellExe $windowsPowerShellExe

        if ($issues.Count -eq 0) {
            Write-Status "Scheduled task '$TaskName' is already configured correctly." 'OK'
            $needsRebuild = $false
        }
        else {
            Write-Status "Scheduled task '$TaskName' needs repair:" 'WARN'
            foreach ($issue in $issues) {
                Write-Status " - $issue" 'WARN'
            }
        }
    }
    catch {
        Write-Status "Scheduled task '$TaskName' does not exist and will be created." 'WARN'
    }

    if (-not $needsRebuild) {
        return $true
    }

    try {
        $action = New-ScheduledTaskAction -Execute $windowsPowerShellExe -Argument $argumentString
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -WeeksInterval 1 -At ([datetime]::Today.Add([timespan]::Parse('01:15')))
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Hours 12)

        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Write-Status "Scheduled task '$TaskName' now points to: $ScriptPath" 'OK'
        Write-Status "Task action: $windowsPowerShellExe $argumentString" 'INFO'
        Write-Status "Task trigger: Sunday at 01:15 as SYSTEM with highest privileges." 'OK'
        return $true
    }
    catch {
        Write-Status "Failed to create/repair scheduled task '$TaskName': $($_.Exception.Message)" 'ERROR'
        return $false
    }
}


try {
    if (-not (Test-IsAdmin)) {
        Write-Status 'This script must be run as Administrator.' 'ERROR'
        $script:FinalStatus = 'Failed'
        $script:ExitCode = 1
        throw 'Administrative privileges are required.'
    }

    $script:ServiceStateBefore = @(Get-ServiceTelemetrySnapshot -Names $script:ServiceNames)
    $script:PolicyConflictsBefore = @(Get-WindowsUpdatePolicyConflicts)

    Invoke-UpdaterBootstrap

Write-Status "Initializing script..." 'INFO'
Set-Windows11UIPreferencesForAllUsers
Enable-ClassicWindows11RightClickMenu

# Restore registry startup values first
# Common defaults used for Windows Update-related services:
# wuauserv = Manual (3)
# bits = Manual (3)
# dosvc = Automatic family (2)
# UsoSvc = Automatic (2)
# WaaSMedicSvc = Manual/triggered on many systems (3)

Set-ServiceStartRegistry -ServiceName 'wuauserv'     -StartValue 3
Set-ServiceStartRegistry -ServiceName 'bits'         -StartValue 3
Set-ServiceStartRegistry -ServiceName 'dosvc'        -StartValue 2
Set-ServiceStartRegistry -ServiceName 'UsoSvc'       -StartValue 2
Set-ServiceStartRegistry -ServiceName 'WaaSMedicSvc' -StartValue 3

# Initial restore and startup
Set-ServiceStartupAndStart -Name 'wuauserv' -StartupType Manual
Set-ServiceStartupAndStart -Name 'bits'     -StartupType Manual
Set-ServiceStartupAndStart -Name 'dosvc'    -StartupType Automatic
Set-ServiceStartupAndStart -Name 'UsoSvc'   -StartupType Automatic

# WaaSMedicSvc can be protected; set registry above, then try starting via sc.exe
try {
    & sc.exe config WaaSMedicSvc start= demand | Out-Null
    Write-Status "Configured WaaSMedicSvc startup via sc.exe" 'OK'
}
catch {
    Write-Status "Could not configure WaaSMedicSvc via sc.exe" 'WARN'
}

try {
    & sc.exe start WaaSMedicSvc | Out-Null
    Write-Status "Attempted to start WaaSMedicSvc" 'INFO'
}
catch {
    Write-Status "Could not start WaaSMedicSvc directly" 'WARN'
}

# Restore Automatic Updates policy
$wuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$auPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'

Set-RegistryDwordSafe -Path $auPolicyPath -Name 'NoAutoUpdate' -Value 0
Set-RegistryDwordSafe -Path $auPolicyPath -Name 'AUOptions'    -Value 3

# Remove common WSUS redirection values if they were previously set
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUServer'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUStatusServer'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'UpdateServiceUrlAlternate'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'SetProxyBehaviorForUpdateDetection'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'DisableWindowsUpdateAccess'
Remove-RegistryValueSafe -Path $auPolicyPath -Name 'UseWUServer'

# Re-enable common update scheduled tasks
$tasks = @(
    @{ Path = '\Microsoft\Windows\WindowsUpdate\';      Name = 'Scheduled Start' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan Static Task' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'USO_UxBroker' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Reboot' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Maintenance Install' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Refresh Settings' },
    @{ Path = '\Microsoft\Windows\WaaSMedic\';          Name = 'PerformRemediation' }
)

foreach ($task in $tasks) {
    Enable-ScheduledTaskSafe -TaskPath $task.Path -TaskName $task.Name
}

# Restart key services in a sensible order
$restartOrder = @('bits', 'dosvc', 'wuauserv', 'UsoSvc')
foreach ($svc in $restartOrder) {
    try {
        Restart-Service -Name $svc -Force -ErrorAction Stop
        Write-Status "Restarted service: $svc" 'OK'
    }
    catch {
        Write-Status "Could not restart $svc : $($_.Exception.Message)" 'WARN'
    }
}

# Verify and retry critical services
$requiredServices = @(
    @{ Name = 'wuauserv'; StartupType = 'Manual' },
    @{ Name = 'bits';     StartupType = 'Manual' },
    @{ Name = 'dosvc';    StartupType = 'Automatic' },
    @{ Name = 'UsoSvc';   StartupType = 'Automatic' }
)

$failedServices = @()

foreach ($requiredService in $requiredServices) {
    $serviceStarted = Ensure-ServiceRunningWithRetry -Name $requiredService.Name -StartupType $requiredService.StartupType -MaxAttempts 4 -WaitPerAttemptSeconds 15
    if (-not $serviceStarted) {
        $failedServices += $requiredService.Name
    }
}

if ($failedServices.Count -gt 0) {
    $failedList = $failedServices -join ', '
    Write-Status "One or more critical Windows Update services failed to start: $failedList" 'ERROR'
    Force-RebootNow -Reason "Windows Update service recovery failed. Services not running: $failedList"
}


# ------------------------------------------------------------
# Scheduled-task reconciliation
# The register script is idempotent: it creates missing tasks and changes only
# managed tasks whose action, trigger, principal, or settings differ.
# ------------------------------------------------------------
$RegisterTasksScript = 'C:\Scripts\Register-Tasks_SYSTEM.ps1'

if (Test-Path -LiteralPath $RegisterTasksScript -PathType Leaf) {
    Write-Status 'Reconciling managed scheduled tasks...' 'INFO'

    try {
        $taskReconcileProcess = Start-Process `
            -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$RegisterTasksScript`"" `
            -Wait `
            -PassThru `
            -WindowStyle Hidden

        if ($taskReconcileProcess.ExitCode -eq 0) {
            $script:TaskReconcileStatus = 'Success'
            Write-Status 'Managed scheduled tasks are current.' 'OK'
        }
        else {
            $script:TaskReconcileStatus = "ExitCode:$($taskReconcileProcess.ExitCode)"
            Write-Status "Task reconciliation exited with code $($taskReconcileProcess.ExitCode)." 'WARN'
        }
    }
    catch {
        $script:TaskReconcileStatus = 'Failed'
        Write-Status "Task reconciliation failed: $($_.Exception.Message)" 'WARN'
    }
}
else {
    $script:TaskReconcileStatus = 'ScriptMissing'
    Write-Status "Task reconciliation script is missing: $RegisterTasksScript" 'WARN'
}

# Final enforcement pass. This runs after any task rebuild so older register scripts
# cannot leave "Run task as soon as possible after a scheduled start is missed" enabled.
Disable-MissedRunBehaviorForMaintenanceTasks

$script:FinalStatus = if ($script:WarningCount -gt 0) { 'SuccessWithWarnings' } else { 'Success' }
$script:ExitCode = 0
Write-Status "Windows Update settings have been restored and critical services are running." 'OK'
Write-Status "No reboot required. Continuing normally." 'INFO'
}
catch {
    if ($script:ExitCode -eq 0) { $script:ExitCode = 1 }
    if ($script:FinalStatus -eq 'Running') { $script:FinalStatus = 'Failed' }
    Write-Status "Unhandled failure: $($_.Exception.Message)" 'ERROR'
}
finally {
    $script:ServiceStateAfter = @(Get-ServiceTelemetrySnapshot -Names $script:ServiceNames)
    Write-ExecutionTelemetry
}

exit $script:ExitCode
