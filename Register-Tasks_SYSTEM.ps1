#requires -Version 5.1
#requires -RunAsAdministrator
<#
.SYNOPSIS
    Reconciles Compton College managed scheduled tasks under SYSTEM.

.DESCRIPTION
    Creates missing tasks and updates only tasks whose managed properties differ.
    Leaves unrelated and Microsoft tasks untouched. Writes structured task
    reconciliation telemetry for Elastic.

.NOTES
    ScriptName:    Register-Tasks_SYSTEM.ps1
    ScriptVersion: 4.2.1
    Designed for:  Windows PowerShell 5.1

    4.2.1:
    - Fixed Windows PowerShell 5.1 generic-list telemetry serialization.
    - Suppressed unapproved-verb warnings when importing Maintenance.Framework.
#>

[CmdletBinding()]
param(
    [string]$ScriptsRoot = 'C:\Scripts',
    [string]$TaskPath = '\',
    [string]$LogDirectory = 'C:\Logs'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ScriptName       = 'Register-Tasks_SYSTEM.ps1'
$script:ScriptVersion    = '4.2.1'
$script:RunId            = [guid]::NewGuid().Guid
$script:StartTime        = Get-Date
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:TaskResults      = New-Object System.Collections.Generic.List[object]
$script:RemovedTasks     = New-Object System.Collections.Generic.List[string]
$script:MissingScripts   = New-Object System.Collections.Generic.List[string]
$script:Mutex            = $null
$script:MutexAcquired    = $false

$script:LogPath    = Join-Path $LogDirectory 'Register-Tasks_SYSTEM.log'
$script:NdjsonPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestPath = Join-Path $LogDirectory 'Register-Tasks_SYSTEM.latest.json'

$WindowsPowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','ACTION','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = '[{0}] [{1,-6}] {2}' -f $timestamp, $Level, $Message

    switch ($Level) {
        'WARN'  { $script:WarningCount++ }
        'ERROR' { $script:ErrorCount++ }
    }

    try {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    } catch {}

    $color = switch ($Level) {
        'ACTION' { 'Yellow' }
        'OK'     { 'Green' }
        'WARN'   { 'DarkYellow' }
        'ERROR'  { 'Red' }
        default  { 'Cyan' }
    }
    Write-Host $line -ForegroundColor $color
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Normalize-TaskPath {
    param([Parameter(Mandatory)][string]$Path)

    $normalized = $Path.Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) { return '\' }
    if (-not $normalized.StartsWith('\')) { $normalized = '\' + $normalized }
    if (-not $normalized.EndsWith('\')) { $normalized += '\' }
    return $normalized
}

function Get-DesiredActionArguments {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$ExtraArguments = ''
    )

    $launcherPath = Join-Path $ScriptsRoot 'Invoke-MaintenanceScript.ps1'
    $escapedExtra = $ExtraArguments.Replace('"','\"')
    $arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -TargetScript "{1}"' -f $launcherPath,$ScriptPath
    if (-not [string]::IsNullOrWhiteSpace($escapedExtra)) {
        $arguments += ' -TargetArguments "' + $escapedExtra + '"'
    }
    return $arguments
}

function Get-TimeText {
    param($DateValue)
    try { return ([datetime]$DateValue).ToString('HH:mm') } catch { return $null }
}

function New-ManagedTaskPrincipal {
    New-ScheduledTaskPrincipal `
        -UserId 'SYSTEM' `
        -LogonType ServiceAccount `
        -RunLevel Highest
}

function New-ManagedTaskSettings {
    New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Hours 12) `
        -MultipleInstances IgnoreNew `
        -RestartCount 2 `
        -RestartInterval (New-TimeSpan -Minutes 5)
}

function Test-CommonTaskProperties {
    param(
        [Parameter(Mandatory)]$ExistingTask,
        [Parameter(Mandatory)][string]$ExpectedArguments
    )

    $issues = New-Object System.Collections.Generic.List[string]
    $actions = @($ExistingTask.Actions)

    if ($actions.Count -ne 1) {
        [void]$issues.Add("Expected one action; found $($actions.Count).")
    } else {
        $action = $actions[0]
        if ([string]$action.Execute -ine $WindowsPowerShellExe) {
            [void]$issues.Add("Executable differs: $($action.Execute)")
        }
        if ([string]$action.Arguments -cne $ExpectedArguments) {
            [void]$issues.Add("Arguments differ: $($action.Arguments)")
        }
    }

    if ([string]$ExistingTask.Principal.UserId -notmatch '^(SYSTEM|S-1-5-18)$') {
        [void]$issues.Add("Principal differs: $($ExistingTask.Principal.UserId)")
    }
    if ([string]$ExistingTask.Principal.RunLevel -ne 'Highest') {
        [void]$issues.Add("Run level differs: $($ExistingTask.Principal.RunLevel)")
    }
    if ($ExistingTask.Settings.StartWhenAvailable) {
        [void]$issues.Add('Run-as-soon-as-possible-after-missed-start is enabled.')
    }
    if ($ExistingTask.Settings.DisallowStartIfOnBatteries) {
        [void]$issues.Add('Task is blocked while on battery power.')
    }
    if ($ExistingTask.Settings.StopIfGoingOnBatteries) {
        [void]$issues.Add('Task stops when switching to battery power.')
    }
    if ([string]$ExistingTask.Settings.MultipleInstances -ne 'IgnoreNew') {
        [void]$issues.Add("Multiple-instance policy differs: $($ExistingTask.Settings.MultipleInstances)")
    }

    return $issues.ToArray()
}

function Test-WeeklyTaskMatches {
    param(
        [Parameter(Mandatory)]$ExistingTask,
        [Parameter(Mandatory)][string]$ExpectedArguments,
        [Parameter(Mandatory)][string]$StartTime
    )

    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($issue in @(Test-CommonTaskProperties -ExistingTask $ExistingTask -ExpectedArguments $ExpectedArguments)) {
        [void]$issues.Add($issue)
    }

    $triggers = @($ExistingTask.Triggers)
    if ($triggers.Count -ne 1) {
        [void]$issues.Add("Expected one weekly trigger; found $($triggers.Count).")
    } else {
        $trigger = $triggers[0]
        if ([string]$trigger.CimClass.CimClassName -notmatch 'Weekly') {
            [void]$issues.Add('Trigger is not weekly.')
        }
        if ([int]$trigger.DaysOfWeek -ne 1) {
            [void]$issues.Add("Weekly day differs: $($trigger.DaysOfWeek)")
        }
        if ([int]$trigger.WeeksInterval -ne 1) {
            [void]$issues.Add("Weeks interval differs: $($trigger.WeeksInterval)")
        }
        $actualTime = Get-TimeText -DateValue $trigger.StartBoundary
        if ($actualTime -ne $StartTime) {
            [void]$issues.Add("Start time differs: $actualTime")
        }
    }

    return $issues.ToArray()
}

function Test-TimeSyncTaskMatches {
    param(
        [Parameter(Mandatory)]$ExistingTask,
        [Parameter(Mandatory)][string]$ExpectedArguments,
        [Parameter(Mandatory)][string[]]$ExpectedTimes
    )

    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($issue in @(Test-CommonTaskProperties -ExistingTask $ExistingTask -ExpectedArguments $ExpectedArguments)) {
        [void]$issues.Add($issue)
    }

    $triggers = @($ExistingTask.Triggers)
    if ($triggers.Count -ne $ExpectedTimes.Count) {
        [void]$issues.Add("Expected $($ExpectedTimes.Count) daily triggers; found $($triggers.Count).")
    }

    foreach ($trigger in $triggers) {
        if ([string]$trigger.CimClass.CimClassName -notmatch 'Daily') {
            [void]$issues.Add('One or more time-sync triggers are not daily.')
            break
        }
    }

    $actualTimes = @(
        $triggers |
            ForEach-Object { Get-TimeText -DateValue $_.StartBoundary } |
            Where-Object { $_ } |
            Sort-Object
    )
    $wantedTimes = @($ExpectedTimes | Sort-Object)

    if (($actualTimes -join ',') -ne ($wantedTimes -join ',')) {
        [void]$issues.Add("Trigger times differ. Actual: $($actualTimes -join ', ')")
    }

    return $issues.ToArray()
}

function Add-TaskResult {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$Schedule,
        [Parameter(Mandatory)][string]$Result,
        [bool]$ScriptExists,
        [string[]]$Issues = @(),
        [string]$ErrorMessage = $null
    )

    $script:TaskResults.Add([pscustomobject]@{
        TaskName      = $Name
        TaskPath      = $TaskPath
        ScriptPath    = $ScriptPath
        ScriptExists  = $ScriptExists
        Schedule      = $Schedule
        Result        = $Result
        Issues        = @($Issues)
        ErrorMessage  = $ErrorMessage
    })
}

function Ensure-WeeklyTask {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$StartTime,
        [string]$ExtraArguments = ''
    )

    $scriptExists = Test-Path -LiteralPath $ScriptPath -PathType Leaf
    if (-not $scriptExists) {
        $script:MissingScripts.Add($ScriptPath)
        Write-Log -Level 'WARN' -Message "Managed script is currently missing; task will still be reconciled: $ScriptPath"
    }

    $arguments = Get-DesiredActionArguments -ScriptPath $ScriptPath -ExtraArguments $ExtraArguments
    $existing = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction SilentlyContinue
    $issues = @()

    if ($existing) {
        $issues = @(Test-WeeklyTaskMatches -ExistingTask $existing -ExpectedArguments $arguments -StartTime $StartTime)
        if ($issues.Count -eq 0) {
            Write-Log -Level 'OK' -Message "Task is already correct: $TaskPath$Name"
            Add-TaskResult -Name $Name -ScriptPath $ScriptPath -Schedule "Sunday $StartTime" `
                -Result 'Current' -ScriptExists $scriptExists
            return
        }

        Write-Log -Level 'ACTION' -Message "Task requires an update: $TaskPath$Name"
        foreach ($issue in $issues) { Write-Log -Level 'WARN' -Message " - $issue" }
    } else {
        Write-Log -Level 'ACTION' -Message "Task is missing and will be created: $TaskPath$Name"
    }

    try {
        $action = New-ScheduledTaskAction -Execute $WindowsPowerShellExe -Argument $arguments
        $startAt = [datetime]::Today.Add([timespan]::Parse($StartTime))
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -WeeksInterval 1 -At $startAt

        Register-ScheduledTask `
            -TaskName $Name `
            -TaskPath $TaskPath `
            -Action $action `
            -Trigger $trigger `
            -Principal (New-ManagedTaskPrincipal) `
            -Settings (New-ManagedTaskSettings) `
            -Force `
            -ErrorAction Stop | Out-Null

        $verified = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction Stop
        $verifyIssues = @(Test-WeeklyTaskMatches -ExistingTask $verified -ExpectedArguments $arguments -StartTime $StartTime)
        if ($verifyIssues.Count -gt 0) {
            throw "Post-registration verification failed: $($verifyIssues -join '; ')"
        }

        $result = if ($existing) { 'Updated' } else { 'Created' }
        Write-Log -Level 'OK' -Message "Task $result and verified: $TaskPath$Name at Sunday $StartTime"
        Add-TaskResult -Name $Name -ScriptPath $ScriptPath -Schedule "Sunday $StartTime" `
            -Result $result -ScriptExists $scriptExists -Issues $issues
    } catch {
        Add-TaskResult -Name $Name -ScriptPath $ScriptPath -Schedule "Sunday $StartTime" `
            -Result 'Failed' -ScriptExists $scriptExists -Issues $issues -ErrorMessage $_.Exception.Message
        throw
    }
}

function Ensure-TimeSyncTask {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$Name = '00. Sync System Time Every 4 Hours'
    )

    $triggerTimes = @('00:00','04:00','08:00','12:00','16:00','20:00')
    $scriptExists = Test-Path -LiteralPath $ScriptPath -PathType Leaf

    if (-not $scriptExists) {
        $script:MissingScripts.Add($ScriptPath)
        Write-Log -Level 'WARN' -Message "Managed script is currently missing; task will still be reconciled: $ScriptPath"
    }

    $arguments = Get-DesiredActionArguments -ScriptPath $ScriptPath
    $existing = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction SilentlyContinue
    $issues = @()

    if ($existing) {
        $issues = @(Test-TimeSyncTaskMatches -ExistingTask $existing `
            -ExpectedArguments $arguments -ExpectedTimes $triggerTimes)

        if ($issues.Count -eq 0) {
            Write-Log -Level 'OK' -Message "Task is already correct: $TaskPath$Name"
            Add-TaskResult -Name $Name -ScriptPath $ScriptPath -Schedule 'Daily every 4 hours' `
                -Result 'Current' -ScriptExists $scriptExists
            return
        }

        Write-Log -Level 'ACTION' -Message "Time-sync task requires an update: $TaskPath$Name"
        foreach ($issue in $issues) { Write-Log -Level 'WARN' -Message " - $issue" }
    } else {
        Write-Log -Level 'ACTION' -Message "Time-sync task is missing and will be created: $TaskPath$Name"
    }

    try {
        $action = New-ScheduledTaskAction -Execute $WindowsPowerShellExe -Argument $arguments
        $triggers = foreach ($timeText in $triggerTimes) {
            $atTime = [datetime]::Today.Add([timespan]::Parse($timeText))
            New-ScheduledTaskTrigger -Daily -At $atTime -DaysInterval 1
        }

        Register-ScheduledTask `
            -TaskName $Name `
            -TaskPath $TaskPath `
            -Action $action `
            -Trigger $triggers `
            -Principal (New-ManagedTaskPrincipal) `
            -Settings (New-ManagedTaskSettings) `
            -Force `
            -ErrorAction Stop | Out-Null

        $verified = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction Stop
        $verifyIssues = @(Test-TimeSyncTaskMatches -ExistingTask $verified `
            -ExpectedArguments $arguments -ExpectedTimes $triggerTimes)
        if ($verifyIssues.Count -gt 0) {
            throw "Post-registration verification failed: $($verifyIssues -join '; ')"
        }

        $result = if ($existing) { 'Updated' } else { 'Created' }
        Write-Log -Level 'OK' -Message "Time-sync task $result and verified: $TaskPath$Name"
        Add-TaskResult -Name $Name -ScriptPath $ScriptPath -Schedule 'Daily every 4 hours' `
            -Result $result -ScriptExists $scriptExists -Issues $issues
    } catch {
        Add-TaskResult -Name $Name -ScriptPath $ScriptPath -Schedule 'Daily every 4 hours' `
            -Result 'Failed' -ScriptExists $scriptExists -Issues $issues -ErrorMessage $_.Exception.Message
        throw
    }
}

function Remove-ObsoleteManagedTask {
    param([Parameter(Mandatory)][string]$Name)

    $task = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction SilentlyContinue
    if (-not $task) { return }

    Write-Log -Level 'ACTION' -Message "Removing superseded managed task: $TaskPath$Name"
    Unregister-ScheduledTask -TaskName $Name -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
    $script:RemovedTasks.Add($Name)
    Write-Log -Level 'OK' -Message "Removed superseded managed task: $TaskPath$Name"
}

function Write-JsonAtomically {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Json
    )

    $temporary = '{0}.{1}.tmp' -f $Path, ([guid]::NewGuid().Guid)
    [IO.File]::WriteAllText($temporary, $Json, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Write-Telemetry {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$ExitCode,
        [string]$FailureMessage = $null
    )

    $endTime = Get-Date
    $results = $script:TaskResults.ToArray()
    $event = [ordered]@{
        '@timestamp'       = $endTime.ToUniversalTime().ToString('o')
        EventType          = 'maintenance.task_reconciliation'
        SchemaVersion      = '1.0'
        RunId              = $script:RunId
        ComputerName       = $env:COMPUTERNAME
        Domain             = $env:USERDOMAIN
        ScriptName         = $script:ScriptName
        ScriptVersion      = $script:ScriptVersion
        Status             = $Status
        ExitCode           = $ExitCode
        StartTime          = $script:StartTime.ToUniversalTime().ToString('o')
        EndTime            = $endTime.ToUniversalTime().ToString('o')
        DurationSeconds    = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 3)
        WarningCount       = $script:WarningCount
        ErrorCount         = $script:ErrorCount
        FailureMessage     = $FailureMessage
        ScriptsRoot        = $ScriptsRoot
        TaskPath           = $TaskPath
        ExpectedTaskCount  = $results.Count
        CurrentCount       = @($results | Where-Object Result -eq 'Current').Count
        CreatedCount       = @($results | Where-Object Result -eq 'Created').Count
        UpdatedCount       = @($results | Where-Object Result -eq 'Updated').Count
        FailedCount        = @($results | Where-Object Result -eq 'Failed').Count
        MissingScriptCount = @($script:MissingScripts.ToArray() | Select-Object -Unique).Count
        MissingScripts     = @($script:MissingScripts.ToArray() | Select-Object -Unique)
        RemovedObsoleteTasks = $script:RemovedTasks.ToArray()
        TaskResults        = $results
    }

    $compact = $event | ConvertTo-Json -Depth 8 -Compress
    $pretty  = $event | ConvertTo-Json -Depth 8

    Write-MaintenanceTelemetryLine -Path $script:NdjsonPath -JsonLine $compact
    Write-JsonAtomically -Path $script:LatestPath -Json $pretty
}

Ensure-Directory -Path $LogDirectory
$TaskPath = Normalize-TaskPath -Path $TaskPath

$exitCode = 0
$status = 'Success'
$failureMessage = $null

try {
    $createdNew = $false
    $script:Mutex = New-Object Threading.Mutex($true, 'Global\Compton_RegisterTasks_SYSTEM', [ref]$createdNew)
    if (-not $createdNew) {
        $script:MutexAcquired = $script:Mutex.WaitOne(0)
        if (-not $script:MutexAcquired) {
            Write-Log -Level 'WARN' -Message 'Another task-reconciliation instance is already running.'
            $status = 'SkippedAlreadyRunning'
            $exitCode = 0
            return
        }
    } else {
        $script:MutexAcquired = $true
    }

    if (-not (Test-IsAdministrator)) {
        throw 'This script must run as Administrator or SYSTEM.'
    }

    Import-Module ScheduledTasks -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $ScriptsRoot -PathType Container)) {
        Write-Log -Level 'WARN' -Message "Scripts root does not currently exist: $ScriptsRoot"
    }

    Write-Log -Message "Reconciling managed tasks under $TaskPath using scripts from $ScriptsRoot"

    # Script 14 runs after the final reboot and Script 13 so the endpoint.health
    # document represents the completed weekly maintenance state.
    $taskDefinitions = @(
        [pscustomobject]@{ Name='01. Check for Updated Scripts';             Script='00_Update-Scripts-FromShare.ps1';                    Time='01:00'; Args='' },
        [pscustomobject]@{ Name='01A. Create Weekly System Restore Point';   Script='12_Enable-SystemRestore-And-Create-RestorePoint.ps1'; Time='01:10'; Args='' },
        [pscustomobject]@{ Name='02. Enable Windows Update Services';        Script='01_Enable_Windows_Update_Services.ps1';              Time='01:30'; Args='' },
        [pscustomobject]@{ Name='03. Remove User Profiles Weekly';           Script='02_Remove_User_Profiles.ps1';                         Time='01:40'; Args='' },
        [pscustomobject]@{ Name='04. Weekend Apps Update';                   Script='03_Weekend_Apps_Update.ps1';                          Time='02:00'; Args='' },
        [pscustomobject]@{ Name='06. Weekend HP Drivers Update';             Script='05_Weekend_HP_Drivers_Update.ps1';                    Time='03:00'; Args='' },
        [pscustomobject]@{ Name='07. Weekend Windows Updates - 1st Pass';    Script='06_Weekend_Windows_Updates.ps1';                      Time='04:00'; Args='' },
        [pscustomobject]@{ Name='08. Force Reboot Install Updates';          Script='07_Force_Reboot_Install_Updates.ps1';                 Time='05:00'; Args='' },
        [pscustomobject]@{ Name='09. Weekend Windows Updates - 2nd Pass';    Script='06_Weekend_Windows_Updates.ps1';                      Time='05:30'; Args='' },
        [pscustomobject]@{ Name='10. Disable Windows Update Services';       Script='09_Disable_Windows_Update_Services.ps1';              Time='06:00'; Args='' },
        [pscustomobject]@{ Name='11. Force Reboot Install Updates 2';        Script='07_Force_Reboot_Install_Updates.ps1';                 Time='06:05'; Args='' },
        [pscustomobject]@{ Name='12. System Repair';                         Script='08_System_Repair.ps1';                                Time='06:15'; Args='' },
        [pscustomobject]@{ Name='13. Maintain SHARP Driver and PaperCut';    Script='11_Install_SharpDriver_And_PaperCut.ps1';             Time='07:00'; Args='' },
        [pscustomobject]@{ Name='14. Force Reboot Install Updates 3';        Script='07_Force_Reboot_Install_Updates.ps1';                 Time='07:30'; Args='' },
        [pscustomobject]@{ Name='15. Configure Autologon and Edge';          Script='13_Configure_Autologon_And_Edge.ps1';                 Time='07:45'; Args='' },
        [pscustomobject]@{ Name='16. Weekly Endpoint Health Inventory';      Script='14_Endpoint_Health_Inventory.ps1';                    Time='08:15'; Args='' }
    )

    $obsoleteManagedTaskNames = @(
        '04. Update Edge Silent',
        '05. Update Edge Silent',
        'Update Edge Silent',
        '13. Force Reboot Install Updates 3',
        '14. Maintain SHARP Driver and PaperCut',
        '13. Configure Autologon and Edge'
    )

    $desiredTaskNames = @($taskDefinitions.Name) + '00. Sync System Time Every 4 Hours'
    foreach ($obsoleteName in $obsoleteManagedTaskNames) {
        if ($obsoleteName -notin $desiredTaskNames) {
            Remove-ObsoleteManagedTask -Name $obsoleteName
        }
    }

    Ensure-TimeSyncTask -ScriptPath (Join-Path $ScriptsRoot '10_Sync_System_Time.ps1')

    foreach ($definition in $taskDefinitions) {
        Ensure-WeeklyTask `
            -Name $definition.Name `
            -ScriptPath (Join-Path $ScriptsRoot $definition.Script) `
            -StartTime $definition.Time `
            -ExtraArguments $definition.Args
    }

    $missingRegisteredTasks = @()
    foreach ($expectedName in $desiredTaskNames) {
        if (-not (Get-ScheduledTask -TaskName $expectedName -TaskPath $TaskPath -ErrorAction SilentlyContinue)) {
            $missingRegisteredTasks += $expectedName
        }
    }

    if ($missingRegisteredTasks.Count -gt 0) {
        throw "Managed tasks missing after reconciliation: $($missingRegisteredTasks -join ', ')"
    }

    if ($script:MissingScripts.Count -gt 0) {
        $status = 'SuccessWithWarnings'
        Write-Log -Level 'WARN' -Message ("Reconciliation succeeded, but {0} managed script file(s) are missing." -f
            @($script:MissingScripts.ToArray() | Select-Object -Unique).Count)
    } else {
        Write-Log -Level 'OK' -Message "All $($desiredTaskNames.Count) managed tasks were reconciled and verified."
    }
}
catch {
    $failureMessage = $_.Exception.Message
    $status = 'Failed'
    $exitCode = 2
    Write-Log -Level 'ERROR' -Message "Task reconciliation failed: $failureMessage"
}
finally {
    try {
        Write-Telemetry -Status $status -ExitCode $exitCode -FailureMessage $failureMessage
    } catch {
        Write-Log -Level 'ERROR' -Message "Unable to write reconciliation telemetry: $($_.Exception.Message)"
        if ($exitCode -eq 0) { $exitCode = 3 }
    }

    if ($script:MutexAcquired -and $script:Mutex) {
        try { $script:Mutex.ReleaseMutex() } catch {}
    }
    if ($script:Mutex) { $script:Mutex.Dispose() }
}

exit $exitCode
