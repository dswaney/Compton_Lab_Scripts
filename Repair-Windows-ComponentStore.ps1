#requires -Version 5.1
<#
.SYNOPSIS
    Repairs the Windows component store and, when necessary, escalates to a
    Windows 11 in-place repair that preserves installed applications and data.

.DESCRIPTION
    Automatically matches the running Windows edition, architecture, and
    language to an index inside install.wim. The script then runs DISM
    CheckHealth, ScanHealth, RestoreHealth, SFC, and a final DISM ScanHealth.
    If both the WIM-only repair and Windows Update fallback fail because repair
    content is unavailable, the script can copy the complete Windows 11 25H2
    setup media locally and launch an unattended in-place repair upgrade.
    Before Windows Setup is launched, enabled Compton maintenance tasks are
    recorded and disabled so another scheduled job cannot interrupt Setup.
    The post-upgrade verification task restores their original enabled state
    after DISM and SFC verification completes.

    Run from an elevated Windows PowerShell session while the computer is
    thawed if Deep Freeze is installed.

.NOTES
    ScriptVersion: 2.2.0
    DefaultSource: \\SERVER\DeploymentShare\Installers\25H2\sources\install.wim
    SetupMedia:    \\SERVER\DeploymentShare\Installers\25H2
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ImagePath = '\\SERVER\DeploymentShare\Installers\25H2\sources\install.wim',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LogDirectory = 'C:\Logs',

    [Parameter()]
    [string]$DismPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$AdkSharePath = '\\SERVER\DeploymentShare\Installers\ADK\Deployment Tools',

    [Parameter()]
    [switch]$SkipSfc,

    [Parameter()]
    [switch]$DisableWindowsUpdateFallback,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SetupMediaShare = '\\SERVER\DeploymentShare\Installers\25H2',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LocalSetupMediaPath = 'C:\Temp\Compton-Windows11-25H2-Repair',

    [Parameter()]
    [bool]$EnableInPlaceRepair = $true,

    [Parameter()]
    [ValidateRange(25, 200)]
    [int]$MinimumFreeSpaceGB = 40,

    [Parameter()]
    [switch]$KeepLocalSetupMedia
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptVersion = '2.2.0'
$script:ComputerName = $env:COMPUTERNAME
$script:StartTime = Get-Date
$script:LogPath = $null
$script:StagedAdkRoot = $null
$script:StateDirectory = 'C:\ProgramData\Compton\Repair-Windows-ComponentStore'
$script:StatePath = Join-Path $script:StateDirectory 'InPlaceRepair-State.json'

function Add-RepairLogContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string[]]$Line,

        [Parameter()]
        [ValidateRange(1, 30)]
        [int]$RetryCount = 12,

        [Parameter()]
        [ValidateRange(50, 5000)]
        [int]$RetryDelayMilliseconds = 250
    )

    if (-not $script:LogPath) {
        return
    }

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        $stream = $null
        $writer = $null

        try {
            $stream = [System.IO.FileStream]::new(
                $script:LogPath,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
            )
            $writer = [System.IO.StreamWriter]::new(
                $stream,
                [System.Text.UTF8Encoding]::new($false)
            )

            foreach ($entry in $Line) {
                $writer.WriteLine($entry)
            }

            $writer.Flush()
            return
        }
        catch [System.IO.IOException] {
            if ($attempt -lt $RetryCount) {
                Start-Sleep -Milliseconds $RetryDelayMilliseconds
                continue
            }

            Write-Warning "Unable to append to repair log after $RetryCount attempts. Repair will continue. Log='$($script:LogPath)'; Error=$($_.Exception.Message)"
            return
        }
        finally {
            if ($writer) {
                $writer.Dispose()
            }
            elseif ($stream) {
                $stream.Dispose()
            }
        }
    }
}

function Write-RepairLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] [{2}] {3}' -f (Get-Date), $script:ComputerName, $Level, $Message

    switch ($Level) {
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        default   { Write-Host $line -ForegroundColor Cyan }
    }

    if ($script:LogPath) {
        Add-RepairLogContent -Line $line
    }
}

function Write-Section {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $separator = '=' * 72
    Write-Host ''
    Write-Host $separator -ForegroundColor Magenta
    Write-Host ('  {0}' -f $Name) -ForegroundColor Magenta
    Write-Host $separator -ForegroundColor Magenta

    if ($script:LogPath) {
        Add-RepairLogContent -Line @('', $separator, "  $Name", $separator)
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Find-DeepFreezeCommandLine {
    [CmdletBinding()]
    param()

    $candidates = @(
        (Join-Path $env:WINDIR 'SysWOW64\DFC.exe'),
        (Join-Path $env:WINDIR 'System32\DFC.exe')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    $command = Get-Command 'DFC.exe' -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    return $null
}

function Get-DeepFreezeState {
    [CmdletBinding()]
    param()

    $service = Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -eq 'DFServ' -or
            $_.DisplayName -match '(?i)Deep Freeze|Faronics'
        } |
        Select-Object -First 1
    $dfcPath = Find-DeepFreezeCommandLine
    $installed = ($null -ne $service -or -not [string]::IsNullOrWhiteSpace($dfcPath))

    if (-not $installed) {
        return [pscustomobject]@{
            Installed = $false
            State     = 'NotInstalled'
            DfcPath   = $null
            ExitCode  = $null
            Output    = @()
            Service   = $null
        }
    }

    if ([string]::IsNullOrWhiteSpace($dfcPath)) {
        return [pscustomobject]@{
            Installed = $true
            State     = 'Unknown'
            DfcPath   = $null
            ExitCode  = $null
            Output    = @('DFC.exe was not found.')
            Service   = $service
        }
    }

    $output = @(& $dfcPath get /ISFROZEN 2>&1 | ForEach-Object { [string]$_ })
    $exitCode = $LASTEXITCODE
    $state = switch ($exitCode) {
        0       { 'Thawed' }
        1       { 'Frozen' }
        default { 'Unknown' }
    }

    return [pscustomobject]@{
        Installed = $true
        State     = $state
        DfcPath   = $dfcPath
        ExitCode  = $exitCode
        Output    = $output
        Service   = $service
    }
}

function ConvertTo-ArchitectureName {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        $Architecture
    )

    $value = [string]$Architecture
    switch -Regex ($value) {
        '^(9|AMD64|x64|64-bit)$' { return 'x64' }
        '^(12|ARM64)$'           { return 'arm64' }
        '^(0|x86|32-bit)$'       { return 'x86' }
        default                  { return $value.ToLowerInvariant() }
    }
}

function Get-PropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [AllowNull()]
        $DefaultValue = $null
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $DefaultValue
    }

    return $property.Value
}

function Get-DismExecutable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x64', 'arm64', 'x86')]
        [string]$TargetArchitecture,

        [Parameter(Mandatory)]
        [version]$MinimumServicingVersion,

        [Parameter()]
        [AllowEmptyString()]
        [string]$RequestedPath
    )

    $candidatePaths = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        $candidatePaths.Add($RequestedPath)
    }

    $adkRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\Assessment and Deployment Kit\Deployment Tools'
    switch ($TargetArchitecture) {
        'x64' {
            $candidatePaths.Add((Join-Path $adkRoot 'amd64\DISM\dism.exe'))
            $candidatePaths.Add((Join-Path $env:SystemRoot 'System32\Dism.exe'))
        }
        'arm64' {
            $candidatePaths.Add((Join-Path $adkRoot 'arm64\DISM\dism.exe'))
            $candidatePaths.Add((Join-Path $env:SystemRoot 'System32\Dism.exe'))
        }
        default {
            $candidatePaths.Add((Join-Path $adkRoot 'x86\DISM\dism.exe'))
            $candidatePaths.Add((Join-Path $env:SystemRoot 'System32\Dism.exe'))
        }
    }

    $candidates = @(
        foreach ($path in @($candidatePaths | Select-Object -Unique)) {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                continue
            }

            $item = Get-Item -LiteralPath $path -ErrorAction Stop
            $versionText = [string]$item.VersionInfo.FileVersion
            $versionMatch = [regex]::Match($versionText, '\d+\.\d+\.\d+\.\d+')
            if (-not $versionMatch.Success) {
                Write-RepairLog -Message "Ignoring DISM candidate with an unreadable version: $path ($versionText)" -Level WARNING
                continue
            }

            [pscustomobject]@{
                Path       = $item.FullName
                Version    = [version]$versionMatch.Value
                IsExplicit = (-not [string]::IsNullOrWhiteSpace($RequestedPath) -and $item.FullName -ieq $RequestedPath)
            }
        }
    )

    foreach ($candidate in $candidates) {
        Write-RepairLog -Message "DISM candidate: Path='$($candidate.Path)'; Version='$($candidate.Version)'"
    }

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath) -and -not ($candidates | Where-Object IsExplicit)) {
        throw "The requested DISM executable is unavailable or has an unreadable version: $RequestedPath"
    }

    $compatibleCandidates = @(
        $candidates | Where-Object { $_.Version -ge $MinimumServicingVersion }
    )
    if ($compatibleCandidates.Count -eq 0) {
        $found = if ($candidates.Count -gt 0) {
            ($candidates | ForEach-Object { "$($_.Path) [$($_.Version)]" }) -join '; '
        }
        else {
            'none'
        }
        throw "No compatible $TargetArchitecture DISM executable was found. Required version is at least $MinimumServicingVersion; found: $found. Install or update the Windows ADK Deployment Tools."
    }

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        $selected = $compatibleCandidates | Where-Object IsExplicit | Select-Object -First 1
    }
    else {
        $selected = $compatibleCandidates | Sort-Object Version -Descending | Select-Object -First 1
    }

    Write-RepairLog -Message "Selected DISM executable: $($selected.Path) (Version $($selected.Version))" -Level SUCCESS
    return $selected
}

function Copy-SharedAdkDism {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x64', 'arm64', 'x86')]
        [string]$TargetArchitecture,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SharePath
    )

    $adkArchitecture = switch ($TargetArchitecture) {
        'x64'   { 'amd64' }
        'arm64' { 'arm64' }
        default { 'x86' }
    }

    $sourceDismDirectory = Join-Path $SharePath "$adkArchitecture\DISM"
    $sourceDismExecutable = Join-Path $sourceDismDirectory 'dism.exe'
    if (-not (Test-Path -LiteralPath $sourceDismExecutable -PathType Leaf)) {
        throw "The shared $adkArchitecture ADK DISM executable is unavailable: $sourceDismExecutable"
    }

    $script:StagedAdkRoot = Join-Path 'C:\Temp' ("Compton-ADK-DISM-{0}" -f ([guid]::NewGuid().ToString('N')))
    $destinationDismDirectory = Join-Path $script:StagedAdkRoot "$adkArchitecture\DISM"
    New-Item -ItemType Directory -Path $destinationDismDirectory -Force | Out-Null

    Write-RepairLog -Message "Staging $adkArchitecture ADK DISM from '$sourceDismDirectory' to '$destinationDismDirectory'."
    & robocopy.exe $sourceDismDirectory $destinationDismDirectory /E /COPY:DAT /DCOPY:DAT /R:2 /W:3 /NP /NFL /NDL 2>&1 |
        ForEach-Object {
            $line = [string]$_
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Add-RepairLogContent -Line $line
            }
        }
    $robocopyExitCode = $LASTEXITCODE

    if ($robocopyExitCode -gt 7) {
        throw "Failed to stage ADK DISM from the file share. RobocopyExitCode=$robocopyExitCode"
    }

    $stagedDismExecutable = Join-Path $destinationDismDirectory 'dism.exe'
    if (-not (Test-Path -LiteralPath $stagedDismExecutable -PathType Leaf)) {
        throw "The staged ADK copy did not contain DISM.exe: $stagedDismExecutable"
    }

    $sourceHash = (Get-FileHash -LiteralPath $sourceDismExecutable -Algorithm SHA256).Hash
    $stagedHash = (Get-FileHash -LiteralPath $stagedDismExecutable -Algorithm SHA256).Hash
    if ($sourceHash -ne $stagedHash) {
        throw 'The staged DISM.exe failed SHA-256 verification against the shared source.'
    }

    $stagedVersion = (Get-Item -LiteralPath $stagedDismExecutable).VersionInfo.FileVersion
    Write-RepairLog -Message "Shared ADK DISM was staged and verified. Version=$stagedVersion; RobocopyExitCode=$robocopyExitCode" -Level SUCCESS
    return $stagedDismExecutable
}

function Remove-StagedAdkDism {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($script:StagedAdkRoot)) {
        return
    }

    if (-not (Test-Path -LiteralPath $script:StagedAdkRoot)) {
        $script:StagedAdkRoot = $null
        return
    }

    try {
        Remove-Item -LiteralPath $script:StagedAdkRoot -Recurse -Force -ErrorAction Stop
        Write-RepairLog -Message "Removed staged ADK files: $($script:StagedAdkRoot)" -Level SUCCESS
        $script:StagedAdkRoot = $null
    }
    catch {
        Write-RepairLog -Message "Unable to remove staged ADK files '$($script:StagedAdkRoot)': $($_.Exception.Message)" -Level WARNING
    }
}

function Invoke-LoggedNativeCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [string[]]$ArgumentList,

        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter()]
        [int[]]$SuccessExitCodes = @(0, 3010),

        [Parameter()]
        [switch]$AllowFailure
    )

    Write-RepairLog -Message "Starting: $Description"
    Write-RepairLog -Message ("Command: {0} {1}" -f $FilePath, ($ArgumentList -join ' '))

    $commandOutput = @(& $FilePath @ArgumentList 2>&1)
    $exitCode = $LASTEXITCODE

    foreach ($outputLine in $commandOutput) {
        $text = [string]$outputLine
        Write-Host $text
        Add-RepairLogContent -Line $text
    }

    if ($exitCode -notin $SuccessExitCodes) {
        if ($AllowFailure) {
            Write-RepairLog -Message "$Description failed with exit code $exitCode; returning the failure to the repair workflow." -Level WARNING
            return [pscustomobject]@{
                Description = $Description
                ExitCode     = $exitCode
                Output       = $commandOutput
                Succeeded    = $false
            }
        }

        throw "$Description failed with exit code $exitCode. See $script:LogPath."
    }

    Write-RepairLog -Message "$Description completed successfully. ExitCode=$exitCode" -Level SUCCESS
    return [pscustomobject]@{
        Description = $Description
        ExitCode     = $exitCode
        Output       = $commandOutput
        Succeeded    = $true
    }
}

function Get-OsVolumeFreeSpaceGB {
    $systemDrive = [string]$env:SystemDrive
    $volume = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction Stop
    return [math]::Round(([double]$volume.FreeSpace / 1GB), 2)
}

function Get-DirectorySizeBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $total = 0L
    Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction Stop |
        ForEach-Object { $total += [long]$_.Length }
    return $total
}

function Suspend-ComptonMaintenanceTasks {
    [CmdletBinding()]
    param()

    $excludedNames = @(
        'Compton - Complete Component Store Repair'
    )
    $disabledTasks = New-Object System.Collections.Generic.List[object]
    $tasks = @(Get-ScheduledTask -ErrorAction Stop)

    foreach ($task in $tasks) {
        if ([string]$task.State -eq 'Disabled' -or $task.TaskName -in $excludedNames) {
            continue
        }

        $actionText = @(
            foreach ($action in @($task.Actions)) {
                '{0} {1}' -f [string]$action.Execute, [string]$action.Arguments
            }
        ) -join ' '

        $isComptonMaintenanceTask = (
            $actionText -match '(?i)C:\\Scripts\\|\\\\filesvr\\labscripts\\' -or
            $task.Description -match '(?i)Compton.*maintenance|lab.*maintenance'
        )
        if (-not $isComptonMaintenanceTask) {
            continue
        }

        try {
            Disable-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop | Out-Null
            $disabledTasks.Add([pscustomobject]@{
                TaskName = [string]$task.TaskName
                TaskPath = [string]$task.TaskPath
            }) | Out-Null
            Write-RepairLog -Message "Temporarily disabled maintenance task: $($task.TaskPath)$($task.TaskName)" -Level SUCCESS
        }
        catch {
            throw "Unable to suspend maintenance task '$($task.TaskPath)$($task.TaskName)': $($_.Exception.Message)"
        }
    }

    return @($disabledTasks)
}

function Restore-ComptonMaintenanceTasks {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        $TaskRecords
    )

    foreach ($record in @($TaskRecords)) {
        $taskName = [string](Get-PropertyValue -InputObject $record -Name 'TaskName' -DefaultValue '')
        $taskPath = [string](Get-PropertyValue -InputObject $record -Name 'TaskPath' -DefaultValue '\')
        if ([string]::IsNullOrWhiteSpace($taskName)) {
            continue
        }

        try {
            $task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue
            if ($task -and [string]$task.State -eq 'Disabled') {
                Enable-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction Stop | Out-Null
            }
            Write-RepairLog -Message "Restored maintenance task: $taskPath$taskName" -Level SUCCESS
        }
        catch {
            Write-RepairLog -Message "Unable to restore maintenance task '$taskPath$taskName': $($_.Exception.Message)" -Level WARNING
        }
    }
}

function Save-InPlaceRepairState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Phase,

        [Parameter()]
        [hashtable]$AdditionalData = @{}
    )

    New-Item -ItemType Directory -Path $script:StateDirectory -Force | Out-Null
    $priorState = Get-InPlaceRepairState
    $priorDisabledTasks = if ($priorState) {
        Get-PropertyValue -InputObject $priorState -Name 'DisabledMaintenanceTasks' -DefaultValue @()
    }
    else {
        @()
    }

    $state = [ordered]@{
        SchemaVersion       = '1.0'
        ScriptVersion       = $script:ScriptVersion
        ComputerName        = $script:ComputerName
        Phase               = $Phase
        UpdatedUtc          = (Get-Date).ToUniversalTime().ToString('o')
        LaunchBootTimeUtc   = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().ToString('o')
        LocalSetupMediaPath = $LocalSetupMediaPath
        LogPath             = $script:LogPath
        DisabledMaintenanceTasks = @($priorDisabledTasks)
    }

    foreach ($key in $AdditionalData.Keys) {
        $state[$key] = $AdditionalData[$key]
    }

    $state | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:StatePath -Encoding UTF8 -Force
}

function Get-InPlaceRepairState {
    if (-not (Test-Path -LiteralPath $script:StatePath -PathType Leaf)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $script:StatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-RepairLog -Message "Ignoring unreadable in-place repair state file '$($script:StatePath)': $($_.Exception.Message)" -Level WARNING
        return $null
    }
}

function Remove-InPlaceRepairArtifacts {
    $repairState = Get-InPlaceRepairState
    if ($repairState) {
        $disabledTasks = Get-PropertyValue -InputObject $repairState -Name 'DisabledMaintenanceTasks' -DefaultValue @()
        Restore-ComptonMaintenanceTasks -TaskRecords $disabledTasks
    }

    $resumeTaskName = 'Compton - Complete Component Store Repair'
    if (Get-ScheduledTask -TaskName $resumeTaskName -ErrorAction SilentlyContinue) {
        try {
            Unregister-ScheduledTask -TaskName $resumeTaskName -Confirm:$false -ErrorAction Stop
            Write-RepairLog -Message "Removed post-repair verification task: $resumeTaskName" -Level SUCCESS
        }
        catch {
            Write-RepairLog -Message "Unable to remove post-repair verification task '$resumeTaskName': $($_.Exception.Message)" -Level WARNING
        }
    }

    if (-not $KeepLocalSetupMedia -and (Test-Path -LiteralPath $LocalSetupMediaPath)) {
        try {
            $normalizedPath = [System.IO.Path]::GetFullPath($LocalSetupMediaPath).TrimEnd('\')
            if ($normalizedPath -notlike 'C:\Temp\Compton-Windows11-25H2-Repair*') {
                throw "Refusing to remove unexpected local setup-media path: $normalizedPath"
            }

            Remove-Item -LiteralPath $normalizedPath -Recurse -Force -ErrorAction Stop
            Write-RepairLog -Message "Removed locally staged Windows setup media: $normalizedPath" -Level SUCCESS
        }
        catch {
            Write-RepairLog -Message "Unable to remove locally staged Windows setup media: $($_.Exception.Message)" -Level WARNING
        }
    }

    if (Test-Path -LiteralPath $script:StatePath -PathType Leaf) {
        Remove-Item -LiteralPath $script:StatePath -Force -ErrorAction SilentlyContinue
    }
}

function Register-InPlaceRepairVerificationTask {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($PSCommandPath) -or -not (Test-Path -LiteralPath $PSCommandPath -PathType Leaf)) {
        throw 'The current script path could not be resolved, so post-repair verification cannot be scheduled.'
    }

    $taskName = 'Compton - Complete Component Store Repair'
    $powerShellArguments = @(
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $PSCommandPath)
    )
    if ($SkipSfc) {
        $powerShellArguments += '-SkipSfc'
    }

    $action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument ($powerShellArguments -join ' ')
    $startupTrigger = New-ScheduledTaskTrigger -AtStartup
    $startupTrigger.Delay = 'PT20M'
    $retryTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddHours(1) `
        -RepetitionInterval (New-TimeSpan -Hours 1) `
        -RepetitionDuration (New-TimeSpan -Days 1)
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 6) `
        -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 15) -MultipleInstances IgnoreNew

    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($startupTrigger, $retryTrigger) `
        -Principal $principal -Settings $settings -Description 'Verifies DISM and SFC after an automated Windows in-place component-store repair.' `
        -Force | Out-Null
    Write-RepairLog -Message "Registered post-repair verification task: $taskName" -Level SUCCESS
}

function Test-AndCompletePendingInPlaceRepair {
    [CmdletBinding()]
    param()

    $state = Get-InPlaceRepairState
    if ($null -eq $state -or [string](Get-PropertyValue -InputObject $state -Name 'Phase' -DefaultValue '') -ne 'InPlaceRepairLaunched') {
        return $false
    }

    Write-Section -Name 'Pending In-Place Repair Verification'
    $launchBootText = [string](Get-PropertyValue -InputObject $state -Name 'LaunchBootTimeUtc' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($launchBootText)) {
        throw "The pending in-place repair state is missing LaunchBootTimeUtc: $($script:StatePath)"
    }
    $launchBootTime = [datetime]::Parse($launchBootText).ToUniversalTime()
    $currentBootTime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime()
    if ($currentBootTime -le $launchBootTime) {
        Write-RepairLog -Message 'An in-place repair was already launched during the current Windows boot. Refusing to launch another repair cycle.' -Level WARNING
        Write-RepairLog -Message "Pending state file: $($script:StatePath)" -Level WARNING
        exit 3010
    }

    $setupProcesses = @(Get-Process -Name setup, setuphost, setupprep, SetupPlatform -ErrorAction SilentlyContinue)
    if ($setupProcesses.Count -gt 0) {
        Write-RepairLog -Message 'Windows Setup is still running. Post-repair verification is deferred until the next script run.' -Level WARNING
        exit 3010
    }

    Write-RepairLog -Message "A newer Windows boot was detected after the in-place repair launch. LaunchBoot=$launchBootTime; CurrentBoot=$currentBootTime" -Level SUCCESS
    $scan = Invoke-LoggedNativeCommand -FilePath (Join-Path $env:SystemRoot 'System32\Dism.exe') -ArgumentList @(
        '/Online', '/Cleanup-Image', '/ScanHealth'
    ) -Description 'Post-upgrade DISM ScanHealth' -AllowFailure

    if (-not $scan.Succeeded) {
        Save-InPlaceRepairState -Phase 'PostRepairVerificationFailed' -AdditionalData @{
            VerificationExitCode = $scan.ExitCode
        }
        throw "The in-place repair completed a reboot, but DISM verification still failed. ExitCode=$($scan.ExitCode)"
    }

    if (-not $SkipSfc) {
        $sfc = Invoke-LoggedNativeCommand -FilePath 'sfc.exe' -ArgumentList @('/scannow') -Description 'Post-upgrade System File Checker' -AllowFailure
        if (-not $sfc.Succeeded) {
            Save-InPlaceRepairState -Phase 'PostRepairVerificationFailed' -AdditionalData @{
                VerificationExitCode = $sfc.ExitCode
            }
            throw "The in-place repair completed, but SFC verification failed. ExitCode=$($sfc.ExitCode)"
        }
    }

    Remove-InPlaceRepairArtifacts
    Write-RepairLog -Message 'The in-place repair reboot and component-store verification completed successfully.' -Level SUCCESS
    return $true
}

function Copy-WindowsSetupMedia {
    [CmdletBinding()]
    param()

    $requiredMediaFiles = @(
        (Join-Path $SetupMediaShare 'setup.exe'),
        (Join-Path $SetupMediaShare 'sources\install.wim'),
        (Join-Path $SetupMediaShare 'sources\boot.wim')
    )
    foreach ($requiredFile in $requiredMediaFiles) {
        if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
            throw "The expanded Windows setup media is incomplete. Required file is missing: $requiredFile"
        }
    }

    $freeSpaceGB = Get-OsVolumeFreeSpaceGB
    $sourceSizeBytes = Get-DirectorySizeBytes -Path $SetupMediaShare
    $sourceSizeGB = [math]::Round(($sourceSizeBytes / 1GB), 2)
    $requiredFreeSpaceGB = [math]::Max($MinimumFreeSpaceGB, [math]::Ceiling($sourceSizeGB + 30))
    Write-RepairLog -Message "Setup media size: $sourceSizeGB GB. Free space on $($env:SystemDrive): $freeSpaceGB GB. Required: $requiredFreeSpaceGB GB."
    if ($freeSpaceGB -lt $requiredFreeSpaceGB) {
        throw "Insufficient free space for an in-place repair. Free=$freeSpaceGB GB; Required=$requiredFreeSpaceGB GB."
    }

    if (Test-Path -LiteralPath $LocalSetupMediaPath) {
        $normalizedPath = [System.IO.Path]::GetFullPath($LocalSetupMediaPath).TrimEnd('\')
        if ($normalizedPath -notlike 'C:\Temp\Compton-Windows11-25H2-Repair*') {
            throw "Refusing to refresh unexpected local setup-media path: $normalizedPath"
        }
    }
    else {
        New-Item -ItemType Directory -Path $LocalSetupMediaPath -Force | Out-Null
    }

    Write-RepairLog -Message "Copying expanded Windows setup media from '$SetupMediaShare' to '$LocalSetupMediaPath'."
    & robocopy.exe $SetupMediaShare $LocalSetupMediaPath /MIR /Z /COPY:DAT /DCOPY:DAT /R:3 /W:5 /XJ /NP 2>&1 |
        ForEach-Object {
            $line = [string]$_
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Add-RepairLogContent -Line $line
            }
        }
    $copyExitCode = $LASTEXITCODE
    if ($copyExitCode -gt 7) {
        throw "Windows setup media copy failed. RobocopyExitCode=$copyExitCode"
    }

    foreach ($relativePath in @('setup.exe', 'sources\install.wim', 'sources\boot.wim')) {
        $sourceFile = Join-Path $SetupMediaShare $relativePath
        $localFile = Join-Path $LocalSetupMediaPath $relativePath
        if (-not (Test-Path -LiteralPath $localFile -PathType Leaf)) {
            throw "Staged setup media is missing required file: $localFile"
        }

        $sourceLength = (Get-Item -LiteralPath $sourceFile).Length
        $localLength = (Get-Item -LiteralPath $localFile).Length
        if ($sourceLength -ne $localLength) {
            throw "Staged setup-media size verification failed for '$relativePath'."
        }
    }

    $sourceWimHash = (Get-FileHash -LiteralPath (Join-Path $SetupMediaShare 'sources\install.wim') -Algorithm SHA256).Hash
    $localWimHash = (Get-FileHash -LiteralPath (Join-Path $LocalSetupMediaPath 'sources\install.wim') -Algorithm SHA256).Hash
    if ($sourceWimHash -ne $localWimHash) {
        throw 'The locally staged install.wim failed SHA-256 verification.'
    }

    Write-RepairLog -Message "Windows setup media was copied and verified. RobocopyExitCode=$copyExitCode; InstallWimSHA256=$localWimHash" -Level SUCCESS
}

function Invoke-InPlaceRepairUpgrade {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [version]$InstalledVersion,

        [Parameter(Mandatory)]
        [version]$MediaVersion,

        [Parameter(Mandatory)]
        [object]$DeepFreezeStatus
    )

    if (-not $EnableInPlaceRepair) {
        throw 'Normal component-store repair failed and automatic in-place repair is disabled.'
    }

    if ($DeepFreezeStatus.Installed -and [string]$DeepFreezeStatus.State -ne 'Thawed') {
        throw "Deep Freeze must report Thawed before an in-place repair can start. CurrentState=$($DeepFreezeStatus.State); DfcExitCode=$($DeepFreezeStatus.ExitCode)"
    }

    if ($MediaVersion -lt $InstalledVersion) {
        throw "The setup-media version '$MediaVersion' is older than installed Windows '$InstalledVersion'. An in-place repair that preserves applications cannot be started."
    }

    Write-Section -Name 'Automatic In-Place Repair Escalation'
    Write-RepairLog -Message 'DISM could not obtain the required repair payload. Escalating to a Windows in-place repair that preserves applications, files, and settings.' -Level WARNING
    Write-RepairLog -Message "Expanded setup-media share: $SetupMediaShare"
    Remove-StagedAdkDism
    Copy-WindowsSetupMedia

    $bitLockerSuspended = $false
    $bitLockerVolume = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction SilentlyContinue
    if ($bitLockerVolume -and [string]$bitLockerVolume.ProtectionStatus -eq 'On') {
        Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount 3 -ErrorAction Stop
        $bitLockerSuspended = $true
        Write-RepairLog -Message 'BitLocker protection was suspended for three reboots.' -Level SUCCESS
    }

    $setupExe = Join-Path $LocalSetupMediaPath 'setup.exe'
    $setupCopyLogs = Join-Path $LogDirectory ("{0}-Windows-Setup-{1}" -f $script:ComputerName, (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
    New-Item -ItemType Directory -Path $setupCopyLogs -Force | Out-Null

    Write-RepairLog -Message 'Suspending Compton maintenance scheduled tasks so Windows Setup is not interrupted.' -Level WARNING
    $disabledMaintenanceTasks = @(Suspend-ComptonMaintenanceTasks)

    Save-InPlaceRepairState -Phase 'InPlaceRepairLaunched' -AdditionalData @{
        InstalledVersion   = [string]$InstalledVersion
        MediaVersion       = [string]$MediaVersion
        SetupCopyLogs      = $setupCopyLogs
        BitLockerSuspended = $bitLockerSuspended
        DisabledMaintenanceTasks = @($disabledMaintenanceTasks)
    }
    Register-InPlaceRepairVerificationTask

    $setupArguments = @(
        '/auto', 'upgrade',
        '/quiet',
        '/eula', 'accept',
        '/dynamicupdate', 'enable',
        '/showoobe', 'none',
        '/copylogs', $setupCopyLogs
    )

    Write-RepairLog -Message "Launching Windows Setup locally. Command: $setupExe $($setupArguments -join ' ')" -Level WARNING
    Write-RepairLog -Message 'Windows Setup will preserve installed applications and files and may restart the computer multiple times. Do not refreeze or power off the computer.' -Level WARNING

    $process = Start-Process -FilePath $setupExe -ArgumentList $setupArguments -PassThru -Wait
    $setupExitCode = $process.ExitCode
    if ($setupExitCode -notin @(0, 3010)) {
        Save-InPlaceRepairState -Phase 'InPlaceRepairLaunchFailed' -AdditionalData @{
            SetupExitCode = $setupExitCode
            SetupCopyLogs = $setupCopyLogs
            DisabledMaintenanceTasks = @($disabledMaintenanceTasks)
        }
        Restore-ComptonMaintenanceTasks -TaskRecords $disabledMaintenanceTasks
        if ($bitLockerSuspended) {
            Resume-BitLocker -MountPoint $env:SystemDrive -ErrorAction SilentlyContinue
        }
        throw "Windows Setup exited before completing the in-place repair. ExitCode=$setupExitCode; SetupLogs=$setupCopyLogs"
    }

    Write-RepairLog -Message "Windows Setup accepted the in-place repair. ExitCode=$setupExitCode. Post-repair verification will occur the next time this script runs." -Level SUCCESS
    exit 3010
}

try {
    if (-not (Test-IsAdministrator)) {
        throw 'This script must be run from an elevated Windows PowerShell session.'
    }

    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $script:LogPath = Join-Path $LogDirectory "$($script:ComputerName)-Repair-Windows-ComponentStore-$timestamp.log"
    New-Item -ItemType File -Path $script:LogPath -Force | Out-Null

    Write-Section -Name 'Initialization'
    Write-RepairLog -Message "Starting component-store repair version $($script:ScriptVersion)."
    Write-RepairLog -Message "Repair source: $ImagePath"
    Write-RepairLog -Message "Expanded Windows setup media: $SetupMediaShare"
    Write-RepairLog -Message "Shared ADK Deployment Tools source: $AdkSharePath"
    Write-RepairLog -Message "Log file: $($script:LogPath)"

    Write-Section -Name 'Faronics Deep Freeze State'
    $deepFreezeStatus = Get-DeepFreezeState
    switch ([string]$deepFreezeStatus.State) {
        'NotInstalled' {
            Write-RepairLog -Message 'Deep Freeze is not installed. Repair may proceed.' -Level INFO
        }
        'Thawed' {
            Write-RepairLog -Message "DFC.exe confirms that this workstation is Thawed. DfcPath='$($deepFreezeStatus.DfcPath)'; DfcExitCode=$($deepFreezeStatus.ExitCode)" -Level SUCCESS
        }
        'Frozen' {
            Write-RepairLog -Message "DFC.exe reports that this workstation is Frozen. DfcPath='$($deepFreezeStatus.DfcPath)'; DfcExitCode=$($deepFreezeStatus.ExitCode)" -Level ERROR
            throw 'Deep Freeze is Frozen. Thaw and restart the workstation before running component-store repair; otherwise repairs would be discarded.'
        }
        default {
            $dfcOutputText = (@($deepFreezeStatus.Output) -join ' | ').Trim()
            if ([string]::IsNullOrWhiteSpace($dfcOutputText)) {
                $dfcOutputText = 'No DFC output was returned.'
            }
            Write-RepairLog -Message "Deep Freeze state could not be verified. DfcPath='$($deepFreezeStatus.DfcPath)'; DfcExitCode=$($deepFreezeStatus.ExitCode); Output=$dfcOutputText" -Level ERROR
            throw 'Deep Freeze appears to be installed, but DFC.exe did not return a recognized Frozen/Thawed result. Repair is blocked for safety.'
        }
    }

    if (Test-AndCompletePendingInPlaceRepair) {
        Write-Section -Name 'Repair Summary'
        Write-RepairLog -Message 'Post-upgrade component-store verification completed successfully.' -Level SUCCESS
        exit 0
    }

    if (-not (Test-Path -LiteralPath $ImagePath -PathType Leaf)) {
        throw "The repair image is unavailable: $ImagePath. Verify network access and share permissions."
    }

    Write-Section -Name 'Installed Windows Detection'
    $installedEdition = (Get-WindowsEdition -Online -ErrorAction Stop).Edition
    $installedArchitecture = ConvertTo-ArchitectureName -Architecture (Get-CimInstance Win32_OperatingSystem).OSArchitecture
    $installedLanguage = (Get-WinSystemLocale).Name
    $currentVersion = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    $installedBuild = [version]("10.0.{0}.{1}" -f $currentVersion.CurrentBuildNumber, $currentVersion.UBR)

    Write-RepairLog -Message "Installed edition: $installedEdition"
    Write-RepairLog -Message "Installed architecture: $installedArchitecture"
    Write-RepairLog -Message "Installed system language: $installedLanguage"
    Write-RepairLog -Message "Installed Windows version: $installedBuild"

    Write-Section -Name 'Automatic WIM Index Selection'
    $imageSummaries = @(Get-WindowsImage -ImagePath $ImagePath -ErrorAction Stop)

    if ($imageSummaries.Count -eq 0) {
        throw "No Windows images were found in $ImagePath."
    }

    # The summary objects returned by Get-WindowsImage do not consistently
    # contain EditionId, DefaultLanguage, or Version on every Windows build.
    # Query each index individually to obtain its complete metadata.
    $allImages = @(
        foreach ($summary in $imageSummaries) {
            $summaryIndex = Get-PropertyValue -InputObject $summary -Name 'ImageIndex'
            if ($null -eq $summaryIndex) {
                throw 'A WIM image summary did not contain an ImageIndex.'
            }

            Get-WindowsImage -ImagePath $ImagePath -Index ([int]$summaryIndex) -ErrorAction Stop
        }
    )

    foreach ($image in $allImages) {
        Write-RepairLog -Message ("Available index {0}: Name='{1}'; EditionId='{2}'; Architecture='{3}'; Language='{4}'; Version='{5}'" -f
            (Get-PropertyValue -InputObject $image -Name 'ImageIndex' -DefaultValue 'Unknown'),
            (Get-PropertyValue -InputObject $image -Name 'ImageName' -DefaultValue 'Unknown'),
            (Get-PropertyValue -InputObject $image -Name 'EditionId' -DefaultValue 'Unknown'),
            (ConvertTo-ArchitectureName -Architecture (Get-PropertyValue -InputObject $image -Name 'Architecture' -DefaultValue 'Unknown')),
            (Get-PropertyValue -InputObject $image -Name 'DefaultLanguage' -DefaultValue 'Unknown'),
            (Get-PropertyValue -InputObject $image -Name 'Version' -DefaultValue '0.0'))
    }

    $candidates = @(
        $allImages | Where-Object {
            [string](Get-PropertyValue -InputObject $_ -Name 'EditionId' -DefaultValue '') -ieq [string]$installedEdition
        }
    )

    if ($candidates.Count -eq 0) {
        $availableEditions = @(
            $allImages |
                ForEach-Object { Get-PropertyValue -InputObject $_ -Name 'EditionId' -DefaultValue 'Unknown' } |
                Sort-Object -Unique
        ) -join ', '
        throw "No WIM index matches installed edition '$installedEdition'. Available editions: $availableEditions"
    }

    $architectureMatches = @(
        $candidates | Where-Object {
            (ConvertTo-ArchitectureName -Architecture (Get-PropertyValue -InputObject $_ -Name 'Architecture' -DefaultValue 'Unknown')) -eq $installedArchitecture
        }
    )
    if ($architectureMatches.Count -gt 0) {
        $candidates = $architectureMatches
    }

    $languageMatches = @(
        $candidates | Where-Object {
            [string](Get-PropertyValue -InputObject $_ -Name 'DefaultLanguage' -DefaultValue '') -ieq [string]$installedLanguage
        }
    )
    if ($languageMatches.Count -gt 0) {
        $candidates = $languageMatches
    }

    if ($candidates.Count -gt 1) {
        Write-RepairLog -Message "Multiple indexes matched. Selecting the highest image version, then the lowest image index." -Level WARNING
    }

    $selectedImage = $candidates |
        Sort-Object -Property @{
            Expression = {
                [version](Get-PropertyValue -InputObject $_ -Name 'Version' -DefaultValue '0.0')
            }
            Descending = $true
        }, @{
            Expression = {
                [int](Get-PropertyValue -InputObject $_ -Name 'ImageIndex' -DefaultValue ([int]::MaxValue))
            }
            Descending = $false
        } |
        Select-Object -First 1

    if (-not $selectedImage) {
        throw 'Automatic WIM index selection did not return an image.'
    }

    $selectedIndex = [int](Get-PropertyValue -InputObject $selectedImage -Name 'ImageIndex')
    $selectedName = [string](Get-PropertyValue -InputObject $selectedImage -Name 'ImageName' -DefaultValue 'Unknown')
    $selectedEdition = [string](Get-PropertyValue -InputObject $selectedImage -Name 'EditionId' -DefaultValue 'Unknown')
    $selectedLanguage = [string](Get-PropertyValue -InputObject $selectedImage -Name 'DefaultLanguage' -DefaultValue 'Unknown')
    $selectedVersion = [version](Get-PropertyValue -InputObject $selectedImage -Name 'Version' -DefaultValue '0.0')
    $selectedArchitecture = ConvertTo-ArchitectureName -Architecture (Get-PropertyValue -InputObject $selectedImage -Name 'Architecture' -DefaultValue 'Unknown')
    Write-RepairLog -Message ("Selected index {0}: Name='{1}'; EditionId='{2}'; Architecture='{3}'; Language='{4}'; Version='{5}'" -f
        $selectedIndex,
        $selectedName,
        $selectedEdition,
        $selectedArchitecture,
        $selectedLanguage,
        $selectedVersion) -Level SUCCESS

    if ($selectedVersion -lt [version]$installedBuild) {
        Write-RepairLog -Message "The selected source version '$selectedVersion' is older than installed version '$installedBuild'. DISM may reject the source if required updated components are absent." -Level WARNING
    }

    $sourceArgument = "wim:$ImagePath`:$selectedIndex"

    Write-Section -Name 'Shared ADK Staging'
    if ([string]::IsNullOrWhiteSpace($DismPath)) {
        $DismPath = Copy-SharedAdkDism -TargetArchitecture $installedArchitecture -SharePath $AdkSharePath
    }
    else {
        Write-RepairLog -Message "Using explicitly supplied DISM path; shared ADK staging was skipped: $DismPath"
    }

    Write-Section -Name 'DISM Tool Selection'
    $minimumDismVersion = if ($selectedVersion -gt $installedBuild) { $selectedVersion } else { $installedBuild }
    $selectedDism = Get-DismExecutable -TargetArchitecture $installedArchitecture -MinimumServicingVersion $minimumDismVersion -RequestedPath $DismPath
    $dismExecutable = $selectedDism.Path

    Write-Section -Name 'DISM Component Store Assessment'
    $null = Invoke-LoggedNativeCommand -FilePath $dismExecutable -ArgumentList @(
        '/Online',
        '/Cleanup-Image',
        '/CheckHealth'
    ) -Description 'DISM CheckHealth'

    $null = Invoke-LoggedNativeCommand -FilePath $dismExecutable -ArgumentList @(
        '/Online',
        '/Cleanup-Image',
        '/ScanHealth'
    ) -Description 'DISM initial ScanHealth'

    Write-Section -Name 'DISM Component Store Repair'
    $restoreResult = Invoke-LoggedNativeCommand -FilePath $dismExecutable -ArgumentList @(
        '/Online',
        '/Cleanup-Image',
        '/RestoreHealth',
        "/Source:$sourceArgument",
        '/LimitAccess'
    ) -Description "DISM RestoreHealth using WIM index $selectedIndex" -AllowFailure

    if (-not $restoreResult.Succeeded) {
        $repairContentMissingExitCodes = @(-2146498283, -2146498529)
        if ($restoreResult.ExitCode -notin $repairContentMissingExitCodes) {
            throw "DISM RestoreHealth using WIM index $selectedIndex failed with exit code $($restoreResult.ExitCode). See $script:LogPath."
        }

        if ($DisableWindowsUpdateFallback) {
            throw "The WIM did not contain the required repair content and Windows Update fallback was disabled. ExitCode=$($restoreResult.ExitCode)"
        }

        Write-Section -Name 'Windows Update Repair-Source Fallback'
        Write-RepairLog -Message 'The WIM is valid but does not contain one or more superseded payloads required by this computer. Retrying RestoreHealth with Windows Update permitted as an additional repair source.' -Level WARNING

        foreach ($serviceName in @('cryptsvc', 'bits', 'wuauserv')) {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($null -eq $service) {
                Write-RepairLog -Message "Windows Update fallback prerequisite service was not found: $serviceName" -Level WARNING
                continue
            }

            if ($service.Status -ne 'Running') {
                try {
                    Start-Service -Name $serviceName -ErrorAction Stop
                    Write-RepairLog -Message "Started Windows Update fallback prerequisite service: $serviceName" -Level SUCCESS
                }
                catch {
                    Write-RepairLog -Message "Unable to start Windows Update fallback prerequisite service '$serviceName': $($_.Exception.Message)" -Level WARNING
                }
            }
        }

        $restoreResult = Invoke-LoggedNativeCommand -FilePath $dismExecutable -ArgumentList @(
            '/Online',
            '/Cleanup-Image',
            '/RestoreHealth',
            "/Source:$sourceArgument"
        ) -Description 'DISM RestoreHealth using WIM plus Windows Update fallback' -AllowFailure

        if (-not $restoreResult.Succeeded) {
            Write-RepairLog -Message "DISM RestoreHealth still failed after both source-WIM and Windows Update attempts. ExitCode=$($restoreResult.ExitCode). Escalating to the authorized in-place repair." -Level WARNING
            Invoke-InPlaceRepairUpgrade -InstalledVersion $installedBuild -MediaVersion $selectedVersion -DeepFreezeStatus $deepFreezeStatus
        }
    }

    $sfcClean = $true
    if (-not $SkipSfc) {
        Write-Section -Name 'System File Checker'
        $sfcResult = Invoke-LoggedNativeCommand -FilePath 'sfc.exe' -ArgumentList @('/scannow') -Description 'System File Checker' -AllowFailure
        $sfcText = (@($sfcResult.Output) -join "`n")
        $sfcClean = ($sfcResult.Succeeded -and $sfcText -match '(?i)did not find any integrity violations|found corrupt files and successfully repaired them')
    }
    else {
        Write-RepairLog -Message 'System File Checker was skipped by parameter.' -Level WARNING
    }

    Write-Section -Name 'Final Component Store Verification'
    $finalScan = Invoke-LoggedNativeCommand -FilePath $dismExecutable -ArgumentList @(
        '/Online',
        '/Cleanup-Image',
        '/ScanHealth'
    ) -Description 'DISM final ScanHealth' -AllowFailure

    $finalScanText = (@($finalScan.Output) -join "`n")
    $dismClean = ($finalScan.Succeeded -and $finalScanText -match '(?i)No component store corruption detected|No component store corruption was detected|The component store is repairable\s*:\s*No')
    if (-not $dismClean -or -not $sfcClean) {
        Write-RepairLog -Message "Persistent corruption remains after source repair. DismClean=$dismClean; SfcClean=$sfcClean. Escalating to the authorized in-place repair." -Level WARNING
        Invoke-InPlaceRepairUpgrade -InstalledVersion $installedBuild -MediaVersion $selectedVersion -DeepFreezeStatus $deepFreezeStatus
    }

    $rebootRequired = ($restoreResult.ExitCode -eq 3010)
    $duration = [math]::Round(((Get-Date) - $script:StartTime).TotalMinutes, 2)

    Write-Section -Name 'Repair Summary'
    Write-RepairLog -Message "Selected WIM index: $selectedIndex ($selectedName)" -Level SUCCESS
    Write-RepairLog -Message "Component-store repair sequence completed. DurationMinutes=$duration; RebootRequired=$rebootRequired" -Level SUCCESS

    if ($deepFreezeStatus.Installed) {
        Write-RepairLog -Message 'Restart and verify the repair while the computer remains thawed, then refreeze it.' -Level WARNING
    }
    elseif ($rebootRequired) {
        Write-RepairLog -Message 'DISM requested a restart. Restart the computer before considering the repair complete.' -Level WARNING
    }

    Remove-StagedAdkDism
    exit 0
}
catch {
    if (-not $script:LogPath) {
        Write-Host "Component-store repair failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    else {
        Write-Section -Name 'Repair Failed'
        Write-RepairLog -Message $_.Exception.Message -Level ERROR
        Write-RepairLog -Message "Review the log file: $($script:LogPath)" -Level ERROR
    }

    Remove-StagedAdkDism
    exit 1
}
