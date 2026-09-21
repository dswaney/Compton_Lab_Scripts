#requires -Version 5.1
#requires -RunAsAdministrator
# =====================================================================
# ScriptName:    Repair-MicrosoftEdgeUpdate.ps1
# ScriptVersion: 1.0.0
# LastUpdated:   2026-09-21
# Purpose:       Detection-first Microsoft Edge Update servicing repair for
#                approved automated remediation. This script does not modify
#                Edge browser policies, profiles, favorites, or user data.
# =====================================================================

<#
.SYNOPSIS
    Repairs Microsoft Edge Update servicing and records structured telemetry.

.DESCRIPTION
    Validates the Edge installation, Edge Update policy, updater services, and
    scheduled tasks. When repair is needed, it restores supported service start
    modes, enables existing Edge Update tasks, and invokes the installed Edge
    updater. An optional, trusted Microsoft-signed Enterprise MSI can be used
    when the installed updater is missing or cannot restore servicing.

    The script is intended to be called by an allowlisted maintenance launcher
    during an approved window. It is deliberately not a numbered or scheduled
    maintenance script.

.PARAMETER CorrelationId
    Correlation identifier supplied by n8n or the maintenance launcher.

.PARAMETER FindingFingerprint
    Stable fingerprint of the Script 14 finding that initiated remediation.

.PARAMETER AttemptNumber
    Current automated-remediation attempt number.

.PARAMETER InstallerPath
    Optional path to a Microsoft Edge Enterprise x64 MSI. The MSI is copied to
    a protected local working directory and its Authenticode signature must be
    valid and issued to Microsoft Corporation before it can run.

.PARAMETER ForceInstallerRepair
    Runs the trusted MSI repair even when Edge Update servicing appears present.

.PARAMETER BypassMaintenanceWindow
    Allows an authorized technician to test outside the configured window. The
    bypass is recorded in the structured result.

.OUTPUTS
    Text log, Maintenance-Telemetry.ndjson event, and
    Repair-MicrosoftEdgeUpdate.latest.json under C:\Logs.

.NOTES
    Exit codes:
      0    Healthy or repaired successfully
      1    Repaired with warnings
      10   Repair failed
      20   Blocked by policy or outside the approved maintenance window
      3010 Repair succeeded and a reboot is required
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$CorrelationId,
    [string]$FindingFingerprint,
    [ValidateRange(1, 20)][int]$AttemptNumber = 1,
    [string]$InstallerPath,
    [switch]$ForceInstallerRepair,
    [switch]$BypassMaintenanceWindow,
    [ValidateRange(30, 3600)][int]$UpdaterTimeoutSeconds = 900,
    [ValidateRange(60, 7200)][int]$InstallerTimeoutSeconds = 1800,
    [ValidateRange(0, 300)][int]$PostRepairWaitSeconds = 20,
    [string]$LogDirectory = 'C:\Logs',
    [string]$WorkingRoot = 'C:\ProgramData\Compton\Remediation\MicrosoftEdgeUpdate'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:ScriptName = 'Repair-MicrosoftEdgeUpdate.ps1'
$script:ScriptVersion = '1.0.0'
$script:RemediationClass = 'MicrosoftEdgeUpdateRepair'
$script:StartTime = Get-Date
$script:RunId = [guid]::NewGuid().Guid
$script:CorrelationId = $null
$script:WarningCount = 0
$script:ErrorCount = 0
$script:FailureMessage = $null
$script:Outcome = 'Running'
$script:FinalStatus = 'Running'
$script:FinalExitCode = 10
$script:RebootRequired = $false
$script:Lock = $null
$script:LogSession = $null
$script:WorkingLogPath = $null
$script:PublishedLogPath = $null
$script:RunWorkingDirectory = Join-Path $WorkingRoot $script:RunId
$script:TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$script:LatestPath = Join-Path $LogDirectory 'Repair-MicrosoftEdgeUpdate.latest.json'
$script:Actions = New-Object System.Collections.Generic.List[object]
$script:BeforeState = $null
$script:AfterState = $null
$script:UpdaterResult = $null
$script:InstallerResult = $null

$frameworkCandidates = @(
    (Join-Path $PSScriptRoot 'Maintenance.Framework.psm1'),
    'C:\Scripts\Maintenance.Framework.psm1'
) | Select-Object -Unique
$frameworkPath = $frameworkCandidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

if (-not $frameworkPath) {
    throw 'Maintenance.Framework.psm1 was not found beside the script or in C:\Scripts.'
}

Import-Module -Name $frameworkPath -Force -DisableNameChecking -ErrorAction Stop
$maintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory
$script:CorrelationId = Get-MaintenanceCorrelationId -CorrelationId $CorrelationId

function Write-RepairLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $line = '{0} [{1}] [{2}] [CorrelationId:{3}] {4}' -f `
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        $env:COMPUTERNAME,
        $Level,
        $script:CorrelationId,
        $Message

    $color = switch ($Level) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host $line -ForegroundColor $color

    if ($script:WorkingLogPath) {
        Add-Content -LiteralPath $script:WorkingLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    }
}

function Add-RepairAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][ValidateSet('Succeeded','Failed','Skipped','NoChange','Warning')][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [AllowNull()]$Data = $null
    )

    $script:Actions.Add([pscustomobject][ordered]@{
        Timestamp = (Get-Date).ToUniversalTime().ToString('o')
        Action    = $Action
        Status    = $Status
        Message   = $Message
        Data      = $Data
    }) | Out-Null
}

function Get-PropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-EdgePolicyState {
    [CmdletBinding()]
    param()

    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate'
    $updateDefault = $null
    $stableUpdate = $null

    if (Test-Path -LiteralPath $policyPath) {
        $policy = Get-ItemProperty -LiteralPath $policyPath -ErrorAction SilentlyContinue
        $updateDefault = Get-PropertyValue -InputObject $policy -Name 'UpdateDefault'
        $stableUpdate = Get-PropertyValue -InputObject $policy -Name 'Update{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}'
    }

    [pscustomobject][ordered]@{
        Path                 = $policyPath
        Present              = Test-Path -LiteralPath $policyPath
        UpdateDefault        = $updateDefault
        StableChannelUpdate  = $stableUpdate
        UpdatesDisabled      = [bool]($updateDefault -eq 0 -or $stableUpdate -eq 0)
    }
}

function Get-EdgeServicingState {
    [CmdletBinding()]
    param()

    $edgeCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) }
    $edgePath = $edgeCandidates | Select-Object -First 1

    $updaterCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) }
    $updaterPath = $updaterCandidates | Select-Object -First 1

    $services = foreach ($serviceName in @('edgeupdate','edgeupdatem','MicrosoftEdgeElevationService')) {
        $service = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $serviceName) -ErrorAction SilentlyContinue
        if ($service) {
            [pscustomobject][ordered]@{
                Name      = [string]$service.Name
                State     = [string]$service.State
                StartMode = [string]$service.StartMode
                PathName  = [string]$service.PathName
            }
        }
    }

    $tasks = @()
    if (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue) {
        $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskName -match 'MicrosoftEdgeUpdate|Microsoft Edge Update' } |
            ForEach-Object {
                $taskInfo = Get-ScheduledTaskInfo -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue
                [pscustomobject][ordered]@{
                    TaskName       = [string]$_.TaskName
                    TaskPath       = [string]$_.TaskPath
                    State          = [string]$_.State
                    LastRunTimeUtc = if ($taskInfo -and $taskInfo.LastRunTime -gt [datetime]::MinValue) {
                        $taskInfo.LastRunTime.ToUniversalTime().ToString('o')
                    }
                    else { $null }
                    LastTaskResult = if ($taskInfo) { [int64]$taskInfo.LastTaskResult } else { $null }
                }
            })
    }

    $policy = Get-EdgePolicyState
    $edgeVersion = $null
    if ($edgePath) {
        $edgeVersion = (Get-Item -LiteralPath $edgePath -ErrorAction Stop).VersionInfo.ProductVersion
    }

    $servicingPresent = [bool](@($services).Count -gt 0 -or @($tasks).Count -gt 0)
    $healthy = [bool]($edgePath -and -not $policy.UpdatesDisabled -and $servicingPresent)

    [pscustomobject][ordered]@{
        CheckedUtc        = (Get-Date).ToUniversalTime().ToString('o')
        EdgeInstalled     = [bool]$edgePath
        EdgePath          = [string]$edgePath
        EdgeVersion       = [string]$edgeVersion
        UpdaterPath       = [string]$updaterPath
        ServicingPresent  = $servicingPresent
        Healthy           = $healthy
        Policy            = $policy
        Services          = [object[]]@($services)
        ScheduledTasks    = [object[]]@($tasks)
    }
}

function Invoke-ProcessWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$ArgumentList,
        [Parameter(Mandatory)][ValidateRange(1, 7200)][int]$TimeoutSeconds,
        [Parameter(Mandatory)][string]$Name
    )

    $stdoutPath = Join-Path $script:RunWorkingDirectory ($Name + '.stdout.log')
    $stderrPath = Join-Path $script:RunWorkingDirectory ($Name + '.stderr.log')
    $started = Get-Date
    $process = $null

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList `
            -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath -ErrorAction Stop

        $completed = $process.WaitForExit($TimeoutSeconds * 1000)
        if (-not $completed) {
            try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
            throw "$Name exceeded the $TimeoutSeconds-second timeout."
        }
        $process.Refresh()

        [pscustomobject][ordered]@{
            Name            = $Name
            FilePath        = $FilePath
            Arguments       = $ArgumentList
            ExitCode        = [int]$process.ExitCode
            TimedOut        = $false
            StartedUtc      = $started.ToUniversalTime().ToString('o')
            EndedUtc        = (Get-Date).ToUniversalTime().ToString('o')
            DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 3)
            StandardOutput  = if (Test-Path -LiteralPath $stdoutPath) { (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue).Trim() } else { $null }
            StandardError   = if (Test-Path -LiteralPath $stderrPath) { (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue).Trim() } else { $null }
        }
    }
    catch {
        [pscustomobject][ordered]@{
            Name            = $Name
            FilePath        = $FilePath
            Arguments       = $ArgumentList
            ExitCode        = $null
            TimedOut        = [bool]($_.Exception.Message -match 'timeout')
            StartedUtc      = $started.ToUniversalTime().ToString('o')
            EndedUtc        = (Get-Date).ToUniversalTime().ToString('o')
            DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 3)
            StandardOutput  = if (Test-Path -LiteralPath $stdoutPath) { (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue).Trim() } else { $null }
            StandardError   = $_.Exception.Message
        }
    }
}

function Repair-EdgeUpdateServices {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $desiredModes = [ordered]@{
        edgeupdate                    = 'delayed-auto'
        edgeupdatem                   = 'demand'
        MicrosoftEdgeElevationService = 'demand'
    }

    foreach ($serviceName in $desiredModes.Keys) {
        $service = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $serviceName) -ErrorAction SilentlyContinue
        if (-not $service) {
            Add-RepairAction -Action "Service:$serviceName" -Status Skipped -Message 'Service is not installed; the updater or MSI repair must recreate it.'
            continue
        }

        $desiredMode = [string]$desiredModes[$serviceName]
        if ($PSCmdlet.ShouldProcess($serviceName, "Set Edge Update service start mode to $desiredMode")) {
            $scOutput = & "$env:SystemRoot\System32\sc.exe" config $serviceName 'start=' $desiredMode 2>&1
            $scExitCode = $LASTEXITCODE
            if ($scExitCode -ne 0) {
                throw "sc.exe could not configure $serviceName. ExitCode=$scExitCode; Output=$($scOutput -join ' ')"
            }
            Add-RepairAction -Action "Service:$serviceName" -Status Succeeded -Message "Configured service start mode as $desiredMode."
            Write-RepairLog "Configured $serviceName start mode as $desiredMode." 'OK'
        }
        else {
            Add-RepairAction -Action "Service:$serviceName" -Status Skipped -Message 'WhatIf prevented the service configuration change.'
        }
    }

    $automaticService = Get-Service -Name 'edgeupdate' -ErrorAction SilentlyContinue
    if ($automaticService -and $automaticService.Status -ne 'Running') {
        if ($PSCmdlet.ShouldProcess('edgeupdate', 'Start Edge Update service')) {
            try {
                Start-Service -Name 'edgeupdate' -ErrorAction Stop
                Add-RepairAction -Action 'StartService:edgeupdate' -Status Succeeded -Message 'Started the automatic Edge Update service.'
                Write-RepairLog 'Started the edgeupdate service.' 'OK'
            }
            catch {
                Add-RepairAction -Action 'StartService:edgeupdate' -Status Warning -Message $_.Exception.Message
                Write-RepairLog "The edgeupdate service did not start: $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Enable-EdgeUpdateTasks {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    if (-not (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue)) {
        Add-RepairAction -Action 'ScheduledTasks' -Status Warning -Message 'ScheduledTasks cmdlets are unavailable.'
        Write-RepairLog 'ScheduledTasks cmdlets are unavailable; existing Edge Update tasks could not be checked.' 'WARN'
        return
    }

    $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { $_.TaskName -match 'MicrosoftEdgeUpdate|Microsoft Edge Update' })

    if ($tasks.Count -eq 0) {
        Add-RepairAction -Action 'ScheduledTasks' -Status Skipped -Message 'No existing Edge Update scheduled tasks were found.'
        Write-RepairLog 'No existing Edge Update scheduled tasks were found; updater or MSI repair must recreate them.' 'WARN'
        return
    }

    foreach ($task in $tasks) {
        $taskIdentity = '{0}{1}' -f $task.TaskPath, $task.TaskName
        if ([string]$task.State -eq 'Disabled') {
            if ($PSCmdlet.ShouldProcess($taskIdentity, 'Enable Edge Update scheduled task')) {
                Enable-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop | Out-Null
                Add-RepairAction -Action "ScheduledTask:$taskIdentity" -Status Succeeded -Message 'Enabled the existing Edge Update scheduled task.'
                Write-RepairLog "Enabled Edge Update scheduled task $taskIdentity." 'OK'
            }
        }
        else {
            Add-RepairAction -Action "ScheduledTask:$taskIdentity" -Status NoChange -Message "Task state is $($task.State)."
        }
    }
}

function Invoke-InstalledEdgeUpdater {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$UpdaterPath)

    if (-not $PSCmdlet.ShouldProcess($UpdaterPath, 'Run Microsoft Edge Update on-demand update check')) {
        Add-RepairAction -Action 'InstalledUpdater' -Status Skipped -Message 'WhatIf prevented the updater invocation.'
        return $null
    }

    Write-RepairLog "Invoking installed Edge updater: $UpdaterPath" 'INFO'
    $result = Invoke-ProcessWithTimeout -FilePath $UpdaterPath `
        -ArgumentList '/ua /installsource scheduler' `
        -TimeoutSeconds $UpdaterTimeoutSeconds `
        -Name 'MicrosoftEdgeUpdate'

    if ($null -ne $result.ExitCode -and $result.ExitCode -eq 0) {
        Add-RepairAction -Action 'InstalledUpdater' -Status Succeeded -Message 'Microsoft Edge Update completed successfully.' -Data $result
        Write-RepairLog 'Microsoft Edge Update completed successfully.' 'OK'
    }
    else {
        Add-RepairAction -Action 'InstalledUpdater' -Status Warning -Message "Microsoft Edge Update did not return success. ExitCode=$($result.ExitCode)." -Data $result
        Write-RepairLog "Microsoft Edge Update did not return success. ExitCode=$($result.ExitCode); Error=$($result.StandardError)" 'WARN'
    }

    return $result
}

function Get-MsiPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('ProductName','Manufacturer','ProductVersion','ProductCode')][string]$Property
    )

    $installer = $null
    $database = $null
    $view = $null
    $record = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.GetType().InvokeMember(
            'OpenDatabase',
            'InvokeMethod',
            $null,
            $installer,
            @($Path, 0)
        )
        $query = "SELECT ``Value`` FROM ``Property`` WHERE ``Property``='$Property'"
        $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @($query))
        $null = $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        if ($record) {
            return [string]$record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
        }
        return $null
    }
    finally {
        foreach ($comObject in @($record, $view, $database, $installer)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch {}
            }
        }
    }
}

function Resolve-TrustedEdgeInstaller {
    [CmdletBinding()]
    param([string]$RequestedPath)

    if ([string]::IsNullOrWhiteSpace($RequestedPath)) { return $null }
    if (-not (Test-Path -LiteralPath $RequestedPath -PathType Leaf)) {
        throw "The configured Edge MSI was not found: $RequestedPath"
    }
    if ([IO.Path]::GetExtension($RequestedPath) -ine '.msi') {
        throw "The Edge installer must be an MSI: $RequestedPath"
    }

    $localInstaller = Join-Path $script:RunWorkingDirectory ([IO.Path]::GetFileName($RequestedPath))
    Copy-Item -LiteralPath $RequestedPath -Destination $localInstaller -Force -ErrorAction Stop

    $signature = Get-AuthenticodeSignature -LiteralPath $localInstaller -ErrorAction Stop
    $subject = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { $null }
    if ($signature.Status -ne 'Valid' -or $subject -notmatch '(?i)Microsoft Corporation') {
        throw "Edge MSI signature validation failed. Status=$($signature.Status); Subject=$subject"
    }

    $productName = Get-MsiPropertyValue -Path $localInstaller -Property ProductName
    $manufacturer = Get-MsiPropertyValue -Path $localInstaller -Property Manufacturer
    $productVersion = Get-MsiPropertyValue -Path $localInstaller -Property ProductVersion
    $productCode = Get-MsiPropertyValue -Path $localInstaller -Property ProductCode
    if ($productName -notmatch '(?i)^Microsoft Edge' -or $manufacturer -notmatch '(?i)^Microsoft Corporation$') {
        throw "MSI product validation failed. ProductName=$productName; Manufacturer=$manufacturer"
    }

    $hash = (Get-FileHash -LiteralPath $localInstaller -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
    Add-RepairAction -Action 'ValidateInstaller' -Status Succeeded -Message 'Validated Microsoft Authenticode signature on the Edge MSI.' -Data ([pscustomobject]@{
        SourcePath = $RequestedPath
        LocalPath  = $localInstaller
        Signer     = $subject
        ProductName = $productName
        Manufacturer = $manufacturer
        ProductVersion = $productVersion
        ProductCode = $productCode
        SHA256     = $hash
    })
    Write-RepairLog "Validated Microsoft-signed Edge MSI. SHA256=$hash" 'OK'
    return $localInstaller
}

function Invoke-EdgeMsiRepair {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$MsiPath)

    $msiLog = Join-Path $LogDirectory ("{0}-MicrosoftEdge-MSI-{1}.log" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
    $arguments = '/i "{0}" /qn /norestart REINSTALL=ALL REINSTALLMODE=vomus /L*v "{1}"' -f $MsiPath, $msiLog

    if (-not $PSCmdlet.ShouldProcess($MsiPath, 'Repair Microsoft Edge using the verified Enterprise MSI')) {
        Add-RepairAction -Action 'EdgeMsiRepair' -Status Skipped -Message 'WhatIf prevented the MSI repair.'
        return $null
    }

    Write-RepairLog "Running verified Microsoft Edge MSI repair. MSI log: $msiLog" 'INFO'
    $result = Invoke-ProcessWithTimeout -FilePath "$env:SystemRoot\System32\msiexec.exe" `
        -ArgumentList $arguments `
        -TimeoutSeconds $InstallerTimeoutSeconds `
        -Name 'MicrosoftEdgeMsiRepair'

    if ($result.ExitCode -in @(0, 1641, 3010)) {
        if ($result.ExitCode -in @(1641, 3010)) { $script:RebootRequired = $true }
        Add-RepairAction -Action 'EdgeMsiRepair' -Status Succeeded -Message "Edge MSI repair completed. ExitCode=$($result.ExitCode)." -Data $result
        Write-RepairLog "Edge MSI repair completed. ExitCode=$($result.ExitCode)." 'OK'
    }
    else {
        Add-RepairAction -Action 'EdgeMsiRepair' -Status Failed -Message "Edge MSI repair failed. ExitCode=$($result.ExitCode). MSI log: $msiLog" -Data $result
        throw "Edge MSI repair failed. ExitCode=$($result.ExitCode). MSI log: $msiLog"
    }

    return $result
}

function New-FinalRecord {
    [CmdletBinding()]
    param()

    $endTime = Get-Date
    [pscustomobject][ordered]@{
        '@timestamp'       = $endTime.ToUniversalTime().ToString('o')
        EventType          = 'maintenance.remediation'
        SchemaVersion      = '1.0'
        ComputerName       = $env:COMPUTERNAME
        Domain             = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { $env:USERDOMAIN }
        ScriptName         = $script:ScriptName
        ScriptVersion      = $script:ScriptVersion
        RemediationClass   = $script:RemediationClass
        FindingFingerprint = $FindingFingerprint
        AttemptNumber      = $AttemptNumber
        RunId              = $script:RunId
        CorrelationId      = $script:CorrelationId
        Status             = $script:FinalStatus
        Outcome            = $script:Outcome
        ExitCode           = $script:FinalExitCode
        StartTime          = $script:StartTime.ToUniversalTime().ToString('o')
        EndTime            = $endTime.ToUniversalTime().ToString('o')
        DurationSeconds    = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 3)
        WarningCount       = $script:WarningCount
        ErrorCount         = $script:ErrorCount
        RebootRequired     = $script:RebootRequired
        MaintenanceWindowBypassed = [bool]$BypassMaintenanceWindow
        FailureMessage     = $script:FailureMessage
        TextLogPath        = $script:PublishedLogPath
        Before             = $script:BeforeState
        After              = $script:AfterState
        UpdaterResult      = $script:UpdaterResult
        InstallerResult    = $script:InstallerResult
        Actions            = [object[]]$script:Actions.ToArray()
    }
}

try {
    if (-not (Test-MaintenanceAdministrator)) {
        throw 'Administrator or SYSTEM privileges are required.'
    }

    $script:LogSession = New-MaintenanceStagedLog `
        -ScriptName $script:ScriptName `
        -LogRoot $LogDirectory `
        -StagingRoot $maintenanceConfig.LogStagingRoot `
        -ComputerName $env:COMPUTERNAME `
        -Timestamp $script:StartTime
    $script:WorkingLogPath = [string]$script:LogSession.WorkingPath
    $script:PublishedLogPath = [string]$script:LogSession.PublishedPath

    New-Item -Path $script:RunWorkingDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    Write-RepairLog "Starting $($script:ScriptName) version $($script:ScriptVersion). RunId=$($script:RunId); Attempt=$AttemptNumber" 'INFO'
    Write-RepairLog "FindingFingerprint=$FindingFingerprint" 'INFO'

    if (-not $BypassMaintenanceWindow) {
        $policyPath = Join-Path (Split-Path -Parent $frameworkPath) 'Maintenance.Policy.json'
        if (-not (Test-Path -LiteralPath $policyPath -PathType Leaf)) {
            $policyPath = 'C:\Scripts\Maintenance.Policy.json'
        }
        $policy = Get-MaintenancePolicy -Path $policyPath
        $window = Test-MaintenanceWindow -ScriptName $script:ScriptName -Policy $policy
        if (-not $window.Allowed) {
            $script:Outcome = 'OutsideMaintenanceWindow'
            $script:FinalStatus = 'Blocked'
            $script:FinalExitCode = 20
            Add-RepairAction -Action 'MaintenanceWindow' -Status Skipped -Message $window.Reason
            Write-RepairLog $window.Reason 'WARN'
            throw [System.OperationCanceledException]::new($window.Reason)
        }
    }
    else {
        Write-RepairLog 'An authorized maintenance-window bypass was requested.' 'WARN'
    }

    # Use a remediation-specific lock. The allowlisted launcher already owns the
    # global maintenance lock when this is orchestrated, while direct technician
    # runs still need protection against duplicate Edge repairs.
    $script:Lock = Enter-MaintenanceExecutionLock `
        -Name 'Global\Compton_MicrosoftEdgeUpdateRepair' `
        -TimeoutSeconds 120
    Write-RepairLog 'Microsoft Edge Update remediation lock acquired.' 'INFO'

    $script:BeforeState = Get-EdgeServicingState
    Write-RepairLog ("Before repair: EdgeInstalled={0}; Version={1}; UpdaterPresent={2}; ServicingPresent={3}; UpdatesDisabled={4}; Healthy={5}" -f `
        $script:BeforeState.EdgeInstalled,
        $script:BeforeState.EdgeVersion,
        [bool]$script:BeforeState.UpdaterPath,
        $script:BeforeState.ServicingPresent,
        $script:BeforeState.Policy.UpdatesDisabled,
        $script:BeforeState.Healthy) 'INFO'

    if ($script:BeforeState.Policy.UpdatesDisabled) {
        $script:Outcome = 'BlockedByPolicy'
        $script:FinalStatus = 'Blocked'
        $script:FinalExitCode = 20
        $script:FailureMessage = 'Microsoft Edge updates are disabled by computer policy. The remediation will not override policy.'
        Add-RepairAction -Action 'PolicyCheck' -Status Skipped -Message $script:FailureMessage -Data $script:BeforeState.Policy
        Write-RepairLog $script:FailureMessage 'ERROR'
    }
    else {
        Repair-EdgeUpdateServices
        Enable-EdgeUpdateTasks

        if ($script:BeforeState.UpdaterPath) {
            $script:UpdaterResult = Invoke-InstalledEdgeUpdater -UpdaterPath $script:BeforeState.UpdaterPath
        }
        else {
            Add-RepairAction -Action 'InstalledUpdater' -Status Skipped -Message 'MicrosoftEdgeUpdate.exe was not found.'
            Write-RepairLog 'MicrosoftEdgeUpdate.exe was not found.' 'WARN'
        }

        if ($PostRepairWaitSeconds -gt 0 -and -not $WhatIfPreference) {
            Write-RepairLog "Waiting $PostRepairWaitSeconds seconds before verification." 'INFO'
            Start-Sleep -Seconds $PostRepairWaitSeconds
        }

        $intermediateState = Get-EdgeServicingState
        $updaterFailed = [bool]($script:UpdaterResult -and $script:UpdaterResult.ExitCode -ne 0)
        $needsInstaller = [bool]($ForceInstallerRepair -or -not $intermediateState.Healthy -or $updaterFailed)

        if ($needsInstaller) {
            if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
                if ($ForceInstallerRepair) {
                    throw 'ForceInstallerRepair was requested, but no InstallerPath was supplied.'
                }
                Write-RepairLog 'MSI fallback is needed, but no InstallerPath was supplied.' 'WARN'
                Add-RepairAction -Action 'EdgeMsiRepair' -Status Skipped -Message 'MSI fallback was needed, but no InstallerPath was supplied.'
            }
            else {
                $trustedInstaller = Resolve-TrustedEdgeInstaller -RequestedPath $InstallerPath
                $script:InstallerResult = Invoke-EdgeMsiRepair -MsiPath $trustedInstaller

                Repair-EdgeUpdateServices
                Enable-EdgeUpdateTasks

                if ($PostRepairWaitSeconds -gt 0 -and -not $WhatIfPreference) {
                    Write-RepairLog "Waiting $PostRepairWaitSeconds seconds after MSI repair before final verification." 'INFO'
                    Start-Sleep -Seconds $PostRepairWaitSeconds
                }
            }
        }

        $script:AfterState = Get-EdgeServicingState
        Write-RepairLog ("After repair: EdgeInstalled={0}; Version={1}; UpdaterPresent={2}; ServicingPresent={3}; UpdatesDisabled={4}; Healthy={5}" -f `
            $script:AfterState.EdgeInstalled,
            $script:AfterState.EdgeVersion,
            [bool]$script:AfterState.UpdaterPath,
            $script:AfterState.ServicingPresent,
            $script:AfterState.Policy.UpdatesDisabled,
            $script:AfterState.Healthy) 'INFO'

        if ($WhatIfPreference) {
            $script:Outcome = 'WhatIf'
            $script:FinalStatus = 'Skipped'
            $script:FinalExitCode = 0
        }
        elseif (-not $script:AfterState.Healthy) {
            $script:Outcome = 'RepairFailed'
            $script:FinalStatus = 'Failed'
            $script:FinalExitCode = 10
            $script:FailureMessage = 'Edge Update servicing did not pass final verification.'
            Write-RepairLog $script:FailureMessage 'ERROR'
        }
        elseif ($script:RebootRequired) {
            $script:Outcome = 'RepairedRebootRequired'
            $script:FinalStatus = 'Success'
            $script:FinalExitCode = 3010
            Write-RepairLog 'Edge Update servicing passed verification; the MSI requested a reboot.' 'WARN'
        }
        elseif ($script:WarningCount -gt 0) {
            $script:Outcome = 'RepairedWithWarnings'
            $script:FinalStatus = 'SuccessWithWarnings'
            $script:FinalExitCode = 1
            Write-RepairLog 'Edge Update servicing passed verification with one or more warnings.' 'WARN'
        }
        elseif ($script:BeforeState.Healthy -and -not $ForceInstallerRepair) {
            $script:Outcome = 'AlreadyHealthy'
            $script:FinalStatus = 'Success'
            $script:FinalExitCode = 0
            Write-RepairLog 'Edge Update servicing was already healthy and remains healthy.' 'OK'
        }
        else {
            $script:Outcome = 'Repaired'
            $script:FinalStatus = 'Success'
            $script:FinalExitCode = 0
            Write-RepairLog 'Edge Update servicing repair completed and passed verification.' 'OK'
        }
    }
}
catch [System.OperationCanceledException] {
    if (-not $script:FailureMessage) { $script:FailureMessage = $_.Exception.Message }
    if ($script:FinalStatus -eq 'Running') {
        $script:Outcome = 'OutsideMaintenanceWindow'
        $script:FinalStatus = 'Blocked'
        $script:FinalExitCode = 20
    }
}
catch {
    $script:FailureMessage = $_.Exception.Message
    $script:Outcome = 'RepairFailed'
    $script:FinalStatus = 'Failed'
    $script:FinalExitCode = 10
    try { Write-RepairLog "Repair failed: $($script:FailureMessage)" 'ERROR' } catch {}
}
finally {
    if ($script:Lock) {
        Exit-MaintenanceExecutionLock -Lock $script:Lock
        try { Write-RepairLog 'Microsoft Edge Update remediation lock released.' 'INFO' } catch {}
    }

    if (-not $script:AfterState) {
        try { $script:AfterState = Get-EdgeServicingState } catch {}
    }

    $record = New-FinalRecord
    try { Write-MaintenanceTelemetry -Record $record -Path $script:TelemetryPath -Depth 16 } catch {
        try { Write-RepairLog "Failed to append remediation telemetry: $($_.Exception.Message)" 'ERROR' } catch {}
    }
    try { Write-MaintenanceJsonAtomically -Path $script:LatestPath -Record $record -Depth 16 | Out-Null } catch {
        try { Write-RepairLog "Failed to write latest remediation JSON: $($_.Exception.Message)" 'ERROR' } catch {}
    }
    try {
        Write-MaintenanceEvent `
            -EventId $(if ($script:FinalExitCode -in @(0,1,3010)) { 1001 } elseif ($script:FinalExitCode -eq 20) { 1700 } else { 1900 }) `
            -EntryType $(if ($script:FinalExitCode -eq 0) { 'Information' } elseif ($script:FinalExitCode -in @(1,20,3010)) { 'Warning' } else { 'Error' }) `
            -Message ("{0} completed. Status={1}; Outcome={2}; ExitCode={3}; RunId={4}; CorrelationId={5}; Fingerprint={6}" -f `
                $script:ScriptName, $script:FinalStatus, $script:Outcome, $script:FinalExitCode,
                $script:RunId, $script:CorrelationId, $FindingFingerprint) | Out-Null
    }
    catch {}

    if ($script:LogSession) {
        try {
            Write-RepairLog ("Completed. Status={0}; Outcome={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
                $script:FinalStatus, $script:Outcome, $script:FinalExitCode,
                $script:WarningCount, $script:ErrorCount) `
                $(if ($script:FinalExitCode -in @(0,1,3010)) { 'OK' } elseif ($script:FinalExitCode -eq 20) { 'WARN' } else { 'ERROR' })
            $publishResult = Publish-MaintenanceLog -LogSession $script:LogSession
            if (-not $publishResult.Published) {
                Write-Warning "Completed remediation log remains in staging: $($publishResult.Path)"
            }
        }
        catch {
            Write-Warning "Failed to publish the completed remediation log: $($_.Exception.Message)"
        }
    }

    if (Test-Path -LiteralPath $script:RunWorkingDirectory -PathType Container) {
        Remove-Item -LiteralPath $script:RunWorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
}

exit $script:FinalExitCode
