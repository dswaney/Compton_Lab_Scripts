# =====================================================================
# ScriptName: 06_Weekend_Windows_Updates.ps1
# ScriptVersion: 2.1.10
# LastUpdated: 2026-08-18
# Changes: v2.1.10 suppresses unusable COM-backed Category values so telemetry
#           records a real category when available, otherwise null.
#           v2.1.9 normalizes PSWindowsUpdate COM-backed KB/category properties and
#           derives the pre-update count from unique install candidates when the
#           pre-scan reports zero even though Install-WindowsUpdate finds updates.
#          v2.1.8 adds Elastic-ready Windows Update result telemetry including
#          installed, failed, remaining, and skipped/deferred update details.
#          v2.1.7 uses Maintenance.Framework v2.4 staged text logging.
# Purpose: Installs Microsoft/Windows updates using PSWindowsUpdate,
#          records detailed YAML and Elastic-compatible JSON telemetry,
#          validates update compliance before and after installation,
#          and records when a reboot is required without initiating one.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$ResetWUComponentsFirst = $false,
    [int]$OperationTimeoutSeconds = 1800,
    [string]$LogFolder = "$env:SystemDrive\Logs"
)

$ErrorActionPreference = 'Stop'

$script:ScriptName       = '06_Weekend_Windows_Updates.ps1'
$script:ScriptVersion    = '2.1.10'
$script:RunId            = [guid]::NewGuid().Guid
$script:RunStart         = Get-Date
$script:ComputerName     = $env:COMPUTERNAME
$script:Domain           = $env:USERDOMAIN
$script:YamlLogPath      = $null
$script:RuntimeLogPath   = $null
$script:PublishedLogPath = $null
$script:LogSession       = $null
$script:TelemetryPath    = Join-Path $LogFolder 'Maintenance-Telemetry.ndjson'
$script:LatestJsonPath   = Join-Path $LogFolder '06_Weekend_Windows_Updates.latest.json'
$script:ActionHistory    = New-Object System.Collections.Generic.List[object]
$script:RawUpdateLines   = New-Object System.Collections.Generic.List[string]
$script:UpdateEntries    = New-Object System.Collections.Generic.List[object]
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:OverallResult    = 'Initializing'
$script:ExitCode         = 0
$script:FailureMessage   = $null
$script:RebootInitiated  = $false
$script:RebootBefore     = $null
$script:RebootAfter      = $null
$script:ModuleVersion    = $null
$script:BuildBefore      = $null
$script:BuildAfter       = $null
$script:LastUpdateBefore = $null
$script:LastUpdateAfter  = $null
$script:UpdatesBefore    = @()
$script:UpdatesAfter     = @()
$script:ResetResult      = [ordered]@{
    Requested = [bool]$ResetWUComponentsFirst
    Completed = $false
    ServicesStopped = @()
    ServicesStarted = @()
    ServiceFailures = @()
    FoldersCleared = @()
    FolderFailures = @()
    RegistryKeysRemoved = @()
    RegistryFailures = @()
}

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -DisableNameChecking -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogFolder

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 06 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}


# Normalizes either a PowerShell array or a generic .NET collection into a
# plain PowerShell Object[] for Windows PowerShell 5.1 compatibility.
function ConvertTo-NormalArray {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $InputObject
    )

    if ($null -eq $InputObject) {
        return @()
    }

    return @(
        $InputObject | ForEach-Object { $_ }
    )
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Write-ImmediateRuntimeLog {
    param([Parameter(Mandatory)][string]$Line)
    try {
        if ($script:RuntimeLogPath) {
            $activeLogDirectory = Split-Path -Parent $script:RuntimeLogPath
            Ensure-Folder -Path $activeLogDirectory
            Add-Content -LiteralPath $script:RuntimeLogPath -Value $Line -Encoding UTF8
        }
    }
    catch {}
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

    if ($normalizedLevel -eq 'WARNING') { $script:WarningCount++ }
    if ($normalizedLevel -eq 'ERROR') { $script:ErrorCount++ }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$($script:ComputerName)] [$normalizedLevel] $Message"

    switch ($normalizedLevel) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
    }

    $script:ActionHistory.Add([pscustomobject]@{
        Time = $timestamp
        Level = $normalizedLevel
        Message = $Message
    }) | Out-Null

    Write-ImmediateRuntimeLog -Line $line
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

function ConvertTo-YamlSafeValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [string]$Value
    }

    $text = [string]$Value
    $text = $text -replace "`r", ' '
    $text = $text -replace "`n", ' '
    $text = $text -replace '"', '\"'
    return '"' + $text + '"'
}

function Initialize-Logging {
    Ensure-Folder -Path $LogFolder

    Archive-MaintenanceLogs `
        -ScriptName $script:ScriptName `
        -LogRoot $LogFolder `
        -AdditionalPatterns @(
            '06_Weekend_Windows_Updates.log',
            '*-06_Weekend_Windows_Updates-*.log'
        ) | Out-Null

    $script:LogSession = New-MaintenanceStagedLog `
        -ScriptName $script:ScriptName `
        -LogRoot $LogFolder `
        -StagingRoot $MaintenanceConfig.LogStagingRoot `
        -ComputerName $script:ComputerName `
        -Timestamp $script:RunStart

    $script:RuntimeLogPath = [string]$script:LogSession.WorkingPath
    $script:PublishedLogPath = [string]$script:LogSession.PublishedPath

    $stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $base = "$($script:ComputerName)-WindowsUpdates-$stamp"
    $script:YamlLogPath = Join-Path $LogFolder "$base.yaml"

    Set-Content -LiteralPath $script:RuntimeLogPath -Encoding UTF8 -Value @(
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$($script:ComputerName)] [INFO] Runtime log initialized in staging.",
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$($script:ComputerName)] [INFO] Run ID: $($script:RunId)",
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$($script:ComputerName)] [INFO] Completed log publish path: $($script:PublishedLogPath)"
    )
}

function Get-WindowsBuildSnapshot {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        return [pscustomobject]@{
            Caption = $os.Caption
            Version = $os.Version
            BuildNumber = $os.BuildNumber
            UBR = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR -ErrorAction SilentlyContinue).UBR
            DisplayVersion = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
        }
    }
    catch {
        Write-Log "Unable to collect Windows build information: $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Get-LastSuccessfulUpdate {
    try {
        $hotfix = Get-HotFix -ErrorAction Stop |
            Where-Object { $_.InstalledOn } |
            Sort-Object InstalledOn -Descending |
            Select-Object -First 1

        if ($hotfix) {
            return [pscustomobject]@{
                KB = $hotfix.HotFixID
                InstalledOn = $hotfix.InstalledOn
                Description = $hotfix.Description
            }
        }
    }
    catch {
        Write-Log "Unable to determine the latest installed hotfix: $($_.Exception.Message)" 'WARN'
    }
    return $null
}

function Get-PendingRebootState {
    $reasons = New-Object System.Collections.Generic.List[string]

    $checks = @(
        @{ Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'; Reason='ComponentBasedServicing' },
        @{ Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'; Reason='WindowsUpdate' },
        @{ Path='HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'; Name='PendingFileRenameOperations'; Reason='PendingFileRenameOperations' },
        @{ Path='HKLM:\SOFTWARE\Microsoft\Updates'; Name='UpdateExeVolatile'; Reason='UpdateExeVolatile' }
    )

    foreach ($check in $checks) {
        try {
            if ($check.ContainsKey('Name')) {
                $value = (Get-ItemProperty -LiteralPath $check.Path -Name $check.Name -ErrorAction SilentlyContinue).($check.Name)
                if ($null -ne $value -and "$value" -ne '' -and "$value" -ne '0') {
                    $reasons.Add($check.Reason) | Out-Null
                }
            }
            elseif (Test-Path -LiteralPath $check.Path) {
                $reasons.Add($check.Reason) | Out-Null
            }
        }
        catch {}
    }

    try {
        if (Get-Command -Name Get-WURebootStatus -ErrorAction SilentlyContinue) {
            if ([bool](Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)) {
                $reasons.Add('PSWindowsUpdate') | Out-Null
            }
        }
    }
    catch {}

    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        if ([bool]$sysInfo.RebootRequired) {
            $reasons.Add('Microsoft.Update.SystemInfo') | Out-Null
        }
    }
    catch {}

    $unique = @($reasons | Sort-Object -Unique)
    return [pscustomobject]@{
        Required = ($unique.Count -gt 0)
        Reasons = $unique
    }
}

function Ensure-PSWindowsUpdate {
    
# Ensure the NuGet PackageManagement provider exists before PowerShellGet
# attempts to install PSWindowsUpdate. This prevents unattended runs from
# stopping at the interactive NuGet bootstrap prompt.
function Ensure-NuGetProvider {
    [CmdletBinding()]
    param()

    try {
        $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($nuget) {
            Write-Log ("NuGet PackageManagement provider is already available. Version: {0}" -f $nuget.Version) 'OK'
            return $true
        }

        Write-Log 'NuGet PackageManagement provider is not installed. Installing silently...' 'INFO'

        [Net.ServicePointManager]::SecurityProtocol = `
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

        $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
            Write-Log 'PSGallery installation policy set to Trusted for unattended maintenance.' 'INFO'
        }

        Install-PackageProvider `
            -Name NuGet `
            -MinimumVersion '2.8.5.201' `
            -Force `
            -ForceBootstrap `
            -Scope AllUsers `
            -Confirm:$false `
            -ErrorAction Stop | Out-Null

        Import-PackageProvider -Name NuGet -Force -ErrorAction SilentlyContinue | Out-Null

        $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $nuget) {
            throw 'NuGet provider installation completed but the provider could not be detected afterward.'
        }

        Write-Log ("NuGet PackageManagement provider installed successfully. Version: {0}" -f $nuget.Version) 'OK'
        return $true
    }
    catch {
        Write-Log ("Unable to install or verify the NuGet PackageManagement provider: {0}" -f $_.Exception.Message) 'ERROR'
        return $false
    }
}

Write-Log 'Ensuring the PSWindowsUpdate module is available...' 'INFO'

if (-not (Ensure-NuGetProvider)) {
    throw 'PSWindowsUpdate prerequisites could not be prepared because the NuGet provider is unavailable.'
}


    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Log 'Installing the NuGet package provider...' 'INFO'
        Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null
    }

    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-Log 'Setting PSGallery to Trusted...' 'INFO'
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
    }
    catch {
        Write-Log "Could not validate PSGallery settings: $($_.Exception.Message)" 'WARN'
    }

    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log 'Installing PSWindowsUpdate...' 'INFO'
        Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Scope AllUsers
    }

    Import-Module PSWindowsUpdate -Force
    $module = Get-Module -Name PSWindowsUpdate | Sort-Object Version -Descending | Select-Object -First 1
    if ($module) { $script:ModuleVersion = $module.Version.ToString() }

    Write-Log "PSWindowsUpdate imported successfully. Version: $($script:ModuleVersion)" 'OK'
}

function Convert-WUItem {
    param(
        [Parameter(Mandatory)][object]$Item,
        [string]$Phase
    )

    $props = $Item.PSObject.Properties.Name

    function Convert-WUValueToText {
        param(
            [object]$Value,
            [ValidateSet('Generic','KB','Category')]
            [string]$Kind = 'Generic'
        )

        if ($null -eq $Value) { return $null }

        # Normal scalar values.
        if ($Value -is [string] -or $Value.GetType().IsPrimitive -or $Value -is [decimal]) {
            $s = [string]$Value
            if ($s -eq 'System.__ComObject') { return $null }
            if ([string]::IsNullOrWhiteSpace($s)) { return $null }
            return $s.Trim()
        }

        # PSWindowsUpdate frequently exposes KBArticleIDs and Categories through
        # COM collections. Enumerate them instead of stringifying the COM wrapper.
        $parts = New-Object System.Collections.Generic.List[string]

        try {
            foreach ($entry in $Value) {
                if ($null -eq $entry) { continue }

                $candidate = $null

                if ($Kind -eq 'Category') {
                    try {
                        if ($entry.PSObject.Properties.Name -contains 'Name') {
                            $candidate = [string]$entry.Name
                        }
                    }
                    catch { }
                }

                if ([string]::IsNullOrWhiteSpace($candidate)) {
                    try { $candidate = [string]$entry } catch { $candidate = $null }
                }

                if (
                    -not [string]::IsNullOrWhiteSpace($candidate) -and
                    $candidate -notmatch 'System\.__ComObject'
                ) {
                    $parts.Add($candidate.Trim()) | Out-Null
                }
            }
        }
        catch { }

        if ($parts.Count -gt 0) {
            return (($parts | Select-Object -Unique) -join ', ')
        }

        # Some COM values expose a Name directly instead of being enumerable.
        if ($Kind -eq 'Category') {
            try {
                if ($Value.PSObject.Properties.Name -contains 'Name') {
                    $name = [string]$Value.Name
                    if (
                        -not [string]::IsNullOrWhiteSpace($name) -and
                        $name -notmatch 'System\.__ComObject'
                    ) {
                        return $name.Trim()
                    }
                }
            }
            catch { }
        }

        # Never allow the COM wrapper type name into Elastic telemetry.
        try {
            $fallback = [string]$Value
            if (
                -not [string]::IsNullOrWhiteSpace($fallback) -and
                $fallback -notmatch 'System\.__ComObject'
            ) {
                return $fallback.Trim()
            }
        }
        catch { }

        return $null
    }

    function Get-FirstPropertyValue {
        param(
            [string[]]$Names,
            [ValidateSet('Generic','KB','Category')]
            [string]$Kind = 'Generic'
        )

        foreach ($name in $Names) {
            if ($props -contains $name -and $null -ne $Item.$name) {
                $value = Convert-WUValueToText -Value $Item.$name -Kind $Kind
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    return $value
                }
            }
        }

        return $null
    }

    $title = Get-FirstPropertyValue -Names @('Title','KBArticleTitle')
    $kb = Get-FirstPropertyValue -Names @('KB','KBArticleIDs','KBArticleID') -Kind 'KB'
    $size = Get-FirstPropertyValue -Names @('Size','MaxDownloadSize')
    $status = Get-FirstPropertyValue -Names @('Status','Result','UpdateState')
    $result = Get-FirstPropertyValue -Names @('Result','Status','HResult')
    $category = Get-FirstPropertyValue -Names @('Categories','Category','CategoryNames') -Kind 'Category'

    # Defensive cleanup for cases where a COM collection stringifies into multiple
    # wrapper names such as "System.__ComObject System.__ComObject".
    if (
        -not [string]::IsNullOrWhiteSpace($category) -and
        $category -match 'System\.__ComObject'
    ) {
        $category = $null
    }

    # If a KB-style property is still just a non-KB artifact, discard it.
    if (
        -not [string]::IsNullOrWhiteSpace($kb) -and
        $kb -notmatch '(?i)\bKB\d+\b'
    ) {
        $kb = $null
    }

    if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($kb)) {
        return $null
    }

    return [pscustomobject]@{
        Phase = $Phase
        Title = $title
        KB = $kb
        Size = $size
        Category = $category
        Status = $status
        Result = $result
        Source = 'MicrosoftUpdate'
    }
}

function Get-AvailableUpdateSnapshot {
    param([Parameter(Mandatory)][string]$Phase)

    Write-Log "Scanning for available updates ($Phase)..." 'INFO'
    $items = @()

    try {
        $results = @(Get-WindowsUpdate -MicrosoftUpdate -ErrorAction Stop)
        foreach ($item in $results) {
            $converted = Convert-WUItem -Item $item -Phase $Phase
            if ($converted) { $items += $converted }
        }

        Write-Log "Available updates found during $Phase scan: $($items.Count)" 'INFO'
    }
    catch {
        Write-Log "The $Phase update scan failed: $($_.Exception.Message)" 'WARN'
    }

    return (ConvertTo-NormalArray -InputObject $items)
}

function Reset-WUComponentsSafe {
    Write-Log 'Resetting Windows Update components...' 'INFO'
    $services = @('wuauserv','bits','cryptsvc','msiserver')

    foreach ($serviceName in $services) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction Stop
            if ($service.Status -ne 'Stopped') {
                Stop-Service -Name $serviceName -Force -ErrorAction Stop
            }
            $script:ResetResult.ServicesStopped += $serviceName
        }
        catch {
            $script:ResetResult.ServiceFailures += "Stop ${serviceName}: $($_.Exception.Message)"
            Write-Log "Failed to stop ${serviceName}: $($_.Exception.Message)" 'WARN'
        }
    }

    Start-Sleep -Seconds 3

    $registryKeys = @(
        'HKLM\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy',
        'HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Update'
    )

    foreach ($key in $registryKeys) {
        try {
            $psPath = 'Registry::' + $key
            if (Test-Path -LiteralPath $psPath) {
                & reg.exe delete $key /f | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    $script:ResetResult.RegistryKeysRemoved += $key
                }
                else {
                    throw "reg.exe returned $LASTEXITCODE"
                }
            }
        }
        catch {
            $script:ResetResult.RegistryFailures += "${key}: $($_.Exception.Message)"
            Write-Log "Failed to remove registry key ${key}: $($_.Exception.Message)" 'WARN'
        }
    }

    foreach ($folder in @("$env:SystemRoot\SoftwareDistribution","$env:SystemRoot\System32\catroot2")) {
        try {
            if (Test-Path -LiteralPath $folder) {
                Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
                $script:ResetResult.FoldersCleared += $folder
            }
        }
        catch {
            $script:ResetResult.FolderFailures += "${folder}: $($_.Exception.Message)"
            Write-Log "Failed to clear ${folder}: $($_.Exception.Message)" 'WARN'
        }
    }

    foreach ($serviceName in $services) {
        try {
            Start-Service -Name $serviceName -ErrorAction Stop
            $script:ResetResult.ServicesStarted += $serviceName
        }
        catch {
            $script:ResetResult.ServiceFailures += "Start ${serviceName}: $($_.Exception.Message)"
            Write-Log "Failed to start ${serviceName}: $($_.Exception.Message)" 'WARN'
        }
    }

    $script:ResetResult.Completed = ($script:ResetResult.ServiceFailures.Count -eq 0)
    Write-Log 'Windows Update component reset completed.' 'OK'
}

function Install-AvailableWindowsUpdates {
    Write-Log 'Starting Windows Update installation job...' 'INFO'

    $job = Start-Job -ScriptBlock {
        $ErrorActionPreference = 'Stop'
        Import-Module PSWindowsUpdate -Force
        Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Verbose *>&1
    }

    $poll = 5
    $heartbeat = 60
    $elapsed = 0
    $completed = $false

    while (-not $completed -and $elapsed -lt $OperationTimeoutSeconds) {
        $state = (Get-Job -Id $job.Id).State
        if ($state -in @('Completed','Failed','Stopped')) {
            $completed = $true
            break
        }

        Start-Sleep -Seconds $poll
        $elapsed += $poll

        if (($elapsed % $heartbeat) -eq 0) {
            Write-Log "Windows Update job is still running after $elapsed seconds. State: $state" 'INFO'
        }
    }

    if (-not $completed) {
        Write-Log "Windows Update exceeded the timeout of $OperationTimeoutSeconds seconds." 'ERROR'
        Stop-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
        throw "Install-WindowsUpdate exceeded timeout of $OperationTimeoutSeconds seconds."
    }

    $jobState = $job.State
    $jobReason = $job.ChildJobs[0].JobStateInfo.Reason
    $results = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
    Remove-Job -Job $job -ErrorAction SilentlyContinue | Out-Null

    foreach ($item in $results) {
        $line = [string]$item
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            $script:RawUpdateLines.Add($line) | Out-Null
            Write-ImmediateRuntimeLog -Line $line
        }

        $converted = Convert-WUItem -Item $item -Phase 'Install'
        if ($converted) {
            $script:UpdateEntries.Add($converted) | Out-Null

            $kb = if ([string]::IsNullOrWhiteSpace([string]$converted.KB)) { 'No KB listed' } else { [string]$converted.KB }
            $title = if ([string]::IsNullOrWhiteSpace([string]$converted.Title)) { 'Untitled update' } else { [string]$converted.Title }
            $status = if ([string]::IsNullOrWhiteSpace([string]$converted.Status)) { 'Unknown' } else { [string]$converted.Status }
            $result = if ([string]::IsNullOrWhiteSpace([string]$converted.Result)) { 'Unknown' } else { [string]$converted.Result }
            $size = if ([string]::IsNullOrWhiteSpace([string]$converted.Size)) { 'Unknown' } else { [string]$converted.Size }
            $category = if ([string]::IsNullOrWhiteSpace([string]$converted.Category)) { 'Unknown' } else { [string]$converted.Category }
            $level = if ($result -match 'Failed|Error' -or $status -match 'Failed|Error') { 'ERROR' } `
                elseif ($result -match 'Installed|Succeeded|Success' -or $status -match 'Installed|Succeeded|Success') { 'SUCCESS' } `
                else { 'INFO' }

            Write-Log ("Windows Update result: KB={0} | Title={1} | Status={2} | Result={3} | Size={4} | Category={5}" -f `
                $kb, $title, $status, $result, $size, $category) $level
        }
    }

    if ($jobState -eq 'Failed') {
        $message = if ($jobReason) { $jobReason.Message } else { 'The update job failed without a detailed reason.' }
        throw $message
    }

    if ($results.Count -eq 0) {
        Write-Log 'Install-WindowsUpdate returned no update objects.' 'WARN'
    }
    else {
        Write-Log "Windows Update job returned $($results.Count) output objects." 'OK'
    }
}


function Get-UpdateIdentityKey {
    param([Parameter(Mandatory)][object]$Update)

    $kb = [string]$Update.KB
    $title = [string]$Update.Title

    if (-not [string]::IsNullOrWhiteSpace($kb)) {
        return ('KB:' + $kb.Trim().ToUpperInvariant())
    }

    return ('TITLE:' + $title.Trim().ToUpperInvariant())
}

function Convert-UpdateForTelemetry {
    param(
        [Parameter(Mandatory)][object]$Update,
        [string]$Classification
    )

    return [pscustomobject]@{
        Title          = $Update.Title
        KB             = $Update.KB
        Category       = $Update.Category
        Size           = $Update.Size
        Status         = $Update.Status
        Result         = $Update.Result
        Source         = $Update.Source
        Classification = $Classification
    }
}

function Get-UpdateSummary {
    [object[]]$installEntries = @(
        foreach ($entry in $script:UpdateEntries) {
            if ($entry.Phase -eq 'Install') { $entry }
        }
    )

    [object[]]$installed = @(
        foreach ($entry in $installEntries) {
            if (
                $entry.Result -match 'Installed|Succeeded|Success' -or
                $entry.Status -match 'Installed|Succeeded|Success'
            ) {
                $entry
            }
        }
    )

    [object[]]$failed = @(
        foreach ($entry in $installEntries) {
            if (
                $entry.Result -match 'Failed|Error' -or
                $entry.Status -match 'Failed|Error'
            ) {
                $entry
            }
        }
    )

    [object[]]$before = ConvertTo-NormalArray -InputObject $script:UpdatesBefore
    [object[]]$after = ConvertTo-NormalArray -InputObject $script:UpdatesAfter

    # Install-WindowsUpdate can occasionally discover updates that the immediately
    # preceding Get-WindowsUpdate snapshot did not return. Build a unique candidate
    # set from the install stream so AvailableBefore never incorrectly reports 0
    # when updates were actually accepted/downloaded/installed.
    $installCandidateMap = @{}
    foreach ($entry in $installEntries) {
        $key = Get-UpdateIdentityKey -Update $entry
        if (-not $installCandidateMap.ContainsKey($key)) {
            $installCandidateMap[$key] = $entry
        }
    }
    [object[]]$installCandidates = @(
        foreach ($key in $installCandidateMap.Keys) {
            $installCandidateMap[$key]
        }
    )

    [int]$scanAvailableBefore = $before.Count
    [int]$effectiveAvailableBefore = [math]::Max($scanAvailableBefore, $installCandidates.Count)

    $failedKeys = @{}
    foreach ($entry in $failed) {
        $failedKeys[(Get-UpdateIdentityKey -Update $entry)] = $true
    }

    $afterKeys = @{}
    foreach ($entry in $after) {
        $afterKeys[(Get-UpdateIdentityKey -Update $entry)] = $true
    }

    # "SkippedOrDeferred" means the update was available before installation and
    # is still available afterward without an explicit failed result. This avoids
    # falsely labeling every remaining update as a hard installation failure.
    [object[]]$skippedOrDeferred = @(
        foreach ($entry in $before) {
            $key = Get-UpdateIdentityKey -Update $entry
            if ($afterKeys.ContainsKey($key) -and -not $failedKeys.ContainsKey($key)) {
                $entry
            }
        }
    )

    [object[]]$installedDetails = @(
        foreach ($entry in $installed) {
            Convert-UpdateForTelemetry -Update $entry -Classification 'Installed'
        }
    )

    [object[]]$failedDetails = @(
        foreach ($entry in $failed) {
            Convert-UpdateForTelemetry -Update $entry -Classification 'Failed'
        }
    )

    [object[]]$remainingDetails = @(
        foreach ($entry in $after) {
            Convert-UpdateForTelemetry -Update $entry -Classification 'Remaining'
        }
    )

    [object[]]$skippedDetails = @(
        foreach ($entry in $skippedOrDeferred) {
            Convert-UpdateForTelemetry -Update $entry -Classification 'SkippedOrDeferred'
        }
    )

    $kbs = @(
        $installed |
        ForEach-Object {
            if ($_.KB) { $_.KB -split ',\s*' }
        } |
        Where-Object { $_ -match '(?i)^KB\d+$' } |
        ForEach-Object { $_.ToUpperInvariant() } |
        Sort-Object -Unique
    )

    return [pscustomobject]@{
        AvailableBefore              = $effectiveAvailableBefore
        AvailableBeforeScan          = $scanAvailableBefore
        InstallCandidatesDetected    = $installCandidates.Count
        Installed                    = $installed.Count
        Failed                       = $failed.Count
        RemainingAfter               = $after.Count
        SkippedOrDeferred            = $skippedOrDeferred.Count

        InstalledKBs                 = $kbs
        InstalledUpdates             = $installedDetails
        FailedUpdates                = $failedDetails
        RemainingUpdates             = $remainingDetails
        SkippedOrDeferredUpdates     = $skippedDetails

        FeatureUpdatesDetected       = @((@($before) + @($installCandidates)) | Where-Object { $_.Title -match 'Feature update|Windows 11, version|Windows 10, version' } | Sort-Object Title -Unique).Count
        CumulativeUpdatesDetected    = @((@($before) + @($installCandidates)) | Where-Object { $_.Title -match 'Cumulative Update' } | Sort-Object Title -Unique).Count
        DefenderUpdatesDetected      = @((@($before) + @($installCandidates)) | Where-Object { $_.Title -match 'Defender|Security Intelligence' } | Sort-Object Title -Unique).Count
    }
}

function Write-YamlLog {
    try {
        $end = Get-Date
        $duration = [math]::Round(($end - $script:RunStart).TotalSeconds, 2)
        $summary = Get-UpdateSummary
        $lines = New-Object System.Collections.Generic.List[string]

        foreach ($pair in @(
            @('computer_name',$script:ComputerName),
            @('domain',$script:Domain),
            @('script_name',$script:ScriptName),
            @('script_version',$script:ScriptVersion),
            @('run_id',$script:RunId),
            @('status',$script:OverallResult),
            @('exit_code',$script:ExitCode),
            @('run_started',$script:RunStart.ToString('o')),
            @('run_finished',$end.ToString('o')),
            @('duration_seconds',$duration),
            @('warning_count',$script:WarningCount),
            @('error_count',$script:ErrorCount),
            @('pswindowsupdate_version',$script:ModuleVersion),
            @('reset_wu_components_first',[bool]$ResetWUComponentsFirst),
            @('reboot_required_before',$script:RebootBefore.Required),
            @('reboot_required_after',$script:RebootAfter.Required),
            @('reboot_initiated',$script:RebootInitiated),
            @('updates_available_before',$summary.AvailableBefore),
            @('updates_installed',$summary.Installed),
            @('updates_failed',$summary.Failed),
            @('updates_remaining_after',$summary.RemainingAfter),
            @('updates_skipped_or_deferred',$summary.SkippedOrDeferred)
        )) {
            $lines.Add("$($pair[0]): $(ConvertTo-YamlSafeValue $pair[1])") | Out-Null
        }

        $lines.Add('installed_kbs:') | Out-Null
        if ($summary.InstalledKBs.Count) {
            foreach ($kb in $summary.InstalledKBs) { $lines.Add("  - $(ConvertTo-YamlSafeValue $kb)") | Out-Null }
        } else { $lines.Add('  - null') | Out-Null }

        $lines.Add('reboot_reasons_after:') | Out-Null
        if ($script:RebootAfter.Reasons.Count) {
            foreach ($reason in $script:RebootAfter.Reasons) { $lines.Add("  - $(ConvertTo-YamlSafeValue $reason)") | Out-Null }
        } else { $lines.Add('  - null') | Out-Null }

        $lines.Add('updates:') | Out-Null
        foreach ($entry in @($script:UpdatesBefore + (ConvertTo-NormalArray -InputObject $script:UpdateEntries) + $script:UpdatesAfter)) {
            $lines.Add("  - phase: $(ConvertTo-YamlSafeValue $entry.Phase)") | Out-Null
            $lines.Add("    title: $(ConvertTo-YamlSafeValue $entry.Title)") | Out-Null
            $lines.Add("    kb: $(ConvertTo-YamlSafeValue $entry.KB)") | Out-Null
            $lines.Add("    category: $(ConvertTo-YamlSafeValue $entry.Category)") | Out-Null
            $lines.Add("    size: $(ConvertTo-YamlSafeValue $entry.Size)") | Out-Null
            $lines.Add("    status: $(ConvertTo-YamlSafeValue $entry.Status)") | Out-Null
            $lines.Add("    result: $(ConvertTo-YamlSafeValue $entry.Result)") | Out-Null
        }

        $lines.Add('actions:') | Out-Null
        foreach ($action in $script:ActionHistory) {
            $lines.Add("  - time: $(ConvertTo-YamlSafeValue $action.Time)") | Out-Null
            $lines.Add("    level: $(ConvertTo-YamlSafeValue $action.Level)") | Out-Null
            $lines.Add("    message: $(ConvertTo-YamlSafeValue $action.Message)") | Out-Null
        }

        if ($script:FailureMessage) {
            $lines.Add("failure_message: $(ConvertTo-YamlSafeValue $script:FailureMessage)") | Out-Null
        }

        Set-Content -LiteralPath $script:YamlLogPath -Value (ConvertTo-NormalArray -InputObject $lines) -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
    }
}

function Write-Telemetry {
    try {
        Ensure-Folder -Path $LogFolder
        $end = Get-Date
        $summary = Get-UpdateSummary

        $event = [ordered]@{
            EventType = 'maintenance.execution'
            ComputerName = $script:ComputerName
            Domain = $script:Domain
            ScriptName = $script:ScriptName
            ScriptVersion = $script:ScriptVersion
            RunId = $script:RunId
            Status = $script:OverallResult
            ExitCode = $script:ExitCode
            StartTime = $script:RunStart.ToString('o')
            EndTime = $end.ToString('o')
            DurationSeconds = [math]::Round(($end - $script:RunStart).TotalSeconds, 2)
            WarningCount = $script:WarningCount
            ErrorCount = $script:ErrorCount
            Timestamp = $end.ToUniversalTime().ToString('o')
            FailureMessage = $script:FailureMessage
            PSWindowsUpdateVersion = $script:ModuleVersion
            ResetWUComponentsFirst = [bool]$ResetWUComponentsFirst
            ResetResult = $script:ResetResult
            WindowsBuildBefore = $script:BuildBefore
            WindowsBuildAfter = $script:BuildAfter
            LastSuccessfulUpdateBefore = $script:LastUpdateBefore
            LastSuccessfulUpdateAfter = $script:LastUpdateAfter
            UpdatesAvailableBefore = $summary.AvailableBefore
            UpdatesAvailableBeforeScan = $summary.AvailableBeforeScan
            InstallCandidatesDetected = $summary.InstallCandidatesDetected
            UpdatesInstalled = $summary.Installed
            UpdatesFailed = $summary.Failed
            UpdatesRemainingAfter = $summary.RemainingAfter
            UpdatesSkippedOrDeferred = $summary.SkippedOrDeferred

            InstalledKBs = $summary.InstalledKBs
            InstalledUpdates = $summary.InstalledUpdates
            FailedUpdates = $summary.FailedUpdates
            RemainingUpdates = $summary.RemainingUpdates
            SkippedOrDeferredUpdates = $summary.SkippedOrDeferredUpdates
            FeatureUpdatesDetected = $summary.FeatureUpdatesDetected
            CumulativeUpdatesDetected = $summary.CumulativeUpdatesDetected
            DefenderUpdatesDetected = $summary.DefenderUpdatesDetected
            PendingRebootBefore = $script:RebootBefore.Required
            PendingRebootReasonsBefore = $script:RebootBefore.Reasons
            PendingRebootAfter = $script:RebootAfter.Required
            PendingRebootReasonsAfter = $script:RebootAfter.Reasons
            RebootInitiated = $script:RebootInitiated
            TextLogPath = $script:PublishedLogPath
            UpdateResults = (ConvertTo-NormalArray -InputObject $script:UpdateEntries)
            AvailableUpdatesBefore = (ConvertTo-NormalArray -InputObject $script:UpdatesBefore)
            RemainingUpdatesAfter = (ConvertTo-NormalArray -InputObject $script:UpdatesAfter)
        }

        $compact = $event | ConvertTo-Json -Depth 12 -Compress
        Write-MaintenanceTelemetryLine -Path $script:TelemetryPath -JsonLine $compact
        $event | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $script:LatestJsonPath -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write JSON telemetry: $($_.Exception.Message)"
    }
}

function Complete-Run {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$Code
    )

    $script:OverallResult = $Status
    $script:ExitCode = $Code

    Write-YamlLog
    Write-Telemetry

    # This must be the final append to the text log. Once the completed file is
    # published into C:\Logs, Elastic may open it immediately.
    Write-Log ("Completed {0}. Status={1}; ExitCode={2}; Warnings={3}; Errors={4}" -f `
        $script:ScriptName,
        $Status,
        $Code,
        $script:WarningCount,
        $script:ErrorCount) $(if ($Code -eq 0 -or $Code -eq 3010) { 'OK' } else { 'ERROR' })

    if ($null -ne $script:LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $script:LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 06 text log for Elastic: {0}" -f $script:PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 06 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

# Main
Initialize-Logging

if (-not (Test-IsAdministrator)) {
    $script:FailureMessage = 'The script must run as Administrator or SYSTEM.'
    Write-Log $script:FailureMessage 'ERROR'
    Complete-Run -Status 'Failed' -Code 1
    exit 1
}

try {
    Write-Log "Starting $($script:ScriptName) version $($script:ScriptVersion)." 'INFO'
    Write-Log "Active staged text log path: $($script:RuntimeLogPath)" 'INFO'
    Write-Log "Completed text log publish path: $($script:PublishedLogPath)" 'INFO'

    Ensure-PSWindowsUpdate

    $script:BuildBefore = Get-WindowsBuildSnapshot
    $script:LastUpdateBefore = Get-LastSuccessfulUpdate
    $script:RebootBefore = Get-PendingRebootState
    $script:UpdatesBefore = @(Get-AvailableUpdateSnapshot -Phase 'Before')

    if ($ResetWUComponentsFirst) {
        Reset-WUComponentsSafe
        Import-Module PSWindowsUpdate -Force
    }

    Install-AvailableWindowsUpdates

    $script:BuildAfter = Get-WindowsBuildSnapshot
    $script:LastUpdateAfter = Get-LastSuccessfulUpdate
    $script:UpdatesAfter = @(Get-AvailableUpdateSnapshot -Phase 'After')
    $script:RebootAfter = Get-PendingRebootState

    $summary = Get-UpdateSummary

    if ($summary.AvailableBeforeScan -lt $summary.AvailableBefore) {
        Write-Log ("Pre-scan reported {0} update(s), but the install stream identified {1} unique update candidate(s). Using {1} for UpdatesAvailableBefore telemetry." -f `
            $summary.AvailableBeforeScan,
            $summary.InstallCandidatesDetected) 'WARN'
    }

    Write-Log ("Elastic Windows Update summary: AvailableBefore={0}; Installed={1}; Failed={2}; RemainingAfter={3}; SkippedOrDeferred={4}" -f `
        $summary.AvailableBefore,
        $summary.Installed,
        $summary.Failed,
        $summary.RemainingAfter,
        $summary.SkippedOrDeferred) 'INFO'

    foreach ($update in $summary.InstalledUpdates) {
        Write-Log ("Installed update: KB={0} | Title={1} | Result={2}" -f $update.KB, $update.Title, $update.Result) 'SUCCESS'
    }
    foreach ($update in $summary.FailedUpdates) {
        Write-Log ("Failed update: KB={0} | Title={1} | Result={2}" -f $update.KB, $update.Title, $update.Result) 'ERROR'
    }
    foreach ($update in $summary.SkippedOrDeferredUpdates) {
        Write-Log ("Skipped/deferred update remains available: KB={0} | Title={1}" -f $update.KB, $update.Title) 'WARNING'
    }

    if ($summary.Failed -gt 0) {
        $status = 'CompletedWithUpdateFailures'
        $code = 3
    }
    elseif ($summary.RemainingAfter -gt 0) {
        $status = 'CompletedWithRemainingUpdates'
        $code = 0
    }
    elseif ($script:RebootAfter.Required) {
        $status = 'SuccessRebootRequired'
        $code = 3010
    }
    elseif ($script:WarningCount -gt 0) {
        $status = 'SuccessWithWarnings'
        $code = 0
    }
    else {
        $status = 'Success'
        $code = 0
    }

    if ($script:RebootAfter.Required) {
        Write-Log 'Windows reports that a reboot is required. This script will not initiate the reboot; the dedicated reboot maintenance script will handle it.' 'INFO'
    }

    Complete-Run -Status $status -Code $code
    exit $code
}
catch {
    $script:FailureMessage = $_.Exception.Message
    Write-Log "Script failed: $($script:FailureMessage)" 'ERROR'

    if (-not $script:BuildAfter) { $script:BuildAfter = Get-WindowsBuildSnapshot }
    if (-not $script:LastUpdateAfter) { $script:LastUpdateAfter = Get-LastSuccessfulUpdate }
    if (-not $script:RebootAfter) { $script:RebootAfter = Get-PendingRebootState }

    Complete-Run -Status 'Failed' -Code 2
    exit 2
}
﻿# =====================================================================
# ScriptName: 06_Weekend_Windows_Updates.ps1
# ScriptVersion: 2.1.0
# LastUpdated: 2026-07-27
# Purpose: Installs Microsoft/Windows updates using PSWindowsUpdate,
#          records detailed YAML and Elastic-compatible JSON telemetry,
#          validates update compliance before and after installation,
#          and explicitly restarts when Windows reports that it is needed.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$ResetWUComponentsFirst = $false,
    [int]$OperationTimeoutSeconds = 1800,
    [int]$RebootDelaySeconds = 30,
    [string]$LogFolder = "$env:SystemDrive\Logs"
)

$ErrorActionPreference = 'Stop'

$script:ScriptName       = '06_Weekend_Windows_Updates.ps1'
$script:ScriptVersion    = '2.1.0'
$script:RunId            = [guid]::NewGuid().Guid
$script:RunStart         = Get-Date
$script:ComputerName     = $env:COMPUTERNAME
$script:Domain           = $env:USERDOMAIN
$script:YamlLogPath      = $null
$script:RuntimeLogPath   = $null
$script:TelemetryPath    = Join-Path $LogFolder 'Maintenance-Telemetry.ndjson'
$script:LatestJsonPath   = Join-Path $LogFolder '06_Weekend_Windows_Updates.latest.json'
$script:ActionHistory    = New-Object System.Collections.Generic.List[object]
$script:RawUpdateLines   = New-Object System.Collections.Generic.List[string]
$script:UpdateEntries    = New-Object System.Collections.Generic.List[object]
$script:WarningCount     = 0
$script:ErrorCount       = 0
$script:OverallResult    = 'Initializing'
$script:ExitCode         = 0
$script:FailureMessage   = $null
$script:RebootInitiated  = $false
$script:RebootBefore     = $null
$script:RebootAfter      = $null
$script:ModuleVersion    = $null
$script:BuildBefore      = $null
$script:BuildAfter       = $null
$script:LastUpdateBefore = $null
$script:LastUpdateAfter  = $null
$script:UpdatesBefore    = @()
$script:UpdatesAfter     = @()
$script:ResetResult      = [ordered]@{
    Requested = [bool]$ResetWUComponentsFirst
    Completed = $false
    ServicesStopped = @()
    ServicesStarted = @()
    ServiceFailures = @()
    FoldersCleared = @()
    FolderFailures = @()
    RegistryKeysRemoved = @()
    RegistryFailures = @()
}

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot 'C:\Logs'

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Write-ImmediateRuntimeLog {
    param([Parameter(Mandatory)][string]$Line)
    try {
        if ($script:RuntimeLogPath) {
            Add-Content -LiteralPath $script:RuntimeLogPath -Value $Line -Encoding UTF8
        }
    }
    catch {}
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN')  { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }

    $script:ActionHistory.Add([pscustomobject]@{
        Time = $timestamp
        Level = $Level
        Message = $Message
    }) | Out-Null

    Write-ImmediateRuntimeLog -Line $line
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

function ConvertTo-YamlSafeValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [string]$Value
    }

    $text = [string]$Value
    $text = $text -replace "`r", ' '
    $text = $text -replace "`n", ' '
    $text = $text -replace '"', '\"'
    return '"' + $text + '"'
}

function Initialize-Logging {
    Ensure-Folder -Path $LogFolder
    $stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $base = "$($script:ComputerName)-WindowsUpdates-$stamp"
    $script:YamlLogPath = Join-Path $LogFolder "$base.yaml"
    $script:RuntimeLogPath = Join-Path $LogFolder "$base.log"

    Set-Content -LiteralPath $script:RuntimeLogPath -Encoding UTF8 -Value @(
        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO ] Runtime log initialized.",
        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO ] Run ID: $($script:RunId)"
    )
}

function Get-WindowsBuildSnapshot {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        return [pscustomobject]@{
            Caption = $os.Caption
            Version = $os.Version
            BuildNumber = $os.BuildNumber
            UBR = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR -ErrorAction SilentlyContinue).UBR
            DisplayVersion = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
        }
    }
    catch {
        Write-Log "Unable to collect Windows build information: $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Get-LastSuccessfulUpdate {
    try {
        $hotfix = Get-HotFix -ErrorAction Stop |
            Where-Object { $_.InstalledOn } |
            Sort-Object InstalledOn -Descending |
            Select-Object -First 1

        if ($hotfix) {
            return [pscustomobject]@{
                KB = $hotfix.HotFixID
                InstalledOn = $hotfix.InstalledOn
                Description = $hotfix.Description
            }
        }
    }
    catch {
        Write-Log "Unable to determine the latest installed hotfix: $($_.Exception.Message)" 'WARN'
    }
    return $null
}

function Get-PendingRebootState {
    $reasons = New-Object System.Collections.Generic.List[string]

    $checks = @(
        @{ Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'; Reason='ComponentBasedServicing' },
        @{ Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'; Reason='WindowsUpdate' },
        @{ Path='HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'; Name='PendingFileRenameOperations'; Reason='PendingFileRenameOperations' },
        @{ Path='HKLM:\SOFTWARE\Microsoft\Updates'; Name='UpdateExeVolatile'; Reason='UpdateExeVolatile' }
    )

    foreach ($check in $checks) {
        try {
            if ($check.ContainsKey('Name')) {
                $value = (Get-ItemProperty -LiteralPath $check.Path -Name $check.Name -ErrorAction SilentlyContinue).($check.Name)
                if ($null -ne $value -and "$value" -ne '' -and "$value" -ne '0') {
                    $reasons.Add($check.Reason) | Out-Null
                }
            }
            elseif (Test-Path -LiteralPath $check.Path) {
                $reasons.Add($check.Reason) | Out-Null
            }
        }
        catch {}
    }

    try {
        if (Get-Command -Name Get-WURebootStatus -ErrorAction SilentlyContinue) {
            if ([bool](Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)) {
                $reasons.Add('PSWindowsUpdate') | Out-Null
            }
        }
    }
    catch {}

    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        if ([bool]$sysInfo.RebootRequired) {
            $reasons.Add('Microsoft.Update.SystemInfo') | Out-Null
        }
    }
    catch {}

    $unique = @($reasons | Sort-Object -Unique)
    return [pscustomobject]@{
        Required = ($unique.Count -gt 0)
        Reasons = $unique
    }
}

function Ensure-PSWindowsUpdate {
    Write-Log 'Ensuring the PSWindowsUpdate module is available...' 'INFO'

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Log 'Installing the NuGet package provider...' 'INFO'
        Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null
    }

    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-Log 'Setting PSGallery to Trusted...' 'INFO'
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
    }
    catch {
        Write-Log "Could not validate PSGallery settings: $($_.Exception.Message)" 'WARN'
    }

    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log 'Installing PSWindowsUpdate...' 'INFO'
        Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Scope AllUsers
    }

    Import-Module PSWindowsUpdate -Force
    $module = Get-Module -Name PSWindowsUpdate | Sort-Object Version -Descending | Select-Object -First 1
    if ($module) { $script:ModuleVersion = $module.Version.ToString() }

    Write-Log "PSWindowsUpdate imported successfully. Version: $($script:ModuleVersion)" 'OK'
}

function Convert-WUItem {
    param(
        [Parameter(Mandatory)][object]$Item,
        [string]$Phase
    )

    $props = $Item.PSObject.Properties.Name

    function Get-FirstPropertyValue {
        param([string[]]$Names)
        foreach ($name in $Names) {
            if ($props -contains $name -and $null -ne $Item.$name) {
                $value = $Item.$name
                if ($value -is [array]) { return (($value | ForEach-Object { [string]$_ }) -join ', ') }
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
            }
        }
        return $null
    }

    $title = Get-FirstPropertyValue @('Title','KBArticleTitle')
    $kb = Get-FirstPropertyValue @('KB','KBArticleIDs','KBArticleID')
    $size = Get-FirstPropertyValue @('Size','MaxDownloadSize')
    $status = Get-FirstPropertyValue @('Status','Result','UpdateState')
    $result = Get-FirstPropertyValue @('Result','Status','HResult')
    $category = Get-FirstPropertyValue @('Categories','Category','CategoryNames')

    if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($kb)) {
        return $null
    }

    return [pscustomobject]@{
        Phase = $Phase
        Title = $title
        KB = $kb
        Size = $size
        Category = $category
        Status = $status
        Result = $result
        Source = 'MicrosoftUpdate'
    }
}

function Get-AvailableUpdateSnapshot {
    param([Parameter(Mandatory)][string]$Phase)

    Write-Log "Scanning for available updates ($Phase)..." 'INFO'
    $items = @()

    try {
        $results = @(Get-WindowsUpdate -MicrosoftUpdate -ErrorAction Stop)
        foreach ($item in $results) {
            $converted = Convert-WUItem -Item $item -Phase $Phase
            if ($converted) { $items += $converted }
        }

        Write-Log "Available updates found during $Phase scan: $($items.Count)" 'INFO'
    }
    catch {
        Write-Log "The $Phase update scan failed: $($_.Exception.Message)" 'WARN'
    }

    return @($items)
}

function Reset-WUComponentsSafe {
    Write-Log 'Resetting Windows Update components...' 'INFO'
    $services = @('wuauserv','bits','cryptsvc','msiserver')

    foreach ($serviceName in $services) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction Stop
            if ($service.Status -ne 'Stopped') {
                Stop-Service -Name $serviceName -Force -ErrorAction Stop
            }
            $script:ResetResult.ServicesStopped += $serviceName
        }
        catch {
            $script:ResetResult.ServiceFailures += "Stop ${serviceName}: $($_.Exception.Message)"
            Write-Log "Failed to stop ${serviceName}: $($_.Exception.Message)" 'WARN'
        }
    }

    Start-Sleep -Seconds 3

    $registryKeys = @(
        'HKLM\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy',
        'HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Update'
    )

    foreach ($key in $registryKeys) {
        try {
            $psPath = 'Registry::' + $key
            if (Test-Path -LiteralPath $psPath) {
                & reg.exe delete $key /f | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    $script:ResetResult.RegistryKeysRemoved += $key
                }
                else {
                    throw "reg.exe returned $LASTEXITCODE"
                }
            }
        }
        catch {
            $script:ResetResult.RegistryFailures += "${key}: $($_.Exception.Message)"
            Write-Log "Failed to remove registry key ${key}: $($_.Exception.Message)" 'WARN'
        }
    }

    foreach ($folder in @("$env:SystemRoot\SoftwareDistribution","$env:SystemRoot\System32\catroot2")) {
        try {
            if (Test-Path -LiteralPath $folder) {
                Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
                $script:ResetResult.FoldersCleared += $folder
            }
        }
        catch {
            $script:ResetResult.FolderFailures += "${folder}: $($_.Exception.Message)"
            Write-Log "Failed to clear ${folder}: $($_.Exception.Message)" 'WARN'
        }
    }

    foreach ($serviceName in $services) {
        try {
            Start-Service -Name $serviceName -ErrorAction Stop
            $script:ResetResult.ServicesStarted += $serviceName
        }
        catch {
            $script:ResetResult.ServiceFailures += "Start ${serviceName}: $($_.Exception.Message)"
            Write-Log "Failed to start ${serviceName}: $($_.Exception.Message)" 'WARN'
        }
    }

    $script:ResetResult.Completed = ($script:ResetResult.ServiceFailures.Count -eq 0)
    Write-Log 'Windows Update component reset completed.' 'OK'
}

function Install-AvailableWindowsUpdates {
    Write-Log 'Starting Windows Update installation job...' 'INFO'

    $job = Start-Job -ScriptBlock {
        $ErrorActionPreference = 'Stop'
        Import-Module PSWindowsUpdate -Force
        Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Verbose *>&1
    }

    $poll = 5
    $heartbeat = 60
    $elapsed = 0
    $completed = $false

    while (-not $completed -and $elapsed -lt $OperationTimeoutSeconds) {
        $state = (Get-Job -Id $job.Id).State
        if ($state -in @('Completed','Failed','Stopped')) {
            $completed = $true
            break
        }

        Start-Sleep -Seconds $poll
        $elapsed += $poll

        if (($elapsed % $heartbeat) -eq 0) {
            Write-Log "Windows Update job is still running after $elapsed seconds. State: $state" 'INFO'
        }
    }

    if (-not $completed) {
        Write-Log "Windows Update exceeded the timeout of $OperationTimeoutSeconds seconds." 'ERROR'
        Stop-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
        throw "Install-WindowsUpdate exceeded timeout of $OperationTimeoutSeconds seconds."
    }

    $jobState = $job.State
    $jobReason = $job.ChildJobs[0].JobStateInfo.Reason
    $results = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
    Remove-Job -Job $job -ErrorAction SilentlyContinue | Out-Null

    foreach ($item in $results) {
        $line = [string]$item
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            $script:RawUpdateLines.Add($line) | Out-Null
            Write-ImmediateRuntimeLog -Line $line
        }

        $converted = Convert-WUItem -Item $item -Phase 'Install'
        if ($converted) {
            $script:UpdateEntries.Add($converted) | Out-Null
        }
    }

    if ($jobState -eq 'Failed') {
        $message = if ($jobReason) { $jobReason.Message } else { 'The update job failed without a detailed reason.' }
        throw $message
    }

    if ($results.Count -eq 0) {
        Write-Log 'Install-WindowsUpdate returned no update objects.' 'WARN'
    }
    else {
        Write-Log "Windows Update job returned $($results.Count) output objects." 'OK'
    }
}

function Get-UpdateSummary {
    $installed = @($script:UpdateEntries | Where-Object {
        $_.Phase -eq 'Install' -and (
            $_.Result -match 'Installed|Succeeded|Success' -or
            $_.Status -match 'Installed|Succeeded|Success'
        )
    })

    $failed = @($script:UpdateEntries | Where-Object {
        $_.Phase -eq 'Install' -and (
            $_.Result -match 'Failed|Error' -or
            $_.Status -match 'Failed|Error'
        )
    })

    $kbs = @($installed | ForEach-Object {
        if ($_.KB) { $_.KB -split ',\s*' }
    } | Where-Object { $_ } | Sort-Object -Unique)

    return [pscustomobject]@{
        AvailableBefore = @($script:UpdatesBefore).Count
        Installed = $installed.Count
        Failed = $failed.Count
        RemainingAfter = @($script:UpdatesAfter).Count
        InstalledKBs = $kbs
        FeatureUpdatesDetected = @($script:UpdatesBefore | Where-Object { $_.Title -match 'Feature update|Windows 11, version|Windows 10, version' }).Count
        CumulativeUpdatesDetected = @($script:UpdatesBefore | Where-Object { $_.Title -match 'Cumulative Update' }).Count
        DefenderUpdatesDetected = @($script:UpdatesBefore | Where-Object { $_.Title -match 'Defender|Security Intelligence' }).Count
    }
}

function Write-YamlLog {
    try {
        $end = Get-Date
        $duration = [math]::Round(($end - $script:RunStart).TotalSeconds, 2)
        $summary = Get-UpdateSummary
        $lines = New-Object System.Collections.Generic.List[string]

        foreach ($pair in @(
            @('computer_name',$script:ComputerName),
            @('domain',$script:Domain),
            @('script_name',$script:ScriptName),
            @('script_version',$script:ScriptVersion),
            @('run_id',$script:RunId),
            @('status',$script:OverallResult),
            @('exit_code',$script:ExitCode),
            @('run_started',$script:RunStart.ToString('o')),
            @('run_finished',$end.ToString('o')),
            @('duration_seconds',$duration),
            @('warning_count',$script:WarningCount),
            @('error_count',$script:ErrorCount),
            @('pswindowsupdate_version',$script:ModuleVersion),
            @('reset_wu_components_first',[bool]$ResetWUComponentsFirst),
            @('reboot_required_before',$script:RebootBefore.Required),
            @('reboot_required_after',$script:RebootAfter.Required),
            @('reboot_initiated',$script:RebootInitiated),
            @('updates_available_before',$summary.AvailableBefore),
            @('updates_installed',$summary.Installed),
            @('updates_failed',$summary.Failed),
            @('updates_remaining_after',$summary.RemainingAfter)
        )) {
            $lines.Add("$($pair[0]): $(ConvertTo-YamlSafeValue $pair[1])") | Out-Null
        }

        $lines.Add('installed_kbs:') | Out-Null
        if ($summary.InstalledKBs.Count) {
            foreach ($kb in $summary.InstalledKBs) { $lines.Add("  - $(ConvertTo-YamlSafeValue $kb)") | Out-Null }
        } else { $lines.Add('  - null') | Out-Null }

        $lines.Add('reboot_reasons_after:') | Out-Null
        if ($script:RebootAfter.Reasons.Count) {
            foreach ($reason in $script:RebootAfter.Reasons) { $lines.Add("  - $(ConvertTo-YamlSafeValue $reason)") | Out-Null }
        } else { $lines.Add('  - null') | Out-Null }

        $lines.Add('updates:') | Out-Null
        foreach ($entry in @($script:UpdatesBefore + @($script:UpdateEntries) + $script:UpdatesAfter)) {
            $lines.Add("  - phase: $(ConvertTo-YamlSafeValue $entry.Phase)") | Out-Null
            $lines.Add("    title: $(ConvertTo-YamlSafeValue $entry.Title)") | Out-Null
            $lines.Add("    kb: $(ConvertTo-YamlSafeValue $entry.KB)") | Out-Null
            $lines.Add("    category: $(ConvertTo-YamlSafeValue $entry.Category)") | Out-Null
            $lines.Add("    size: $(ConvertTo-YamlSafeValue $entry.Size)") | Out-Null
            $lines.Add("    status: $(ConvertTo-YamlSafeValue $entry.Status)") | Out-Null
            $lines.Add("    result: $(ConvertTo-YamlSafeValue $entry.Result)") | Out-Null
        }

        $lines.Add('actions:') | Out-Null
        foreach ($action in $script:ActionHistory) {
            $lines.Add("  - time: $(ConvertTo-YamlSafeValue $action.Time)") | Out-Null
            $lines.Add("    level: $(ConvertTo-YamlSafeValue $action.Level)") | Out-Null
            $lines.Add("    message: $(ConvertTo-YamlSafeValue $action.Message)") | Out-Null
        }

        if ($script:FailureMessage) {
            $lines.Add("failure_message: $(ConvertTo-YamlSafeValue $script:FailureMessage)") | Out-Null
        }

        Set-Content -LiteralPath $script:YamlLogPath -Value $lines -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
    }
}

function Write-Telemetry {
    try {
        Ensure-Folder -Path $LogFolder
        $end = Get-Date
        $summary = Get-UpdateSummary

        $event = [ordered]@{
            EventType = 'maintenance.execution'
            ComputerName = $script:ComputerName
            Domain = $script:Domain
            ScriptName = $script:ScriptName
            ScriptVersion = $script:ScriptVersion
            RunId = $script:RunId
            Status = $script:OverallResult
            ExitCode = $script:ExitCode
            StartTime = $script:RunStart.ToString('o')
            EndTime = $end.ToString('o')
            DurationSeconds = [math]::Round(($end - $script:RunStart).TotalSeconds, 2)
            WarningCount = $script:WarningCount
            ErrorCount = $script:ErrorCount
            Timestamp = $end.ToUniversalTime().ToString('o')
            FailureMessage = $script:FailureMessage
            PSWindowsUpdateVersion = $script:ModuleVersion
            ResetWUComponentsFirst = [bool]$ResetWUComponentsFirst
            ResetResult = $script:ResetResult
            WindowsBuildBefore = $script:BuildBefore
            WindowsBuildAfter = $script:BuildAfter
            LastSuccessfulUpdateBefore = $script:LastUpdateBefore
            LastSuccessfulUpdateAfter = $script:LastUpdateAfter
            UpdatesAvailableBefore = $summary.AvailableBefore
            UpdatesInstalled = $summary.Installed
            UpdatesFailed = $summary.Failed
            UpdatesRemainingAfter = $summary.RemainingAfter
            InstalledKBs = $summary.InstalledKBs
            FeatureUpdatesDetected = $summary.FeatureUpdatesDetected
            CumulativeUpdatesDetected = $summary.CumulativeUpdatesDetected
            DefenderUpdatesDetected = $summary.DefenderUpdatesDetected
            PendingRebootBefore = $script:RebootBefore.Required
            PendingRebootReasonsBefore = $script:RebootBefore.Reasons
            PendingRebootAfter = $script:RebootAfter.Required
            PendingRebootReasonsAfter = $script:RebootAfter.Reasons
            RebootInitiated = $script:RebootInitiated
            RebootDelaySeconds = $RebootDelaySeconds
            UpdateResults = @($script:UpdateEntries)
            AvailableUpdatesBefore = @($script:UpdatesBefore)
            RemainingUpdatesAfter = @($script:UpdatesAfter)
        }

        $compact = $event | ConvertTo-Json -Depth 12 -Compress
        Write-MaintenanceTelemetryLine -Path $script:TelemetryPath -JsonLine $compact
        $event | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $script:LatestJsonPath -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write JSON telemetry: $($_.Exception.Message)"
    }
}

function Complete-Run {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$Code
    )
    $script:OverallResult = $Status
    $script:ExitCode = $Code
    Write-YamlLog
    Write-Telemetry
}

function Invoke-ExplicitReboot {
    param([int]$Delay = 30)

    $comment = 'Restarting to complete Windows Update installation.'
    $arguments = @('/r','/t',$Delay.ToString(),'/d','p:2:17','/c',"`"$comment`"",' /f')
    # Remove the accidental leading space from the final argument.
    $arguments[$arguments.Count - 1] = '/f'

    Write-Log "Issuing reboot command with a $Delay-second delay." 'INFO'
    & "$env:SystemRoot\System32\shutdown.exe" @arguments

    if ($LASTEXITCODE -ne 0) {
        throw "shutdown.exe returned exit code $LASTEXITCODE"
    }

    $script:RebootInitiated = $true
    Write-Log 'Reboot command issued successfully.' 'OK'
}

# Main
Initialize-Logging

if (-not (Test-IsAdministrator)) {
    $script:FailureMessage = 'The script must run as Administrator or SYSTEM.'
    Write-Log $script:FailureMessage 'ERROR'
    Complete-Run -Status 'Failed' -Code 1
    exit 1
}

try {
    Write-Log "Starting $($script:ScriptName) version $($script:ScriptVersion)." 'INFO'

    Ensure-PSWindowsUpdate

    $script:BuildBefore = Get-WindowsBuildSnapshot
    $script:LastUpdateBefore = Get-LastSuccessfulUpdate
    $script:RebootBefore = Get-PendingRebootState
    $script:UpdatesBefore = @(Get-AvailableUpdateSnapshot -Phase 'Before')

    if ($ResetWUComponentsFirst) {
        Reset-WUComponentsSafe
        Import-Module PSWindowsUpdate -Force
    }

    Install-AvailableWindowsUpdates

    $script:BuildAfter = Get-WindowsBuildSnapshot
    $script:LastUpdateAfter = Get-LastSuccessfulUpdate
    $script:UpdatesAfter = @(Get-AvailableUpdateSnapshot -Phase 'After')
    $script:RebootAfter = Get-PendingRebootState

    $summary = Get-UpdateSummary

    if ($summary.Failed -gt 0) {
        $status = 'CompletedWithUpdateFailures'
        $code = 3
    }
    elseif ($summary.RemainingAfter -gt 0) {
        $status = 'CompletedWithRemainingUpdates'
        $code = 0
    }
    elseif ($script:RebootAfter.Required) {
        $status = 'SuccessRebootRequired'
        $code = 3010
    }
    elseif ($script:WarningCount -gt 0) {
        $status = 'SuccessWithWarnings'
        $code = 0
    }
    else {
        $status = 'Success'
        $code = 0
    }

    Complete-Run -Status $status -Code $code

    if ($script:RebootAfter.Required) {
        Invoke-ExplicitReboot -Delay $RebootDelaySeconds
        # Rewrite telemetry so RebootInitiated is captured.
        Complete-Run -Status 'SuccessRebootInitiated' -Code 3010
        exit 3010
    }

    exit $code
}
catch {
    $script:FailureMessage = $_.Exception.Message
    Write-Log "Script failed: $($script:FailureMessage)" 'ERROR'

    if (-not $script:BuildAfter) { $script:BuildAfter = Get-WindowsBuildSnapshot }
    if (-not $script:LastUpdateAfter) { $script:LastUpdateAfter = Get-LastSuccessfulUpdate }
    if (-not $script:RebootAfter) { $script:RebootAfter = Get-PendingRebootState }

    Complete-Run -Status 'Failed' -Code 2
    exit 2
}
