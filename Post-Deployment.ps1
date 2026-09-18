# =====================================================================
# ScriptName: Post-Deployment.ps1
# ScriptVersion: 2.3.0
# LastUpdated: 2026-09-18
# Changes:       v2.3.0 adds comprehensive Microsoft Copilot removal through
#                the canonical module on the central maintenance-script share.
#                The module does not depend on C:\Scripts already existing.
#                v2.1.0 adds idempotent PaperCut Print Deploy and Action1 Agent
#                installation sections. Each installer is verified after it runs,
#                reports reboot-required MSI results, and cannot stop later
#                post-deployment sections when its source is unavailable or fails.
#                v2.0.0 makes the workflow general purpose. It preserves all
#                scheduled tasks, safely refreshes C:\Scripts from the central
#                share, runs Register-Tasks_SYSTEM.ps1, conditionally activates
#                Office LTSC 2024, updates applications and Windows, and logs
#                section failures without stopping the remaining operations.
# =====================================================================

[CmdletBinding()]
param(
    # Public template values use placeholder server/share names. Replace these
    # paths with organization-approved deployment sources before production use.
    [Alias('HPDriversWorkingRoot')]
    [string]$DriverWorkingRoot = 'C:\Temp\DriverUpdates',
    [string]$LogFolder = 'C:\Logs',
    [switch]$IncludeSoftware = $false,
    [bool]$IncludeBiosAndFirmwareUpdates = $true,
    [string]$DellCommandUpdateSharePath = '\\SERVER\DeploymentShare\Dell-Command-Update-Windows-Universal-Application.exe',
    [string]$DotNetDesktopRuntimeSharePath = '\\SERVER\DeploymentShare\windowsdesktop-runtime-8-win-x64.exe',
    [string]$DotNetDesktopRuntimeInternetUrl = 'https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe',
    [string]$DellCommandUpdateWingetId = 'Dell.CommandUpdate.Universal',
    [int]$CleanupRetryCount = 12,
    [int]$CleanupRetryDelaySeconds = 10,
    [int]$RebootDelaySeconds = 15,
    [bool]$RebootWhenComplete = $true,
    [string]$ScriptsSourcePath = '\\SERVER\DeploymentShare\Scripts',
    [string]$LocalScriptsPath = 'C:\Scripts',
    [bool]$ActivateInstalledOffice = $true,
    [bool]$InstallPaperCut = $true,
    [string]$PaperCutMsiPath = '\\PRINT-SERVER\PrintDeploy\pc-print-deploy-client.msi',
    [int]$PaperCutInstallTimeoutSeconds = 900,
    [bool]$InstallAction1 = $true,
    [string]$Action1AgentPath = 'C:\Windows\Action1\action1_agent.exe',
    [string]$Action1InstallerPath = '\\SERVER\DeploymentShare\Action1-Agent.msi',
    [int]$Action1InstallTimeoutSeconds = 900,
    [bool]$RemoveMicrosoftCopilot = $true,
    [string]$CopilotModuleSharePath = '\\SERVER\DeploymentShare\Scripts\Maintenance.Copilot.psm1',
    [bool]$UpdateApplications = $true,
    [bool]$InstallWindowsUpdates = $true,
    [int]$WindowsUpdateTimeoutMinutes = 90
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:RunStart = Get-Date
$script:ComputerName = $env:COMPUTERNAME
$script:LogPath = $null
$script:OverallResult = 'Unknown'
$script:FailureMessage = $null
$script:DetectedSoftPaqs = New-Object System.Collections.Generic.List[object]
$script:InstalledSoftPaqResults = New-Object System.Collections.Generic.List[object]
$script:BitLockerSuspendedByScript = $false
$script:SectionResults = New-Object System.Collections.Generic.List[object]
$script:RebootRequired = $false

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

    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
    }
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Initialize-Log {
    Ensure-Folder -Path $LogFolder
    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $script:LogPath = Join-Path $LogFolder "$($script:ComputerName)-Post-Deployment-$timestamp.log"
    New-Item -Path $script:LogPath -ItemType File -Force | Out-Null
    Write-Log "Log file: $($script:LogPath)" 'INFO'
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

function Get-DriverVendor {
    try {
        $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
        Write-Log "Detected system manufacturer: $manufacturer" 'INFO'

        if ($manufacturer -match 'Dell') {
            return 'Dell'
        }
        if ($manufacturer -match 'HP|Hewlett-Packard') {
            return 'HP'
        }

        throw "Unsupported computer manufacturer: $manufacturer"
    }
    catch {
        throw "Unable to determine a supported HP or Dell driver vendor. $($_.Exception.Message)"
    }
}

function Save-PowerSettings {
    $settings = [ordered]@{}

    $settings.DisplayTimeoutDC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\DC\{3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}'
    ).SettingIndexValue / 60

    $settings.DisplayTimeoutAC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\AC\{3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}'
    ).SettingIndexValue / 60

    $settings.SleepTimeoutDC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\DC\{29f6c1db-86da-48c5-9fdb-f2b67b1f44da}'
    ).SettingIndexValue / 60

    $settings.SleepTimeoutAC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\AC\{29f6c1db-86da-48c5-9fdb-f2b67b1f44da}'
    ).SettingIndexValue / 60

    return [PSCustomObject]$settings
}

function Set-UnlimitedPowerTimeouts {
    Write-Log 'Temporarily disabling monitor and sleep timeouts...' 'INFO'
    powercfg -change -monitor-timeout-dc 0 | Out-Null
    powercfg -change -monitor-timeout-ac 0 | Out-Null
    powercfg -change -standby-timeout-dc 0 | Out-Null
    powercfg -change -standby-timeout-ac 0 | Out-Null
}

function Restore-PowerSettings {
    param($Saved)

    if ($null -eq $Saved) { return }

    Write-Log 'Restoring previous power timeout settings...' 'INFO'
    powercfg -change -monitor-timeout-dc $Saved.DisplayTimeoutDC | Out-Null
    powercfg -change -monitor-timeout-ac $Saved.DisplayTimeoutAC | Out-Null
    powercfg -change -standby-timeout-dc $Saved.SleepTimeoutDC | Out-Null
    powercfg -change -standby-timeout-ac $Saved.SleepTimeoutAC | Out-Null
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 12,
        [int]$RetryDelaySeconds = 10
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "Folder already absent: $Path" 'OK'
        return $true
    }

    Write-Log "Attempting to remove folder: $Path" 'INFO'

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Start-Sleep -Seconds 2
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $Path)) {
                Write-Log "Folder removed successfully: $Path" 'OK'
                return $true
            }
        }
        catch {
            Write-Log "Cleanup attempt $i/$RetryCount failed for $Path : $($_.Exception.Message)" 'WARN'
        }

        if ($i -lt $RetryCount) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    Write-Log "Folder still exists after cleanup attempts: $Path" 'ERROR'
    return $false
}

function Initialize-HPNetworkAccess {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log 'Enabled TLS 1.2 for PowerShell Gallery access.' 'OK'
    }
    catch {
        Write-Log "Could not explicitly set TLS 1.2: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-NuGetProvider {
    try {
        $provider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $provider) {
            Write-Log 'Installing NuGet package provider...' 'INFO'
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
            Write-Log 'NuGet package provider installed successfully.' 'OK'
        }
        else {
            Write-Log 'NuGet package provider already available.' 'INFO'
        }
    }
    catch {
        throw "Failed to ensure NuGet provider. $($_.Exception.Message)"
    }
}

function Ensure-PSGalleryTrusted {
    try {
        $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
        if (-not $repo) {
            Write-Log 'PSGallery repository is not registered. Attempting Register-PSRepository -Default...' 'WARN'
            Register-PSRepository -Default -ErrorAction Stop
            $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
        }

        if ($repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop
            Write-Log 'PSGallery repository set to Trusted.' 'OK'
        }
        else {
            Write-Log 'PSGallery repository already Trusted.' 'INFO'
        }
    }
    catch {
        Write-Log "Could not validate/set PSGallery trust: $($_.Exception.Message)" 'WARN'
    }
}

function Install-ModuleIfPossible {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$AllowClobber
    )

    $args = @{
        Name        = $Name
        Scope       = 'AllUsers'
        Force       = $true
        ErrorAction = 'Stop'
    }

    if ($AllowClobber) {
        $args['AllowClobber'] = $true
    }

    Install-Module @args | Out-Null
}

function Ensure-PackageTooling {
    Write-Log 'Ensuring PowerShell package tooling is current enough for HPCMSL...' 'INFO'

    Initialize-HPNetworkAccess
    Ensure-NuGetProvider
    Ensure-PSGalleryTrusted

    try {
        Install-ModuleIfPossible -Name 'PowerShellGet' -AllowClobber
        Write-Log 'PowerShellGet installed/updated.' 'OK'
    }
    catch {
        Write-Log "PowerShellGet update failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Install-ModuleIfPossible -Name 'Microsoft.PowerShell.PSResourceGet'
        Write-Log 'Microsoft.PowerShell.PSResourceGet installed/updated.' 'OK'
    }
    catch {
        Write-Log "PSResourceGet install/update failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Import-Module PowerShellGet -Force -ErrorAction Stop
    }
    catch {
        Write-Log "Could not import PowerShellGet: $($_.Exception.Message)" 'WARN'
    }

    try {
        Import-Module Microsoft.PowerShell.PSResourceGet -Force -ErrorAction Stop
    }
    catch {
        Write-Log "Could not import PSResourceGet yet: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-HPCMSL {
    Write-Log 'Ensuring HP CMSL is available...' 'INFO'

    Ensure-PackageTooling

    if (-not (Get-Module -ListAvailable -Name HPCMSL)) {
        try {
            if (Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue) {
                Write-Log 'Installing HPCMSL with Install-PSResource...' 'INFO'
                Install-PSResource -Name HPCMSL -Scope AllUsers -TrustRepository -Quiet -AcceptLicense -ErrorAction Stop | Out-Null
            }
            else {
                Write-Log 'Install-PSResource not available. Falling back to Install-Module for HPCMSL...' 'WARN'
                Install-ModuleIfPossible -Name 'HPCMSL' -AllowClobber
            }
        }
        catch {
            Write-Log "Primary HPCMSL install attempt failed: $($_.Exception.Message)" 'WARN'
            Write-Log 'Trying fallback HPCMSL install with Install-Module...' 'INFO'
            Install-ModuleIfPossible -Name 'HPCMSL' -AllowClobber
        }
    }
    else {
        Write-Log 'HPCMSL already present on system.' 'INFO'
    }

    try {
        Import-Module HPCMSL -Force -ErrorAction Stop
    }
    catch {
        try {
            Import-Module HP.Softpaq -Force -ErrorAction Stop
        }
        catch {
            throw "HPCMSL/HP.Softpaq could not be imported after installation. $($_.Exception.Message)"
        }
    }

    Write-Log 'HP CMSL imported successfully.' 'OK'
}

function Add-DetectedSoftPaq {
    param(
        [string]$Id,
        [string]$Name,
        [string]$Version,
        [string]$Category
    )

    $script:DetectedSoftPaqs.Add([PSCustomObject]@{
        Id       = $Id
        Name     = $Name
        Version  = $Version
        Category = $Category
    }) | Out-Null
}

function Add-InstalledSoftPaqResult {
    param(
        [string]$Id,
        [string]$Name,
        [string]$Version,
        [string]$Status,
        [string]$Message
    )

    $script:InstalledSoftPaqResults.Add([PSCustomObject]@{
        Id      = $Id
        Name    = $Name
        Version = $Version
        Status  = $Status
        Message = $Message
    }) | Out-Null
}

function Get-HPSoftpaqCategories {
    $categories = @('Driver')

    if ($IncludeBiosAndFirmwareUpdates) {
        $categories += @('BIOS', 'Firmware')
    }

    if ($IncludeSoftware) {
        $categories += @('Diagnostic', 'Dock', 'Software', 'Utility')
    }

    return $categories
}

function Get-DriverList {
    $categories = Get-HPSoftpaqCategories
    Write-Log "Querying HP SoftPaq list for categories: $($categories -join ', ')" 'INFO'

    $list = @(
        foreach ($categoryName in $categories) {
            try {
                @(Get-SoftpaqList -Category $categoryName -ErrorAction Stop)
            }
            catch {
                Write-Log "HP SoftPaq category '$categoryName' could not be queried: $($_.Exception.Message)" 'WARN'
            }
        }
    ) | Sort-Object -Property Id -Unique

    if (-not $list) {
        Write-Log 'No applicable HP SoftPaq updates were returned.' 'OK'
        return @()
    }

    foreach ($item in $list) {
        $category = $null
        if ($item.PSObject.Properties.Name -contains 'Category') {
            $category = [string]$item.Category
        }

        Add-DetectedSoftPaq -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Category $category
        Write-Log "Detected: [$($item.Id)] $($item.Name) Version $($item.Version) Category [$category]" 'INFO'
    }

    return @($list)
}

function Install-SoftpaqList {
    param([Parameter(Mandatory)][object[]]$Softpaqs)

    $failures = 0

    foreach ($item in $Softpaqs) {
        try {
            $category = $null
            if ($item.PSObject.Properties.Name -contains 'Category') {
                $category = [string]$item.Category
            }

            Write-Log "Installing SoftPaq [$($item.Id)] $($item.Name)..." 'INFO'

            Get-Softpaq -Number $item.Id -Action SilentInstall | Out-Null

            Write-Log "Installed SoftPaq [$($item.Id)] $($item.Name)." 'OK'
            Add-InstalledSoftPaqResult -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Status 'Succeeded' -Message 'Installed successfully'
        }
        catch {
            $msg = $_.Exception.Message
            Write-Log "Failed SoftPaq [$($item.Id)] $($item.Name): $msg" 'WARN'
            Add-InstalledSoftPaqResult -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Status 'Failed' -Message $msg
            $failures++
        }
    }

    return $failures
}

function Invoke-HPDriverUpdate {
    param([Parameter(Mandatory)][string]$WorkingRoot)

    Ensure-Folder -Path $WorkingRoot

    $savedPower = $null
    $originalLocation = (Get-Location).Path
    $failures = 0

    try {
        Write-Log 'Initializing HP driver update section...' 'INFO'
        Write-Log "Working root: $WorkingRoot" 'INFO'

        $savedPower = Save-PowerSettings
        Set-UnlimitedPowerTimeouts
        Ensure-HPCMSL

        Set-Location -Path $WorkingRoot

        $softpaqs = Get-DriverList
        if ($softpaqs.Count -eq 0) {
            Write-Log 'No HP SoftPaq updates were applicable.' 'OK'
            return 0
        }

        $failures = Install-SoftpaqList -Softpaqs $softpaqs
    }
    finally {
        try {
            Set-Location -Path $originalLocation
        }
        catch {
        }

        try {
            Restore-PowerSettings -Saved $savedPower
        }
        catch {
            Write-Log "Failed restoring power settings: $($_.Exception.Message)" 'WARN'
        }
    }

    if ($failures -gt 0) {
        throw "$failures HP SoftPaq installation(s) failed."
    }

    Write-Log 'HP driver update section completed successfully.' 'OK'
    return 0
}

function Suspend-BitLockerForVendorFirmware {
    [CmdletBinding()]
    param()

    if (-not $IncludeBiosAndFirmwareUpdates) {
        return
    }

    $bitLockerCommand = Get-Command -Name Get-BitLockerVolume -ErrorAction SilentlyContinue
    if (-not $bitLockerCommand) {
        Write-Log 'BitLocker cmdlets are unavailable; no BitLocker suspension was required or performed.' 'WARN'
        return
    }

    $volume = Get-BitLockerVolume -MountPoint 'C:' -ErrorAction Stop
    if ($volume.VolumeStatus -eq 'FullyDecrypted') {
        Write-Log 'BitLocker is not enabled on C:. No suspension is needed for BIOS/firmware updates.' 'INFO'
        return
    }

    if ([string]$volume.ProtectionStatus -match 'Off|0') {
        Write-Log 'BitLocker protection on C: is already suspended.' 'INFO'
        return
    }

    Write-Log 'Suspending BitLocker protection on C: for one reboot before BIOS/firmware updates...' 'INFO'
    Suspend-BitLocker -MountPoint 'C:' -RebootCount 1 -ErrorAction Stop | Out-Null
    $verification = Get-BitLockerVolume -MountPoint 'C:' -ErrorAction Stop
    if ([string]$verification.ProtectionStatus -notmatch 'Off|0') {
        throw 'BitLocker protection remained enabled after the suspension request. BIOS/firmware updates were blocked.'
    }

    $script:BitLockerSuspendedByScript = $true
    Write-Log 'BitLocker protection is suspended for one reboot.' 'OK'
}

function Test-DotNetDesktopRuntime8 {
    [CmdletBinding()]
    param([version]$MinimumVersion = [version]'8.0.25')

    $dotnetExe = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
    if (-not (Test-Path -LiteralPath $dotnetExe -PathType Leaf)) {
        return $false
    }

    try {
        foreach ($line in @(& $dotnetExe --list-runtimes 2>$null)) {
            if ($line -match '^Microsoft\.WindowsDesktop\.App\s+(?<Version>\d+\.\d+\.\d+)\s+') {
                try {
                    if ([version]$Matches.Version -ge $MinimumVersion) {
                        return $true
                    }
                }
                catch {}
            }
        }
    }
    catch {
        Write-Log "Unable to query installed .NET desktop runtimes: $($_.Exception.Message)" 'WARN'
    }

    return $false
}

function Invoke-InstallerAndValidateExitCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [Parameter(Mandatory)][string]$DisplayName,
        [int[]]$SuccessExitCodes = @(0, 3010, 1641)
    )

    $process = Start-Process `
        -FilePath $FilePath `
        -ArgumentList $ArgumentList `
        -Wait `
        -PassThru `
        -NoNewWindow `
        -ErrorAction Stop

    Write-Log "$DisplayName installer exit code: $($process.ExitCode)" 'INFO'
    if ([int]$process.ExitCode -notin $SuccessExitCodes) {
        throw "$DisplayName installation failed with exit code $($process.ExitCode)."
    }
}

function Get-PaperCutProduct {
    $uninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $product = Get-ItemProperty -Path $uninstallRoots -ErrorAction SilentlyContinue |
        Where-Object {
            $displayName = $_.PSObject.Properties['DisplayName']
            $null -ne $displayName -and
            -not [string]::IsNullOrWhiteSpace([string]$displayName.Value) -and
            [string]$displayName.Value -match 'PaperCut.*Print Deploy|Print Deploy Client'
        } |
        Select-Object -First 1

    if ($null -eq $product) {
        return [pscustomobject]@{
            Installed      = $false
            DisplayName    = $null
            DisplayVersion = $null
        }
    }

    $versionProperty = $product.PSObject.Properties['DisplayVersion']
    return [pscustomobject]@{
        Installed      = $true
        DisplayName    = [string]$product.DisplayName
        DisplayVersion = if ($null -ne $versionProperty) { [string]$versionProperty.Value } else { $null }
    }
}

function Install-PaperCutPrintDeploy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [ValidateRange(60, 7200)][int]$TimeoutSeconds = 900
    )

    $installed = Get-PaperCutProduct
    if ($installed.Installed) {
        Write-Log "PaperCut is already installed: $($installed.DisplayName) $($installed.DisplayVersion)" 'OK'
        return
    }

    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) {
        throw "PaperCut is not installed and its MSI is unavailable: $InstallerPath"
    }

    $msiLogPath = Join-Path $LogFolder ("{0}-PaperCut-Install-{1}.log" -f $script:ComputerName, (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
    $arguments = "/i `"$InstallerPath`" /qn /norestart REBOOT=ReallySuppress /L*v `"$msiLogPath`""
    Write-Log "Installing PaperCut Print Deploy from $InstallerPath" 'INFO'

    $process = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\msiexec.exe') `
        -ArgumentList $arguments -PassThru -NoNewWindow -ErrorAction Stop
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        try { $process.Kill() } catch { }
        throw "PaperCut installation exceeded the $TimeoutSeconds-second timeout. MSI log: $msiLogPath"
    }
    $process.Refresh()
    $exitCode = [int]$process.ExitCode
    Write-Log "PaperCut installer exit code: $exitCode. MSI log: $msiLogPath" 'INFO'

    if ($exitCode -in @(3010, 1641)) {
        $script:RebootRequired = $true
        Write-Log 'PaperCut installation completed and requires a reboot.' 'WARN'
    }
    elseif ($exitCode -ne 0) {
        throw "PaperCut installation failed with exit code $exitCode. MSI log: $msiLogPath"
    }

    $deadline = (Get-Date).AddSeconds(60)
    do {
        Start-Sleep -Seconds 3
        $installed = Get-PaperCutProduct
        if ($installed.Installed) { break }
    } while ((Get-Date) -lt $deadline)

    if (-not $installed.Installed) {
        throw "PaperCut MSI returned success, but the product was not found in the uninstall registry. MSI log: $msiLogPath"
    }

    Write-Log "PaperCut installed successfully: $($installed.DisplayName) $($installed.DisplayVersion)" 'OK'
}

function Install-Action1Agent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AgentPath,
        [Parameter(Mandatory)][string]$InstallerPath,
        [ValidateRange(60, 7200)][int]$TimeoutSeconds = 900
    )

    if (Test-Path -LiteralPath $AgentPath -PathType Leaf) {
        Write-Log "Action1 Agent is already installed at $AgentPath" 'OK'
        return
    }

    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) {
        throw "Action1 Agent is not installed and its MSI is unavailable: $InstallerPath"
    }

    $msiLogPath = Join-Path $LogFolder ("{0}-Action1-Install-{1}.log" -f $script:ComputerName, (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
    $arguments = "/i `"$InstallerPath`" /qn /norestart REBOOT=ReallySuppress /L*v `"$msiLogPath`""
    Write-Log "Installing Action1 Agent from $InstallerPath" 'INFO'

    $process = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\msiexec.exe') `
        -ArgumentList $arguments -PassThru -NoNewWindow -ErrorAction Stop
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        try { $process.Kill() } catch { }
        throw "Action1 Agent installation exceeded the $TimeoutSeconds-second timeout. MSI log: $msiLogPath"
    }
    $process.Refresh()
    $exitCode = [int]$process.ExitCode
    Write-Log "Action1 Agent installer exit code: $exitCode. MSI log: $msiLogPath" 'INFO'

    if ($exitCode -in @(3010, 1641)) {
        $script:RebootRequired = $true
        Write-Log 'Action1 Agent installation completed and requires a reboot.' 'WARN'
    }
    elseif ($exitCode -ne 0) {
        throw "Action1 Agent installation failed with exit code $exitCode. MSI log: $msiLogPath"
    }

    $deadline = (Get-Date).AddSeconds(60)
    do {
        Start-Sleep -Seconds 3
        if (Test-Path -LiteralPath $AgentPath -PathType Leaf) { break }
    } while ((Get-Date) -lt $deadline)

    if (-not (Test-Path -LiteralPath $AgentPath -PathType Leaf)) {
        throw "Action1 Agent MSI returned success, but $AgentPath was not detected. MSI log: $msiLogPath"
    }

    Write-Log "Action1 Agent installed successfully at $AgentPath" 'OK'
}

function Save-InternetFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][uri]$Uri,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][string]$DisplayName
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Test-Path -LiteralPath $DestinationPath) {
        Remove-Item -LiteralPath $DestinationPath -Force -ErrorAction SilentlyContinue
    }

    try {
        Start-BitsTransfer -Source $Uri.AbsoluteUri -Destination $DestinationPath -ErrorAction Stop
    }
    catch {
        Invoke-WebRequest `
            -Uri $Uri.AbsoluteUri `
            -OutFile $DestinationPath `
            -UseBasicParsing `
            -TimeoutSec 300 `
            -ErrorAction Stop
    }

    if (
        -not (Test-Path -LiteralPath $DestinationPath -PathType Leaf) -or
        (Get-Item -LiteralPath $DestinationPath).Length -lt 1MB
    ) {
        throw "The downloaded $DisplayName installer is missing or unexpectedly small."
    }
}

function Ensure-DotNetDesktopRuntimeForDcu {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkingRoot)

    if (Test-DotNetDesktopRuntime8 -MinimumVersion ([version]'8.0.25')) {
        Write-Log '.NET Desktop Runtime 8.0.25 or newer is already installed.' 'OK'
        return
    }

    $localInstaller = Join-Path $WorkingRoot 'windowsdesktop-runtime-8-win-x64.exe'
    if (Test-Path -LiteralPath $DotNetDesktopRuntimeSharePath -PathType Leaf) {
        Write-Log 'Copying .NET Desktop Runtime from the Labscripts share...' 'INFO'
        Copy-Item -LiteralPath $DotNetDesktopRuntimeSharePath -Destination $localInstaller -Force -ErrorAction Stop
    }
    else {
        Write-Log 'The .NET share installer is unavailable; using the Microsoft Internet fallback.' 'WARN'
        Save-InternetFile `
            -Uri ([uri]$DotNetDesktopRuntimeInternetUrl) `
            -DestinationPath $localInstaller `
            -DisplayName '.NET Desktop Runtime 8 x64'
    }

    Invoke-InstallerAndValidateExitCode `
        -FilePath $localInstaller `
        -ArgumentList @('/install','/quiet','/norestart') `
        -DisplayName '.NET Desktop Runtime 8 x64'

    if (-not (Test-DotNetDesktopRuntime8 -MinimumVersion ([version]'8.0.25'))) {
        throw '.NET Desktop Runtime installation completed, but version 8.0.25 or newer was not detected.'
    }
}

function Get-DellCommandUpdateCliPath {
    [CmdletBinding()]
    param()

    $programFilesPaths = @(
        [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles),
        [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($candidate in @(
        $programFilesPaths | ForEach-Object {
            Join-Path $_ 'Dell\CommandUpdate\dcu-cli.exe'
        }
    )) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }

    return $null
}

function Ensure-DellCommandUpdateInstalled {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkingRoot)

    $dcuCliPath = Get-DellCommandUpdateCliPath
    if (-not [string]::IsNullOrWhiteSpace([string]$dcuCliPath)) {
        Write-Log "Dell Command | Update is already installed: $dcuCliPath" 'OK'
        return $dcuCliPath
    }

    Ensure-DotNetDesktopRuntimeForDcu -WorkingRoot $WorkingRoot

    if (Test-Path -LiteralPath $DellCommandUpdateSharePath -PathType Leaf) {
        Write-Log 'Installing Dell Command | Update from the Labscripts share...' 'INFO'
        $localDcuInstaller = Join-Path $WorkingRoot 'Dell-Command-Update-Windows-Universal-Application.exe'
        Copy-Item -LiteralPath $DellCommandUpdateSharePath -Destination $localDcuInstaller -Force -ErrorAction Stop
        Invoke-InstallerAndValidateExitCode `
            -FilePath $localDcuInstaller `
            -ArgumentList @('/s') `
            -DisplayName 'Dell Command | Update'
    }
    else {
        Write-Log 'The Dell Command Update share installer is unavailable; using winget as the fallback.' 'WARN'
        $winget = Get-Command -Name 'winget.exe' -ErrorAction SilentlyContinue
        if (-not $winget) {
            throw 'Dell Command | Update is not on the file share and winget.exe is unavailable.'
        }

        $wingetArguments = @(
            'install','--id',$DellCommandUpdateWingetId,'--exact','--silent',
            '--accept-package-agreements','--accept-source-agreements','--disable-interactivity'
        )
        $wingetProcess = Start-Process `
            -FilePath $winget.Source `
            -ArgumentList $wingetArguments `
            -Wait `
            -PassThru `
            -NoNewWindow `
            -ErrorAction Stop

        if ([int]$wingetProcess.ExitCode -ne 0) {
            throw "winget failed to install Dell Command | Update. Exit code: $($wingetProcess.ExitCode)"
        }
    }

    Start-Sleep -Seconds 3
    $dcuCliPath = Get-DellCommandUpdateCliPath
    if ([string]::IsNullOrWhiteSpace([string]$dcuCliPath)) {
        throw 'Dell Command | Update installed, but dcu-cli.exe was not found afterward.'
    }

    return $dcuCliPath
}

function Get-DellDCUService {
    [CmdletBinding()]
    param()

    foreach ($name in @(
        'DellClientManagementService',
        'DellCommandUpdate',
        'DellUpdateService'
    )) {
        $service = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($service) {
            return $service
        }
    }

    return Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'Dell.*Client.*Management|Dell.*Command.*Update' } |
        Select-Object -First 1
}

function Ensure-DellDCUService {
    [CmdletBinding()]
    param([int]$TimeoutSeconds = 30)

    Write-Log 'Validating Dell Client Management Service...' 'INFO'
    $service = Get-DellDCUService
    if (-not $service) {
        throw 'Dell Client Management Service was not found. Dell Command | Update may need to be repaired or reinstalled.'
    }

    try {
        $serviceConfiguration = Get-CimInstance `
            -ClassName Win32_Service `
            -Filter ("Name='{0}'" -f $service.Name) `
            -ErrorAction Stop
        if ($serviceConfiguration.StartMode -eq 'Disabled') {
            Write-Log 'Dell service startup type is Disabled. Setting it to Manual.' 'WARN'
            Set-Service -Name $service.Name -StartupType Manual -ErrorAction Stop
        }
    }
    catch {
        Write-Log "Could not validate or set the Dell service startup type: $($_.Exception.Message)" 'WARN'
    }

    $service.Refresh()
    if ($service.Status -ne 'Running') {
        Write-Log 'Starting Dell Client Management Service...' 'INFO'
        try {
            Start-Service -Name $service.Name -ErrorAction Stop
        }
        catch {
            Write-Log "Start-Service failed; attempting sc.exe start. $($_.Exception.Message)" 'WARN'
            & sc.exe start $service.Name | Out-Null
        }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Seconds 2
        $service = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq 'Running') {
            Write-Log 'Dell Client Management Service is running.' 'OK'
            return
        }
    } while ((Get-Date) -lt $deadline)

    throw 'Dell Client Management Service did not reach the Running state before timeout.'
}

function Invoke-DellDCUCommandWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DcuCli,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$OperationName,
        [int[]]$AcceptableExitCodes = @(0),
        [int]$MaxAttempts = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log "Dell DCU $OperationName attempt $attempt of $MaxAttempts..." 'INFO'
        $process = Start-Process `
            -FilePath $DcuCli `
            -ArgumentList $Arguments `
            -Wait `
            -PassThru `
            -NoNewWindow `
            -ErrorAction Stop

        $exitCode = [int]$process.ExitCode
        Write-Log "Dell DCU $OperationName exit code: $exitCode" 'INFO'
        if ($AcceptableExitCodes -contains $exitCode) {
            return $process
        }

        if ($exitCode -eq 3000) {
            Write-Log 'Dell DCU returned 3000; recovering the Dell Client Management Service before retrying.' 'WARN'
        }
        else {
            Write-Log "Dell DCU $OperationName returned non-success exit code $exitCode." 'WARN'
        }

        if ($attempt -lt $MaxAttempts) {
            Ensure-DellDCUService
            Start-Sleep -Seconds 5
            continue
        }

        throw "Dell DCU $OperationName failed after $MaxAttempts attempt(s). Last exit code: $exitCode"
    }
}

function Invoke-DellDriverUpdate {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkingRoot)

    Ensure-Folder -Path $WorkingRoot
    $savedPower = $null

    try {
        Write-Log 'Initializing Dell vendor update section...' 'INFO'
        $dellUpdateTypes = if ($IncludeBiosAndFirmwareUpdates) {
            'driver,bios,firmware'
        }
        else {
            'driver'
        }
        Write-Log "Dell update types selected: $dellUpdateTypes" 'INFO'
        $savedPower = Save-PowerSettings
        Set-UnlimitedPowerTimeouts

        $dcuCli = Ensure-DellCommandUpdateInstalled -WorkingRoot $WorkingRoot
        Write-Log "Using Dell Command | Update CLI: $dcuCli" 'OK'
        Ensure-DellDCUService

        $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $dcuScanLog = Join-Path $LogFolder "$($script:ComputerName)-Dell-DCU-Scan-$timestamp.log"
        $dcuApplyLog = Join-Path $LogFolder "$($script:ComputerName)-Dell-DCU-Apply-$timestamp.log"

        $configureArguments = '/configure -silent -scheduleAuto -lockSettings=disable'
        [void](Invoke-DellDCUCommandWithRetry `
            -DcuCli $dcuCli `
            -Arguments $configureArguments `
            -OperationName 'Configure' `
            -MaxAttempts 2)

        $scanArguments = "/scan -silent -updateType=$dellUpdateTypes -outputLog=`"$dcuScanLog`""
        [void](Invoke-DellDCUCommandWithRetry `
            -DcuCli $dcuCli `
            -Arguments $scanArguments `
            -OperationName 'Scan' `
            -MaxAttempts 2)

        $applyArguments = "/applyUpdates -silent -updateType=$dellUpdateTypes -reboot=disable -outputLog=`"$dcuApplyLog`""
        [void](Invoke-DellDCUCommandWithRetry `
            -DcuCli $dcuCli `
            -Arguments $applyArguments `
            -OperationName 'ApplyUpdates' `
            -MaxAttempts 2)

        Write-Log "Dell vendor updates completed. Logs: $dcuScanLog ; $dcuApplyLog" 'OK'
    }
    finally {
        try {
            Restore-PowerSettings -Saved $savedPower
        }
        catch {
            Write-Log "Failed restoring power settings after Dell updates: $($_.Exception.Message)" 'WARN'
        }
    }
}

function Disable-MicrosoftCopilot {
    [CmdletBinding()]
    param()

    if (-not (Test-Path -LiteralPath $CopilotModuleSharePath -PathType Leaf)) {
        throw "The shared Copilot-removal module was not found: $CopilotModuleSharePath"
    }
    Import-Module -Name $CopilotModuleSharePath -Force -ErrorAction Stop

    $copilotResult = Invoke-ComprehensiveCopilotRemoval `
        -Logger { param($Message,$Level) Write-Log $Message $Level } `
        -SkipExplorerRestart

    if ($copilotResult.RebootRequired) {
        $script:RebootRequired = $true
    }
    if ($copilotResult.Failures -gt 0) {
        throw "Copilot removal completed with status '$($copilotResult.Status)' and $($copilotResult.Failures) failure(s)."
    }
}

function Invoke-PostDeploymentSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Operation
    )

    $started = Get-Date
    Write-Log "===== Starting section: $Name =====" 'INFO'
    try {
        & $Operation
        $elapsed = [math]::Round(((Get-Date) - $started).TotalMinutes, 2)
        [void]$script:SectionResults.Add([pscustomobject]@{
            Name = $Name; Result = 'Succeeded'; Minutes = $elapsed; Error = $null
        })
        Write-Log "===== Section succeeded: $Name ($elapsed minute(s)) =====" 'OK'
        return $true
    }
    catch {
        $elapsed = [math]::Round(((Get-Date) - $started).TotalMinutes, 2)
        $message = $_.Exception.Message
        [void]$script:SectionResults.Add([pscustomobject]@{
            Name = $Name; Result = 'Failed'; Minutes = $elapsed; Error = $message
        })
        Write-Log "===== Section failed: $Name ($elapsed minute(s)): $message =====" 'ERROR'
        return $false
    }
}

function Invoke-RobocopyChecked {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )

    $robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
    $arguments = @(
        $Source,
        $Destination,
        '*.ps1', '*.psm1', '*.json', '*.yml', '*.yaml',
        '/E', '/R:2', '/W:5', '/COPY:DAT', '/DCOPY:DAT', '/XJ',
        '/XD', 'Installers', 'Retired', '.git',
        '/NP', '/NFL', '/NDL'
    )

    $output = & $robocopy @arguments 2>&1
    $exitCode = [int]$LASTEXITCODE
    foreach ($line in @($output)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            Write-Log "ROBOCOPY: $line" 'INFO'
        }
    }

    # Robocopy codes 0 through 7 are successful or informational.
    if ($exitCode -ge 8) {
        throw "Robocopy failed with exit code $exitCode while copying '$Source' to '$Destination'."
    }
}

function Update-LocalMaintenanceScripts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
        throw "The maintenance-script share is unavailable: $SourcePath"
    }

    $stagingRoot = Join-Path $env:SystemDrive 'Temp\Compton-PostDeployment'
    $stagingPath = Join-Path $stagingRoot ("Scripts-{0}" -f [guid]::NewGuid().ToString('N'))
    Ensure-Folder -Path $stagingPath
    $originalLocation = (Get-Location).Path
    $originalLocationNormalized = $originalLocation.TrimEnd('\')
    $destinationNormalized = $DestinationPath.TrimEnd('\')
    $locationWasInsideDestination = (
        $originalLocationNormalized.Equals($destinationNormalized, [StringComparison]::OrdinalIgnoreCase) -or
        $originalLocationNormalized.StartsWith("$destinationNormalized\", [StringComparison]::OrdinalIgnoreCase)
    )
    $localCopyCompleted = $false

    try {
        Write-Log "Staging current maintenance files from $SourcePath..." 'INFO'
        Invoke-RobocopyChecked -Source $SourcePath -Destination $stagingPath

        $stagedRegisterScript = Join-Path $stagingPath 'Register-Tasks_SYSTEM.ps1'
        if (-not (Test-Path -LiteralPath $stagedRegisterScript -PathType Leaf)) {
            throw "The staged content does not contain Register-Tasks_SYSTEM.ps1. Existing local scripts were not changed."
        }

        # Destructive replacement occurs only after the share was reached and
        # a complete staging copy containing the registration script exists.
        if ($locationWasInsideDestination) {
            Set-Location -LiteralPath $stagingRoot
        }
        if (Test-Path -LiteralPath $DestinationPath) {
            Write-Log "Removing the existing local scripts folder: $DestinationPath" 'INFO'
            Remove-Item -LiteralPath $DestinationPath -Recurse -Force -ErrorAction Stop
        }
        Ensure-Folder -Path $DestinationPath
        Invoke-RobocopyChecked -Source $stagingPath -Destination $DestinationPath
        $localCopyCompleted = $true

        $registerScript = Join-Path $DestinationPath 'Register-Tasks_SYSTEM.ps1'
        Write-Log "Running scheduled-task registration: $registerScript" 'INFO'
        $powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $registrationOutput = & $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $registerScript 2>&1
        $registrationExitCode = [int]$LASTEXITCODE
        foreach ($line in @($registrationOutput)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
                Write-Log "REGISTER-TASKS: $line" 'INFO'
            }
        }
        if ($registrationExitCode -ne 0) {
            throw "Register-Tasks_SYSTEM.ps1 returned exit code $registrationExitCode."
        }
        Write-Log 'Local maintenance scripts were refreshed and scheduled tasks were registered.' 'OK'
    }
    finally {
        if ($locationWasInsideDestination -and (Test-Path -LiteralPath $DestinationPath)) {
            Set-Location -LiteralPath $DestinationPath -ErrorAction SilentlyContinue
        }
        elseif (Test-Path -LiteralPath $originalLocation) {
            Set-Location -LiteralPath $originalLocation -ErrorAction SilentlyContinue
        }

        if ($localCopyCompleted -and (Test-Path -LiteralPath $stagingPath)) {
            Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        elseif (Test-Path -LiteralPath $stagingPath) {
            Write-Log "The staged script copy was retained for recovery: $stagingPath" 'WARN'
        }
    }
}

function Test-Office2024Installed {
    [CmdletBinding()]
    param()

    foreach ($path in @(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
    )) {
        $configuration = Get-ItemProperty -LiteralPath $path -ErrorAction SilentlyContinue
        if ($null -ne $configuration) {
            $releaseIds = '{0} {1}' -f ([string]$configuration.ProductReleaseIds),([string]$configuration.ProductReleaseIdsToAdd)
            if ($releaseIds -match '(?i)2024') {
                return $true
            }
        }
    }

    foreach ($root in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )) {
        $matches = @(Get-ItemProperty -Path $root -ErrorAction SilentlyContinue | Where-Object {
            $name = [string]$_.DisplayName
            $name -match '(?i)(?:Microsoft Office|Office LTSC).*2024' -and
            $name -notmatch '(?i)(?:Visio|Project|Language Pack|Proofing)'
        })
        if ($matches.Count -gt 0) {
            return $true
        }
    }

    return $false
}

function Get-OfficeOsppPath {
    [CmdletBinding()]
    param()

    $programFilesPaths = @(
        [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles),
        [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($programFilesPath in $programFilesPaths) {
        foreach ($relativePath in @(
            'Microsoft Office\root\Office16\OSPP.VBS',
            'Microsoft Office\Office16\OSPP.VBS'
        )) {
            $candidate = Join-Path $programFilesPath $relativePath
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return $candidate
            }
        }
    }

    return $null
}

function Invoke-OfficeOspp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OsppPath,
        [Parameter(Mandatory)][ValidateSet('/dstatusall','/act')][string]$Argument
    )

    $cscriptPath = Join-Path $env:SystemRoot 'System32\cscript.exe'
    $output = & $cscriptPath '//Nologo' $OsppPath $Argument 2>&1 | Out-String

    return [pscustomobject]@{
        ExitCode = [int]$LASTEXITCODE
        Output   = [string]$output
    }
}

function Get-OfficeSuiteLicenseState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$StatusOutput)

    $licenseBlocks = @([regex]::Split($StatusOutput, '(?m)^-{20,}\s*$'))
    $officeSuiteBlocks = @(
        $licenseBlocks | Where-Object {
            $_ -match '(?im)^\s*LICENSE (?:NAME|DESCRIPTION):.*Office' -and
            $_ -match '(?im)^\s*LICENSE (?:NAME|DESCRIPTION):.*(?:Office\s*24|2024)' -and
            $_ -notmatch '(?im)^\s*LICENSE (?:NAME|DESCRIPTION):.*(?:Project|Visio)'
        }
    )
    $licensedBlocks = @(
        $officeSuiteBlocks | Where-Object {
            $_ -match '(?im)^\s*LICENSE STATUS:\s*---LICENSED---\s*$'
        }
    )

    return [pscustomobject]@{
        LicenseFound = ($officeSuiteBlocks.Count -gt 0)
        IsActivated  = ($licensedBlocks.Count -gt 0)
        LicenseCount = $officeSuiteBlocks.Count
    }
}

function Enable-InstalledOfficeActivation {
    [CmdletBinding()]
    param()

    $osppPath = Get-OfficeOsppPath
    if ([string]::IsNullOrWhiteSpace([string]$osppPath)) {
        throw 'Office activation could not run because OSPP.VBS was not found. Confirm that volume-licensed Office is installed.'
    }

    Write-Log "Checking machine-wide Office activation with: $osppPath" 'INFO'
    $initialStatus = Invoke-OfficeOspp -OsppPath $osppPath -Argument '/dstatusall'
    if ($initialStatus.ExitCode -ne 0) {
        throw "Office activation status check failed with exit code $($initialStatus.ExitCode)."
    }

    $initialState = Get-OfficeSuiteLicenseState -StatusOutput $initialStatus.Output
    if (-not $initialState.LicenseFound) {
        throw 'OSPP.VBS did not return an installed Office suite volume-license record.'
    }

    if ($initialState.IsActivated) {
        Write-Log 'The installed Office suite is already activated for this computer and all users.' 'OK'
        return
    }

    Write-Log 'Office is not activated. Starting machine-wide Office activation...' 'INFO'
    $activationResult = Invoke-OfficeOspp -OsppPath $osppPath -Argument '/act'
    if ($activationResult.ExitCode -ne 0) {
        Write-Log "OSPP.VBS /act returned exit code $($activationResult.ExitCode); verifying final license state before failing." 'WARN'
    }

    Start-Sleep -Seconds 5
    $verificationStatus = Invoke-OfficeOspp -OsppPath $osppPath -Argument '/dstatusall'
    if ($verificationStatus.ExitCode -ne 0) {
        throw "Office activation verification failed with exit code $($verificationStatus.ExitCode)."
    }

    $verifiedState = Get-OfficeSuiteLicenseState -StatusOutput $verificationStatus.Output
    if (-not $verifiedState.IsActivated) {
        $errorCodeMatch = [regex]::Match(
            [string]$activationResult.Output,
            '(?i)(?:ERROR CODE|error):?\s*(0x[0-9A-F]{8})'
        )
        $errorSuffix = if ($errorCodeMatch.Success) {
            " Microsoft error code: $($errorCodeMatch.Groups[1].Value)."
        }
        else {
            " OSPP.VBS exit code: $($activationResult.ExitCode)."
        }
        throw "Office remained unactivated after the activation attempt.$errorSuffix"
    }

    Write-Log 'Office activation completed and the machine-wide LICENSED state was verified for all users.' 'OK'
}

function Get-WingetPath {
    [CmdletBinding()]
    param()

    $command = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    $candidate = Get-ChildItem -Path 'C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe' -ErrorAction SilentlyContinue |
        Sort-Object VersionInfo -Descending |
        Select-Object -First 1
    if ($null -ne $candidate) {
        return $candidate.FullName
    }

    return $null
}

function Invoke-WingetLogged {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WingetPath,
        [Parameter(Mandatory)][string[]]$Arguments,
        [Parameter(Mandatory)][string]$Operation
    )

    Write-Log "Running winget operation: $Operation" 'INFO'
    $output = & $WingetPath @Arguments 2>&1
    $exitCode = [int]$LASTEXITCODE
    foreach ($line in @($output)) {
        $text = ([string]$line).Trim()
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            Write-Log "WINGET: $text" 'INFO'
        }
    }
    if ($exitCode -ne 0) {
        throw "Winget operation '$Operation' returned exit code $exitCode."
    }
}

function Update-InstalledApplications {
    [CmdletBinding()]
    param()

    $wingetPath = Get-WingetPath
    if ([string]::IsNullOrWhiteSpace([string]$wingetPath)) {
        throw 'Windows Package Manager (winget.exe) was not found.'
    }

    try {
        Invoke-WingetLogged -WingetPath $wingetPath -Operation 'source update' -Arguments @(
            'source', 'update', '--disable-interactivity'
        )
    }
    catch {
        Write-Log "Winget source refresh failed; the application upgrade will still be attempted. $($_.Exception.Message)" 'WARN'
    }
    Invoke-WingetLogged -WingetPath $wingetPath -Operation 'upgrade all applications' -Arguments @(
        'upgrade', '--all', '--silent', '--include-unknown',
        '--accept-package-agreements', '--accept-source-agreements',
        '--disable-interactivity'
    )

    $officeUpdater = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    if (Test-Path -LiteralPath $officeUpdater -PathType Leaf) {
        Write-Log 'Starting the Microsoft Office Click-to-Run update...' 'INFO'
        $officeProcess = Start-Process -FilePath $officeUpdater -ArgumentList @(
            '/update', 'USER', 'displaylevel=False', 'forceappshutdown=True'
        ) -PassThru -WindowStyle Hidden -ErrorAction Stop
        if (-not $officeProcess.WaitForExit(30 * 60 * 1000)) {
            try { $officeProcess.Kill() } catch { }
            throw 'The Microsoft Office update did not finish within 30 minutes.'
        }
        if ($officeProcess.ExitCode -ne 0) {
            throw "The Microsoft Office Click-to-Run updater returned exit code $($officeProcess.ExitCode)."
        }
        Write-Log 'Microsoft Office Click-to-Run update completed.' 'OK'
    }
    else {
        Write-Log 'Microsoft Office Click-to-Run is not installed; its update was skipped.' 'INFO'
    }
    Write-Log 'Application update scan and installation completed.' 'OK'
}

function Ensure-WindowsUpdateModule {
    [CmdletBinding()]
    param()

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    foreach ($serviceName in @('bits','cryptsvc','wuauserv')) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($null -eq $service) {
            Write-Log "Windows Update prerequisite service was not found: $serviceName" 'WARN'
            continue
        }
        try {
            if ($serviceName -eq 'wuauserv') {
                Set-Service -Name $serviceName -StartupType Manual -ErrorAction Stop
            }
            elseif ($service.StartType -eq 'Disabled') {
                Set-Service -Name $serviceName -StartupType Manual -ErrorAction Stop
            }
            Start-Service -Name $serviceName -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log "Could not prepare service $serviceName : $($_.Exception.Message)" 'WARN'
        }
    }

    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Write-Log 'Installing the NuGet package provider...' 'INFO'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
    }

    $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($null -ne $gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
    }

    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log 'Installing the PSWindowsUpdate module...' 'INFO'
        Install-Module -Name PSWindowsUpdate -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop
}

function Install-AllWindowsUpdates {
    [CmdletBinding()]
    param([int]$TimeoutMinutes = 90)

    if ($TimeoutMinutes -lt 5) {
        $TimeoutMinutes = 5
    }
    Ensure-WindowsUpdateModule

    Write-Log 'Scanning Microsoft Update and installing all applicable updates...' 'INFO'
    $job = Start-Job -ScriptBlock {
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Verbose -ErrorAction Stop *>&1
    }
    try {
        $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
        $nextHeartbeat = (Get-Date).AddMinutes(1)
        while ($job.State -in @('NotStarted','Running')) {
            if ((Get-Date) -ge $deadline) {
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                throw "Windows Update exceeded the $TimeoutMinutes minute timeout."
            }
            if ((Get-Date) -ge $nextHeartbeat) {
                Write-Log 'Windows Update is still running...' 'INFO'
                $nextHeartbeat = (Get-Date).AddMinutes(1)
            }
            Start-Sleep -Seconds 5
            $job = Get-Job -Id $job.Id
        }

        $output = @(Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue)
        foreach ($item in $output) {
            $text = ($item | Out-String).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                Write-Log "WINDOWS-UPDATE: $text" 'INFO'
            }
        }
        if ($job.State -ne 'Completed') {
            $reason = [string]$job.ChildJobs[0].JobStateInfo.Reason
            throw "Windows Update job ended in state '$($job.State)'. $reason"
        }
    }
    finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }

    try {
        $systemInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        $script:RebootRequired = [bool]$systemInfo.RebootRequired
    }
    catch {
        Write-Log "Could not query the Windows Update reboot status: $($_.Exception.Message)" 'WARN'
    }
    Write-Log "Windows Update completed. RebootRequired=$($script:RebootRequired)" 'OK'
}

function Invoke-FinalReboot {
    param([int]$DelaySeconds = 15)

    if ($DelaySeconds -lt 0) {
        $DelaySeconds = 0
    }

    Write-Log "Rebooting system in $DelaySeconds second(s)..." 'WARN'
    Start-Sleep -Seconds $DelaySeconds
    Restart-Computer -Force
}

try {
    Initialize-Log
    Write-Log 'Starting Post-Deployment.ps1 version 2.3.0.' 'INFO'

    if (-not (Test-IsAdministrator)) {
        throw 'Please run this script as Administrator.'
    }

    [void](Invoke-PostDeploymentSection -Name 'HP or Dell driver, BIOS, and firmware updates' -Operation {
        $driverVendor = Get-DriverVendor
        Suspend-BitLockerForVendorFirmware
        switch ($driverVendor) {
            'HP'   { Invoke-HPDriverUpdate -WorkingRoot $DriverWorkingRoot }
            'Dell' { Invoke-DellDriverUpdate -WorkingRoot $DriverWorkingRoot }
            default { throw "Unsupported driver vendor returned: $driverVendor" }
        }
    })

    [void](Invoke-PostDeploymentSection -Name 'Refresh C:\Scripts and register scheduled tasks' -Operation {
        Update-LocalMaintenanceScripts -SourcePath $ScriptsSourcePath -DestinationPath $LocalScriptsPath
    })

    if ($InstallPaperCut) {
        [void](Invoke-PostDeploymentSection -Name 'PaperCut Print Deploy installation' -Operation {
            Install-PaperCutPrintDeploy `
                -InstallerPath $PaperCutMsiPath `
                -TimeoutSeconds $PaperCutInstallTimeoutSeconds
        })
    }
    else {
        Write-Log 'PaperCut Print Deploy installation was disabled by configuration.' 'WARN'
    }

    if ($InstallAction1) {
        [void](Invoke-PostDeploymentSection -Name 'Action1 Agent installation' -Operation {
            Install-Action1Agent `
                -AgentPath $Action1AgentPath `
                -InstallerPath $Action1InstallerPath `
                -TimeoutSeconds $Action1InstallTimeoutSeconds
        })
    }
    else {
        Write-Log 'Action1 Agent installation was disabled by configuration.' 'WARN'
    }

    if ($RemoveMicrosoftCopilot) {
        [void](Invoke-PostDeploymentSection -Name 'Remove and disable Microsoft Copilot' -Operation {
            Disable-MicrosoftCopilot
        })
    }
    else {
        Write-Log 'Microsoft Copilot removal was disabled by configuration.' 'WARN'
    }

    if ($ActivateInstalledOffice) {
        [void](Invoke-PostDeploymentSection -Name 'Office LTSC 2024 activation' -Operation {
            if (Test-Office2024Installed) {
                Write-Log 'Office LTSC 2024 was detected. Checking activation.' 'INFO'
                Enable-InstalledOfficeActivation
            }
            else {
                Write-Log 'Office LTSC 2024 is not installed. Activation was skipped.' 'OK'
            }
        })
    }
    else {
        Write-Log 'Office LTSC 2024 activation was disabled by configuration.' 'WARN'
    }

    if ($UpdateApplications) {
        [void](Invoke-PostDeploymentSection -Name 'Application updates with winget' -Operation {
            Update-InstalledApplications
        })
    }
    else {
        Write-Log 'Application updates were disabled by configuration.' 'WARN'
    }

    if ($InstallWindowsUpdates) {
        [void](Invoke-PostDeploymentSection -Name 'Microsoft Windows updates' -Operation {
            Install-AllWindowsUpdates -TimeoutMinutes $WindowsUpdateTimeoutMinutes
        })
    }
    else {
        Write-Log 'Windows updates were disabled by configuration.' 'WARN'
    }

    [void](Invoke-PostDeploymentSection -Name 'Driver working-folder cleanup' -Operation {
        $removed = Remove-WorkingFolderRobust -Path $DriverWorkingRoot -RetryCount $CleanupRetryCount -RetryDelaySeconds $CleanupRetryDelaySeconds
        if (-not $removed) {
            throw "The driver working folder could not be removed: $DriverWorkingRoot"
        }
    })

    $failedSections = @($script:SectionResults | Where-Object Result -EQ 'Failed')
    Write-Log '===== Post-deployment section summary =====' 'INFO'
    foreach ($result in $script:SectionResults) {
        $summary = "$($result.Name): $($result.Result); Minutes=$($result.Minutes)"
        if (-not [string]::IsNullOrWhiteSpace([string]$result.Error)) {
            $summary += "; Error=$($result.Error)"
        }
        Write-Log $summary $(if ($result.Result -eq 'Succeeded') { 'OK' } else { 'ERROR' })
    }

    if ($failedSections.Count -gt 0) {
        $script:OverallResult = 'CompletedWithErrors'
        Write-Log "Post-deployment processing completed with $($failedSections.Count) failed section(s). Review this log: $($script:LogPath)" 'WARN'
    }
    else {
        $script:OverallResult = 'Succeeded'
        Write-Log 'All post-deployment sections completed successfully.' 'OK'
    }

    if ($RebootWhenComplete) {
        Invoke-FinalReboot -DelaySeconds $RebootDelaySeconds
    }
    else {
        Write-Log "Automatic reboot is disabled. WindowsUpdateRebootRequired=$($script:RebootRequired)" 'WARN'
    }

    if ($failedSections.Count -gt 0) { exit 2 }
    exit 0
}
catch {
    $script:FailureMessage = $_.Exception.Message
    $script:OverallResult = 'Failed'
    if ([string]::IsNullOrWhiteSpace([string]$script:LogPath)) {
        try { Initialize-Log } catch { }
    }
    Write-Log "Fatal script failure: $($script:FailureMessage)" 'ERROR'
    exit 1
}
