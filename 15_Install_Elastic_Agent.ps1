#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs Elastic Agent only on approved lab computers without enrolling it
    in Fleet or enabling log forwarding.

.DESCRIPTION
    - Uses a configurable computer-name prefix allowlist.
    - Defaults to the IB1-103 lab prefix.
    - Installs Elastic Agent from an approved ZIP package.
    - Does not use a Fleet URL or enrollment token.
    - Installs with a minimal standalone policy containing no inputs.
    - Stops and disables the Elastic Agent service after installation.
    - Is safe to run repeatedly; already-installed systems are skipped.

.NOTES
    Script:  15_Install_Elastic_Agent.ps1
    Version: 1.0.0
    Requires: Windows PowerShell 5.1, administrative/SYSTEM context
#>

[CmdletBinding()]
param(
    [switch]$ForceReinstall,
    [switch]$UseInternetDownload
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# CONFIGURATION
# ============================================================================

# Computer names must begin with one of these values.
# Add additional labs later, for example: 'IB1-103', 'SSB-122', 'SSB-114'
[string[]]$AllowedComputerPrefixes = @(
    'IB1-103'
)

# Keep this version aligned with the ZIP placed on the deployment share.
[string]$ElasticAgentVersion = '9.4.4'

# Preferred and fallback ZIP locations.
[string]$PreferredInstallerPath = "\\filesvr\Labscripts\ElasticAgent\elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"
[string]$FallbackInstallerPath  = "\\10.2.3.30\Labscripts\ElasticAgent\elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"

# Optional official-download fallback. It is used only with -UseInternetDownload.
[string]$DownloadUri = "https://artifacts.elastic.co/downloads/beats/elastic-agent/elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"

# Optional SHA-512 value from Elastic's matching .sha512 file.
# Leave blank to skip package-hash enforcement.
[string]$ExpectedSHA512 = ''

[string]$ElasticServiceName = 'Elastic Agent'
[string]$InstalledAgentPath = 'C:\Program Files\Elastic\Agent\elastic-agent.exe'
[string]$WorkingRoot        = 'C:\ProgramData\Compton\ElasticAgentInstall'
[string]$LogDirectory       = 'C:\Logs'
[string]$LogPath            = Join-Path $LogDirectory '15_Install_Elastic_Agent.log'
[int]$MaximumLogSizeMB      = 10
[int]$LogRetentionDays      = 60

# ============================================================================
# FUNCTIONS
# ============================================================================

function Initialize-Directory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Invoke-LogMaintenance {
    [CmdletBinding()]
    param()

    Initialize-Directory -Path $LogDirectory

    if (Test-Path -LiteralPath $LogPath -PathType Leaf) {
        $logItem = Get-Item -LiteralPath $LogPath
        if ($logItem.Length -ge ($MaximumLogSizeMB * 1MB)) {
            $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
            $archivePath = Join-Path $LogDirectory "15_Install_Elastic_Agent-$stamp.log"
            Move-Item -LiteralPath $LogPath -Destination $archivePath -Force
        }
    }

    Get-ChildItem -LiteralPath $LogDirectory -Filter '15_Install_Elastic_Agent-*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )

    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:COMPUTERNAME, $Level, $Message
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Test-ApprovedComputer {
    [CmdletBinding()]
    param()

    foreach ($prefix in $AllowedComputerPrefixes) {
        if (-not [string]::IsNullOrWhiteSpace($prefix) -and
            $env:COMPUTERNAME.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Test-IsAdministrator {
    [CmdletBinding()]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-InstalledElasticAgent {
    [CmdletBinding()]
    param()

    $service = Get-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue
    $exeExists = Test-Path -LiteralPath $InstalledAgentPath -PathType Leaf

    [pscustomobject]@{
        ServiceExists = ($null -ne $service)
        Service       = $service
        ExecutableExists = $exeExists
        IsInstalled   = (($null -ne $service) -and $exeExists)
    }
}

function Resolve-InstallerPackage {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DestinationPath)

    $source = $null

    if (Test-Path -LiteralPath $PreferredInstallerPath -PathType Leaf) {
        $source = $PreferredInstallerPath
    }
    elseif (Test-Path -LiteralPath $FallbackInstallerPath -PathType Leaf) {
        $source = $FallbackInstallerPath
    }

    if ($null -ne $source) {
        Write-Log -Message "Copying Elastic Agent package from $source"
        Copy-Item -LiteralPath $source -Destination $DestinationPath -Force
        return
    }

    if (-not $UseInternetDownload) {
        throw "Elastic Agent package was not found at the preferred or fallback share. Internet download is disabled."
    }

    Write-Log -Message "Downloading Elastic Agent package from the official Elastic artifact site."
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($DownloadUri, $DestinationPath)
}

function Test-PackageHash {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($ExpectedSHA512)) {
        Write-Log -Level WARNING -Message 'ExpectedSHA512 is blank; package hash enforcement was skipped.'
        return
    }

    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA512).Hash
    if (-not $actual.Equals($ExpectedSHA512.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Elastic Agent package SHA-512 mismatch. Expected $ExpectedSHA512 but found $actual."
    }

    Write-Log -Message 'Elastic Agent package SHA-512 validation passed.'
}

function Set-DisabledStandalonePolicy {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ExtractedAgentDirectory)

    # This is a valid, intentionally inactive standalone policy:
    # - No inputs are configured, so it collects no workstation logs.
    # - Agent self-monitoring is disabled.
    # - The output points only to localhost and is unused because there are no inputs.
    $policy = @'
outputs:
  default:
    type: elasticsearch
    hosts:
      - 'http://127.0.0.1:9200'

inputs: []

agent.monitoring:
  enabled: false
  logs: false
  metrics: false
'@

    $policyPath = Join-Path $ExtractedAgentDirectory 'elastic-agent.yml'
    Set-Content -LiteralPath $policyPath -Value $policy -Encoding ASCII -Force
    Write-Log -Message 'Created a disabled standalone policy with no log inputs.'
}

function Disable-ElasticAgentService {
    [CmdletBinding()]
    param()

    $service = Get-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        throw "The '$ElasticServiceName' service was not found after installation."
    }

    if ($service.Status -ne 'Stopped') {
        Stop-Service -Name $ElasticServiceName -Force -ErrorAction Stop
        $service.WaitForStatus('Stopped', (New-TimeSpan -Seconds 30))
    }

    Set-Service -Name $ElasticServiceName -StartupType Disabled
    Write-Log -Message 'Elastic Agent service was stopped and set to Disabled.'
}

# ============================================================================
# MAIN
# ============================================================================

$exitCode = 1

try {
    Invoke-LogMaintenance
    Write-Log -Message 'Elastic Agent installation check started.'

    if (-not (Test-IsAdministrator)) {
        throw 'This script must run as Administrator or SYSTEM.'
    }

    if (-not (Test-ApprovedComputer)) {
        Write-Log -Message "Computer is not in an approved lab prefix. Allowed prefixes: $($AllowedComputerPrefixes -join ', '). No changes were made."
        $exitCode = 0
        return
    }

    Write-Log -Message "Computer matched the approved Elastic Agent pilot prefix list."

    $installed = Get-InstalledElasticAgent
    if ($installed.IsInstalled -and -not $ForceReinstall) {
        Disable-ElasticAgentService
        Write-Log -Level SUCCESS -Message 'Elastic Agent is already installed. The service remains disabled and no enrollment was performed.'
        $exitCode = 0
        return
    }

    if (($installed.ServiceExists -or $installed.ExecutableExists) -and $ForceReinstall) {
        Write-Log -Level WARNING -Message 'ForceReinstall was requested. Removing the existing Elastic Agent installation.'

        if (Test-Path -LiteralPath $InstalledAgentPath -PathType Leaf) {
            $uninstallProcess = Start-Process -FilePath $InstalledAgentPath `
                -ArgumentList @('uninstall', '--force', '--skip-fleet-audit') `
                -Wait -PassThru -WindowStyle Hidden

            if ($uninstallProcess.ExitCode -ne 0) {
                throw "Existing Elastic Agent uninstall returned exit code $($uninstallProcess.ExitCode)."
            }
        }
    }

    if (Test-Path -LiteralPath $WorkingRoot) {
        Remove-Item -LiteralPath $WorkingRoot -Recurse -Force
    }
    Initialize-Directory -Path $WorkingRoot

    $zipPath = Join-Path $WorkingRoot "elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"
    $extractRoot = Join-Path $WorkingRoot 'Extracted'

    Resolve-InstallerPackage -DestinationPath $zipPath
    Unblock-File -LiteralPath $zipPath -ErrorAction SilentlyContinue
    Test-PackageHash -Path $zipPath

    Initialize-Directory -Path $extractRoot
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force

    $agentExecutable = Get-ChildItem -LiteralPath $extractRoot -Filter 'elastic-agent.exe' -File -Recurse |
        Select-Object -First 1

    if ($null -eq $agentExecutable) {
        throw 'elastic-agent.exe was not found in the extracted package.'
    }

    $agentDirectory = $agentExecutable.Directory.FullName
    Get-ChildItem -LiteralPath $agentDirectory -File -Recurse | Unblock-File -ErrorAction SilentlyContinue
    Set-DisabledStandalonePolicy -ExtractedAgentDirectory $agentDirectory

    Write-Log -Message "Installing Elastic Agent $ElasticAgentVersion without Fleet enrollment."
    $installProcess = Start-Process -FilePath $agentExecutable.FullName `
        -ArgumentList @('install', '--non-interactive') `
        -WorkingDirectory $agentDirectory `
        -Wait -PassThru -WindowStyle Hidden

    if ($installProcess.ExitCode -ne 0) {
        throw "Elastic Agent installer returned exit code $($installProcess.ExitCode)."
    }

    Disable-ElasticAgentService

    $verification = Get-InstalledElasticAgent
    if (-not $verification.IsInstalled) {
        throw 'Elastic Agent installation verification failed.'
    }

    $installedVersion = (& $InstalledAgentPath version 2>$null | Out-String).Trim()
    Write-Log -Level SUCCESS -Message "Elastic Agent installed successfully. Version response: $installedVersion"
    Write-Log -Level SUCCESS -Message 'Fleet enrollment was not performed, no log inputs are configured, and the Elastic Agent service is disabled.'

    $exitCode = 0
}
catch {
    try {
        Write-Log -Level ERROR -Message $_.Exception.Message
    }
    catch {
        Write-Error $_.Exception.Message
    }
    $exitCode = 1
}
finally {
    if (Test-Path -LiteralPath $WorkingRoot -PathType Container) {
        Remove-Item -LiteralPath $WorkingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

exit $exitCode
