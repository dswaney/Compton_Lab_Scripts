#requires -Version 5.1
#requires -RunAsAdministrator
# ScriptVersion: 2.1.5

<#
.SYNOPSIS
    Installs and enrolls Elastic Agent only on approved lab computers using
    the configured Fleet Server and enrollment token.

.DESCRIPTION
    - Uses a configurable computer-name prefix allowlist.
    - Defaults to the IB1-103 lab prefix.
    - Installs Elastic Agent from the internal ZIP package, with automatic official Elastic Internet fallback.
    - Uses the current Fleet Server URL and enrollment token from Install.ps1.
    - Performs a fully unattended Fleet installation with no Y/N prompt.
    - Leaves the Elastic Agent service installed and running under Fleet management.
    - Is safe to run repeatedly; already-installed systems are health-checked and skipped when present.

.NOTES
    Script:  15_Install_Elastic_Agent.ps1
    Version: 2.1.5
    LastUpdated: 2026-08-19
    Changes: v2.1.5 adds structured, mapping-safe Elastic Agent deployment telemetry,
             explicit result/failure-stage reporting, package-source/hash/enrollment state,
             existing/final agent health snapshots, and enrollment-token redaction.
             v2.1.4 uses Maintenance.Framework v2.4 staged text logging.
    Requires: Windows PowerShell 5.1, administrative/SYSTEM context
#>

[CmdletBinding()]
param(
    [switch]$ForceReinstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '2.1.5'

# ============================================================================
# CONFIGURATION
# ============================================================================

# Computer names must begin with one of these values.
# Add additional labs later, for example: 'IB1-103', 'SSB-122', 'SSB-114'
[string[]]$AllowedComputerPrefixes = @(
    'IB1-103'
)

# Keep this version aligned with the ZIP placed on the deployment share.
[string]$ElasticAgentVersion = '9.5.0'

# Fleet enrollment settings copied from the current Install.ps1.
# NOTE: The enrollment token is sensitive. Restrict read access to this script/share.
[string]$FleetServerUrl   = 'https://10.2.12.4:8220'
[string]$EnrollmentToken  = ''

# Fleet Server is currently using the Quick Start self-signed TLS certificate.
# Keep this $true until Fleet Server is moved to a certificate trusted by the lab PCs.
[bool]$UseInsecureFleetTls = $true

# Preferred and fallback ZIP locations.
[string]$PreferredInstallerPath = "\\filesvr\Labscripts\ElasticAgent\elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"
[string]$FallbackInstallerPath  = "\\10.2.3.30\Labscripts\ElasticAgent\elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"

# Official Elastic download fallback. Used automatically when both internal shares are unavailable.
[string]$DownloadUri = "https://artifacts.elastic.co/downloads/beats/elastic-agent/elastic-agent-$ElasticAgentVersion-windows-x86_64.zip"

# Optional SHA-512 value from Elastic's matching .sha512 file.
# Leave blank to skip package-hash enforcement.
[string]$ExpectedSHA512 = ''

[string]$ElasticServiceName = 'Elastic Agent'
[string]$InstalledAgentPath = 'C:\Program Files\Elastic\Agent\elastic-agent.exe'
[string]$WorkingRoot        = 'C:\ProgramData\Compton\ElasticAgentInstall'
[string]$LogDirectory       = 'C:\Logs'
[string]$LogPath            = $null
[string]$PublishedLogPath   = $null
$LogSession                 = $null
[int]$MaximumLogSizeMB      = 10
[int]$LogRetentionDays      = 60

[string]$TelemetryPath       = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
[string]$LatestTelemetryPath = Join-Path $LogDirectory '15_Install_Elastic_Agent.latest.json'

$script:StartTime    = Get-Date
$script:WarningCount = 0
$script:ErrorCount   = 0
$script:CurrentStage = 'Initialization'
$script:FailureStage = $null
$script:FailureMessage = $null
$script:MatchedPrefix = $null
$script:ExistingHealth = $null
$script:FinalHealth = $null
$script:PackageSource = $null
$script:PackageHashEnforced = -not [string]::IsNullOrWhiteSpace($ExpectedSHA512)
$script:PackageHashVerified = $false
$script:InstallationAttempted = $false
$script:EnrollmentAttempted = $false
$script:InstallerExitCode = $null
$script:VersionResponse = $null
$script:ExistingHealthRestartAttempted = $false
$script:ExistingHealthRestored = $false
$script:ReinstallAttempted = $false
$script:OverallResult = 'Unknown'

$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 15 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName '15_Install_Elastic_Agent.ps1' `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '15_Install_Elastic_Agent.log',
        '*-15_Install_Elastic_Agent-*.log'
    ) | Out-Null

$LogSession = New-MaintenanceStagedLog `
    -ScriptName '15_Install_Elastic_Agent.ps1' `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $env:COMPUTERNAME `
    -Timestamp (Get-Date)

$LogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

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
    # Current and legacy script 15 logs were archived before the staged log was created.
    # Retention of C:\Logs\Old Logs is managed centrally by Maintenance.Framework.
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )

    if ($Level -eq 'WARNING') { $script:WarningCount++ }
    elseif ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:COMPUTERNAME, $Level, $Message
    $activeLogDirectory = Split-Path -Parent $LogPath
    Initialize-Directory -Path $activeLogDirectory
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    Write-Host $line
}


function New-StringArrayForJson {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    [string[]]$items = @(
        $InputObject |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Convert-AgentHealthForTelemetry {
    [CmdletBinding()]
    param([AllowNull()]$Health)

    if ($null -eq $Health) {
        return $null
    }

    [ordered]@{
        Installed      = [bool]$Health.Installed
        ServiceRunning = [bool]$Health.ServiceRunning
        ServiceStatus  = $Health.ServiceStatus
        StatusHealthy  = [bool]$Health.StatusHealthy
        Healthy        = [bool]$Health.Healthy
        StatusExitCode = $Health.StatusExitCode
        StatusOutput   = New-StringArrayForJson -InputObject $Health.StatusOutput
        Message        = $Health.Message
    }
}

function Write-JsonAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Json
    )

    $temporary = '{0}.{1}.tmp' -f $Path, ([guid]::NewGuid().Guid)
    [IO.File]::WriteAllText($temporary, $Json, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Write-ElasticAgentTelemetry {
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ExitCode)

    $eventTime = Get-Date
    $domain = $env:USERDOMAIN
    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.Domain) { $domain = [string]$computerSystem.Domain }
    }
    catch { }

    $status = if ($ExitCode -ne 0) {
        'Failed'
    }
    elseif ($script:OverallResult -eq 'ExistingUnhealthyNeedsInvestigation' -or
            $script:OverallResult -eq 'InstalledButUnhealthy') {
        'SuccessWithWarnings'
    }
    else {
        'Success'
    }

    $event = [ordered]@{
        '@timestamp'   = $eventTime.ToUniversalTime().ToString('o')
        EventType      = 'maintenance.elastic_agent'
        SchemaVersion  = '1.0'
        ComputerName   = $env:COMPUTERNAME
        Domain         = $domain
        ScriptName     = '15_Install_Elastic_Agent.ps1'
        ScriptVersion  = $ScriptVersion
        Status         = $status
        OverallResult  = $script:OverallResult
        ExitCode       = $ExitCode
        StartTime      = $script:StartTime.ToUniversalTime().ToString('o')
        EndTime        = $eventTime.ToUniversalTime().ToString('o')
        DurationSeconds = [math]::Round(($eventTime - $script:StartTime).TotalSeconds, 3)
        WarningCount   = $script:WarningCount
        ErrorCount     = $script:ErrorCount
        FailureStage   = $script:FailureStage
        FailureMessage = $script:FailureMessage
        TextLogPath    = $PublishedLogPath

        Targeting = [ordered]@{
            Approved        = ($null -ne $script:MatchedPrefix)
            MatchedPrefix   = $script:MatchedPrefix
            AllowedPrefixes = New-StringArrayForJson -InputObject $AllowedComputerPrefixes
        }

        ExistingAgent = Convert-AgentHealthForTelemetry -Health $script:ExistingHealth

        Installation = [ordered]@{
            ForceReinstall          = [bool]$ForceReinstall
            Attempted               = [bool]$script:InstallationAttempted
            ReinstallAttempted      = [bool]$script:ReinstallAttempted
            PackageSource           = $script:PackageSource
            PackageVersion          = $ElasticAgentVersion
            PreferredInstallerPath  = $PreferredInstallerPath
            FallbackInstallerPath   = $FallbackInstallerPath
            DownloadUri             = $DownloadUri
            PackageHashEnforced     = [bool]$script:PackageHashEnforced
            PackageHashVerified     = [bool]$script:PackageHashVerified
            FleetServerUrl          = $FleetServerUrl
            UseInsecureFleetTls     = [bool]$UseInsecureFleetTls
            EnrollmentAttempted     = [bool]$script:EnrollmentAttempted
            EnrollmentTokenProvided = -not [string]::IsNullOrWhiteSpace($EnrollmentToken)
            EnrollmentTokenLogged   = $false
            InstallerExitCode       = $script:InstallerExitCode
        }

        Actions = [ordered]@{
            ExistingAgentDetected          = ($null -ne $script:ExistingHealth -and [bool]$script:ExistingHealth.Installed)
            ExistingHealthRestartAttempted = [bool]$script:ExistingHealthRestartAttempted
            ExistingHealthRestored         = [bool]$script:ExistingHealthRestored
            ReinstallAttempted              = [bool]$script:ReinstallAttempted
            EnrollmentAttempted             = [bool]$script:EnrollmentAttempted
        }

        FinalAgent = Convert-AgentHealthForTelemetry -Health $script:FinalHealth
        VersionResponse = $script:VersionResponse
    }

    Write-Log -Message ("Elastic Agent telemetry summary: Status={0}; Result={1}; Targeted={2}; ExistingHealthy={3}; InstallAttempted={4}; PackageSource={5}; FinalHealthy={6}; ExitCode={7}" -f `
        $event.Status,
        $event.OverallResult,
        $event.Targeting.Approved,
        $(if ($event.ExistingAgent) { $event.ExistingAgent.Healthy } else { $null }),
        $event.Installation.Attempted,
        $event.Installation.PackageSource,
        $(if ($event.FinalAgent) { $event.FinalAgent.Healthy } else { $null }),
        $event.ExitCode)

    $compact = $event | ConvertTo-Json -Depth 12 -Compress
    $pretty  = $event | ConvertTo-Json -Depth 12

    Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compact
    Write-JsonAtomically -Path $LatestTelemetryPath -Json $pretty
}

function Test-ApprovedComputer {
    [CmdletBinding()]
    param()

    foreach ($prefix in $AllowedComputerPrefixes) {
        if (-not [string]::IsNullOrWhiteSpace($prefix) -and
            $env:COMPUTERNAME.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $script:MatchedPrefix = [string]$prefix
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


function Get-ElasticAgentHealth {
    [CmdletBinding()]
    param()

    $installed = Get-InstalledElasticAgent

    if (-not $installed.IsInstalled) {
        return [pscustomobject]@{
            Installed      = $false
            ServiceRunning = $false
            StatusHealthy  = $false
            Healthy        = $false
            ServiceStatus  = $null
            StatusExitCode = $null
            StatusOutput   = New-StringArrayForJson -InputObject @()
            Message        = 'Elastic Agent is not fully installed.'
        }
    }

    $service = Get-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue

    if ($service -and $service.Status -ne 'Running') {
        Write-Log -Level WARNING -Message "Elastic Agent service is installed but currently $($service.Status). Attempting to start it."

        try {
            Set-Service -Name $ElasticServiceName -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name $ElasticServiceName -ErrorAction Stop
            $service.WaitForStatus('Running', [timespan]::FromSeconds(20))
        }
        catch {
            Write-Log -Level WARNING -Message "Unable to start Elastic Agent service: $($_.Exception.Message)"
        }

        $service = Get-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue
    }

    $serviceRunning = ($service -and $service.Status -eq 'Running')
    $statusOutput = @()
    $statusExitCode = $null
    $statusHealthy = $false

    if (Test-Path -LiteralPath $InstalledAgentPath -PathType Leaf) {
        try {
            $statusOutput = @(& $InstalledAgentPath status 2>&1 | ForEach-Object { [string]$_ })
            $statusExitCode = $LASTEXITCODE
            $joined = ($statusOutput -join "`n")

            $statusHealthy = (
                $statusExitCode -eq 0 -and
                $joined -match '(?im)\bHEALTHY\b' -and
                $joined -notmatch '(?im)\bFAILED\b|\bDEGRADED\b|\bUNHEALTHY\b'
            )
        }
        catch {
            $statusOutput = @($_.Exception.Message)
            $statusExitCode = -1
        }
    }

    $healthy = ($serviceRunning -and $statusHealthy)

    [pscustomobject]@{
        Installed      = $true
        ServiceRunning = $serviceRunning
        StatusHealthy  = $statusHealthy
        Healthy        = $healthy
        ServiceStatus  = if ($service) { [string]$service.Status } else { 'Missing' }
        StatusExitCode = $statusExitCode
        StatusOutput   = New-StringArrayForJson -InputObject $statusOutput
        Message        = if ($healthy) {
            'Elastic Agent is installed and reporting HEALTHY.'
        }
        elseif (-not $serviceRunning) {
            'Elastic Agent is installed, but its Windows service is not running.'
        }
        else {
            'Elastic Agent service is running, but elastic-agent status is not reporting HEALTHY.'
        }
    }
}

function Wait-ElasticAgentHealth {
    [CmdletBinding()]
    param(
        [ValidateRange(1,20)][int]$Attempts = 8,
        [ValidateRange(1,60)][int]$DelaySeconds = 5
    )

    $lastHealth = $null

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        $lastHealth = Get-ElasticAgentHealth

        if ($lastHealth.Healthy) {
            return $lastHealth
        }

        if ($attempt -lt $Attempts) {
            Write-Log -Message ("Elastic Agent health check {0}/{1} is not yet HEALTHY. ServiceStatus={2}; StatusExitCode={3}. Retrying in {4} seconds." -f `
                $attempt, $Attempts, $lastHealth.ServiceStatus, $lastHealth.StatusExitCode, $DelaySeconds)
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    return $lastHealth
}

function Test-ExistingElasticAgentAndExit {
    [CmdletBinding()]
    param()

    $health = Get-ElasticAgentHealth

    if (-not $health.Installed) {
        return $false
    }

    Write-Log -Message "Existing Elastic Agent detected. ServiceStatus=$($health.ServiceStatus); StatusExitCode=$($health.StatusExitCode)."

    foreach ($line in @($health.StatusOutput | Select-Object -First 15)) {
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            Write-Log -Message "Elastic status: $line"
        }
    }

    if ($health.Healthy) {
        Write-Log -Level SUCCESS -Message 'Elastic Agent is already installed and working properly. No installation action is required.'
        return $true
    }

    Write-Log -Level WARNING -Message $health.Message
    $script:ExistingHealthRestartAttempted = $true

    try {
        Restart-Service -Name $ElasticServiceName -Force -ErrorAction Stop
        (Get-Service -Name $ElasticServiceName -ErrorAction Stop).WaitForStatus('Running', [timespan]::FromSeconds(20))
    }
    catch {
        Write-Log -Level WARNING -Message "Elastic Agent service restart attempt failed: $($_.Exception.Message)"
    }

    $retry = Wait-ElasticAgentHealth -Attempts 6 -DelaySeconds 5

    if ($retry.Healthy) {
        $script:ExistingHealthRestored = $true
        Write-Log -Level SUCCESS -Message 'Elastic Agent was already installed. Health was restored after restarting the service; agent is now working properly.'
        return $true
    }

    Write-Log -Level WARNING -Message 'Elastic Agent is installed but did not verify as healthy after a service restart. Installation will not be duplicated; manual/Fleet investigation is recommended.'
    return $true
}

function Resolve-InstallerPackage {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DestinationPath)

    $source = $null

    if (Test-Path -LiteralPath $PreferredInstallerPath -PathType Leaf -ErrorAction SilentlyContinue) {
        $source = $PreferredInstallerPath
        Write-Log -Level SUCCESS -Message "Elastic Agent package found at preferred share: $PreferredInstallerPath"
    }
    elseif (Test-Path -LiteralPath $FallbackInstallerPath -PathType Leaf -ErrorAction SilentlyContinue) {
        $source = $FallbackInstallerPath
        Write-Log -Level WARNING -Message "Preferred Elastic Agent share was unavailable. Using fallback share: $FallbackInstallerPath"
    }

    if ($null -ne $source) {
        Write-Log -Message "Copying Elastic Agent package from $source"
        Copy-Item -LiteralPath $source -Destination $DestinationPath -Force -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $DestinationPath -PathType Leaf)) {
            throw "Elastic Agent package copy completed without error, but the local ZIP was not created: $DestinationPath"
        }

        return 'Share'
    }

    Write-Log -Level WARNING -Message 'Elastic Agent package was not accessible from either internal share. Falling back to the official Elastic artifact site.'
    Write-Log -Message "Downloading Elastic Agent $ElasticAgentVersion from $DownloadUri"

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        Invoke-WebRequest `
            -Uri $DownloadUri `
            -OutFile $DestinationPath `
            -UseBasicParsing `
            -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $DestinationPath -PathType Leaf)) {
            throw "Elastic Agent Internet download completed without error, but the ZIP was not created: $DestinationPath"
        }

        $downloadedLength = (Get-Item -LiteralPath $DestinationPath -ErrorAction Stop).Length
        if ($downloadedLength -lt 1MB) {
            throw "Elastic Agent Internet download appears incomplete. Downloaded size: $downloadedLength bytes."
        }

        Write-Log -Level SUCCESS -Message "Elastic Agent package downloaded successfully from the official Elastic artifact site. Size=$downloadedLength bytes."
        return 'Internet'
    }
    catch {
        throw "Elastic Agent package was unavailable from both internal shares and the official Internet download also failed: $($_.Exception.Message)"
    }
}

function Test-PackageHash {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($ExpectedSHA512)) {
        Write-Log -Level WARNING -Message 'ExpectedSHA512 is blank; package hash enforcement was skipped.'
        return $false
    }

    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA512).Hash
    if (-not $actual.Equals($ExpectedSHA512.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Elastic Agent package SHA-512 mismatch. Expected $ExpectedSHA512 but found $actual."
    }

    Write-Log -Message 'Elastic Agent package SHA-512 validation passed.'
    return $true
}

# ============================================================================
# MAIN
# ============================================================================

$exitCode = 1

try {
    $script:CurrentStage = 'Initialization'
    Invoke-LogMaintenance
    Write-Log -Message "Elastic Agent installation check started. ScriptVersion=$ScriptVersion."

    $script:CurrentStage = 'PrivilegeCheck'
    if (-not (Test-IsAdministrator)) {
        throw 'This script must run as Administrator or SYSTEM.'
    }

    $script:CurrentStage = 'Targeting'
    if (-not (Test-ApprovedComputer)) {
        Write-Log -Message "Computer is not in an approved lab prefix. Allowed prefixes: $($AllowedComputerPrefixes -join ', '). No changes were made."
        $script:OverallResult = 'NotTargeted'
        $script:FinalHealth = Get-ElasticAgentHealth
        $exitCode = 0
        return
    }

    Write-Log -Message "Computer matched the approved Elastic Agent pilot prefix list."

    $script:CurrentStage = 'ExistingAgentCheck'
    $installed = Get-InstalledElasticAgent
    $script:ExistingHealth = Get-ElasticAgentHealth

    if ($installed.IsInstalled -and -not $ForceReinstall) {
        $existingWasHealthy = [bool]$script:ExistingHealth.Healthy

        if (Test-ExistingElasticAgentAndExit) {
            $script:FinalHealth = Get-ElasticAgentHealth

            if ($existingWasHealthy) {
                $script:OverallResult = 'AlreadyHealthy'
            }
            elseif ($script:FinalHealth.Healthy -and $script:ExistingHealthRestored) {
                $script:OverallResult = 'ExistingHealthRestored'
            }
            else {
                $script:OverallResult = 'ExistingUnhealthyNeedsInvestigation'
            }

            $exitCode = 0
            return
        }
    }

    if (($installed.ServiceExists -or $installed.ExecutableExists) -and $ForceReinstall) {
        $script:CurrentStage = 'UninstallExistingAgent'
        $script:ReinstallAttempted = $true
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

    $script:CurrentStage = 'ResolveInstallerPackage'
    $script:PackageSource = Resolve-InstallerPackage -DestinationPath $zipPath
    Unblock-File -LiteralPath $zipPath -ErrorAction SilentlyContinue

    $script:CurrentStage = 'ValidatePackageHash'
    $script:PackageHashVerified = [bool](Test-PackageHash -Path $zipPath)

    $script:CurrentStage = 'ExtractPackage'
    Initialize-Directory -Path $extractRoot
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force

    $agentExecutable = Get-ChildItem -LiteralPath $extractRoot -Filter 'elastic-agent.exe' -File -Recurse |
        Select-Object -First 1

    if ($null -eq $agentExecutable) {
        throw 'elastic-agent.exe was not found in the extracted package.'
    }

    $agentDirectory = $agentExecutable.Directory.FullName
    Get-ChildItem -LiteralPath $agentDirectory -File -Recurse | Unblock-File -ErrorAction SilentlyContinue

    $script:CurrentStage = 'InstallAndEnroll'
    $script:InstallationAttempted = $true
    $script:EnrollmentAttempted = $true
    Write-Log -Message "Installing Elastic Agent $ElasticAgentVersion and enrolling with Fleet Server $FleetServerUrl."

    # --non-interactive suppresses installation questions/prompts.
    # --force suppresses configuration-overwrite confirmation. This script only
    # reaches this point on a fresh install or after -ForceReinstall removed the
    # previous installation, avoiding duplicate-agent behavior.
    $installArguments = @(
        'install',
        "--url=$FleetServerUrl",
        "--enrollment-token=$EnrollmentToken",
        '--non-interactive',
        '--force'
    )

    # Quick Start Fleet Server uses a self-signed certificate. Remote agents must
    # use --insecure until Fleet Server is configured with a certificate trusted
    # by the Windows endpoints.
    if ($UseInsecureFleetTls) {
        $installArguments += '--insecure'
        Write-Log -Level WARNING -Message 'Fleet Server TLS certificate verification is disabled for enrollment because UseInsecureFleetTls is enabled.'
    }

    # Capture the Elastic Agent install/enrollment output so a failed deployment
    # records the underlying Elastic error instead of only returning exit code 1.
    $installerStdOut = Join-Path $WorkingRoot 'elastic-agent-install.stdout.log'
    $installerStdErr = Join-Path $WorkingRoot 'elastic-agent-install.stderr.log'

    $installProcess = Start-Process -FilePath $agentExecutable.FullName `
        -ArgumentList $installArguments `
        -WorkingDirectory $agentDirectory `
        -Wait -PassThru -WindowStyle Hidden `
        -RedirectStandardOutput $installerStdOut `
        -RedirectStandardError $installerStdErr

    $script:InstallerExitCode = $installProcess.ExitCode
    $capturedInstallerOutput = @()
    if (Test-Path -LiteralPath $installerStdOut -PathType Leaf) {
        $capturedInstallerOutput += Get-Content -LiteralPath $installerStdOut -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $installerStdErr -PathType Leaf) {
        $capturedInstallerOutput += Get-Content -LiteralPath $installerStdErr -ErrorAction SilentlyContinue
    }

    $script:CurrentStage = 'VerifyInstalledAgent'
    $verification = Wait-ElasticAgentHealth -Attempts 8 -DelaySeconds 5
    $script:FinalHealth = $verification

    if ($installProcess.ExitCode -ne 0 -and -not $verification.Healthy) {
        $capturedInstallerOutput | Select-Object -Last 30 | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                Write-Log -Level ERROR -Message "Elastic Agent: $_"
            }
        }
        throw "Elastic Agent installer returned exit code $($installProcess.ExitCode), and the installed agent did not become HEALTHY during verification."
    }

    $capturedInstallerOutput | Select-Object -Last 10 | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_)) {
            Write-Log -Message "Elastic Agent: $_"
        }
    }

    if ($installProcess.ExitCode -ne 0 -and $verification.Healthy) {
        Write-Log -Level WARNING -Message ("Elastic Agent installer returned exit code {0}, but the installed service subsequently reported HEALTHY. The installer result was treated as a transient daemon-startup race." -f $installProcess.ExitCode)
    }

    if (-not $verification.Installed) {
        throw 'Elastic Agent installation verification failed.'
    }

    if (-not $verification.Healthy) {
        Write-Log -Level WARNING -Message "Elastic Agent installation completed, but health verification did not report HEALTHY. ServiceStatus=$($verification.ServiceStatus); StatusExitCode=$($verification.StatusExitCode)."
        foreach ($line in @($verification.StatusOutput | Select-Object -First 15)) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Log -Message "Elastic status: $line"
            }
        }
    }
    else {
        Write-Log -Level SUCCESS -Message 'Elastic Agent installation verified: service is running and elastic-agent status reports HEALTHY.'
    }

    try {
        Set-Service -Name $ElasticServiceName -StartupType Automatic -ErrorAction Stop
        Start-Service -Name $ElasticServiceName -ErrorAction SilentlyContinue
        $serviceState = (Get-Service -Name $ElasticServiceName -ErrorAction Stop).Status
        Write-Log -Message "Elastic Agent service state after installation: $serviceState"
    }
    catch {
        Write-Log -Level WARNING -Message "Elastic Agent installed, but service verification encountered an issue: $($_.Exception.Message)"
    }

    $installedVersion = (& $InstalledAgentPath version 2>$null | Out-String).Trim()
    $script:VersionResponse = $installedVersion
    Write-Log -Level SUCCESS -Message "Elastic Agent installed successfully. Version response: $installedVersion"
    Write-Log -Level SUCCESS -Message 'Fleet enrollment was requested successfully and the Elastic Agent service is left enabled for Fleet management.'

    $script:FinalHealth = Get-ElasticAgentHealth
    if ($script:FinalHealth.Healthy) {
        $script:OverallResult = if ($script:ReinstallAttempted) { 'ReinstalledAndHealthy' } else { 'InstalledAndHealthy' }
    }
    else {
        $script:OverallResult = 'InstalledButUnhealthy'
    }

    $exitCode = 0
}
catch {
    $script:FailureStage = $script:CurrentStage
    $script:FailureMessage = $_.Exception.Message
    $script:OverallResult = 'Failed'
    try {
        $script:FinalHealth = Get-ElasticAgentHealth
    }
    catch { }

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

    # Emit one structured deployment event for Elastic. The enrollment token itself
    # is never written; telemetry records only whether one was provided.
    try {
        if ($null -eq $script:FinalHealth) {
            $script:FinalHealth = Get-ElasticAgentHealth
        }
        Write-ElasticAgentTelemetry -ExitCode $exitCode
    }
    catch {
        try {
            Write-Log -Level ERROR -Message ("Elastic Agent telemetry finalization failed: {0}" -f $_.Exception.Message)
        }
        catch { }
        if ($exitCode -eq 0) { $exitCode = 4 }
    }

    # Final append before the completed immutable text log enters C:\Logs.
    try {
        Write-Log -Message ("Finalizing 15_Install_Elastic_Agent.ps1. ExitCode={0}" -f $exitCode) `
            -Level $(if ($exitCode -eq 0) { 'SUCCESS' } else { 'ERROR' })
    }
    catch {
        Write-Warning ("Unable to write final script 15 staged log entry: {0}" -f $_.Exception.Message)
    }

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 15 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 15 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $exitCode
