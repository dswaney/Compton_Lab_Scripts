# =====================================================================
# ScriptName: 13_Configure_Autologon_And_Edge.ps1
# ScriptVersion: 2.1.3
# LastUpdated: 2026-08-17
# Changes: v2.1.3 normalizes Windows 11 product naming, makes telemetry collections explicitly JSON-array safe,
#          and adds a concise Elastic configuration summary while preserving password redaction.
#          v2.1.2 uses Maintenance.Framework v2.4 staged text logging.
# Purpose: Configure lab autologon and launch Microsoft Edge InPrivate
#          for approved computer-name patterns, with verification and
#          structured maintenance telemetry. The autologon password is Base64-obfuscated in the script and decoded only at runtime.
# =====================================================================

[CmdletBinding()]
param(
    [string[]]$ComputerNamePatterns = @(
        'SSB-122-*',
        'SSB-114*'
    ),

    [string]$DefaultUserName = 'CC-Student',

    [ValidateNotNullOrEmpty()]
    [string]$DefaultPasswordBase64 = 'Q0MkdHVkM250IQ==',

    [string]$DefaultDomainName = 'Compton.edu',

    [ValidateNotNullOrEmpty()]
    [string]$EdgeUrl = 'https://www.compton.edu',

    [string]$LogDirectory = 'C:\Logs',

    [switch]$AllowHttpEdgeUrl
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptName = '13_Configure_Autologon_And_Edge.ps1'
$ScriptVersion = '2.1.3'
$RunId = [guid]::NewGuid().Guid
$StartTime = Get-Date
$ComputerName = $env:COMPUTERNAME
$DomainName = $env:USERDOMAIN
$RunningAccount = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$LogPath = $null
$PublishedLogPath = $null
$LogSession = $null
$TelemetryPath = Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson'
$LatestTelemetryPath = Join-Path $LogDirectory '13_Configure_Autologon_And_Edge.latest.json'
$WinlogonPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
$RunKeyPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$EdgeRunValueName = 'LaunchComptonEdge'
$AllUsersStartupPath = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup'
$LegacyChromeShortcutPath = Join-Path $AllUsersStartupPath 'Google Chrome.lnk'

$script:WarningCount = 0
$script:ErrorCount = 0
$script:ExitCode = 0
$script:FinalStatus = 'Success'
$script:OverallResult = 'Unknown'
$script:FailureMessage = $null
$script:Targeted = $false
$script:MatchedPattern = $null
$script:ChangesMade = 0
$script:VerificationFailures = New-Object System.Collections.Generic.List[string]
$script:RegistryChanges = New-Object System.Collections.Generic.List[object]
$script:EdgePath = $null
$script:EdgeCommand = $null
$script:LegacyShortcutRemoved = $false
$script:LegacyShortcutPresentBefore = $false
$script:LegacyShortcutPresentAfter = $false
$script:BeforeState = $null
$script:AfterState = $null
$script:DecodedPassword = $null

# Load the shared framework from the same directory as this script.
$MaintenanceFrameworkPath = 'C:\Scripts\Maintenance.Framework.psm1'
Import-Module -Name $MaintenanceFrameworkPath -Force -ErrorAction Stop
$MaintenanceConfig = Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogDirectory

$requiredFrameworkVersion = [version]'2.4.0'
$currentFrameworkVersion = [version](Get-MaintenanceFrameworkVersion)

if ($currentFrameworkVersion -lt $requiredFrameworkVersion) {
    throw "Script 13 requires Maintenance.Framework.psm1 version $requiredFrameworkVersion or newer. Installed version: $currentFrameworkVersion"
}

Archive-MaintenanceLogs `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -AdditionalPatterns @(
        '13_Configure_Autologon_And_Edge.log',
        '*-13_Configure_Autologon_And_Edge-*.log'
    ) | Out-Null

$LogSession = New-MaintenanceStagedLog `
    -ScriptName $ScriptName `
    -LogRoot $LogDirectory `
    -StagingRoot $MaintenanceConfig.LogStagingRoot `
    -ComputerName $ComputerName `
    -Timestamp $StartTime

$LogPath = [string]$LogSession.WorkingPath
$PublishedLogPath = [string]$LogSession.PublishedPath

function Initialize-LogDirectory {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') { $script:WarningCount++ }
    if ($Level -eq 'ERROR') { $script:ErrorCount++ }

    $normalizedLevel = switch ($Level) {
        'OK'   { 'SUCCESS' }
        'WARN' { 'WARNING' }
        default { $Level }
    }
    $logComputerName = if ($ComputerName) { $ComputerName } else { 'UNKNOWN' }
    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $logComputerName, $normalizedLevel, $Message

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }

    try {
        $activeLogDirectory = Split-Path -Parent $LogPath
        if (-not (Test-Path -LiteralPath $activeLogDirectory -PathType Container)) {
            New-Item -Path $activeLogDirectory -ItemType Directory -Force | Out-Null
        }
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Unable to write to the log file: $($_.Exception.Message)"
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-WindowsInformation {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $currentVersion = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop

        $productName = [string]$currentVersion.ProductName
        $buildText = [string]$os.BuildNumber
        $buildNumber = 0
        [void][int]::TryParse($buildText, [ref]$buildNumber)

        if ($buildNumber -ge 22000 -and $productName -match '^Windows 10') {
            $productName = $productName -replace '^Windows 10', 'Windows 11'
        }

        [pscustomobject]@{
            ProductName    = $productName
            EditionId      = [string]$currentVersion.EditionID
            DisplayVersion = [string]$currentVersion.DisplayVersion
            Version        = [string]$os.Version
            BuildNumber    = $buildText
            UBR            = [int]$currentVersion.UBR
            FullBuild      = '{0}.{1}' -f $buildText, $currentVersion.UBR
        }
    }
    catch {
        Write-Log "Unable to collect Windows version information: $($_.Exception.Message)" 'WARN'
        [pscustomobject]@{
            ProductName = $null; EditionId = $null; DisplayVersion = $null
            Version = $null; BuildNumber = $null; UBR = $null; FullBuild = $null
        }
    }
}

function Get-MatchedComputerPattern {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $Name -like $pattern) {
            return $pattern
        }
    }
    return $null
}

function Get-RegistryStringValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $value = (Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop).$Name
        return [string]$value
    }
    catch {
        return $null
    }
}

function Set-VerifiedRegistryString {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value,
        [switch]$Sensitive
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    $before = Get-RegistryStringValue -Path $Path -Name $Name
    $changed = ([string]$before -cne [string]$Value)

    if ($changed) {
        New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
        $script:ChangesMade++
        if ($Sensitive) {
            Write-Log "Updated protected registry value: $Path\$Name" 'OK'
        }
        else {
            Write-Log "Updated registry value: $Path\$Name" 'OK'
        }
    }
    else {
        Write-Log "Registry value is already correct: $Path\$Name"
    }

    $after = Get-RegistryStringValue -Path $Path -Name $Name
    $verified = ([string]$after -ceq [string]$Value)

    if (-not $verified) {
        $script:VerificationFailures.Add("Registry verification failed: $Path\$Name")
        throw "Registry value verification failed: $Path\$Name"
    }

    $script:RegistryChanges.Add([pscustomobject]@{
        Path          = $Path
        Name          = $Name
        Changed       = $changed
        Verified      = $verified
        Sensitive     = [bool]$Sensitive
        BeforePresent = ($null -ne $before)
        AfterPresent  = ($null -ne $after)
        BeforeValue   = $(if ($Sensitive) { $null } else { $before })
        AfterValue    = $(if ($Sensitive) { $null } else { $after })
    })
}

function Get-DecodedAutologonPassword {
    try {
        $bytes = [Convert]::FromBase64String($DefaultPasswordBase64)
        $password = [Text.Encoding]::UTF8.GetString($bytes)

        if ([string]::IsNullOrWhiteSpace($password)) {
            throw 'The decoded autologon password is blank.'
        }

        return $password
    }
    catch {
        throw "Unable to decode the configured autologon password: $($_.Exception.Message)"
    }
    finally {
        $bytes = $null
    }
}

function Test-ConfigurationInput {
    if (-not $ComputerNamePatterns -or @($ComputerNamePatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0) {
        throw 'At least one nonblank computer-name pattern is required.'
    }
    if ([string]::IsNullOrWhiteSpace($DefaultUserName)) {
        throw 'DefaultUserName cannot be blank.'
    }
    if ([string]::IsNullOrWhiteSpace($DefaultPasswordBase64)) {
        throw 'DefaultPasswordBase64 cannot be blank when autologon is enabled.'
    }

    try {
        $decodedBytes = [Convert]::FromBase64String($DefaultPasswordBase64)
        if ($decodedBytes.Length -eq 0) {
            throw 'The decoded password is empty.'
        }
    }
    catch {
        throw 'DefaultPasswordBase64 is not valid Base64 or decodes to an empty value.'
    }
    finally {
        $decodedBytes = $null
    }
    if ([string]::IsNullOrWhiteSpace($DefaultDomainName)) {
        throw 'DefaultDomainName cannot be blank.'
    }

    $uri = $null
    if (-not [uri]::TryCreate($EdgeUrl, [UriKind]::Absolute, [ref]$uri)) {
        throw "EdgeUrl is not a valid absolute URL: $EdgeUrl"
    }
    if ($uri.Scheme -notin @('https','http')) {
        throw "EdgeUrl must use HTTP or HTTPS: $EdgeUrl"
    }
    if ($uri.Scheme -eq 'http' -and -not $AllowHttpEdgeUrl) {
        throw 'EdgeUrl uses HTTP. Use HTTPS or explicitly supply -AllowHttpEdgeUrl.'
    }
}

function Get-AutologonState {
    $autoAdmin = Get-RegistryStringValue -Path $WinlogonPath -Name 'AutoAdminLogon'
    $user = Get-RegistryStringValue -Path $WinlogonPath -Name 'DefaultUserName'
    $domain = Get-RegistryStringValue -Path $WinlogonPath -Name 'DefaultDomainName'
    $force = Get-RegistryStringValue -Path $WinlogonPath -Name 'ForceAutoLogon'
    $password = Get-RegistryStringValue -Path $WinlogonPath -Name 'DefaultPassword'

    [pscustomobject]@{
        AutoAdminLogonEnabled = ($autoAdmin -eq '1')
        DefaultUserName       = $user
        DefaultDomainName     = $domain
        ForceAutoLogonEnabled = ($force -eq '1')
        PasswordPresent       = (-not [string]::IsNullOrEmpty($password))
        PasswordMatches       = ($null -ne $script:DecodedPassword -and [string]$password -ceq [string]$script:DecodedPassword)
        ConfigurationMatches  = (
            $autoAdmin -eq '1' -and
            $user -ceq $DefaultUserName -and
            $domain -ceq $DefaultDomainName -and
            $force -eq '1' -and
            [string]$password -ceq [string]$script:DecodedPassword
        )
    }
}

function Set-AutologonConfiguration {
    Write-Log "Ensuring Windows autologon is configured for '$DefaultDomainName\$DefaultUserName'."

    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'AutoAdminLogon' -Value '1'
    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'DefaultUserName' -Value $DefaultUserName
    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'DefaultPassword' -Value $script:DecodedPassword -Sensitive
    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'DefaultDomainName' -Value $DefaultDomainName
    Set-VerifiedRegistryString -Path $WinlogonPath -Name 'ForceAutoLogon' -Value '1'

    $state = Get-AutologonState
    if (-not $state.ConfigurationMatches) {
        throw 'Autologon registry values did not pass final verification.'
    }
    Write-Log "Autologon configuration verified for '$DefaultDomainName\$DefaultUserName'." 'OK'
}

function Remove-LegacyChromeStartupShortcut {
    $script:LegacyShortcutPresentBefore = Test-Path -LiteralPath $LegacyChromeShortcutPath -PathType Leaf
    Write-Log 'Checking for the legacy All Users Google Chrome startup shortcut.'

    if ($script:LegacyShortcutPresentBefore) {
        Remove-Item -LiteralPath $LegacyChromeShortcutPath -Force -ErrorAction Stop
        $script:ChangesMade++
        $script:LegacyShortcutRemoved = $true
        Write-Log "Removed legacy Chrome startup shortcut: $LegacyChromeShortcutPath" 'OK'
    }
    else {
        Write-Log 'Legacy Google Chrome startup shortcut is not present.'
    }

    $script:LegacyShortcutPresentAfter = Test-Path -LiteralPath $LegacyChromeShortcutPath -PathType Leaf
    if ($script:LegacyShortcutPresentAfter) {
        $script:VerificationFailures.Add('Legacy Chrome startup shortcut remains present.')
        throw "The legacy Chrome startup shortcut could not be removed: $LegacyChromeShortcutPath"
    }
}

function Get-EdgePath {
    $candidatePaths = @()
    if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
        $candidatePaths += (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe')
    }
    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
        $candidatePaths += (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    }

    return $candidatePaths |
        Select-Object -Unique |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1
}

function Get-EdgeRunState {
    $configuredCommand = Get-RegistryStringValue -Path $RunKeyPath -Name $EdgeRunValueName
    $edgePath = Get-EdgePath
    $edgeVersion = $null
    if ($edgePath) {
        try { $edgeVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($edgePath).ProductVersion } catch { }
    }
    [pscustomobject]@{
        EdgeInstalled      = [bool]$edgePath
        EdgePath           = $edgePath
        EdgeVersion        = $edgeVersion
        RunValuePresent    = (-not [string]::IsNullOrWhiteSpace($configuredCommand))
        RunCommand         = $configuredCommand
        CommandMatches     = ($null -ne $script:EdgeCommand -and $configuredCommand -ceq $script:EdgeCommand)
        UrlPresent         = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand.Contains($EdgeUrl))
        InPrivateEnabled   = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand -match '(?i)--inprivate(?:\s|$)')
        NewWindowEnabled   = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand -match '(?i)--new-window(?:\s|$)')
        MaximizedEnabled   = (-not [string]::IsNullOrWhiteSpace($configuredCommand) -and $configuredCommand -match '(?i)--start-maximized(?:\s|$)')
    }
}

function Write-ConfigurationStateSummary {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)]$State
    )

    Write-Log ("Autologon {0}: Enabled={1}; User={2}; Domain={3}; ForceAutoLogon={4}; PasswordPresent={5}; PasswordMatches={6}; ConfigurationMatches={7}." -f `
        $Label, $State.Autologon.AutoAdminLogonEnabled, $State.Autologon.DefaultUserName, `
        $State.Autologon.DefaultDomainName, $State.Autologon.ForceAutoLogonEnabled, `
        $State.Autologon.PasswordPresent, $State.Autologon.PasswordMatches, $State.Autologon.ConfigurationMatches)

    Write-Log ("Edge startup {0}: Installed={1}; Version={2}; RunValuePresent={3}; CommandMatches={4}; UrlPresent={5}; InPrivate={6}; NewWindow={7}; Maximized={8}." -f `
        $Label, $State.Edge.EdgeInstalled, $State.Edge.EdgeVersion, $State.Edge.RunValuePresent, `
        $State.Edge.CommandMatches, $State.Edge.UrlPresent, $State.Edge.InPrivateEnabled, `
        $State.Edge.NewWindowEnabled, $State.Edge.MaximizedEnabled)
}

function Set-EdgeAutoLaunch {
    $script:EdgePath = Get-EdgePath
    if (-not $script:EdgePath) {
        throw 'Microsoft Edge was not found in either Program Files location.'
    }

    $escapedUrl = $EdgeUrl.Replace('"','')
    $script:EdgeCommand = '"{0}" --inprivate --new-window --start-maximized "{1}"' -f $script:EdgePath, $escapedUrl
    Set-VerifiedRegistryString -Path $RunKeyPath -Name $EdgeRunValueName -Value $script:EdgeCommand

    $edgeState = Get-EdgeRunState
    if (-not $edgeState.CommandMatches -or -not $edgeState.EdgeInstalled) {
        $script:VerificationFailures.Add('Edge auto-launch configuration verification failed.')
        throw 'Edge auto-launch configuration did not pass final verification.'
    }

    Write-Log "Configured and verified Edge auto-launch for every user: $EdgeUrl" 'OK'
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

function New-ObjectArrayForJson {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    [object[]]$items = @(
        $InputObject |
        ForEach-Object { $_ }
    )

    if ($items.Count -eq 0) {
        return ,([object[]]@())
    }

    return ,([object[]]$items)
}

function Write-Telemetry {
    param([Parameter(Mandatory)][object]$WindowsInfo)

    try {
        Initialize-LogDirectory
        $endTime = Get-Date
        $duration = [math]::Round(($endTime - $StartTime).TotalSeconds, 3)

        $event = [ordered]@{
            '@timestamp'             = $endTime.ToUniversalTime().ToString('o')
            EventType                = 'maintenance.execution'
            ComputerName             = $ComputerName
            Domain                   = $DomainName
            ScriptName               = $ScriptName
            ScriptVersion            = $ScriptVersion
            RunId                    = $RunId
            StartTime                = $StartTime.ToUniversalTime().ToString('o')
            EndTime                  = $endTime.ToUniversalTime().ToString('o')
            DurationSeconds          = $duration
            Status                   = $script:FinalStatus
            OverallResult            = $script:OverallResult
            ExitCode                 = $script:ExitCode
            ErrorCount               = $script:ErrorCount
            WarningCount             = $script:WarningCount
            FailureMessage           = $script:FailureMessage
            TextLogPath              = $PublishedLogPath
            RunningAccount           = $RunningAccount
            RunningAsSystem          = ($RunningAccount -eq 'NT AUTHORITY\SYSTEM')
            IsAdministrator          = (Test-IsAdministrator)
            Targeted                 = $script:Targeted
            MatchedPattern           = $script:MatchedPattern
            ConfiguredPatterns       = New-StringArrayForJson -InputObject $ComputerNamePatterns
            ChangesMade              = $script:ChangesMade
            VerificationFailureCount = $script:VerificationFailures.Count
            VerificationFailures     = New-StringArrayForJson -InputObject $script:VerificationFailures
            Windows                  = $WindowsInfo
            Configuration            = [ordered]@{
                DefaultUserName   = $DefaultUserName
                DefaultDomainName = $DefaultDomainName
                PasswordProvided  = (-not [string]::IsNullOrWhiteSpace($DefaultPasswordBase64))
                PasswordStorage   = 'Base64ObfuscatedInScript'
                PasswordValueLogged = $false
                EdgeUrl           = $EdgeUrl
                EdgeRunValueName  = $EdgeRunValueName
            }
            Before                   = $script:BeforeState
            After                    = $script:AfterState
            RegistryChanges          = New-ObjectArrayForJson -InputObject $script:RegistryChanges
            LegacyChromeShortcut     = [ordered]@{
                Path          = $LegacyChromeShortcutPath
                PresentBefore = $script:LegacyShortcutPresentBefore
                Removed       = $script:LegacyShortcutRemoved
                PresentAfter  = $script:LegacyShortcutPresentAfter
            }
            Edge                     = [ordered]@{
                ExecutablePath = $script:EdgePath
                RunCommand     = $script:EdgeCommand
            }
        }

        try {
            $autologonVerified = $false
            $edgeVerified = $false
            if ($event.After) {
                if ($event.After.Autologon) { $autologonVerified = [bool]$event.After.Autologon.ConfigurationMatches }
                if ($event.After.Edge) { $edgeVerified = [bool]$event.After.Edge.CommandMatches }
            }

            Write-Log ("Elastic configuration summary: Targeted={0}; MatchedPattern={1}; ChangesMade={2}; VerificationFailures={3}; AutologonVerified={4}; EdgeVerified={5}; LegacyChromeRemoved={6}; Result={7}" -f `
                $event.Targeted,
                $event.MatchedPattern,
                $event.ChangesMade,
                $event.VerificationFailureCount,
                $autologonVerified,
                $edgeVerified,
                $event.LegacyChromeShortcut.Removed,
                $event.OverallResult) 'INFO'
        }
        catch { }

        $compactJson = $event | ConvertTo-Json -Depth 10 -Compress
        Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $compactJson

        $prettyJson = $event | ConvertTo-Json -Depth 10
        $tempPath = "$LatestTelemetryPath.tmp.$RunId"
        Set-Content -LiteralPath $tempPath -Value $prettyJson -Encoding UTF8
        Move-Item -LiteralPath $tempPath -Destination $LatestTelemetryPath -Force
    }
    catch {
        Write-Warning "Unable to write structured telemetry: $($_.Exception.Message)"
    }
}

$windowsInfo = Get-WindowsInformation

try {
    Initialize-LogDirectory
    Write-Log "===== Autologon and Edge configuration v$ScriptVersion started ====="
    Write-Log "Text log: $LogPath"
    Write-Log "Computer name: $ComputerName"
    Write-Log "Configured computer-name patterns: $($ComputerNamePatterns -join ', ')"

    if (-not (Test-IsAdministrator)) {
        throw 'Please run this script as Administrator or through a SYSTEM scheduled task.'
    }

    Test-ConfigurationInput

    $script:MatchedPattern = Get-MatchedComputerPattern -Name $ComputerName -Patterns $ComputerNamePatterns
    $script:Targeted = (-not [string]::IsNullOrWhiteSpace($script:MatchedPattern))

    if (-not $script:Targeted) {
        $script:BeforeState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
        $script:AfterState = $script:BeforeState
        Write-ConfigurationStateSummary -Label 'unchanged (not targeted)' -State $script:BeforeState
        $script:OverallResult = 'NotTargeted'
        $script:FinalStatus = 'Success'
        Write-Log "Computer '$ComputerName' does not match the configured pattern list. No changes were made."
        Write-Log '===== Script completed: computer not targeted =====' 'OK'
    }
    else {
        Write-Log "Computer '$ComputerName' matches pattern '$($script:MatchedPattern)'." 'OK'

        $script:DecodedPassword = Get-DecodedAutologonPassword

        $script:BeforeState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
        Write-ConfigurationStateSummary -Label 'before' -State $script:BeforeState

        Set-AutologonConfiguration
        Remove-LegacyChromeStartupShortcut
        Set-EdgeAutoLaunch

        $script:AfterState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
        Write-ConfigurationStateSummary -Label 'after' -State $script:AfterState
        Write-Log ("Legacy Chrome startup shortcut: PresentBefore={0}; Removed={1}; PresentAfter={2}." -f `
            $script:LegacyShortcutPresentBefore, $script:LegacyShortcutRemoved, $script:LegacyShortcutPresentAfter)

        if (-not $script:AfterState.Autologon.ConfigurationMatches) {
            $script:VerificationFailures.Add('Final autologon configuration does not match the requested values.')
        }
        if (-not $script:AfterState.Edge.CommandMatches) {
            $script:VerificationFailures.Add('Final Edge Run value does not match the requested command.')
        }
        if ($script:LegacyShortcutPresentAfter) {
            $script:VerificationFailures.Add('Legacy Chrome startup shortcut remains present.')
        }

        if ($script:VerificationFailures.Count -gt 0) {
            throw "One or more final verification checks failed: $($script:VerificationFailures -join '; ')"
        }

        if ($script:ChangesMade -eq 0) {
            $script:OverallResult = 'AlreadyCompliant'
        }
        else {
            $script:OverallResult = 'ConfiguredAndVerified'
        }
        $script:FinalStatus = 'Success'
        Write-Log 'Autologon, legacy Chrome cleanup, and Edge startup settings were verified successfully.' 'OK'
        Write-Log '===== Script completed successfully =====' 'OK'
    }
}
catch {
    $script:ExitCode = 1
    $script:FinalStatus = 'Failed'
    $script:OverallResult = 'Failed'
    $script:FailureMessage = $_.Exception.Message
    Write-Log "Configuration failed: $($_.Exception.Message)" 'ERROR'

    try {
        $script:AfterState = [pscustomobject]@{
            Autologon = Get-AutologonState
            Edge      = Get-EdgeRunState
        }
    }
    catch {
        Write-Log "Unable to collect final configuration state: $($_.Exception.Message)" 'WARN'
    }

    Write-Log '===== Script completed with errors =====' 'ERROR'
}
finally {
    # Clear the decoded credential before final telemetry/log publication.
    $script:DecodedPassword = $null
    [GC]::Collect()

    try {
        Write-Telemetry -WindowsInfo $windowsInfo
    }
    catch {
        Write-Log "Telemetry write failed: $($_.Exception.Message)" 'ERROR'
        if ($script:ExitCode -eq 0) {
            $script:FinalStatus = 'TelemetryFailure'
            $script:OverallResult = 'TelemetryFailure'
            $script:ExitCode = 4
        }
    }

    # Final text-log append before the immutable completed file enters C:\Logs.
    Write-Log ("Finalizing {0}. Status={1}; Result={2}; ExitCode={3}; Warnings={4}; Errors={5}" -f `
        $ScriptName,
        $script:FinalStatus,
        $script:OverallResult,
        $script:ExitCode,
        $script:WarningCount,
        $script:ErrorCount) $(if ($script:ExitCode -eq 0) { 'OK' } else { 'ERROR' })

    if ($null -ne $LogSession) {
        $publishResult = Publish-MaintenanceLog -LogSession $LogSession

        if ($publishResult.Published) {
            Write-Host ("Published completed script 13 text log for Elastic: {0}" -f $PublishedLogPath) -ForegroundColor Green
        }
        else {
            Write-Warning ("Script 13 completed text log remains in staging because publication failed: {0}" -f $publishResult.Path)
        }
    }
}

exit $script:ExitCode
