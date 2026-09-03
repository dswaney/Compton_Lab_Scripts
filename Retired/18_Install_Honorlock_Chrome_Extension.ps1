# ScriptName: 18_Install_Honorlock_Chrome_Extension.ps1
# ScriptVersion: 1.1.2
# LastUpdated: 2026-08-27
# Changes: v1.1.2 updates the internal script identity, log names, and telemetry
#          dataset after renaming the file from script 20 to script 18.
#          v1.1.1 corrects the computer-name pattern array that caused a
#          PowerShell parser error.
# Purpose: Force-install the Honorlock Chrome extension machine-wide through
#          Google Chrome enterprise policy without overwriting other forced
#          extension entries. Deployment can be limited to configured computer
#          name prefixes or wildcard patterns.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [string]$LogFolder = 'C:\Logs',

    # Leave empty to target every computer. Examples:
    # @('IB1-103-*')
    # @('IB1-103-*','SSB-122-*','SSB-114*')
    [string[]]$ComputerNamePatterns = @(
        'SSC-216*',
        'AHB-146*'
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptName = '18_Install_Honorlock_Chrome_Extension.ps1'
$script:ScriptVersion = '1.1.2'
$script:ComputerName = $env:COMPUTERNAME
$script:StartTime = Get-Date
$script:RunId = [guid]::NewGuid().Guid
$script:Status = 'failed'
$script:ExitCode = 1
$script:PolicyChanged = $false
$script:PolicyVerified = $false
$script:PolicyValueName = $null
$script:Targeted = $false
$script:MatchedComputerPattern = $null
$script:OverallResult = 'NotStarted'

$script:HonorlockExtensionId = 'hnbmpkmhjackfpkpcbapafmpepgmmddc'
$script:ChromeUpdateUrl = 'https://clients2.google.com/service/update2/crx'
$script:ForceInstallValue = '{0};{1}' -f $script:HonorlockExtensionId, $script:ChromeUpdateUrl
$script:ForceListPath = 'HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist'

$script:LogPath = Join-Path $LogFolder '18_Install_Honorlock_Chrome_Extension.log'
$script:LatestTelemetryPath = Join-Path $LogFolder '18_Install_Honorlock_Chrome_Extension.latest.json'
$script:TelemetryPath = Join-Path $LogFolder 'Maintenance-Telemetry.ndjson'

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level = 'INFO'
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '{0} [{1}] [{2}] {3}' -f $timestamp, $script:ComputerName, $Level, $Message
    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        default   { 'Cyan' }
    }

    Write-Host $line -ForegroundColor $color
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'This script must run elevated or as Local System.'
    }

    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        throw 'Run this script with 64-bit PowerShell so the Chrome policy is written to the correct registry view.'
    }
}

function Get-MatchedComputerPattern {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowEmptyCollection()][string[]]$Patterns
    )

    $configuredPatterns = @(
        $Patterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    # An empty list intentionally targets all computers.
    if ($configuredPatterns.Count -eq 0) {
        return '*'
    }

    foreach ($pattern in $configuredPatterns) {
        if ($Name -like $pattern) {
            return [string]$pattern
        }
    }

    return $null
}

function Get-ForceListProperties {
    if (-not (Test-Path -LiteralPath $script:ForceListPath)) {
        return @()
    }

    return @(
        (Get-ItemProperty -LiteralPath $script:ForceListPath -ErrorAction Stop).PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' }
    )
}

function Set-HonorlockForceInstallPolicy {
    if (-not (Test-Path -LiteralPath $script:ForceListPath)) {
        New-Item -Path $script:ForceListPath -Force | Out-Null
        Write-Log ("Created Chrome force-install policy path: {0}" -f $script:ForceListPath) 'INFO'
    }

    $properties = @(Get-ForceListProperties)
    $extensionPattern = '^{0}(;|$)' -f [regex]::Escape($script:HonorlockExtensionId)
    $existingProperty = @(
        $properties | Where-Object { [string]$_.Value -match $extensionPattern }
    ) | Select-Object -First 1

    if ($null -ne $existingProperty) {
        $script:PolicyValueName = [string]$existingProperty.Name
        if ([string]$existingProperty.Value -ne $script:ForceInstallValue) {
            New-ItemProperty -LiteralPath $script:ForceListPath -Name $script:PolicyValueName `
                -Value $script:ForceInstallValue -PropertyType String -Force | Out-Null
            $script:PolicyChanged = $true
            Write-Log ("Updated existing Honorlock force-install entry '{0}'." -f $script:PolicyValueName) 'INFO'
        }
        else {
            Write-Log ("Honorlock is already present in Chrome's force-install policy as entry '{0}'." -f $script:PolicyValueName) 'INFO'
        }
    }
    else {
        $usedNumbers = @(
            $properties |
            Where-Object { $_.Name -match '^\d+$' } |
            ForEach-Object { [int]$_.Name }
        )

        $nextNumber = 1
        while ($usedNumbers -contains $nextNumber) {
            $nextNumber++
        }

        $script:PolicyValueName = [string]$nextNumber
        New-ItemProperty -LiteralPath $script:ForceListPath -Name $script:PolicyValueName `
            -Value $script:ForceInstallValue -PropertyType String -Force | Out-Null
        $script:PolicyChanged = $true
        Write-Log ("Added Honorlock to Chrome's force-install policy as entry '{0}'." -f $script:PolicyValueName) 'INFO'
    }

    $verifiedProperties = @(Get-ForceListProperties)
    $verified = @(
        $verifiedProperties | Where-Object { [string]$_.Value -eq $script:ForceInstallValue }
    )

    if ($verified.Count -lt 1) {
        throw 'Honorlock force-install policy verification failed after the registry update.'
    }

    $script:PolicyVerified = $true
    Write-Log ("Verified Honorlock force-install policy. ExtensionId={0}; Entry={1}" -f $script:HonorlockExtensionId, $script:PolicyValueName) 'SUCCESS'
}

function Add-SharedTextLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Line,
        [int]$MaximumAttempts = 10
    )

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        $stream = $null
        $writer = $null
        try {
            $stream = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
            $writer.WriteLine($Line)
            $writer.Flush()
            return
        }
        catch [System.IO.IOException] {
            if ($attempt -ge $MaximumAttempts) { throw }
            Start-Sleep -Milliseconds (250 * $attempt)
        }
        finally {
            if ($null -ne $writer) { $writer.Dispose() }
            elseif ($null -ne $stream) { $stream.Dispose() }
        }
    }
}

function Write-Telemetry {
    $endTime = Get-Date
    $installState = switch ($script:OverallResult) {
        'PolicyConfigured' { 'Policy configured; Chrome installs/updates the extension at policy refresh or browser startup.' }
        'NotTargeted'      { 'Skipped because the computer name did not match the configured pattern list.' }
        'WhatIf'           { 'WhatIf evaluation only; no policy changes were made.' }
        'Failed'           { 'Policy deployment failed. Review the text log.' }
        default            { 'No completed deployment result was recorded.' }
    }

    $event = [pscustomobject][ordered]@{
        EventType       = 'maintenance.execution'
        EventDataset    = 'compton.maintenance.script18'
        ComputerName    = $script:ComputerName
        ScriptName      = $script:ScriptName
        ScriptVersion   = $script:ScriptVersion
        RunId           = $script:RunId
        Status          = $script:Status
        ExitCode        = $script:ExitCode
        StartTime       = $script:StartTime.ToString('o')
        EndTime         = $endTime.ToString('o')
        DurationSeconds = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 2)
        Timestamp       = $endTime.ToUniversalTime().ToString('o')
        Browser         = 'Google Chrome'
        ExtensionName   = 'Honorlock'
        ExtensionId     = $script:HonorlockExtensionId
        PolicyChanged   = $script:PolicyChanged
        PolicyVerified  = $script:PolicyVerified
        PolicyValueName = $script:PolicyValueName
        ComputerNamePatterns = @($ComputerNamePatterns)
        Targeted        = $script:Targeted
        MatchedPattern  = $script:MatchedComputerPattern
        OverallResult   = $script:OverallResult
        InstallState    = $installState
        LogPath         = $script:LogPath
    }

    $json = $event | ConvertTo-Json -Depth 5 -Compress
    Add-SharedTextLine -Path $script:TelemetryPath -Line $json
    $event | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:LatestTelemetryPath -Encoding UTF8 -Force
}

try {
    Write-Log ("Starting Honorlock Chrome extension policy deployment. Version={0}; RunId={1}" -f $script:ScriptVersion, $script:RunId) 'INFO'

    $script:MatchedComputerPattern = Get-MatchedComputerPattern `
        -Name $script:ComputerName `
        -Patterns $ComputerNamePatterns

    if ([string]::IsNullOrWhiteSpace($script:MatchedComputerPattern)) {
        $script:Targeted = $false
        $script:OverallResult = 'NotTargeted'
        $script:Status = 'success'
        $script:ExitCode = 0
        $configuredText = if (@($ComputerNamePatterns).Count -gt 0) {
            @($ComputerNamePatterns) -join ', '
        }
        else {
            '(none)'
        }
        Write-Log ("Computer '{0}' does not match the configured pattern list [{1}]. No changes were made." -f $script:ComputerName, $configuredText) 'INFO'
        return
    }

    $script:Targeted = $true
    Write-Log ("Computer is targeted. Matched pattern: {0}" -f $script:MatchedComputerPattern) 'SUCCESS'
    Assert-Administrator

    if (-not $PSCmdlet.ShouldProcess($script:ComputerName, 'Force-install the Honorlock extension in Google Chrome for all users')) {
        $script:Status = 'what_if'
        $script:ExitCode = 0
        $script:OverallResult = 'WhatIf'
        return
    }

    Set-HonorlockForceInstallPolicy

    $script:Status = 'success'
    $script:ExitCode = 0
    $script:OverallResult = 'PolicyConfigured'
    Write-Log 'Honorlock policy deployment completed. Chrome will install or update the extension when policy refreshes or Chrome next starts.' 'SUCCESS'
}
catch {
    $script:Status = 'failed'
    $script:ExitCode = 1
    $script:OverallResult = 'Failed'
    try { Write-Log ("Honorlock policy deployment failed: {0}" -f $_.Exception.Message) 'ERROR' }
    catch { Write-Error $_.Exception.Message }
}
finally {
    try { Write-Telemetry }
    catch { try { Write-Log ("Telemetry write failed: {0}" -f $_.Exception.Message) 'WARNING' } catch {} }
}

exit $script:ExitCode
