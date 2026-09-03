# ScriptName: 17_Set_Browser_Homepage.ps1
# ScriptVersion: 1.0.1
# LastUpdated: 2026-08-27
# Changes: v1.0.1 updates the internal script identity, log names, and telemetry
#          dataset after renaming the file from script 19 to script 17.
# Purpose: Configure Firefox, Google Chrome, and Microsoft Edge to use the
#          Compton College website as the machine-wide homepage and startup
#          page. Disable Chrome browser/profile sign-in prompts and sync while
#          preserving normal sign-in to websites.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [ValidatePattern('^https://')]
    [string]$HomepageUrl = 'https://www.compton.edu',

    [string]$LogFolder = 'C:\Logs'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptName = '17_Set_Browser_Homepage.ps1'
$script:ScriptVersion = '1.0.1'
$script:ComputerName = $env:COMPUTERNAME
$script:StartTime = Get-Date
$script:RunId = [guid]::NewGuid().Guid
$script:ChangedCount = 0
$script:VerifiedCount = 0
$script:Status = 'failed'
$script:ExitCode = 1
$script:LogPath = Join-Path $LogFolder '17_Set_Browser_Homepage.log'
$script:LatestTelemetryPath = Join-Path $LogFolder '17_Set_Browser_Homepage.latest.json'
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
    Write-Host $line -ForegroundColor $(switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        default   { 'Cyan' }
    })
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'This script must run elevated or as Local System.'
    }

    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        throw 'Run this script with 64-bit PowerShell so browser policies are written to the correct registry view.'
    }
}

function Set-PolicyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Browser,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [Parameter(Mandatory)][ValidateSet('String','DWord')][string]$Type
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    $currentValue = $null
    $valueExists = $false
    try {
        $currentValue = Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop
        $valueExists = $true
    }
    catch [System.Management.Automation.PSArgumentException] {}
    catch [System.Management.Automation.ItemNotFoundException] {}

    $expectedValue = if ($Type -eq 'DWord') { [int]$Value } else { [string]$Value }
    $alreadyCorrect = $valueExists -and ([string]$currentValue -eq [string]$expectedValue)

    if (-not $alreadyCorrect) {
        New-ItemProperty -LiteralPath $Path -Name $Name -Value $expectedValue -PropertyType $Type -Force | Out-Null
        $script:ChangedCount++
        Write-Log ("{0}: Set {1} = {2}." -f $Browser, $Name, $expectedValue) 'INFO'
    }
    else {
        Write-Log ("{0}: {1} is already correctly configured." -f $Browser, $Name) 'INFO'
    }

    $verifiedValue = Get-ItemPropertyValue -LiteralPath $Path -Name $Name -ErrorAction Stop
    if ([string]$verifiedValue -ne [string]$expectedValue) {
        throw ("{0}: Verification failed for {1}. Expected '{2}', found '{3}'." -f $Browser, $Name, $expectedValue, $verifiedValue)
    }

    $script:VerifiedCount++
}

function Remove-UnwantedPolicyValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Browser,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$AllowedNames
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    $properties = @(
        (Get-ItemProperty -LiteralPath $Path -ErrorAction Stop).PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' }
    )

    foreach ($property in $properties) {
        if ($property.Name -notin $AllowedNames) {
            Remove-ItemProperty -LiteralPath $Path -Name $property.Name -Force -ErrorAction Stop
            $script:ChangedCount++
            Write-Log ("{0}: Removed unwanted startup URL policy entry '{1}'." -f $Browser, $property.Name) 'INFO'
        }
    }
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
    $event = [pscustomobject][ordered]@{
        EventType       = 'maintenance.execution'
        EventDataset    = 'compton.maintenance.script17'
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
        HomepageUrl     = $HomepageUrl
        ChangedCount    = $script:ChangedCount
        VerifiedCount   = $script:VerifiedCount
        ChromeSignin    = 'disabled'
        LogPath         = $script:LogPath
    }

    $json = $event | ConvertTo-Json -Depth 5 -Compress
    Add-SharedTextLine -Path $script:TelemetryPath -Line $json
    $event | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:LatestTelemetryPath -Encoding UTF8 -Force
}

try {
    Assert-Administrator
    Write-Log ("Starting browser homepage configuration. Version={0}; Homepage={1}; RunId={2}" -f $script:ScriptVersion, $HomepageUrl, $script:RunId) 'INFO'

    if (-not $PSCmdlet.ShouldProcess($script:ComputerName, "Set Firefox, Chrome, and Edge homepage policies to $HomepageUrl")) {
        $script:Status = 'what_if'
        $script:ExitCode = 0
        return
    }

    # Google Chrome machine policies.
    $chromePath = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    $chromeStartupPath = Join-Path $chromePath 'RestoreOnStartupURLs'
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'HomepageLocation' -Value $HomepageUrl -Type String
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'HomepageIsNewTabPage' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'ShowHomeButton' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'RestoreOnStartup' -Value 4 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromeStartupPath -Name '1' -Value $HomepageUrl -Type String
    Remove-UnwantedPolicyValues -Browser 'Chrome' -Path $chromeStartupPath -AllowedNames @('1')

    # Disable Chrome browser/profile sign-in and sync prompts. This does not
    # block users from signing in to websites in Chrome.
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'BrowserSignin' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'SyncDisabled' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'SigninInterceptionEnabled' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Chrome' -Path $chromePath -Name 'PromotionalTabsEnabled' -Value 0 -Type DWord

    # Microsoft Edge machine policies.
    $edgePath = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    $edgeStartupPath = Join-Path $edgePath 'RestoreOnStartupURLs'
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'HomepageLocation' -Value $HomepageUrl -Type String
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'HomepageIsNewTabPage' -Value 0 -Type DWord
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'ShowHomeButton' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Edge' -Path $edgePath -Name 'RestoreOnStartup' -Value 4 -Type DWord
    Set-PolicyValue -Browser 'Edge' -Path $edgeStartupPath -Name '1' -Value $HomepageUrl -Type String
    Remove-UnwantedPolicyValues -Browser 'Edge' -Path $edgeStartupPath -AllowedNames @('1')

    # Mozilla Firefox machine policies.
    $firefoxHomepagePath = 'HKLM:\SOFTWARE\Policies\Mozilla\Firefox\Homepage'
    Set-PolicyValue -Browser 'Firefox' -Path $firefoxHomepagePath -Name 'URL' -Value $HomepageUrl -Type String
    Set-PolicyValue -Browser 'Firefox' -Path $firefoxHomepagePath -Name 'Locked' -Value 1 -Type DWord
    Set-PolicyValue -Browser 'Firefox' -Path $firefoxHomepagePath -Name 'StartPage' -Value 'homepage' -Type String

    $script:Status = 'success'
    $script:ExitCode = 0
    Write-Log ("Browser policies completed successfully. Changed={0}; Verified={1}. Close and reopen each browser to apply the policies." -f $script:ChangedCount, $script:VerifiedCount) 'SUCCESS'
}
catch {
    $script:Status = 'failed'
    $script:ExitCode = 1
    try { Write-Log ("Browser policy configuration failed: {0}" -f $_.Exception.Message) 'ERROR' }
    catch { Write-Error $_.Exception.Message }
}
finally {
    try { Write-Telemetry }
    catch { try { Write-Log ("Telemetry write failed: {0}" -f $_.Exception.Message) 'WARNING' } catch {} }
}

exit $script:ExitCode
