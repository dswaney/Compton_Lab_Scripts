#requires -Version 5.1
#requires -RunAsAdministrator
# ScriptName: Show-RecoveryNotice.ps1
# ScriptVersion: 1.5.4
# Startup-only display helper. Never runs repairs, changes DFC state, or reboots.
[CmdletBinding()]
param([string]$ConfigPath=(Join-Path $PSScriptRoot 'RecoveryNotice.config.json'))
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$incident=$null
$noticeResult=$null
function Write-NoticeReport {
    param([string]$Status,[string]$Message,[string]$Severity='Info')
    $event=[ordered]@{
        '@timestamp'=[datetime]::UtcNow.ToString('o');EventType='endpoint.reboot_recovery.notice'
        ComputerName=$env:COMPUTERNAME;ScriptName='Show-RecoveryNotice.ps1';ScriptVersion='1.5.4'
        RunId=[guid]::NewGuid().ToString();IncidentId=$incident;Status=$Status;Severity=$Severity
        Message=$Message;DisplayResult=$noticeResult
    }
    try {
        $root=Join-Path $env:SystemDrive 'Logs'
        New-Item -Path $root -ItemType Directory -Force | Out-Null
        $json=$event | ConvertTo-Json -Depth 6 -Compress
        [IO.File]::WriteAllText((Join-Path $root 'RecoveryNotice.latest.json'),$json,(New-Object Text.UTF8Encoding($false)))
        # Display first; framework loading and telemetry publication come AFTER it.
        $framework=Join-Path $env:SystemDrive 'Scripts\Maintenance.Framework.psm1'
        if (Test-Path -LiteralPath $framework -PathType Leaf) {
            Import-Module -Name $framework -Force -DisableNameChecking -ErrorAction Stop
            Write-MaintenanceTelemetryLine -Path (Join-Path $root 'Maintenance-Telemetry.ndjson') -JsonLine $json
        }
    }
    catch {Write-Warning ('Notice diagnostic publication failed: '+$_.Exception.Message)}
    Write-Host ($Status+': '+$Message)
}
try {
    $config=[IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
    if ($config.SchemaVersion -ne 1 -or $config.ComputerName -ne $env:COMPUTERNAME) {throw 'Notice configuration identity is invalid.'}
    if (-not $config.Enabled) {exit 0}
    $stateRoot=[string]$config.PersistentStateRoot
    $stateUri=$null
    if (-not $stateRoot.StartsWith('\\') -or -not [Uri]::TryCreate($stateRoot,[UriKind]::Absolute,[ref]$stateUri)) {throw 'Notice state root must be a valid UNC path.'}
    if (-not $stateUri.IsUnc -or $stateUri.Segments.Count -lt 2) {throw 'Notice state root must include a server and share.'}
    if (-not [IO.Path]::IsPathRooted([string]$config.DfcPath) -or -not (Test-Path -LiteralPath $config.DfcPath -PathType Leaf)) {throw 'Configured DFC executable is unavailable.'}
    $native=New-Object Diagnostics.Process
    try {
        $native.StartInfo.FileName=[string]$config.DfcPath
        $native.StartInfo.Arguments='get /ISFROZEN'
        $native.StartInfo.UseShellExecute=$false
        $native.StartInfo.CreateNoWindow=$true
        if (-not $native.Start()) {throw 'DFC state query did not start.'}
        if (-not $native.WaitForExit(5000)) {try {$native.Kill()} catch {};throw 'Notice DFC query exceeded five seconds.'}
        $code=$native.ExitCode
        if ($null -eq $code) {throw 'Notice DFC query returned no exit code.'}
    }
    finally {$native.Dispose()}
    if ($code -eq 1) {Write-NoticeReport 'NoticeSkippedFrozen' 'Computer is frozen; no display settings changed.';exit 0}
    if ($code -ne 0) {throw ('Deep Freeze state is unknown: '+$code)}
    $machineFolder=Join-Path ([string]$config.PersistentStateRoot) $env:COMPUTERNAME
    $statePath=Join-Path $machineFolder 'RecoveryState.json'
    # Retry network availability without running detection/collection/maintenance.
    # The task also has a two-minute execution limit for blocked SMB operations.
    $clock=[Diagnostics.Stopwatch]::StartNew()
    $state=$null
    do {
        try {
            if (Test-Path -LiteralPath $statePath -PathType Leaf) {
                $accountDomain='COMPTON'
                if ($null -ne $config.PSObject.Properties['PersistentStateAccountDomain']) {$accountDomain=[string]$config.PersistentStateAccountDomain}
                $account=New-Object Security.Principal.NTAccount($accountDomain,($env:COMPUTERNAME+'$'))
                $expectedOwner=$account.Translate([Security.Principal.SecurityIdentifier]).Value
                $actualOwner=(Get-Acl -LiteralPath $machineFolder -ErrorAction Stop).GetOwner([Security.Principal.SecurityIdentifier]).Value
                if ($actualOwner -ne $expectedOwner) {throw 'Persistent folder owner mismatch; notice receipt is not trusted.'}
                $state=[IO.File]::ReadAllText($statePath) | ConvertFrom-Json
                break
            }
        }
        catch { }
        if ($clock.Elapsed.TotalSeconds -ge 45) {break}
        Start-Sleep -Seconds 2
    } while ($true)
    if ($null -eq $state) {Write-NoticeReport 'NoticeNoReceipt' 'No accessible active receipt within the startup retry window; watchdog display remains the fallback.';exit 0}
    if ($state.SchemaVersion -ne 1 -or $state.ComputerName -ne $env:COMPUTERNAME) {throw 'Recovery receipt identity is invalid.'}
    $incident=([guid]::Parse([string]$state.IncidentId)).ToString()
    if ($state.Phase -notin @('ThawRequested','AfterThaw','UpdatingScripts','AwaitingEvaluation','ClosingServices','ForceRequested')) {
        Write-NoticeReport 'NoticeSkippedInactive' ('No active repair display required for phase '+$state.Phase);exit 0
    }
    $preview=Join-Path $machineFolder 'RepairScreenPreview\Context.json'
    if (Test-Path -LiteralPath $preview -PathType Leaf) {
        $context=[IO.File]::ReadAllText($preview) | ConvertFrom-Json
        if ($context.ComputerName -ne $env:COMPUTERNAME -or $context.Phase -ne 'Restored') {throw 'An unrestored display preview is present.'}
    }
    Import-Module -Name (Join-Path $PSScriptRoot 'RepairScreen.psm1') -Force -DisableNameChecking
    $noticeResult=Enable-ComptonRepairScreen -IncidentFolder (Join-Path $machineFolder $incident) -IncidentId $incident -ComputerName $env:COMPUTERNAME -Title ([string]$config.Title) -Message ([string]$config.Message)
    if ($noticeResult.Phase -eq 'Restored') {Write-NoticeReport 'NoticeAlreadyRestored' 'Final display restoration already completed; notice was not reapplied.';exit 0}
    Write-NoticeReport 'StartupNoticeConfigured' 'Repair display requested by the dedicated startup task. Windows controls sign-in UI refresh.'
    exit 0
}
catch {
    Write-NoticeReport 'StartupNoticeFailed' $_.Exception.Message 'Warning'
    exit 2
}
