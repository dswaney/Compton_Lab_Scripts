#requires -Version 5.1
<#
.SYNOPSIS
    Detects repeated Script 07 maintenance reboots; optionally performs bounded recovery.
.NOTES
    ScriptVersion: 1.5.4
    LastUpdated: 2026-10-29
    Default invocation is diagnostic-only. Install while THAWED.
    Do not share this file after entering your password.
#>
[CmdletBinding()]
param(
    [switch]$Deploy,
    [switch]$Install,
    [switch]$EnableRecovery,
    [switch]$Recover,
    [switch]$Force,
    [switch]$ValidateOnly,
    [switch]$AcknowledgeIncident,
    [switch]$Uninstall,
    [switch]$CollectEvidenceOnly,
    [string]$ExpectedIncidentId
)

# =========================== EDIT THESE SETTINGS ===========================
# Use single quotes. If the password contains an apostrophe, double it: abc''def
$DeepFreezeCommandLinePassword = ''

# REQUIRED for automatic recovery: a real, persistent UNC folder.
# Example ONLY: '\\filesvr\MaintenanceRecovery$'
# SYSTEM creates its private per-computer folder using the computer account.
# Configure the Samba self-provisioning ACLs before fleet deployment.
# Folder ownership must match this computer account before state is trusted.
$PersistentStateRoot = '\\SERVER\ComptonRebootRecovery$'
$PersistentStateAccountDomain = 'COMPTON'

$MinimumCompletedRestarts = 3
$LookbackMinutes = 30
$CooldownHours = 24
$BootSettleSeconds = 60
$AutomaticFreezeAfterVerifiedRepair = $true
$FinalFreezeRebootDelaySeconds = 10
$NormalSundayStart = '00:15:00'
$NormalSundayEnd = '21:45:00'
$RunUpdaterAfterThaw = $true
$EvidenceCollectionTimeoutSeconds = 120

# Temporary repair display; no license/Shared PC mode changes are made.
$EnableRepairScreen = $true
$RepairScreenTitle = 'IT REPAIR IN PROGRESS'
$RepairScreenMessage = 'Please wait while IT repairs this computer. Do not turn it off. For help, call ext. 2181.'

# Approved hashes from the actual scripts reviewed in this conversation.
# Add a new hash only after reviewing a replacement script's behavior.
$ApprovedUpdaterHashes = @('99888F0BB500925EDCDC2D6A1F08C1E0F923D791BE7CD2CFEE6FDB9276D076A8','A0848FD185FB81C33ED358FC59D18532B3B072652349D5E1F0504B3080A7AF04')
$ApprovedScript07Hashes = @('48A1CE608236264D02F5812550CE13D40C847615D108D147C335D0A3DAFA5B3F')
$ApprovedScript09Hashes = @('EF87737B620259BBCBD3A73526A567027BC00223C6EA6F133808B5C6A1AC42DA')
# ==========================================================================

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:Version = '1.5.4'
$script:Name = 'Watch-RebootRecovery.ps1'
$script:RunId = [guid]::NewGuid().ToString()
$script:Incident = $null
$script:LoopEvidence = $null
$script:EvidenceCount = 0
$script:EvidenceBytes = 0
$script:EvidenceTruncated = $false
$script:EvidenceWriter = $null
$script:State = $null
$script:StatePath = $null
$script:DomainName = [string]$env:USERDOMAIN
$script:TelemetryReady = $false
$script:Boot = $null
$script:PublishedStatus = $null
$script:Mutex = $null
$script:MutexHeld = $false
$LocalRoot = Join-Path $env:ProgramData 'Compton\RebootRecovery'
$InstalledScript = Join-Path $LocalRoot $script:Name
$LogRoot = 'C:\Logs'
$TelemetryPath = Join-Path $LogRoot 'Maintenance-Telemetry.ndjson'
$LatestPath = Join-Path $LogRoot 'RebootRecovery.latest.json'
$TaskName = 'Compton. Reboot Loop Recovery Watchdog'
$Pwsh = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$Updater = 'C:\Scripts\00_Update-Scripts-FromShare.ps1'
$Script07 = 'C:\Scripts\07_Force_Reboot_Install_Updates.ps1'
$Script09 = 'C:\Scripts\09_Disable_Windows_Update_Services.ps1'
$Framework = 'C:\Scripts\Maintenance.Framework.psm1'
$RepairScreenModule = Join-Path $PSScriptRoot 'RepairScreen.psm1'
$StartupNoticeTaskName='Compton. Recovery Notice at Startup'
$StartupNoticeSource=Join-Path $PSScriptRoot 'Show-RecoveryNotice.ps1'

function Test-NormalMaintenance {
    param([datetime]$Now = (Get-Date), [timespan]$Start = ([timespan]'00:15:00'), [timespan]$End = ([timespan]'21:45:00'))
    return ($Now.DayOfWeek -eq [DayOfWeek]::Sunday -and $Now.TimeOfDay -ge $Start -and $Now.TimeOfDay -lt $End)
}

function Get-LoopDecision {
    param([object[]]$Requests, [datetime[]]$Boots, [datetime]$Now, [int]$Minutes = 30, [int]$Threshold = 3)
    $start = $Now.AddMinutes(-$Minutes)
    $usedBoots = @{}
    $usedRequests = @{}
    $matched = @()
    foreach ($boot in @($Boots | Sort-Object)) {
        if ($boot -lt $start -or $boot -gt $Now) {continue}
        $key=$boot.ToUniversalTime().ToString('o')
        if ($usedBoots.ContainsKey($key)) {continue}
        $usedBoots[$key]=$true
        # The last shutdown request before this boot must be a known maintenance
        # request. An intervening user/Deep Freeze restart invalidates the match.
        $last=@($Requests | Where-Object {
            $_.TimeCreated -le $boot -and $_.TimeCreated -ge $boot.AddMinutes(-10)
        } | Sort-Object TimeCreated -Descending | Select-Object -First 1)
        if ($last.Count -eq 0) {continue}
        $request=$last[0]
        if ($request.Process -notmatch '(?i)(^|\\)shutdown\.exe(?:\s|\(|$)') {continue}
        if ($request.Comment -notin @(
            'Pending reboot flags detected during maintenance closeout.',
            'Reboot flags remain after initial reboot. Issuing second reboot.',
            'Reboot flags remain after two reboots. Issuing third and final automatic reboot.',
            'State reset occurred. Performing initial forced reboot for update cycle.'
        )) {continue}
        $requestKey=$request.TimeCreated.ToUniversalTime().ToString('o')+'|'+$request.Process+'|'+$request.Comment
        if ($usedRequests.ContainsKey($requestKey)) {continue}
        $usedRequests[$requestKey]=$true
        $matched+=$request
    }
    return [pscustomobject]@{ Detected = ($matched.Count -ge $Threshold); CompletedRestarts = $matched.Count; Requests = @($matched) }
}

function Read-Events {
    param([hashtable]$Filter, [int]$Maximum = 1000)
    try { return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $Maximum -ErrorAction Stop) }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { return @() }
        throw
    }
}

function Protect-Text {
    param([AllowEmptyString()][string]$Text)
    if (-not [string]::IsNullOrEmpty($DeepFreezeCommandLinePassword)) {
        $Text=$Text.Replace($DeepFreezeCommandLinePassword,'[REDACTED]')
        $escaped=ConvertTo-Json -InputObject $DeepFreezeCommandLinePassword -Compress
        if ($escaped.Length -gt 2) {$Text=$Text.Replace($escaped.Substring(1,$escaped.Length-2),'[REDACTED]')}
    }
    return $Text
}

function Get-RebootEvidence {
    $since = (Get-Date).AddMinutes(-$LookbackMinutes - 10)
    $requests = @(Read-Events -Filter @{LogName='System'; Id=1074; StartTime=$since} | ForEach-Object {
        $xml = [xml]$_.ToXml()
        $fields = @{}
        foreach ($entry in @($xml.Event.EventData.Data)) { $fields[[string]$entry.Name] = [string]$entry.InnerText }
        # USER32 1074: param1 process, param5 shutdown type, param6 comment.
        [pscustomobject]@{ TimeCreated=$_.TimeCreated; Process=$fields['param1']; Comment=$fields['param6']; RecordId=$_.RecordId }
    })
    $boots = @(Read-Events -Filter @{LogName='System'; ProviderName='Microsoft-Windows-Kernel-General'; Id=12; StartTime=$since} |
        ForEach-Object { [datetime]$_.TimeCreated })
    return Get-LoopDecision -Requests $requests -Boots $boots -Now (Get-Date) -Minutes $LookbackMinutes -Threshold $MinimumCompletedRestarts
}

function Assert-Administrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run as Administrator or SYSTEM.' }
}

function Assert-NoReparsePath {
    param([string]$Path)
    $item = $Path
    while (-not [string]::IsNullOrWhiteSpace($item)) {
        if (Test-Path -LiteralPath $item) {
            if ((Get-Item -LiteralPath $item -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw 'A protected local path contains a reparse point.'
            }
        }
        $parent = Split-Path -Parent $item
        if ($parent -eq $item) { break }
        $item = $parent
    }
}

function Set-PrivateLocalDirectory {
    Assert-NoReparsePath -Path $LocalRoot
    New-Item -Path $LocalRoot -ItemType Directory -Force | Out-Null
    $acl = New-Object Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true, $false)
    $acl.SetOwner((New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))
    foreach ($sidText in @('S-1-5-18','S-1-5-32-544')) {
        $sid = New-Object Security.Principal.SecurityIdentifier($sidText)
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $LocalRoot -AclObject $acl
}

function Quote-NativeArgument {
    param([AllowEmptyString()][string]$Value)
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function Find-Dfc {
    foreach ($path in @(
        "$env:WINDIR\System32\DFC.exe",
        "$env:WINDIR\SysWOW64\DFC.exe",
        "${env:ProgramFiles(x86)}\Faronics\Deep Freeze\Install C-0\DFC.exe"
    )) {
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
    }
    throw 'DFC.exe is unavailable. No automatic thaw or reboot will be attempted.'
}

function Start-RecoveryProcess {
    param([string]$FilePath,[string]$Arguments,[string]$OutputPath,[string]$ErrorPath)
    # Own the native Process from creation through exit. Avoid Start-Process's
    # redirected-process ExitCode issue; an unknown code must never become 0.
    $process=New-Object Diagnostics.Process
    $process.StartInfo.FileName=$FilePath
    $process.StartInfo.Arguments=$Arguments
    $process.StartInfo.UseShellExecute=$false
    $process.StartInfo.CreateNoWindow=$true
    $process.StartInfo.RedirectStandardOutput=-not [string]::IsNullOrEmpty($OutputPath)
    $process.StartInfo.RedirectStandardError=-not [string]::IsNullOrEmpty($ErrorPath)
    try {
        if (-not $process.Start()) {throw 'Native process failed to start.'}
        if ($process.StartInfo.RedirectStandardOutput) {
            $process | Add-Member -NotePropertyName RecoveryOutputTask -NotePropertyValue ($process.StandardOutput.ReadToEndAsync())
            $process | Add-Member -NotePropertyName RecoveryOutputPath -NotePropertyValue $OutputPath
        }
        if ($process.StartInfo.RedirectStandardError) {
            $process | Add-Member -NotePropertyName RecoveryErrorTask -NotePropertyValue ($process.StandardError.ReadToEndAsync())
            $process | Add-Member -NotePropertyName RecoveryErrorPath -NotePropertyValue $ErrorPath
        }
        return $process
    }
    catch {try {if (-not $process.HasExited) {$process.Kill()}} catch {}; $process.Dispose();throw}
}

function Save-RecoveryProcessOutput {
    param($Process)
    foreach ($kind in @('Output','Error')) {
        $taskProperty=$Process.PSObject.Properties['Recovery'+$kind+'Task']
        $pathProperty=$Process.PSObject.Properties['Recovery'+$kind+'Path']
        if ($null -ne $taskProperty -and $null -ne $pathProperty) {
            try {
                if ($taskProperty.Value.Wait(5000)) {
                    [IO.File]::WriteAllText([string]$pathProperty.Value,(Protect-Text ([string]$taskProperty.Value.Result)),(New-Object Text.UTF8Encoding($false)))
                }
            }
            catch {Write-Warning ('Could not save protected child '+$kind+' output.')}
        }
    }
}

function Invoke-Dfc {
    param([switch]$SetThawNextBoot,[switch]$SetFreezeNextBoot)
    $dfc = Find-Dfc
    if ($SetThawNextBoot -and $SetFreezeNextBoot) {throw 'Choose only one Deep Freeze next-boot state.'}
    $arguments = 'get /ISFROZEN'
    if ($SetFreezeNextBoot) {
        if ([string]::IsNullOrWhiteSpace($DeepFreezeCommandLinePassword)) {throw 'Configure the Deep Freeze command-line password before automatic refreezing.'}
        $arguments=(Quote-NativeArgument $DeepFreezeCommandLinePassword)+' /FREEZENEXTBOOT'
    }
    if ($SetThawNextBoot) {
        if ([string]::IsNullOrWhiteSpace($DeepFreezeCommandLinePassword)) { throw 'The Deep Freeze command-line password is not configured.' }
        $arguments = (Quote-NativeArgument $DeepFreezeCommandLinePassword) + ' /THAWNEXTBOOT'
    }
    $stdout = Join-Path $LocalRoot ([guid]::NewGuid().ToString() + '.out')
    $stderr = $stdout + '.err'
    $process = $null
    try {
        $process = Start-RecoveryProcess -FilePath $dfc -Arguments $arguments -OutputPath $stdout -ErrorPath $stderr
        if (-not $process.WaitForExit(10000)) {
            try { $process.Kill() } catch { }
            throw 'DFC command timed out. Command arguments and output were withheld.'
        }
        $process.WaitForExit()
        if ($null -eq $process.ExitCode) {throw 'Native process exit code is unavailable; operation refused.'}
        return [int]$process.ExitCode
    }
    finally {
        if ($null -ne $process) {Save-RecoveryProcessOutput -Process $process; $process.Dispose()}
        Remove-Item -LiteralPath $stdout,$stderr -Force -ErrorAction SilentlyContinue
    }
}

function Get-FreezeState {
    $code = Invoke-Dfc
    if ($code -eq 0) { return 'Thawed' }
    if ($code -eq 1) { return 'Frozen' }
    throw "Deep Freeze state is unknown (DFC exit $code). Recovery refused."
}

function Assert-PinnedScript {
    param([string]$Path, [string[]]$AllowedHashes)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw ('Required script is missing: ' + $Path) }
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($hash -notin $AllowedHashes) { throw ('Script hash is not approved: ' + (Split-Path -Leaf $Path) + '. Review its new version before adding its hash to this watchdog.') }
    $tokens = $null; $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) { throw ('PowerShell parser errors in ' + (Split-Path -Leaf $Path)) }
}

function Save-Json {
    param([string]$Path, $Value)
    $text = ConvertTo-Json -InputObject $Value -Depth 18
    $temp = $Path + '.' + [guid]::NewGuid().ToString() + '.tmp'
    try {
        $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($text)
        $stream = New-Object IO.FileStream($temp, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $stream.Write($bytes,0,$bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
        if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($temp,$Path,[System.Management.Automation.Language.NullString]::Value) }
        if (-not (Test-Path -LiteralPath $Path)) { [IO.File]::Move($temp,$Path) }
        if ([IO.File]::ReadAllText($Path) -cne $text) { throw 'Persistent write readback did not match.' }
    }
    finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
}

function Open-PersistentState {
    if ($PersistentStateRoot -notmatch '^\\\\[^\\]+\\[^\\]+(?:\\.*)?$' -or $PersistentStateRoot -match '(?i)EXAMPLE|YOURSERVER|CHANGEME') {
        throw 'Configure PersistentStateRoot with a real, durable UNC folder before enabling recovery.'
    }
    $machineFolder = Join-Path $PersistentStateRoot $env:COMPUTERNAME
    if (-not (Test-Path -LiteralPath $machineFolder -PathType Container)) {
        if ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -ne 'S-1-5-18') {
            throw 'First-time persistent folder creation must run as SYSTEM so SMB uses the computer account.'
        }
        [void][IO.Directory]::CreateDirectory($machineFolder)
    }
    # A create-enabled root permits name squatting. Reject another account's
    # folder before reading or writing any receipt, even if it grants us access.
    $account=New-Object Security.Principal.NTAccount($PersistentStateAccountDomain,($env:COMPUTERNAME+'$'))
    $expectedOwner=$account.Translate([Security.Principal.SecurityIdentifier]).Value
    $actualOwner=(Get-Acl -LiteralPath $machineFolder -ErrorAction Stop).GetOwner([Security.Principal.SecurityIdentifier]).Value
    if ($actualOwner -ne $expectedOwner) {throw 'Persistent folder owner does not match this computer account. Administrator review is required.'}
    $probe = Join-Path $machineFolder ([guid]::NewGuid().ToString() + '.probe')
    try {
        [IO.File]::WriteAllText($probe,'probe')
        if ([IO.File]::ReadAllText($probe) -ne 'probe') { throw 'State storage readback failed.' }
    }
    finally { Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue }
    $script:StatePath = Join-Path $machineFolder 'RecoveryState.json'
    if (Test-Path -LiteralPath $script:StatePath -PathType Leaf) {
        $candidate=Get-Content -LiteralPath $script:StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($field in @('SchemaVersion','ComputerName','Phase','StartedUtc','IncidentId','ThawAttempts','ForceAttempts','SuspendedTasks','UpdaterAttempted','UpdaterVerified','EvidencePath','BundlePath','EvidenceReplayDone','EvidenceTruncated','EvidenceCoverage','ForceStartedUtc','ForceBootUtc','InitialBootUtc','FailureMessage','CompletedUtc','DetectedRestarts')) {
            if ($null -eq $candidate.PSObject.Properties[$field]) {throw 'Persistent recovery state has missing fields. Automatic reset is refused.'}
        }
        if ($candidate.SchemaVersion -ne 1 -or $candidate.ComputerName -ne $env:COMPUTERNAME -or
            $candidate.Phase -notin @('Detected','ThawRequested','AfterThaw','UpdatingScripts','AwaitingEvaluation','ClosingServices','ForceRequested','FreezeRequested','Completed','NeedsAttention') -or
            $candidate.ThawAttempts -notin @(0,1) -or $candidate.ForceAttempts -notin @(0,1)) {
            throw 'Persistent recovery state is invalid. Automatic reset is refused.'
        }
        $started=[datetime]::Parse($candidate.StartedUtc)
        if ($started.ToUniversalTime() -gt (Get-Date).ToUniversalTime().AddMinutes(5)) {throw 'Persistent recovery state has a future timestamp.'}
        $incidentGuid=[guid]::Parse($candidate.IncidentId)
        $expectedFolder=Join-Path $machineFolder $incidentGuid.ToString()
        if ($candidate.EvidencePath -and $candidate.EvidencePath -ine (Join-Path $expectedFolder 'Evidence.ndjson')) {throw 'Unexpected evidence path in recovery state.'}
        if ($candidate.BundlePath -and $candidate.BundlePath -ine (Join-Path $expectedFolder 'Diagnostics.zip')) {throw 'Unexpected bundle path in recovery state.'}
        # Upgrade older receipt fields without resetting any consumed attempt.
        foreach ($field in @('FinalFreezeAttempts','FinalFreezeBootUtc','FinalFreezeRequestedUtc')) {
            if ($null -eq $candidate.PSObject.Properties[$field]) {
                $default=$null
                if ($field -eq 'FinalFreezeAttempts') {$default=0}
                $candidate | Add-Member -NotePropertyName $field -NotePropertyValue $default
            }
        }
        if ($candidate.FinalFreezeAttempts -notin @(0,1)) {throw 'Invalid final-freeze attempt count.'}
        # Assign only AFTER validation; malformed receipts must never be overwritten.
        $script:State=$candidate
    }
    return $machineFolder
}

function Open-StateWithRetry {
    for ($attempt=1;$attempt -le 3;$attempt++) {
        try {return Open-PersistentState}
        catch {
            if ($attempt -ge 3) {throw}
            Start-Sleep -Seconds 10
        }
    }
}

function Save-State { Save-Json -Path $script:StatePath -Value $script:State }

function Send-Event {
    param([string]$Status, [string]$Message, $Details = $null, [string]$Type = 'endpoint.reboot_recovery')
    $Message=Protect-Text $Message
    $event = [ordered]@{
        '@timestamp'=(Get-Date).ToUniversalTime().ToString('o'); EventType=$Type; SchemaVersion=1
        ComputerName=$env:COMPUTERNAME; Domain=$script:DomainName; ScriptName=$script:Name; ScriptVersion=$script:Version
        RunId=$script:RunId; IncidentId=$script:Incident; Status=$Status
        Severity=$(if ($Status -match '^(LoopDetected|NeedsAttention|Error|Cooldown-NeedsAttention)$') { 'Critical' } else { 'Info' })
        Message=$Message; RecoveryEnabled=[bool]($Recover -or $CollectEvidenceOnly -or ($Install -and $EnableRecovery)); DetailsJson=$(if ($null -ne $Details) { ConvertTo-Json -InputObject $Details -Depth 12 -Compress } else { $null })
    }
    $append=$true
    $lastSent=$null
    if ($Status -in @('AuditNoLoop','NoLoopDetected','AuditSundayExcluded','DeferredSundayMaintenance','Cooldown-Completed','Cooldown-NeedsAttention')) {
        try {
            if (Test-Path -LiteralPath $LatestPath) {
                $previous=Get-Content -LiteralPath $LatestPath -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($null -ne $previous.PSObject.Properties['TelemetryLastSentUtc']) {$lastSent=$previous.TelemetryLastSentUtc}
                if ($lastSent -and $previous.Status -eq $Status -and $previous.DetailsJson -eq $event.DetailsJson -and ([datetime]$lastSent).ToUniversalTime() -gt (Get-Date).ToUniversalTime().AddMinutes(-30)) {$append=$false}
            }
        }
        catch {$append=$true}
    }
    if ($script:TelemetryReady -and $append) {
        $lastSent=(Get-Date).ToUniversalTime().ToString('o')
        $event.TelemetryLastSentUtc=$lastSent
        Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine (ConvertTo-Json -InputObject $event -Depth 5 -Compress)
    }
    $event.TelemetryLastSentUtc=$lastSent
    Save-Json -Path $LatestPath -Value $event
    $script:PublishedStatus = $Status
    Write-Host ('{0}: {1}' -f $Status,$Message)
}

function Add-Evidence {
    param([string]$Kind,[string]$Source,$Value)
    $text = Protect-Text (ConvertTo-Json -InputObject $Value -Depth 12 -Compress)
    if ($text.Length -gt 12000) {
        $text=ConvertTo-Json -InputObject @{Truncated=$true;OriginalCharacters=$text.Length;Text=$text.Substring(0,10000)} -Compress
        $script:EvidenceTruncated=$true
    }
    $entry = [ordered]@{
        '@timestamp'=(Get-Date).ToUniversalTime().ToString('o'); EventType='endpoint.reboot_recovery.evidence'; SchemaVersion=1
        ComputerName=$env:COMPUTERNAME; ScriptName=$script:Name; ScriptVersion=$script:Version
        IncidentId=$script:Incident; RunId=$script:RunId; EvidenceKind=$Kind; Source=$Source; DetailsJson=$text
    }
    $json = ConvertTo-Json -InputObject $entry -Depth 4 -Compress
    $size = [Text.Encoding]::UTF8.GetByteCount($json)
    if ($script:EvidenceCount -ge 2000 -or ($script:EvidenceBytes + $size) -gt 2097152) {
        $script:EvidenceTruncated=$true
        return
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $entry.EvidenceFingerprint=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($script:Incident+'|'+$Kind+'|'+$Source+'|'+$text)))).Replace('-','') }
    finally { $sha.Dispose() }
    $json=ConvertTo-Json -InputObject $entry -Depth 4 -Compress
    if ($null -ne $script:EvidenceWriter) { $script:EvidenceWriter.WriteLine($json); $script:EvidenceWriter.Flush() }
    if ($script:TelemetryReady) { Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $json }
    $script:EvidenceCount++; $script:EvidenceBytes += $size
}

function Read-TelemetryTail {
    param([string]$Path,[int]$MaximumBytes=4194304,[int]$MaximumLines=20000)
    $stream=New-Object IO.FileStream($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $offset=[Math]::Max(0,$stream.Length-$MaximumBytes)
        [void]$stream.Seek($offset,[IO.SeekOrigin]::Begin)
        $reader=New-Object IO.StreamReader($stream)
        try {
            if ($offset -gt 0) {[void]$reader.ReadLine();$script:EvidenceTruncated=$true}
            return @($reader.ReadToEnd() -split '\r?\n' | Select-Object -Last $MaximumLines)
        }
        finally {$reader.Dispose()}
    }
    finally {$stream.Dispose()}
}

function Get-PendingFlags {
    $flags = @()
    foreach ($name in @('RebootPending','RebootInProgress','PackagesPending')) {
        $path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\' + $name
        if (Test-Path -LiteralPath $path -ErrorAction Stop) { $flags += [pscustomobject]@{Name=('CBS'+$name); Details=$path} }
    }
    foreach ($name in @('RebootRequired','PostRebootReporting')) {
        $path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\' + $name
        if (Test-Path -LiteralPath $path -ErrorAction Stop) { $flags += [pscustomobject]@{Name=('WindowsUpdate'+$name); Details=$path} }
    }
    $session = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop
    foreach ($name in @('PendingFileRenameOperations','PendingFileRenameOperations2')) {
        if ($null -ne $session.PSObject.Properties[$name]) {
            $flags += [pscustomobject]@{Name=$name; Details=(@($session.$name) -join ' | ')}
        }
    }
    $updates=$null
    if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Updates' -ErrorAction Stop) {
        $updates=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Updates' -ErrorAction Stop
    }
    if ($null -ne $updates -and $null -ne $updates.PSObject.Properties['UpdateExeVolatile'] -and [string]$updates.UpdateExeVolatile -ne '0') {
        $flags += [pscustomobject]@{Name='UpdateExeVolatile'; Details=[string]$updates.UpdateExeVolatile}
    }
    $wu = New-Object -ComObject Microsoft.Update.SystemInfo
    if ([bool]$wu.RebootRequired) { $flags += [pscustomobject]@{Name='WindowsUpdateCOM'; Details='Microsoft.Update.SystemInfo.RebootRequired=true'} }
    return @($flags)
}

function Read-PlainScriptHeader {
    param([string]$Path, [int]$MaximumLines=12)
    # StreamReader produces undecorated strings. Get-Content adds provider
    # properties which Windows PowerShell 5.1 can recursively serialize.
    $reader=[IO.StreamReader]::new($Path)
    try {
        for ($i=0;$i -lt $MaximumLines;$i++) {
            $line=$reader.ReadLine()
            if ($null -eq $line) {break}
            $line
        }
    }
    finally {$reader.Dispose()}
}

function Collect-IncidentEvidence {
    param([string]$MachineFolder, $Loop)
    $incidentFolder = Join-Path $MachineFolder $script:Incident
    New-Item -Path $incidentFolder -ItemType Directory -Force | Out-Null
    $evidencePath = Join-Path $incidentFolder 'Evidence.ndjson'
    $script:EvidenceWriter = New-Object IO.StreamWriter($evidencePath,$false,(New-Object Text.UTF8Encoding($false)))
    $coverage = New-Object System.Collections.Generic.List[string]
    $logQueue = New-Object System.Collections.Generic.List[object]
    $captured = Join-Path $incidentFolder 'Diagnostics'
    New-Item -Path $captured -ItemType Directory -Force | Out-Null
    $totalCopied = 0
    $now = Get-Date
    $sunday = $now.Date.AddDays(-[int]$now.DayOfWeek)
    if ($now.DayOfWeek -eq [DayOfWeek]::Sunday -and $now.TimeOfDay -lt [timespan]$NormalSundayStart) { $sunday=$sunday.AddDays(-7) }
    try {
        Send-Event 'EvidenceProgress' 'Collecting Deep Freeze and pending-reboot context.'
        Add-Evidence 'Detection' 'System/1074 + Kernel-General/12' $Loop
        $dfcPath=Find-Dfc
        Add-Evidence 'DeepFreezeContext' 'Current DFC state and executable' @{State=(Get-FreezeState);DfcPath=$dfcPath;DfcFileVersion=(Get-Item -LiteralPath $dfcPath).VersionInfo.FileVersion;Boot=$script:Boot;ConfiguredSundayStart=$NormalSundayStart;ConfiguredSundayEnd=$NormalSundayEnd;Note='Current observation does not prove the last Sunday thaw state.'}
        Add-Evidence 'OriginalRebootTaskStates' 'Persistent incident receipt' @($script:State.SuspendedTasks)
        Add-Evidence 'PendingReboot' 'Registry and Windows Update COM' @(Get-PendingFlags)
        Add-Evidence 'SundayWindow' 'Configured local maintenance dates' @{Start=$sunday.ToString('o'); End=$sunday.AddDays(1).ToString('o'); Note='Available evidence only; absence is not proof of a missed run.'}
        $names = @('00_Update-Scripts-FromShare','01_Enable_Windows_Update_Services','06_Weekend_Windows_Updates','07_Force_Reboot_Install_Updates','09_Disable_Windows_Update_Services','15_Sunday_Update_Closeout','Register-Tasks_SYSTEM')
        foreach ($name in $names) {
            Send-Event 'EvidenceProgress' ('Collecting logs and script metadata: '+$name)
            $candidates = @()
            $latest = Join-Path $LogRoot ($name+'.latest.json')
            if (Test-Path -LiteralPath $latest -PathType Leaf) { $candidates += Get-Item -LiteralPath $latest }
            $available = @(Get-ChildItem -LiteralPath $LogRoot -Filter ('*'+$name+'*.log') -File -ErrorAction SilentlyContinue)
            $candidates += @($available | Sort-Object LastWriteTime -Descending | Select-Object -First 2)
            $candidates += @($available | Where-Object {$_.LastWriteTime -ge $sunday -and $_.LastWriteTime -lt $sunday.AddDays(1)} | Sort-Object LastWriteTime -Descending | Select-Object -First 2)
            if ($candidates.Count -eq 0) { $coverage.Add('No available latest report/text log for '+$name) }
            foreach ($file in @($candidates | Sort-Object FullName -Unique)) {
                Send-Event 'EvidenceProgress' ('Reading diagnostic file: '+$file.FullName)
                if ($totalCopied -ge 8388608) { $script:EvidenceTruncated=$true; break }
                try {
                    $lines = @(Read-TelemetryTail -Path $file.FullName -MaximumBytes 1048576 -MaximumLines 500 | ForEach-Object {Protect-Text ([string]$_)})
                    $destination = Join-Path $captured $file.Name
                    if ($file.Length -le 1048576 -and ($totalCopied+$file.Length) -le 8388608) {
                        $redacted=Protect-Text ([IO.File]::ReadAllText($file.FullName))
                        [IO.File]::WriteAllText($destination,$redacted,(New-Object Text.UTF8Encoding($false)))
                        $totalCopied += $file.Length
                    }
                    if (-not (Test-Path -LiteralPath $destination)) {
                        $tailText=$lines -join [Environment]::NewLine
                        $tailSize=[Text.Encoding]::UTF8.GetByteCount($tailText)
                        if ($tailSize -le 1048576 -and ($totalCopied+$tailSize) -le 8388608) {
                            [IO.File]::WriteAllText(($destination+'.tail.txt'),$tailText,(New-Object Text.UTF8Encoding($false)))
                            $totalCopied+=$tailSize
                        }
                        $script:EvidenceTruncated=$true
                    }
                    Add-Evidence 'FileMetadata' $file.FullName @{Length=$file.Length; LastWriteTime=$file.LastWriteTime.ToString('o'); Selection='Latest report or recent/last-Sunday text log; at most 500 lines sent'}
                    if ($file.Extension -eq '.json') {
                        $snapshot=Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
                        $selected=[ordered]@{}
                        foreach ($field in @('EventType','ComputerName','ScriptName','ScriptVersion','RunId','Status','ExitCode','StartTime','EndTime','Timestamp','FailureMessage','SourceShare','WarningCount','ErrorCount','PendingReboot','PendingRebootReasons','OverallResult','StartupResumeInvocation','PreviousWindowsUpdateContext','RebootCauseDetails','windows_update_control','Results')) {
                            if ($null -ne $snapshot.PSObject.Properties[$field]) {$selected[$field]=$snapshot.$field}
                        }
                        Add-Evidence 'LatestReport' $file.FullName $selected
                    }
                    if ($file.Extension -ne '.json') {
                        foreach ($line in $lines) {
                            $value=[string]$line
                            if ($value.Length -gt 12000) {$value=$value.Substring(0,12000);$script:EvidenceTruncated=$true}
                            $logQueue.Add([pscustomobject]@{Source=$file.FullName;Line=$value})
                        }
                    }
                }
                catch { $coverage.Add('Could not collect '+$file.Name+': '+$_.Exception.Message) }
            }
            $source = Join-Path 'C:\Scripts' ($name+'.ps1')
            if (Test-Path -LiteralPath $source -PathType Leaf) {
                $file=Get-Item -LiteralPath $source
                Send-Event 'EvidenceProgress' ('Hashing installed script: '+$source)
                $sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
                Send-Event 'EvidenceProgress' ('Reading plain-text script header: '+$source)
                $plainHeader=@(Read-PlainScriptHeader -Path $source)
                Send-Event 'EvidenceProgress' ('Serializing installed-script evidence: '+$source)
                Add-Evidence 'InstalledScript' $source @{Length=[long]$file.Length; LastWriteTime=$file.LastWriteTime.ToString('o'); SHA256=$sourceHash; Header=$plainHeader}
            }
        }
        Send-Event 'EvidenceProgress' 'Reading recent maintenance telemetry.'
        # Bounded scan of existing telemetry for last-Sunday execution summaries.
        $seenSunday=0
        if (Test-Path -LiteralPath $TelemetryPath -PathType Leaf) {
            foreach ($line in @(Read-TelemetryTail -Path $TelemetryPath)) {
                try {
                    $row=$line | ConvertFrom-Json -ErrorAction Stop
                    if ($null -eq $row.PSObject.Properties['ScriptName'] -or [string]$row.ScriptName -notmatch '^(00_Update|01_Enable|06_Weekend|07_Force|09_Disable|15_Sunday|Register-Tasks)') {continue}
                    $stamp=$null
                    foreach ($field in @('StartTime','Timestamp','@timestamp')) {
                        if ($null -ne $row.PSObject.Properties[$field] -and $row.$field) {$stamp=$row.$field;break}
                    }
                    if ($null -eq $stamp) {continue}
                    $when=([datetime]::Parse([string]$stamp)).ToLocalTime()
                    if ($when -ge $sunday -and $when -lt $sunday.AddDays(1)) {Add-Evidence 'SundayTelemetry' $TelemetryPath $row;$seenSunday++}
                }
                catch { $coverage.Add('An existing telemetry line could not be parsed or published.'); break }
            }
        }
        $coverage.Add('Last-Sunday execution summaries found in the newest 4 MiB / 20000 telemetry lines: '+$seenSunday)
        Send-Event 'EvidenceProgress' 'Collecting scheduled-task configuration.'
        $taskData = @(Get-ScheduledTask | Where-Object {
            (Get-TaskArgumentText -Task $_) -match '00_Update|01_Enable_Windows|06_Weekend_Windows|07_Force_Reboot|09_Disable_Windows|15_Sunday_Update|-ActionId\s+"?(UpdateScripts|EnableWindowsUpdateServices|WindowsUpdatePass|ForceRebootInstallUpdates|SundayUpdateCloseout)'
        } | ForEach-Object {
            $info=$_ | Get-ScheduledTaskInfo
            # Export the full definition as plain XML text rather than serializing
            # raw CIM actions/triggers and their provider object graphs.
            $definition=[string]::Join([Environment]::NewLine,[string[]]@(Export-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction Stop))
            [pscustomobject]@{Name=[string]$_.TaskName; Path=[string]$_.TaskPath; State=[string]$_.State; Account=[string]$_.Principal.UserId; Enabled=[bool]$_.Settings.Enabled; StartWhenAvailable=[bool]$_.Settings.StartWhenAvailable; LastRunTime=$info.LastRunTime.ToString('o'); LastTaskResult=[long]$info.LastTaskResult; DefinitionXml=$definition}
        })
        foreach ($task in $taskData) { Add-Evidence 'ScheduledTask' $task.Name $task }
        $taskData | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $captured 'ScheduledTasks.json') -Encoding UTF8
        foreach ($window in @(@{Name='LastSunday';Start=$sunday;End=$sunday.AddDays(1)},@{Name='Recent';Start=$now.AddHours(-2);End=$now})) {
            foreach ($channel in @('System','Microsoft-Windows-WindowsUpdateClient/Operational','Microsoft-Windows-TaskScheduler/Operational')) {
                try {
                    $filter=@{LogName=$channel;StartTime=$window.Start;EndTime=$window.End}
                    if ($channel -eq 'System') { $filter.Id=@(12,13,19,20,21,41,1074,6005,6006,6008,7031,7040) }
                    Send-Event 'EvidenceProgress' ('Reading event channel: '+$channel+'; window='+$window.Name)
                    $events=@(Read-Events -Filter $filter -Maximum 100)
                    if ($events.Count -eq 100) { $script:EvidenceTruncated=$true }
                    if ($events.Count -eq 0) { $coverage.Add('No retained '+$window.Name+' events in '+$channel) }
                    $rows=@($events | ForEach-Object { [pscustomobject]@{TimeCreated=$_.TimeCreated.ToString('o');Id=$_.Id;Provider=$_.ProviderName;RecordId=$_.RecordId;Message=$_.Message} })
                    foreach ($row in $rows) { Add-Evidence ('EventLog-'+$window.Name) $channel $row }
                    $rows | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $captured (($window.Name+'-'+$channel) -replace '[^A-Za-z0-9_-]','_')) -Encoding UTF8
                }
                catch { $coverage.Add('Event channel unavailable: '+$channel) }
            }
        }
        foreach ($item in $logQueue) {Add-Evidence 'LogLine' $item.Source $item.Line}
    }
    finally {
        $coverage.Add('Evidence budget: 2 MiB / 2000 records; each value at most 12000 characters. Files: at most 500 lines sent each. Event queries: at most 100 newest events per channel/window.')
        $coverage.Add('EvidenceTruncated='+$script:EvidenceTruncated)
        $coverage.Add('No source/password configuration for the watchdog was collected. Archived log ZIPs and full WindowsUpdate ETL/CBS logs are not collected by this version.')
        $coverage | Set-Content -LiteralPath (Join-Path $captured 'Coverage.txt') -Encoding UTF8
        $script:EvidenceWriter.Dispose(); $script:EvidenceWriter=$null
    }
    # The source files are named explicitly above; never archive the credential-bearing script.
    $bundle=Join-Path $incidentFolder 'Diagnostics.zip'
    Send-Event 'EvidenceProgress' 'Compressing incident diagnostics.'
    Compress-Archive -LiteralPath $captured -DestinationPath $bundle -Force
    $script:State.EvidencePath=$evidencePath
    $script:State.BundlePath=$bundle
    $script:State.EvidenceTruncated=$script:EvidenceTruncated
    $script:State.EvidenceCoverage=@($coverage)
    Save-State
    Send-Event 'EvidenceCaptured' 'Recovery diagnostics saved on persistent storage and appended to the existing Elastic telemetry file.' @{BundlePath=$bundle; EvidenceRecords=$script:EvidenceCount; Truncated=$script:EvidenceTruncated; Coverage=@($coverage)}
}

function Invoke-BoundedEvidenceCollection {
    param([string]$MachineFolder, $Loop)
    # Run potentially blocking event/file/SMB collection in a separate process.
    # Parent retains the recovery mutex. Child is evidence-only and may not thaw,
    # update scripts, suspend tasks, or request a reboot.
    $stdout=Join-Path $LocalRoot ('EvidenceCollector-'+$script:Incident+'.out')
    $stderr=Join-Path $LocalRoot ('EvidenceCollector-'+$script:Incident+'.err')
    $arguments='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File '+
        (Quote-NativeArgument $PSCommandPath)+' -CollectEvidenceOnly -ExpectedIncidentId '+
        (Quote-NativeArgument $script:Incident)
    $process=$null
    $lastProgress=$null
    try {
        Write-Host ('Collecting incident evidence; time limit '+$EvidenceCollectionTimeoutSeconds+' seconds.')
        $process=Start-RecoveryProcess -FilePath $Pwsh -Arguments $arguments -OutputPath $stdout -ErrorPath $stderr
        $clock=[Diagnostics.Stopwatch]::StartNew()
        while (-not $process.WaitForExit(1000)) {
            if ($clock.Elapsed.TotalSeconds -ge $EvidenceCollectionTimeoutSeconds) {
                try {$process.Kill();[void]$process.WaitForExit(5000)} catch {}
                $lastStage='Unavailable'
                try {$lastStage=[string](Get-Content -LiteralPath $LatestPath -Raw | ConvertFrom-Json).Message} catch {}
                throw ('Evidence collection exceeded '+$EvidenceCollectionTimeoutSeconds+' seconds. No thaw or reboot requested. Last stage: '+$lastStage)
            }
            if ([int]$clock.Elapsed.TotalSeconds % 10 -eq 0) {
                $message='Evidence collection is running.'
                try {$message=[string](Get-Content -LiteralPath $LatestPath -Raw | ConvertFrom-Json).Message} catch {}
                if ($message -ne $lastProgress) {Write-Host $message;$lastProgress=$message}
            }
        }
        $process.WaitForExit()
        $process.WaitForExit()
        if ($null -eq $process.ExitCode) {throw 'Native process exit code is unavailable; operation refused.'}
        if ($process.ExitCode -ne 0) {throw ('Evidence collector failed with exit code '+$process.ExitCode+'. Inspect the latest recovery report and protected collector output.')}
        [void](Open-PersistentState)
        if ($script:State.IncidentId -ne $script:Incident -or -not $script:State.EvidencePath -or -not $script:State.BundlePath) {throw 'Evidence collector did not publish a complete matching incident receipt.'}
    }
    catch {
        # The child can have committed evidence fields even if the parent fails.
        # Reload before the outer handler latches the incident, preserving them.
        try {
            [void](Open-PersistentState)
            if ($script:State.IncidentId -ne $script:Incident) {$script:StatePath=$null;throw 'Incident changed during evidence collection.'}
        }
        catch { }
        throw
    }
    finally {if ($null -ne $process) {Save-RecoveryProcessOutput -Process $process;$process.Dispose()}}
}

function Replay-EvidenceAfterThaw {
    if ($script:State.EvidenceReplayDone) { return }
    if (-not $script:TelemetryReady -or -not (Test-Path -LiteralPath $script:State.EvidencePath -PathType Leaf)) { throw 'Persistent incident evidence is unavailable for replay.' }
    foreach ($line in Get-Content -LiteralPath $script:State.EvidencePath -Encoding UTF8) {
        if (-not [string]::IsNullOrWhiteSpace($line)) { Write-MaintenanceTelemetryLine -Path $TelemetryPath -JsonLine $line }
    }
    $script:State.EvidenceReplayDone=$true
    Save-State
    Send-Event 'EvidenceReplayed' 'Pre-reboot evidence appended again while thawed so the agent can retain and ship it.' @{BundlePath=$script:State.BundlePath; DeduplicationField='EvidenceFingerprint'; Note='Local append is not proof of Elastic ingestion.';Coverage=$script:State.EvidenceCoverage;Truncated=$script:State.EvidenceTruncated}
}

function Get-TaskArgumentText {
    param($Task)
    $values=@(foreach ($action in @($Task.Actions)) {
        if ($null -ne $action.PSObject.Properties['Arguments']) {[string]$action.Arguments}
    })
    return ($values -join ' ')
}

function Get-RebootTasks {
    return @(Get-ScheduledTask | Where-Object {
        $argsText=Get-TaskArgumentText -Task $_
        ($argsText -match '(?i)-File\s+"?C:\\Scripts\\07_Force_Reboot_Install_Updates\.ps1"?(?:\s|$)' -or
         $argsText -match '(?i)-ActionId\s+"?ForceRebootInstallUpdates"?(?:\s|$)') -and
        $argsText -notmatch '(?i)-(Force|ManualResume)(?:\s|$)'
    })
}

function Suspend-RebootTasks {
    foreach ($task in @(Get-RebootTasks)) {
        $key=$task.TaskPath+$task.TaskName
        if (@($script:State.SuspendedTasks | Where-Object {($_.Path+$_.Name) -eq $key}).Count -eq 0) {
            $script:State.SuspendedTasks=@($script:State.SuspendedTasks)+@([pscustomobject]@{Name=$task.TaskName;Path=$task.TaskPath;WasEnabled=[bool]$task.Settings.Enabled})
            Save-State
        }
        if ($task.Settings.Enabled) { Disable-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop | Out-Null }
        if ($task.State -eq 'Running') { Stop-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop }
    }
    # Stop only recognized, non-manual Script 07/launcher PowerShell processes.
    foreach ($process in @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop)) {
        $command=[string]$process.CommandLine
        if ($process.ProcessId -eq $PID -or $command -match '(?i)-(Force|ManualResume)(?:\s|$)') { continue }
        $direct=$command -match '(?i)-File\s+"?C:\\Scripts\\07_Force_Reboot_Install_Updates\.ps1"?(?:\s|$)'
        $launcher=$command -match '(?i)-File\s+"?C:\\Scripts\\Invoke-MaintenanceScript\.ps1"?\s+-ActionId\s+"?ForceRebootInstallUpdates"?(?:\s|$)'
        if ($direct -or $launcher) { Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop }
    }
}

function Restore-RebootTasks {
    Assert-PinnedScript -Path $Script07 -AllowedHashes $ApprovedScript07Hashes
    foreach ($saved in @($script:State.SuspendedTasks)) {
        $task=Get-ScheduledTask -TaskName $saved.Name -TaskPath $saved.Path -ErrorAction Stop
        $stillMatches=@(Get-RebootTasks | Where-Object {$_.TaskName -eq $saved.Name -and $_.TaskPath -eq $saved.Path}).Count -gt 0
        if (-not $stillMatches) { throw 'A suspended task changed to an unrecognized action. Restore it manually.' }
        if ($saved.WasEnabled) { Enable-ScheduledTask -TaskName $saved.Name -TaskPath $saved.Path -ErrorAction Stop | Out-Null }
        if (-not $saved.WasEnabled) { Disable-ScheduledTask -TaskName $saved.Name -TaskPath $saved.Path -ErrorAction Stop | Out-Null }
    }
}

function Invoke-ChildScript {
    param([string]$Path,[string]$ExtraArguments='',[int]$TimeoutSeconds=600)
    $arguments='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File '+(Quote-NativeArgument $Path)
    if ($ExtraArguments) { $arguments+=' '+$ExtraArguments }
    $process=Start-RecoveryProcess -FilePath $Pwsh -Arguments $arguments
    try {
        if (-not $process.WaitForExit($TimeoutSeconds*1000)) {
            try {$process.Kill()} catch { }
            throw ('Child script timed out: '+(Split-Path -Leaf $Path)+'. No automatic retry will occur.')
        }
        $process.WaitForExit()
        if ($null -eq $process.ExitCode) {throw 'Native process exit code is unavailable; operation refused.'}
        return [int]$process.ExitCode
    }
    finally {$process.Dispose()}
}

function Assert-BackgroundClosed {
    $wrong=@(Get-CimInstance Win32_Service -ErrorAction Stop | Where-Object {
        $_.Name -in @('wuauserv','BITS','UsoSvc') -and ($_.State -ne 'Stopped' -or $_.StartMode -ne 'Disabled')
    })
    if ($wrong.Count -gt 0) { throw 'Background update services are not all stopped and disabled. Manual investigation is required.' }
    $rebootTasks=@(Get-ScheduledTask -ErrorAction Stop | Where-Object {$_.TaskName -eq 'Reboot' -and $_.TaskPath -eq '\Microsoft\Windows\UpdateOrchestrator\'})
    if (@($rebootTasks | Where-Object {$_.Settings.Enabled}).Count -gt 0) { throw 'The Windows Update Reboot task remains enabled.' }
}

function Initialize-RepairScreenModule {
    if (-not (Test-Path -LiteralPath $RepairScreenModule -PathType Leaf)) {throw 'RepairScreen.psm1 is missing. Extract the complete package and reinstall while thawed.'}
    Import-Module -Name $RepairScreenModule -Global -Force -DisableNameChecking -ErrorAction Stop
}

function Update-RepairDisplay {
    if (-not $EnableRepairScreen -or $null -eq $script:StatePath -or $null -eq $script:State) {return}
    if ((Get-FreezeState) -ne 'Thawed') {return}
    try {
        $folder=Join-Path (Split-Path -Parent $script:StatePath) $script:Incident
        $result=Enable-ComptonRepairScreen -IncidentFolder $folder -IncidentId $script:Incident -ComputerName $env:COMPUTERNAME -Title $RepairScreenTitle -Message $RepairScreenMessage
        $status='RepairScreenConfigured'
        if (-not $result.CurrentSettingsMatch) {$status='RepairScreenPolicyChanged'}
        Send-Event $status 'Repair display method selected from Windows edition; appearance is refreshed by Windows at lock/sign-in.' $result
    }
    catch {
        Send-Event 'RepairScreenUnavailable' 'Repair screen setup could not be completed; continuing bounded repair. The saved display backup is retained for restoration.' @{Failure=$_.Exception.Message}
    }
}

function Restore-RepairDisplay {
    # Restore a prior snapshot even if the feature was disabled after it was applied.
    if ($null -eq $script:StatePath -or $null -eq $script:State) {return}
    $folder=Join-Path (Split-Path -Parent $script:StatePath) $script:Incident
    $result=Restore-ComptonRepairScreen -IncidentFolder $folder -IncidentId $script:Incident -ComputerName $env:COMPUTERNAME
    if (-not $result.Restored) {throw 'The original sign-in display has not been restored.'}
    Send-Event 'RepairScreenRestored' 'Original sign-in settings were restored or newer external policy settings preserved while thawed.' $result
}

function Complete-Recovery {
    if ((Get-FreezeState) -ne 'Thawed') { throw 'Thawed state was lost during recovery.' }
    if (@(Get-PendingFlags).Count -gt 0) { throw 'Pending reboot indicators remain. No additional reboot is permitted for this incident.' }
    Assert-BackgroundClosed
    foreach ($path in @('C:\Logs\00_Update-Scripts-FromShare.latest.json','C:\Logs\09_Disable_Windows_Update_Services.latest.json','C:\Logs\07_ManualUpdateRepair.latest.json')) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            try {Add-Evidence 'RepairResult' $path (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)} catch { }
        }
    }
    Restore-RebootTasks
    $script:State.FailureMessage=$null
    if ($AutomaticFreezeAfterVerifiedRepair) {
        if ($script:State.FinalFreezeAttempts -ge 1) {throw 'The final frozen restart attempt was already consumed.'}
        $script:State.FinalFreezeAttempts=1
        $script:State.FinalFreezeBootUtc=$script:Boot.ToUniversalTime().ToString('o')
        $script:State.FinalFreezeRequestedUtc=[datetime]::UtcNow.ToString('o')
        $script:State.Phase='FreezeRequested'
        Save-State
        $code=Invoke-Dfc -SetFreezeNextBoot
        if ($code -ne 0) {throw "Deep Freeze refused FREEZENEXTBOOT (exit $code). No final restart requested."}
        Restore-RepairDisplay
        Send-Event 'FrozenRestartRequested' 'Repair verified and original display/tasks restored. Restarting into frozen operation.' @{DelaySeconds=$FinalFreezeRebootDelaySeconds;IncidentId=$script:Incident;ForceAttempts=$script:State.ForceAttempts;FinalFreezeAttempts=$script:State.FinalFreezeAttempts}
        & "$env:WINDIR\System32\shutdown.exe" /r /f /t $FinalFreezeRebootDelaySeconds /d p:4:1 /c 'Compton reboot recovery: verified repair complete; return to frozen operation.' | Out-Null
        $shutdownCode=$LASTEXITCODE
        if ($shutdownCode -ne 0) {throw "Final frozen restart request failed (exit $shutdownCode). No automatic repeat."}
        return
    }
    Restore-RepairDisplay
    $script:State.Phase='Completed'
    $script:State.CompletedUtc=(Get-Date).ToUniversalTime().ToString('o')
    Save-State
    Send-Event 'ReadyForManualFreeze' 'Recovery verified; normal reboot tasks restored. The computer remains THAWED for technician review.' @{IncidentId=$script:Incident;ForceAttempts=$script:State.ForceAttempts;BundlePath=$script:State.BundlePath;ReadyForManualFreeze=$true}
}

function New-RecoveryState {
    param($Loop)
    return [pscustomobject]@{
        SchemaVersion=1;ComputerName=$env:COMPUTERNAME;IncidentId=[guid]::NewGuid().ToString()
        StartedUtc=(Get-Date).ToUniversalTime().ToString('o');Phase='Detected';InitialBootUtc=$script:Boot.ToUniversalTime().ToString('o')
        ThawAttempts=0;ForceAttempts=0;FinalFreezeAttempts=0;FinalFreezeBootUtc=$null;FinalFreezeRequestedUtc=$null;ForceStartedUtc=$null;ForceBootUtc=$null;UpdaterAttempted=$false;UpdaterVerified=$false
        SuspendedTasks=@();EvidencePath=$null;BundlePath=$null;EvidenceReplayDone=$false;EvidenceTruncated=$false;EvidenceCoverage=@()
        FailureMessage=$null;CompletedUtc=$null;DetectedRestarts=$Loop.CompletedRestarts
    }
}

function Run-Recovery {
    param([string]$MachineFolder,$Loop)
    if ($script:State.Phase -eq 'FreezeRequested') {
        if ($script:Boot.ToUniversalTime() -le ([datetimeoffset]$script:State.FinalFreezeBootUtc).UtcDateTime) {
            if ([datetime]::UtcNow -gt ([datetimeoffset]$script:State.FinalFreezeRequestedUtc).UtcDateTime.AddMinutes(3)) {throw 'Final frozen restart did not occur within three minutes. No repeat requested.'}
            Send-Event 'AwaitingFrozenRestart' 'Final frozen restart already requested; no further repair or reboot will be queued.'
            return
        }
        if ((Get-FreezeState) -ne 'Frozen') {throw 'Final restart occurred but Deep Freeze is not frozen. Manual attention required; no retry.'}
        $script:State.Phase='Completed';$script:State.CompletedUtc=[datetime]::UtcNow.ToString('o');Save-State
        Send-Event 'CompletedFrozen' 'Repair verified and computer confirmed frozen after the final restart. Classroom use can resume.' @{IncidentId=$script:Incident;ThawAttempts=$script:State.ThawAttempts;ForceAttempts=$script:State.ForceAttempts;FinalFreezeAttempts=$script:State.FinalFreezeAttempts}
        return
    }
    if ((Get-Date).ToUniversalTime() -gt (([datetimeoffset]$script:State.StartedUtc).UtcDateTime).AddMinutes(60)) {throw 'The active recovery incident is older than one hour. Manual review is required.'}
    for ($transition=0;$transition -lt 8;$transition++) {
        if ((Test-NormalMaintenance -Now (Get-Date) -Start ([timespan]$NormalSundayStart) -End ([timespan]$NormalSundayEnd)) -or
            (Test-NormalMaintenance -Now (Get-Date).AddMinutes(15) -Start ([timespan]$NormalSundayStart) -End ([timespan]$NormalSundayEnd))) {
            Send-Event 'DeferredSundayMaintenance' 'Automatic loop repair is excluded during the normal Sunday maintenance window.'
            return
        }
        $freeze=Get-FreezeState
        if ($freeze -eq 'Thawed' -and $script:State.Phase -in @('AfterThaw','UpdatingScripts','AwaitingEvaluation','ClosingServices','ForceRequested')) {Update-RepairDisplay}
        switch ($script:State.Phase) {
            'Detected' {
                if ($RunUpdaterAfterThaw) { Assert-PinnedScript -Path $Updater -AllowedHashes $ApprovedUpdaterHashes }
                if (-not $RunUpdaterAfterThaw) { Assert-PinnedScript -Path $Script07 -AllowedHashes $ApprovedScript07Hashes }
                if ($freeze -eq 'Frozen' -and [string]::IsNullOrWhiteSpace($DeepFreezeCommandLinePassword)) { throw 'Configure the Deep Freeze command-line password before enabling automatic thaw.' }
                Suspend-RebootTasks
                Abort-MaintenanceCountdown
                Invoke-BoundedEvidenceCollection -MachineFolder $MachineFolder -Loop $Loop
                if ($freeze -eq 'Thawed') {
                    $script:State.Phase='AfterThaw'; Save-State
                    continue
                }
                if ($script:State.ThawAttempts -ge 1) { throw 'The one allowed thaw attempt was already consumed.' }
                # Record attempt BEFORE the native command so a crash cannot reset the budget.
                $script:State.ThawAttempts=1; $script:State.Phase='ThawRequested'; Save-State
                $code=Invoke-Dfc -SetThawNextBoot
                if ($code -ne 0) { throw "DFC refused THAWNEXTBOOT (exit $code). No reboot requested by this watchdog." }
                Send-Event 'ThawRestartRequested' 'Deep Freeze accepted THAWNEXTBOOT. A single restart will enter thawed repair.' @{DelaySeconds=60;IncidentId=$script:Incident;BundlePath=$script:State.BundlePath}
                $arguments=@('/r','/t','60','/d','p:4:1','/c',('Compton reboot recovery: enter thawed repair. Incident '+$script:Incident))
                & "$env:WINDIR\System32\shutdown.exe" @arguments | Out-Null
                $shutdownCode=$LASTEXITCODE
                if ($shutdownCode -ne 0) { throw "Thaw restart request failed (shutdown exit $shutdownCode). No automatic retry." }
                return
            }
            'ThawRequested' {
                if ($script:Boot.ToUniversalTime() -le ([datetimeoffset]$script:State.InitialBootUtc).UtcDateTime) {
                    if ((Get-Date).ToUniversalTime() -gt (([datetimeoffset]$script:State.StartedUtc).UtcDateTime).AddMinutes(10)) { throw 'The thawing restart did not occur within ten minutes.' }
                    Send-Event 'AwaitingThawRestart' 'A thawing restart is already requested. No second request will be sent.'
                    return
                }
                if ($freeze -ne 'Thawed') { throw 'Computer rebooted but is still frozen. The thaw budget is exhausted.' }
                $script:State.Phase='AfterThaw'; Save-State
                Send-Event 'ThawedVerified' 'The next Windows boot is verified and Deep Freeze reports thawed.'
                continue
            }
            'AfterThaw' {
                if ($freeze -ne 'Thawed') { throw 'Recovery cannot continue while frozen.' }
                Replay-EvidenceAfterThaw
                Suspend-RebootTasks
                $manualPath='C:\ProgramData\MISMaintenance\State\07_ManualUpdateRepair.json'
                if (Test-Path -LiteralPath $manualPath) { throw 'A Script 07 manual repair transaction already exists. Review it before starting a second repair.' }
                if ($RunUpdaterAfterThaw) {
                    if ($script:State.UpdaterAttempted) { throw 'The updater attempt was already consumed; automatic repeat refused.' }
                    Assert-PinnedScript -Path $Updater -AllowedHashes $ApprovedUpdaterHashes
                    $script:State.UpdaterAttempted=$true;$script:State.Phase='UpdatingScripts';Save-State
                    Send-Event 'UpdatingScripts' 'Running the approved Script 00 once while thawed.'
                    $exitCode=Invoke-ChildScript -Path $Updater -TimeoutSeconds 600
                    if ($exitCode -ne 0) { throw "Script 00 returned exit $exitCode. Recovery will not run Script 07." }
                }
                Assert-PinnedScript -Path $Script07 -AllowedHashes $ApprovedScript07Hashes
                Assert-PinnedScript -Path $Script09 -AllowedHashes $ApprovedScript09Hashes
                # Registrar may have restored reboot tasks; quarantine them again.
                Suspend-RebootTasks
                $script:State.UpdaterVerified=$true;$script:State.Phase='AwaitingEvaluation';Save-State
                continue
            }
            'UpdatingScripts' { throw 'The previous updater run was interrupted. No automatic repeat or reboot will be requested.' }
            'AwaitingEvaluation' {
                if ($freeze -ne 'Thawed') { throw 'Computer is frozen; repair refused.' }
                Suspend-RebootTasks
                if (((Get-Date)-$script:Boot).TotalSeconds -lt $BootSettleSeconds) {
                    Send-Event 'WaitingForServicing' 'Waiting for the boot-settle interval before inspecting pending work.'
                    return
                }
                Assert-PinnedScript -Path $Script07 -AllowedHashes $ApprovedScript07Hashes
                Assert-PinnedScript -Path $Script09 -AllowedHashes $ApprovedScript09Hashes
                $pending=@(Get-PendingFlags)
                Send-Event 'PendingWorkEvaluated' 'Pending reboot checks completed while thawed.' $pending
                # Always run the approved one-reboot Script 07 -Force pass,
                # even when the watchdog's preliminary flag set is empty.
                if ($script:State.ForceAttempts -ge 1) { throw 'The one permitted Script 07 -Force attempt was consumed.' }
                $script:State.ForceAttempts=1;$script:State.ForceStartedUtc=(Get-Date).ToUniversalTime().ToString('o')
                $script:State.ForceBootUtc=$script:Boot.ToUniversalTime().ToString('o');$script:State.Phase='ForceRequested';Save-State
                Send-Event 'SingleForceRepairRequested' 'Invoking the pinned current Script 07 -Force exactly once before automatic refreezing. Its startup continuation only verifies.'
                $exitCode=Invoke-ChildScript -Path $Script07 -ExtraArguments '-Force' -TimeoutSeconds 900
                if ($exitCode -notin @(0,3010)) { throw "Script 07 -Force returned exit $exitCode. No second force attempt is permitted." }
                return
            }
            'ClosingServices' { throw 'The background closeout pass was interrupted. No automatic repeat or reboot will be requested.' }
            'ForceRequested' {
                if ($freeze -ne 'Thawed') { throw 'The force-repair boot is frozen; further recovery is refused.' }
                if ((Get-Date).ToUniversalTime() -gt (([datetimeoffset]$script:State.ForceStartedUtc).UtcDateTime).AddMinutes(25)) { throw 'Script 07 startup verification did not complete within 25 minutes. No further reboot will be requested.' }
                if ($script:Boot.ToUniversalTime() -le ([datetimeoffset]$script:State.ForceBootUtc).UtcDateTime) {
                    Send-Event 'AwaitingSingleRepairRestart' 'The one Script 07 -Force invocation is already consumed; waiting for a newer boot.'
                    return
                }
                $reportPath='C:\Logs\07_ManualUpdateRepair.latest.json'
                if (-not (Test-Path -LiteralPath $reportPath)) { Send-Event 'AwaitingScript07Verification' 'Waiting for the Script 07 verification-only startup report.';return }
                $report=Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
                if ($report.ComputerName -ne $env:COMPUTERNAME -or ([datetime]$report.Timestamp).ToUniversalTime() -lt ([datetimeoffset]$script:State.ForceStartedUtc).UtcDateTime) { throw 'Script 07 verification report is stale or belongs to a different computer.' }
                if ($report.Status -eq 'NeedsAttention') { throw 'Script 07 could not verify repair. Inspect its manual repair report and closeout logs.' }
                if ($report.Status -ne 'ReadyForManualFreeze' -or -not $report.ReadyForManualFreeze -or -not $report.RebootVerified -or $report.RebootsRequested -ne 1) {
                    Send-Event 'AwaitingScript07Verification' 'Script 07 startup verification has not finished; no additional reboot requested.'
                    return
                }
                Complete-Recovery
                return
            }
            'Completed' { return }
            'NeedsAttention' { return }
            default { throw 'Unknown recovery phase.' }
        }
    }
    throw 'Recovery transition limit reached.'
}

function Abort-MaintenanceCountdown {
    $recent=@(Read-Events -Filter @{LogName='System';Id=1074;StartTime=$script:Boot} -Maximum 20 | Sort-Object TimeCreated -Descending | Select-Object -First 1)
    if ($recent.Count -eq 0) { return }
    $xml=[xml]$recent[0].ToXml()
    $data=@{}
    foreach ($entry in @($xml.Event.EventData.Data)) {$data[[string]$entry.Name]=[string]$entry.InnerText}
    if ([string]$data['param1'] -notmatch '(?i)shutdown\.exe' -or [string]$data['param6'] -notmatch '^(Pending reboot flags detected during maintenance closeout\.|Reboot flags remain after|State reset occurred\.)') { return }
    $oldPreference=$ErrorActionPreference
    try {
        $ErrorActionPreference='Continue'
        & "$env:WINDIR\System32\shutdown.exe" /a 2>$null | Out-Null
        $code=$LASTEXITCODE
    }
    finally {$ErrorActionPreference=$oldPreference}
    if ($code -notin @(0,1116)) { throw "Could not cancel the confirmed maintenance restart countdown (exit $code)." }
}

function Initialize-Telemetry {
    Assert-NoReparsePath -Path $LogRoot
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    if (Test-Path -LiteralPath $Framework -PathType Leaf) {
        Import-Module -Name $Framework -Global -Force -DisableNameChecking -ErrorAction Stop
        $null=Initialize-MaintenanceEnvironment -ScriptRoot 'C:\Scripts' -LogRoot $LogRoot
        $script:TelemetryReady=($null -ne (Get-Command Write-MaintenanceTelemetryLine -ErrorAction SilentlyContinue))
    }
}

function Assert-NoRepairPreview {
    param([string]$MachineFolder)
    $preview=Join-Path $MachineFolder 'RepairScreenPreview\Context.json'
    if (Test-Path -LiteralPath $preview -PathType Leaf) {
        $context=Get-Content -LiteralPath $preview -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($context.ComputerName -ne $env:COMPUTERNAME -or $context.Phase -ne 'Restored') {throw 'Restore the active repair-screen preview before enabling or running automatic recovery.'}
    }
}

function Install-Watchdog {
    if ((Get-FreezeState) -ne 'Thawed') { throw 'Install only while Deep Freeze is THAWED so the task and protected script persist.' }
    if (-not (Test-Path -LiteralPath $StartupNoticeSource -PathType Leaf)) {throw 'Show-RecoveryNotice.ps1 is missing. Extract the complete package.'}
    $noticeTokens=$null;$noticeErrors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($StartupNoticeSource,[ref]$noticeTokens,[ref]$noticeErrors)
    if ($noticeErrors.Count -gt 0) {throw 'Startup notice helper has parser errors.'}
    $moduleTokens=$null;$moduleErrors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($RepairScreenModule,[ref]$moduleTokens,[ref]$moduleErrors)
    if ($moduleErrors.Count -gt 0) {throw 'Repair display module has parser errors.'}
    $installedModule=Join-Path $LocalRoot 'RepairScreen.psm1'
    if ([IO.Path]::GetFullPath($RepairScreenModule) -ine [IO.Path]::GetFullPath($installedModule)) {Copy-Item -LiteralPath $RepairScreenModule -Destination $installedModule -Force}
    if ($EnableRecovery) {
        if ($DeepFreezeCommandLinePassword.Length -gt 63 -or $DeepFreezeCommandLinePassword -match '[\r\n]') {throw 'The DFC password must be a single line of at most 63 characters.'}
        if ([string]::IsNullOrWhiteSpace($DeepFreezeCommandLinePassword)) { throw 'Fill in the password section before installing recovery mode.' }
        $machineFolder=Open-PersistentState
        Assert-NoRepairPreview -MachineFolder $machineFolder
        if (-not $script:TelemetryReady) { throw 'The shared telemetry writer is unavailable.' }
        if ($RunUpdaterAfterThaw) { Assert-PinnedScript -Path $Updater -AllowedHashes $ApprovedUpdaterHashes }
        Assert-PinnedScript -Path $Script09 -AllowedHashes $ApprovedScript09Hashes
    }
    $sourcePath=$PSCommandPath
    # Function scope PSCommandPath is this file, not the parent shell's command.
    if ([IO.Path]::GetFullPath($sourcePath) -ine [IO.Path]::GetFullPath($InstalledScript)) {
        Copy-Item -LiteralPath $sourcePath -Destination $InstalledScript -Force
    }
    $fileAcl=New-Object Security.AccessControl.FileSecurity
    $fileAcl.SetAccessRuleProtection($true,$false)
    $fileAcl.SetOwner((New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))
    foreach ($sidText in @('S-1-5-18','S-1-5-32-544')) {
        $rule=New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier($sidText)),'FullControl','Allow')
        $fileAcl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $InstalledScript -AclObject $fileAcl
    $arguments='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File '+(Quote-NativeArgument $InstalledScript)
    if ($EnableRecovery) {$arguments+=' -Recover'}
    $action=New-ScheduledTaskAction -Execute $Pwsh -Argument $arguments
    $startup=New-ScheduledTaskTrigger -AtStartup
    $startup.Delay='PT0S'
    $periodic=New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 1)
    $principal=New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings=New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 35) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    $settings.StartWhenAvailable=$false
    Register-ScheduledTask -TaskName $TaskName -TaskPath '\' -Action $action -Trigger @($startup,$periodic) -Principal $principal -Settings $settings -Force | Out-Null
    $noticeInstalled=Join-Path $LocalRoot 'Show-RecoveryNotice.ps1'
    if ([IO.Path]::GetFullPath($StartupNoticeSource) -ine [IO.Path]::GetFullPath($noticeInstalled)) {Copy-Item -LiteralPath $StartupNoticeSource -Destination $noticeInstalled -Force}
    $noticeConfig=Join-Path $LocalRoot 'RecoveryNotice.config.json'
    Save-Json -Path $noticeConfig -Value ([ordered]@{SchemaVersion=1;ComputerName=$env:COMPUTERNAME;Enabled=[bool]($EnableRecovery -and $EnableRepairScreen);PersistentStateRoot=$PersistentStateRoot;PersistentStateAccountDomain=$PersistentStateAccountDomain;DfcPath=(Find-Dfc);Title=$RepairScreenTitle;Message=$RepairScreenMessage})
    if ($EnableRecovery -and $EnableRepairScreen) {
        $noticeArguments='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File '+(Quote-NativeArgument $noticeInstalled)+' -ConfigPath '+(Quote-NativeArgument $noticeConfig)
        $noticeAction=New-ScheduledTaskAction -Execute $Pwsh -Argument $noticeArguments
        $noticeTrigger=New-ScheduledTaskTrigger -AtStartup
        $noticeTrigger.Delay='PT0S'
        $noticeSettings=New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        $noticeSettings.StartWhenAvailable=$false
        Register-ScheduledTask -TaskName $StartupNoticeTaskName -TaskPath '\' -Action $noticeAction -Trigger $noticeTrigger -Principal $principal -Settings $noticeSettings -Force | Out-Null
        Write-Host ('Installed startup notice task: '+$StartupNoticeTaskName)
    }
    else {Unregister-ScheduledTask -TaskName $StartupNoticeTaskName -TaskPath '\' -Confirm:$false -ErrorAction SilentlyContinue}
    Send-Event 'Installed' 'Installed the protected SYSTEM watchdog. Existing maintenance schedules were not changed.' @{TaskName=$TaskName;AutomaticRecovery=[bool]$EnableRecovery;InstalledScript=$InstalledScript;IntervalMinutes=1}
    Write-Host ('Installed task: '+$TaskName)
}

# Parse-only mode returns before importing a framework, contacting a share,
# changing tasks, running DFC, or writing any report.
if ($ValidateOnly) {
    $tokens=$null;$errors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($PSCommandPath,[ref]$tokens,[ref]$errors)
    if ($errors.Count -gt 0) {$errors | ForEach-Object {$_.Message};exit 1}
    if (-not (Test-Path -LiteralPath $RepairScreenModule -PathType Leaf)) {Write-Host 'RepairScreen.psm1 is missing.';exit 1}
    [void][Management.Automation.Language.Parser]::ParseFile($RepairScreenModule,[ref]$tokens,[ref]$errors)
    if ($errors.Count -gt 0) {$errors | ForEach-Object {$_.Message};exit 1}
    if (-not (Test-Path -LiteralPath $StartupNoticeSource -PathType Leaf)) {Write-Host 'Show-RecoveryNotice.ps1 is missing.';exit 1}
    [void][Management.Automation.Language.Parser]::ParseFile($StartupNoticeSource,[ref]$tokens,[ref]$errors)
    if ($errors.Count -gt 0) {$errors | ForEach-Object {$_.Message};exit 1}
    Write-Host 'PowerShell parser: PASS for watchdog and display module. No operational actions performed.'
    exit 0
}

try {
    if ([int][bool]$Install+[int][bool]$Uninstall+[int][bool]$Recover+[int][bool]$AcknowledgeIncident -gt 1) { throw 'Choose only one of Install, Uninstall, Recover, or AcknowledgeIncident.' }
    if ($EnableRecovery -and -not $Install) { throw 'EnableRecovery is used only with Install. Use Recover to run one recovery check.' }
    if ($Force -and ($Install -or $Uninstall -or $AcknowledgeIncident)) {throw 'Force is a one-time recovery test; do not combine it with installation, removal, or acknowledgement.'}
    if ($Force) {$Recover=$true}
    Assert-Administrator
    if ($MinimumCompletedRestarts -lt 3 -or $LookbackMinutes -lt 10 -or $CooldownHours -lt 24) { throw 'Recovery thresholds may not be weakened below three restarts, ten minutes, or a 24-hour cooldown.' }
    if ($EvidenceCollectionTimeoutSeconds -lt 30 -or $EvidenceCollectionTimeoutSeconds -gt 300) {throw 'Evidence collection time limit must be between 30 and 300 seconds.'}
    Set-PrivateLocalDirectory
    Initialize-Telemetry
    Initialize-RepairScreenModule
    if ($CollectEvidenceOnly) {
        if ($Install -or $Uninstall -or $Recover -or $Force -or $AcknowledgeIncident -or $EnableRecovery) {throw 'Evidence-only mode cannot be combined with operational modes.'}
        $expected=[guid]::Parse($ExpectedIncidentId).ToString()
        $machineFolder=Open-PersistentState
        if ($null -eq $script:State -or $script:State.IncidentId -ne $expected -or $script:State.Phase -ne 'Detected' -or $script:State.ThawAttempts -ne 0 -or $script:State.ForceAttempts -ne 0) {throw 'Evidence-only child requires the matching pre-thaw incident receipt.'}
        $script:Incident=$expected
        $script:Boot=[datetime](Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime
        $loop=[pscustomobject]@{Detected=($script:State.DetectedRestarts -ge $MinimumCompletedRestarts);CompletedRestarts=$script:State.DetectedRestarts;Requests=@();Note='Original request details remain in parent detection telemetry.'}
        Collect-IncidentEvidence -MachineFolder $machineFolder -Loop $loop
        exit 0
    }
    $script:Mutex=New-Object Threading.Mutex($false,'Global\Compton_RebootRecoveryWatchdog')
    try {$script:MutexHeld=$script:Mutex.WaitOne(0)} catch [Threading.AbandonedMutexException] {$script:MutexHeld=$true}
    if (-not $script:MutexHeld) {Write-Host 'DeploymentDeferred: Another watchdog instance is active.';exit 0}
    $script:Boot=[datetime](Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime
    if ($Uninstall) {
        if ((Get-FreezeState) -ne 'Thawed') {throw 'Uninstall while Deep Freeze is thawed so task removal persists.'}
        Unregister-ScheduledTask -TaskName $StartupNoticeTaskName -TaskPath '\' -Confirm:$false -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath '\' -Confirm:$false -ErrorAction SilentlyContinue
        if (@(Get-ScheduledTask -ErrorAction Stop | Where-Object {$_.TaskName -in @($TaskName,$StartupNoticeTaskName) -and $_.TaskPath -eq '\'}).Count -gt 0) {throw 'Watchdog task removal did not verify.'}
        Send-Event 'Uninstalled' 'Watchdog task removed. Existing incident evidence and the protected script were retained.'
        exit 0
    }

    if ($Deploy) {
        if ($Install -or $Uninstall -or $Recover -or $Force -or $AcknowledgeIncident -or $CollectEvidenceOnly -or $EnableRecovery) {throw 'Deploy cannot be combined with other operational modes.'}
        if ([string]::IsNullOrWhiteSpace($DeepFreezeCommandLinePassword) -or [string]::IsNullOrWhiteSpace($PersistentStateRoot)) {
            Write-Host 'DeploymentDeferred: Configure the password and persistent share in the source watchdog.'
            exit 0
        }
        try {$deploymentFreeze=Get-FreezeState} catch {Write-Host 'DeploymentDeferred: DFC state is unavailable.';exit 0}
        if ($deploymentFreeze -ne 'Thawed') {Write-Host 'DeploymentDeferred: Install while thawed.';exit 0}
        try {$deploymentFolder=Open-PersistentState} catch {Write-Host 'DeploymentDeferred: Persistent folder is unavailable or invalid.';exit 0}
        if ($null -ne $script:State -and $script:State.Phase -ne 'Completed') {
            Write-Host 'DeploymentDeferred: An active or latched incident requires review before replacement.'
            exit 0
        }
        $Install=$true
        $EnableRecovery=$true
    }
    if ($Install) {Install-Watchdog;exit 0}
    if ($AcknowledgeIncident) {
        if (-not $script:TelemetryReady) {throw 'The telemetry writer is required.'}
        [void](Open-PersistentState)
        if ($null -eq $script:State) {throw 'No persistent incident exists.'}
        if ($script:State.Phase -notin @('NeedsAttention','Completed')) {throw 'An active recovery transaction cannot be acknowledged.'}
        $script:Incident=$script:State.IncidentId
        Assert-PinnedScript -Path $Script09 -AllowedHashes $ApprovedScript09Hashes
        Complete-Recovery
        exit 0
    }
    $normal=Test-NormalMaintenance -Now (Get-Date) -Start ([timespan]$NormalSundayStart) -End ([timespan]$NormalSundayEnd)
    $earlyMachineFolder=$null
    $resumeExisting=$false
    if ($Recover -and -not $normal) {
        # Read the durable transaction before expensive history queries. An
        # existing recovery needs continuation, not a new detection decision.
        $earlyMachineFolder=Open-StateWithRetry
        $resumeExisting=($null -ne $script:State -and $script:State.Phase -notin @('Completed','NeedsAttention'))
        if ($resumeExisting) {
            $script:Incident=$script:State.IncidentId
            if ($script:State.Phase -in @('ThawRequested','AfterThaw','UpdatingScripts','AwaitingEvaluation','ClosingServices','ForceRequested')) {Update-RepairDisplay}
        }
    }
    if ($resumeExisting) {
        $loop=[pscustomobject]@{Detected=($script:State.DetectedRestarts -ge $MinimumCompletedRestarts);CompletedRestarts=$script:State.DetectedRestarts;Requests=@();Note='Resuming existing incident; original reboot evidence retained in incident diagnostics.'}
    }
    else {$loop=Get-RebootEvidence}
    $script:LoopEvidence=$loop
    if (-not $Recover) {
        $status='AuditNoLoop'
        if ($loop.Detected -and -not $normal) {$status='LoopDetected'}
        if ($normal) {$status='AuditSundayExcluded'}
        $persistentStatus='NotConfigured'
        if (-not [string]::IsNullOrWhiteSpace($PersistentStateRoot)) {
            try {[void](Open-PersistentState);$persistentStatus='AvailableInThisExecutionContext'}
            catch {$persistentStatus='Unavailable: '+(Protect-Text $_.Exception.Message)}
        }
        $freezeStatus='Unknown'
        try {$freezeStatus=Get-FreezeState} catch { }
        $edition=Get-ComptonRepairEdition
        $message='Diagnostic-only check. No task, service, Deep Freeze setting, or reboot was changed.'
        Send-Event $status $message @{CompletedMaintenanceRestarts=$loop.CompletedRestarts;Threshold=$MinimumCompletedRestarts;LookbackMinutes=$LookbackMinutes;NormalSundayWindow=$normal;Requests=$loop.Requests;FreezeState=$freezeStatus;PersistentStorage=$persistentStatus;ExecutionIdentity=[Security.Principal.WindowsIdentity]::GetCurrent().Name;TelemetryReady=$script:TelemetryReady;WindowsEdition=$edition;RepairScreenEnabled=$EnableRepairScreen}
        if ($loop.Detected -and -not $normal) {
            $script:State=New-RecoveryState -Loop $loop
            $script:Incident=$script:State.IncidentId
            $script:StatePath=Join-Path $LocalRoot 'AuditRecoveryState.json'
            Save-State
            Collect-IncidentEvidence -MachineFolder $LocalRoot -Loop $loop
            Send-Event 'AuditEvidenceCaptured' 'Diagnostic-only evidence captured locally. Enable persistent storage before enabling automated recovery.' @{BundlePath=$script:State.BundlePath}
        }
        exit 0
    }
    if (-not $script:TelemetryReady) {throw 'Recovery requires the shared Elastic telemetry writer.'}
    if ($normal) {Send-Event 'DeferredSundayMaintenance' 'Normal Sunday maintenance is excluded from automatic reboot-loop repair.';exit 0}
    $machineFolder=$earlyMachineFolder
    if ([string]::IsNullOrEmpty($machineFolder)) {$machineFolder=Open-StateWithRetry}
    Assert-NoRepairPreview -MachineFolder $machineFolder
    if ($null -ne $script:State -and $script:State.Phase -in @('Completed','NeedsAttention')) {
        $script:Incident=$script:State.IncidentId
        if ($script:State.Phase -eq 'NeedsAttention' -or (Get-Date).ToUniversalTime() -lt (([datetimeoffset]$script:State.StartedUtc).UtcDateTime).AddHours($CooldownHours)) {
            Send-Event ('Cooldown-'+$script:State.Phase) 'The incident is complete or latched for technician attention. Automatic retries are blocked; completed incidents retain a 24-hour cooldown.' @{IncidentId=$script:Incident;Phase=$script:State.Phase;Failure=$script:State.FailureMessage;BundlePath=$script:State.BundlePath}
            exit 0
        }
        $script:State=$null
    }
    if ($null -eq $script:State) {
        if (-not $loop.Detected -and -not $Force) {Send-Event 'NoLoopDetected' 'No repeated maintenance reboot loop meets the configured threshold.' @{CompletedMaintenanceRestarts=$loop.CompletedRestarts};exit 0}
        if ($Force) {
            # A frozen boot discards local state; an enabled recovery task must resume
            # this transaction using its persistent receipt. Never schedule Force.
            $resumeTask=Get-ScheduledTask -TaskName $TaskName -TaskPath '\' -ErrorAction Stop
            if ($resumeTask.State -eq 'Disabled' -or -not $resumeTask.Settings.Enabled -or
                (Get-TaskArgumentText -Task $resumeTask) -notmatch '(?i)(?:^|\s)-Recover(?:\s|$)') {
                throw 'Install with -Install -EnableRecovery before testing with -Force.'
            }
            if ((Get-FileHash -LiteralPath $InstalledScript -Algorithm SHA256).Hash -ne
                (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash) {
                throw 'Reinstall this configured version with -Install -EnableRecovery before testing.'
            }
            Send-Event 'ManualForceRequested' 'Administrator requested a recovery test without the repeated-reboot threshold. Sunday exclusion, cooldown, incident latch, and reboot limits remain active.' @{ActualCompletedRestarts=$loop.CompletedRestarts}
        }
        $script:State=New-RecoveryState -Loop $loop
        Save-State
        $script:Incident=$script:State.IncidentId
        if ($Force) {Send-Event 'ManualTestStarted' 'Manual recovery test started; no reboot loop is asserted.' $loop}
        if (-not $Force) {Send-Event 'LoopDetected' 'Repeated completed Script 07 maintenance restarts detected outside the Sunday maintenance window.' $loop}
    }
    $script:Incident=$script:State.IncidentId
    Run-Recovery -MachineFolder $machineFolder -Loop $loop
    exit 0
}
catch {
    $safeMessage=$_.Exception.Message
    if (-not [string]::IsNullOrEmpty($DeepFreezeCommandLinePassword)) {$safeMessage=$safeMessage.Replace($DeepFreezeCommandLinePassword,'[REDACTED]')}
    if ($null -ne $script:State -and $null -ne $script:StatePath -and ($Recover -or $AcknowledgeIncident)) {
        $script:State.Phase='NeedsAttention';$script:State.FailureMessage=$safeMessage
        try {Save-State} catch { }
    }
    try {Send-Event 'NeedsAttention' $safeMessage @{Phase=$(if ($null -ne $script:State) {$script:State.Phase} else {'Preflight'});NoAutomaticRetry=$true;LoopEvidence=$script:LoopEvidence}}
    catch {Write-Host ('Recovery refused: '+$safeMessage);Write-Warning 'Telemetry could not be written. No further action was requested.'}
    exit 2
}
finally {
    if ($null -ne $script:EvidenceWriter) {$script:EvidenceWriter.Dispose()}
    if ($script:MutexHeld) {try {$script:Mutex.ReleaseMutex()} catch { }}
    if ($null -ne $script:Mutex) {$script:Mutex.Dispose()}
}
