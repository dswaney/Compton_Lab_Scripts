# Offline branch tests: extracted functions, mocked Windows APIs. No host registry/task/BitLocker changes.
$ErrorActionPreference='Stop'
if ($PSVersionTable.PSVersion.Major -lt 7 -or -not $IsLinux) { throw 'Run only in an isolated Linux PowerShell 7 process; Windows APIs are mocked.' }
$root=$PSScriptRoot
$tokens=$null;$errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root '15_Sunday_Update_Closeout.ps1'),[ref]$tokens,[ref]$errors)
if($errors){throw ($errors|Out-String)}
$functions=$ast.FindAll({param($n)$n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$false)
foreach($f in $functions){. ([scriptblock]::Create($f.Extent.Text))}
$temp=Join-Path ([IO.Path]::GetTempPath()) ('compton-secureboot-test-'+[guid]::NewGuid())
New-Item -ItemType Directory $temp|Out-Null
$env:ProgramData=$temp;$env:WINDIR=Join-Path $temp 'Windows';$env:SystemDrive='C:'
$ScriptsRoot=$temp;$LogDirectory=$temp;$SecureBootAuditOnly=$false;$SecureBootStartup=$false
New-Item -ItemType Directory (Join-Path $env:WINDIR 'System32') -Force|Out-Null
# Native reboot call is replaced with a test executable that only exits successfully.
if($IsLinux){$shutdown=Join-Path $env:WINDIR 'System32/shutdown.exe';"#!/bin/sh`nexit 0"|Set-Content $shutdown;& chmod +x $shutdown}
function Get-Date {return $script:now}
function Get-CimInstance {param($ClassName,$Filter,$ErrorAction)
 switch($ClassName){
 'Win32_ComputerSystem'{[pscustomobject]@{Name='PILOT'}}
 'Win32_ComputerSystemProduct'{[pscustomobject]@{UUID='test-uuid'}}
 'Win32_Volume'{[pscustomobject]@{DeviceID='volume-guid'}}
 'Win32_OperatingSystem'{[pscustomobject]@{LastBootUpTime=$script:boot}}
 'Win32_Service'{if($script:deepFreeze){[pscustomobject]@{Name='DFServ'}}}
 'Win32_SystemDriver'{}
 }
}
function Confirm-SecureBootUEFI {param($ErrorAction)if($script:firmwareDenied){throw 'Unable to set proper privileges. Access was denied.'};return $true}
function Get-ItemProperty {param($LiteralPath,$ErrorAction)
 if($LiteralPath.EndsWith('\State')){return [pscustomobject]@{UEFISecureBootEnabled=1}}
 if($LiteralPath.EndsWith('\Servicing')){[pscustomobject]@{UEFICA2023Status=$script:status;UEFICA2023Error=$script:servicingError;WindowsUEFICA2023Capable=2}}
 else{[pscustomobject]@{AvailableUpdates=$script:mask;AvailableUpdatesPolicy=0}}
}
function Get-BitLockerVolume {param($MountPoint,$ErrorAction)[pscustomobject]@{ProtectionStatus=$script:protection;VolumeStatus='FullyEncrypted';KeyProtector=@([pscustomobject]@{KeyProtectorType='RecoveryPassword'})}}
function Initialize-SecureBootStateDirectory {$dir=Join-Path $env:ProgramData 'Compton-SecureBootMaintenance';New-Item -ItemType Directory -Path $dir -Force|Out-Null;return $dir}
function Test-Window {if((Get-Date).DayOfWeek -ne 'Sunday'){throw 'Outside Sunday'};if($script:frozen){throw 'Frozen'}}
function Register-SecureBootResumeTask {$script:registered++}
function Remove-SecureBootResumeTask {$script:removed++}
function Get-ScheduledTask {param($TaskPath,$TaskName,$ErrorAction)[pscustomobject]@{State='Ready'}}
function Suspend-BitLocker {param($MountPoint,$RebootCount,$ErrorAction)if($RebootCount -ne 2){throw 'Unbounded suspension'};$script:suspended++;$script:protection='Off'}
function Resume-BitLocker {param($MountPoint,$ErrorAction)$script:resumed++;$script:protection='On'}
function New-ItemProperty {param($LiteralPath,$Name,$PropertyType,$Value,[switch]$Force,$ErrorAction)$script:writes++;if($Value -ne 0x5944){throw 'Unexpected update flags'};$script:mask=$Value}
function Start-ScheduledTask {param($TaskPath,$TaskName,$ErrorAction)$script:started++;if($script:completeOnTask){$script:status='Updated';$script:mask=0x4000}else{$script:mask=0x4100}}
function Start-Sleep {param($Seconds)}
function Get-WinEvent {param($FilterHashtable,$ErrorAction)if($script:eventReadFails){throw 'Denied'};if($script:recoveryEvents){[pscustomobject]@{Id=24652}}}
function Reset-Test {
 Remove-Item (Join-Path $temp 'Compton-SecureBootMaintenance') -Recurse -Force -ErrorAction SilentlyContinue
 $script:firmwareDenied=$false;$script:now=[datetime]'2026-10-18T10:00:00';$script:status='NotStarted';$script:mask=0;$script:servicingError=0;$script:protection='On';$script:deepFreeze=$false;$script:frozen=$false;$script:boot=[datetime]'2026-10-18T09:00:00';$script:completeOnTask=$false;$script:recoveryEvents=$false;$script:eventReadFails=$false
 $script:writes=0;$script:suspended=0;$script:resumed=0;$script:registered=0;$script:removed=0;$script:started=0
 $script:SecureBootAuditOnly=$false;$script:SecureBootStartup=$false
 $p=[ordered]@{SecureBootMaintenance=@{Enabled=$true;PilotComputers=@('PILOT');RecoveryKeyVerifiedComputers=@('PILOT');ThawedThroughVerificationComputers=@('PILOT');MinimumDeepFreezeVersion='10.20.220.5844';LatestStart='18:00';RebootCutoff='20:15';MaximumReboots=3}}
 $p|ConvertTo-Json -Depth 6|Set-Content (Join-Path $temp 'Maintenance.Policy.json')
}
$script:checks=0
function Check($Condition,[string]$Message){if(!$Condition){throw $Message};$script:checks++}
try{
 Reset-Test;$script:status='Updated';$script:mask=0x4000;$r=Invoke-SecureBootMaintenance
 Check ($r.Ready -and $script:writes -eq 0 -and $script:suspended -eq 0) 'AlreadyUpdated must not mutate certificate/protection state'
 Reset-Test;$script:SecureBootAuditOnly=$true;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'NeedsUpdate' -and $script:writes -eq 0 -and $script:registered -eq 0) 'Audit must not remediate'
 Reset-Test;$script:frozen=$true;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:writes -eq 0 -and $script:suspended -eq 0) 'Frozen must block'
 Reset-Test;$script:deepFreeze=$true;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:writes -eq 0) 'Unknown Deep Freeze version must block encrypted remediation'
 Reset-Test;$p=Get-Content (Join-Path $temp 'Maintenance.Policy.json') -Raw|ConvertFrom-Json;$p.SecureBootMaintenance.RecoveryKeyVerifiedComputers=@();$p|ConvertTo-Json -Depth 6|Set-Content (Join-Path $temp 'Maintenance.Policy.json');$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:suspended -eq 0) 'Missing recovery verification must block'
 Reset-Test;$script:servicingError=1;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:writes -eq 0) 'Servicing errors must block'
 Reset-Test;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'RebootRequested' -and $script:suspended -eq 1 -and $script:writes -eq 1) ('Initial update/reboot failed: '+($r|ConvertTo-Json -Depth 4))
 $r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'RebootAlreadyRequested' -and $script:writes -eq 1 -and $script:suspended -eq 1) ('Repeated pass must not reset flags or resuspend: '+($r|ConvertTo-Json -Depth 5))
 $script:boot=[datetime]'2026-10-18T10:05:00';$script:completeOnTask=$true;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'RebootRequested' -and $script:resumed -eq 1 -and $script:protection -eq 'On') ('Completion must resume and request protected boot: '+$r.Message)
 $script:boot=[datetime]'2026-10-18T10:10:00';$r=Invoke-SecureBootMaintenance
 Check ($r.Ready -and $r.Status -eq 'Updated_ProtectedBootVerified' -and $script:removed -eq 1) ('Protected boot must finish transaction: '+$r.Message)
 Reset-Test;$script:protection='Off';$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:resumed -eq 0) 'Do not take over unrelated suspension'
 Reset-Test;$script:SecureBootStartup=$true;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'NoActiveTransaction' -and $script:suspended -eq 0) 'Orphan startup must not initiate updates'
 Reset-Test;$null=Invoke-SecureBootMaintenance;$script:boot=[datetime]'2026-10-18T10:05:00';$script:completeOnTask=$true;$null=Invoke-SecureBootMaintenance;$script:boot=[datetime]'2026-10-18T10:10:00';$script:recoveryEvents=$true;$r=Invoke-SecureBootMaintenance
 Check (!$r.Ready -and $r.Status -eq 'BlockedOrFailed') 'Recovery on verification boot must not be declared ready'
 Reset-Test;$null=Invoke-SecureBootMaintenance;$script:boot=[datetime]'2026-10-18T10:05:00';$script:completeOnTask=$true;$null=Invoke-SecureBootMaintenance;$script:boot=[datetime]'2026-10-18T10:10:00';$script:eventReadFails=$true;$r=Invoke-SecureBootMaintenance
 Check (!$r.Ready -and $r.Message -match 'Unable to read BitLocker') 'Unreadable event log must fail verification'
 Reset-Test;$script:now=[datetime]'2026-10-19T10:00:00';$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:suspended -eq 0 -and $script:writes -eq 0) 'Weekday must never initiate servicing'
 Reset-Test;$script:now=[datetime]'2026-10-18T18:01:00';$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:suspended -eq 0) 'Late new update must be blocked'
 Reset-Test;$null=Invoke-SecureBootMaintenance;$script:now=[datetime]'2026-10-18T20:15:00';$script:boot=[datetime]'2026-10-18T20:00:00';$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:resumed -eq 1 -and $script:protection -eq 'On') 'Deadline must clean up owned suspension without reboot'
 Reset-Test;$null=Invoke-SecureBootMaintenance;$script:now=[datetime]'2026-10-19T09:00:00';$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $script:resumed -eq 1) 'Expired transaction must clean up protection'
 Reset-Test;$null=Invoke-SecureBootMaintenance;$file=Join-Path $temp 'Compton-SecureBootMaintenance/state.json';$s=Get-Content $file -Raw|ConvertFrom-Json;$s.RebootsRequested=3;$s.LastRequestedBoot='other';$s|ConvertTo-Json -Depth 8|Set-Content $file;$script:boot=[datetime]'2026-10-18T10:05:00';$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and $r.Message -match 'budget exhausted' -and $script:resumed -eq 1) 'Reboot budget must stop and resume protection'
 $telemetryPath=Join-Path $temp 'Maintenance-Telemetry.ndjson'
 $reader=[IO.File]::Open($telemetryPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
 try{$delivery=Write-CloseoutTelemetry -Event @{Test='ReaderSharing'} -Directory $temp;Check ($delivery -eq 'Appended') 'Telemetry append must coexist with reader'}finally{$reader.Dispose()}
 $exclusive=[IO.File]::Open($telemetryPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
 try{$delivery=Write-CloseoutTelemetry -Event @{Test='ExclusiveLock'} -Directory $temp -WarningAction SilentlyContinue;Check ($delivery -like 'Pending: *' -and (Test-Path -LiteralPath $delivery.Substring(9))) 'Exclusive telemetry lock must preserve a pending event'}finally{$exclusive.Dispose()}
 Reset-Test;$script:status='Updated';$script:mask=0x4000;$script:firmwareDenied=$true;$script:SecureBootAuditOnly=$true;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'Updated_FirmwareQueryRestricted' -and !$r.Ready -and $r.Snapshot.RegistryDeploymentComplete -and $r.Snapshot.RegistrySecureBootEnabled -and !$r.Snapshot.FirmwareQueryVerified -and $null -eq $r.Snapshot.Enabled -and $r.BitLockerProtection -eq 'On' -and $script:writes -eq 0) 'Restricted firmware audit must retain registry/volume evidence and avoid claiming full verification'
 Reset-Test;$script:status='Updated';$script:mask=0x4000;$script:firmwareDenied=$true;$r=Invoke-SecureBootMaintenance
 Check ($r.Status -eq 'BlockedOrFailed' -and !$r.Ready -and $script:writes -eq 0 -and $script:suspended -eq 0 -and $script:registered -eq 0) 'Firmware restriction must continue to block remediation even with Updated registry status'
 "Passed $script:checks offline branch tests. Windows pilot still required."
}finally{Remove-Item $temp -Recurse -Force}
