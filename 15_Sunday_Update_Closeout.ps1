#requires -Version 5.1
#requires -RunAsAdministrator
# ScriptVersion: 1.1.4
# LastUpdated: 2026-10-19
[CmdletBinding()]
param([string]$ScriptsRoot='C:\Scripts',[switch]$VerifyOnly,[string]$LogDirectory='C:\Logs',[switch]$SecureBootOnly,[switch]$SecureBootAuditOnly,[switch]$SecureBootStartup)
$ErrorActionPreference='Stop'
New-Item -Path $LogDirectory -ItemType Directory -Force|Out-Null
$start=Get-Date;$result=[ordered]@{ComputerName=$env:COMPUTERNAME;ScriptVersion='1.1.4';StartTime=$start.ToString('o');Status='Running';ReadyToFreeze=$false;RemainingUpdates=$null;RebootRequired=$null;FailureMessage=$null;FreezeScheduleControlled=$false}
$log=Join-Path $LogDirectory ($env:COMPUTERNAME+'-15_Sunday_Update_Closeout-'+$start.ToString('yyyyMMdd-HHmmss')+'.log')
function Log([string]$Text){$line=(Get-Date -Format s)+' '+$Text;$line|Add-Content -LiteralPath $log -Encoding UTF8;Write-Host $line}
function Test-Window {
 $now=Get-Date
 if($now.DayOfWeek -ne 'Sunday' -or $now.TimeOfDay -ge [timespan]'20:45:00'){throw 'Outside Sunday closeout window; no update or reboot action permitted.'}
 $df=@(Get-CimInstance Win32_Service -ErrorAction Stop|Where-Object Name -in @('DFServ','DeepFrz'))
 $driver=@(Get-CimInstance Win32_SystemDriver -ErrorAction Stop|Where-Object Name -match '^(DfDiskLo|DeepFrz)')
 if($df.Count -or $driver.Count){
  $paths=@("$env:WINDIR\System32\DFC.exe","$env:WINDIR\SysWOW64\DFC.exe","${env:ProgramFiles(x86)}\Faronics\Deep Freeze\Install C-0\DFC.exe")
  $exe=$paths|Where-Object {Test-Path -LiteralPath $_}|Select-Object -First 1
  if(!$exe){throw 'Deep Freeze detected; DFC state query unavailable. Closeout deferred.'}
  $p=Start-Process -FilePath $exe -ArgumentList @('get','/ISFROZEN') -PassThru -WindowStyle Hidden
  if(!$p.WaitForExit(10000)){try{$p.Kill()}catch{};throw 'Deep Freeze state query timed out.'}
  $p.Refresh();$code=$p.ExitCode;$p.Dispose()
  if($code -ne 0){throw "Deep Freeze is frozen or unknown (DFC=$code); closeout deferred."}
 }
}
function Run-Script([string]$Name,[string[]]$Extra=@()){
 Test-Window
 $path=Join-Path $ScriptsRoot $Name
 if(!(Test-Path -LiteralPath $path)){throw "Missing script: $path"}
 $exe=Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
 & $exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $path @Extra | Out-Host
 $code=$LASTEXITCODE;Log "$Name exit=$code";return $code
}
function Servicing-RebootRequired {
 $system=New-Object -ComObject Microsoft.Update.SystemInfo
 return ([bool]$system.RebootRequired -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'))
}
# Open explicitly with read/write/delete sharing so Elastic readers can coexist.
# An exclusive external lock still requires retries or a separate durable spool file.
function Write-CloseoutTelemetry {
 param([object]$Event,[string]$Directory)
 $line=($Event|ConvertTo-Json -Depth 12 -Compress)+[Environment]::NewLine
 $bytes=(New-Object System.Text.UTF8Encoding($false)).GetBytes($line)
 $path=Join-Path $Directory 'Maintenance-Telemetry.ndjson'
 for($attempt=0;$attempt -lt 5;$attempt++){
  $stream=$null
  try{
   $stream=[IO.File]::Open($path,[IO.FileMode]::Append,[IO.FileAccess]::Write,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
   $stream.Write($bytes,0,$bytes.Length);$stream.Flush()
   return 'Appended'
  }catch [IO.IOException]{
   if($attempt -lt 4){Start-Sleep -Milliseconds 250}
  }finally{if($null -ne $stream){$stream.Dispose()}}
 }
 $pending=Join-Path $Directory ('SecureBoot-Telemetry-Pending-'+[guid]::NewGuid().ToString('N')+'.ndjson')
 [IO.File]::WriteAllBytes($pending,$bytes)
 Write-Warning ('Shared telemetry file remained locked. Event saved locally to '+$pending+'. This separate file is not guaranteed to be collected by Elastic.')
 return ('Pending: '+$pending)
}

# Secure Boot coordinator. No certificate blobs, TPM clears, PCR changes or DBX revocations.
function Get-SecureBootSnapshot {
 $base='HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
 $b=Get-ItemProperty -LiteralPath $base -ErrorAction Stop
 $v=Get-ItemProperty -LiteralPath ($base+'\Servicing') -ErrorAction SilentlyContinue
 $recorded=Get-ItemProperty -LiteralPath ($base+'\State') -ErrorAction SilentlyContinue
 $recordedEnabled=$null
 if($null -ne $recorded.UEFISecureBootEnabled){$recordedEnabled=([int]$recorded.UEFISecureBootEnabled -eq 1)}
 $enabled=$null;$firmwareVerified=$false;$firmwareError=$null
 try{$enabled=[bool](Confirm-SecureBootUEFI -ErrorAction Stop);$firmwareVerified=$true}catch{$firmwareError=$_.Exception.Message}
 $errorCode=0;if($null -ne $v.UEFICA2023Error){$errorCode=[long]$v.UEFICA2023Error}
 $mask=0;if($null -ne $b.AvailableUpdates){$mask=[long]$b.AvailableUpdates}
 $registryComplete=($v.UEFICA2023Status -eq 'Updated' -and $errorCode -eq 0 -and $mask -in @(0,0x4000))
 [pscustomobject]@{Enabled=$enabled;FirmwareQueryVerified=$firmwareVerified;FirmwareQueryError=$firmwareError;RegistrySecureBootEnabled=$recordedEnabled;RegistryDeploymentComplete=$registryComplete;Status=[string]$v.UEFICA2023Status;ErrorCode=$errorCode;ErrorEvent=$v.UEFICA2023ErrorEvent;AvailableUpdates=$mask;AvailableUpdatesHex=('0x{0:X}' -f $mask);Capable=$v.WindowsUEFICA2023Capable;Complete=($firmwareVerified -and $enabled -eq $true -and $registryComplete)}
}
function Test-SecureBootPolicyHost($Policy,[string]$List,[string]$Computer) {
 return @($Policy.$List | Where-Object { [string]$_ -ieq $Computer }).Count -gt 0
}
function Assert-SecureBootWindow($Policy,[switch]$NewUpdate) {
 Test-Window # Includes live DFC /ISFROZEN proof, not just a calendar assumption.
 $now=Get-Date
 $cutoff=[timespan]::Parse([string]$Policy.RebootCutoff)
 if($now.AddSeconds(60).TimeOfDay -ge $cutoff){throw 'Secure Boot reboot cutoff reached; manual follow-up required.'}
 if($NewUpdate -and $now.TimeOfDay -ge [timespan]::Parse([string]$Policy.LatestStart)){throw 'Too late to begin a new Secure Boot update this Sunday.'}
}
function Initialize-SecureBootStateDirectory {
 $dir=Join-Path $env:ProgramData 'Compton-SecureBootMaintenance'
 if(Test-Path -LiteralPath $dir){
  $item=Get-Item -LiteralPath $dir -Force
  if($item.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Secure Boot state directory must not be a reparse point.'}
 }else{New-Item -ItemType Directory -Path $dir -Force|Out-Null}
 $acl=New-Object System.Security.AccessControl.DirectorySecurity
 $acl.SetAccessRuleProtection($true,$false)
 $acl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
 foreach($sid in @('S-1-5-18','S-1-5-32-544')){
  $identity=New-Object System.Security.Principal.SecurityIdentifier($sid)
  $rule=New-Object System.Security.AccessControl.FileSystemAccessRule($identity,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
  $acl.AddAccessRule($rule)
 }
 Set-Acl -LiteralPath $dir -AclObject $acl -ErrorAction Stop
 return $dir
}
function Save-SecureBootState($State,[string]$Path) {
 $temp=$Path+'.tmp'
 $State|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $temp -Encoding UTF8
 Move-Item -LiteralPath $temp -Destination $Path -Force
}
function Register-SecureBootResumeTask {
 $exe=Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
 $path=Join-Path $ScriptsRoot '15_Sunday_Update_Closeout.ps1'
 if($path.Contains('"') -or $ScriptsRoot.Contains('"')){throw 'Invalid scripts path.'}
 $action=New-ScheduledTaskAction -Execute $exe -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$path+'" -ScriptsRoot "'+$ScriptsRoot+'" -SecureBootOnly -SecureBootStartup')
 $trigger=New-ScheduledTaskTrigger -AtStartup;$trigger.Delay='PT2M'
 $deadline=New-ScheduledTaskTrigger -Daily -At '20:15'
 $settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew
 $principal=New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
 Register-ScheduledTask -TaskName 'Compton Secure Boot Update Resume' -Action $action -Trigger @($trigger,$deadline) -Settings $settings -Principal $principal -Force|Out-Null
}
function Remove-SecureBootResumeTask {
 if(Get-ScheduledTask -TaskName 'Compton Secure Boot Update Resume' -ErrorAction SilentlyContinue){Unregister-ScheduledTask -TaskName 'Compton Secure Boot Update Resume' -Confirm:$false -ErrorAction Stop}
}
function Resume-OwnedSecureBootProtection($State,[string]$StatePath) {
 if($State -and $State.OwnsSuspension){
  Resume-BitLocker -MountPoint $env:SystemDrive -ErrorAction Stop|Out-Null
  $volume=Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
  if([string]$volume.ProtectionStatus -ne 'On'){throw 'Unable to verify BitLocker protection resumed.'}
  $State.OwnsSuspension=$false
  $State.Phase='ProtectionResumed';$State.ResumeBoot=(Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().Ticks.ToString()
  Save-SecureBootState $State $StatePath
 }
}
function Request-SecureBootReboot($State,$Policy,[string]$StatePath) {
 Assert-SecureBootWindow $Policy
 $boot=(Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().Ticks.ToString()
 if($State.LastRequestedBoot -eq $boot){return 'RebootAlreadyRequested'}
 $max=[int]$Policy.MaximumReboots
 if($max -lt 1 -or $max -gt 3){throw 'MaximumReboots must be between one and three.'}
 if([int]$State.RebootsRequested -ge $max){throw 'Secure Boot reboot budget exhausted; no further automatic reboot.'}
 $State.RebootsRequested=[int]$State.RebootsRequested+1;$State.LastRequestedBoot=$boot
 Save-SecureBootState $State $StatePath
 Assert-SecureBootWindow $Policy
 & "$env:WINDIR\System32\shutdown.exe" /r /t 30 /c 'Compton Sunday Secure Boot certificate maintenance'|Out-Null
 if($LASTEXITCODE -ne 0){throw "Secure Boot reboot request failed: $LASTEXITCODE"}
 return 'RebootRequested'
}
function Invoke-SecureBootMaintenance {
 $report=[ordered]@{EventType='maintenance.secure_boot';ComputerName=$env:COMPUTERNAME;ScriptName='15_Sunday_Update_Closeout.ps1';ScriptVersion='1.1.4';'@timestamp'=(Get-Date).ToUniversalTime().ToString('o');Status='Checking';Ready=$false;Message=$null;Snapshot=$null;BitLockerProtection=$null;DeepFreezeVersion=$null;Phase=$null;RebootsRequested=0;FreezeScheduleControlled=$false}
 $state=$null;$statePath=$null;$held=$false;$mutex=$null
 try{
  if(![Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem){throw 'Run the coordinator in native 64-bit Windows PowerShell.'}
  $computer=(Get-CimInstance Win32_ComputerSystem).Name
  $uuid=(Get-CimInstance Win32_ComputerSystemProduct).UUID
  $report.ComputerName=$computer
  $snapshot=Get-SecureBootSnapshot;$report.Snapshot=$snapshot
  $volume=Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
  $report.BitLockerProtection=[string]$volume.ProtectionStatus
  $volumeIdentity=[string](Get-CimInstance Win32_Volume -Filter ("DriveLetter='"+$env:SystemDrive+"'") -ErrorAction Stop).DeviceID
  if(!$volumeIdentity){throw 'Unable to establish a stable OS volume identity.'}
  $encrypted=([string]$volume.VolumeStatus -ne 'FullyDecrypted')
  if($encrypted -and [string]$volume.VolumeStatus -ne 'FullyEncrypted'){throw 'Encryption/decryption is in progress; defer Secure Boot servicing.'}
  $dfPath=Join-Path $env:WINDIR 'System32\drivers\DeepFrz.sys'
  $dfServices=@(Get-CimInstance Win32_Service -ErrorAction Stop|Where-Object Name -in @('DFServ','DeepFrz'))
  $dfDrivers=@(Get-CimInstance Win32_SystemDriver -ErrorAction Stop|Where-Object Name -match '^(DfDiskLo|DeepFrz)')
  $dfDetected=((Test-Path -LiteralPath $dfPath) -or $dfServices.Count -gt 0 -or $dfDrivers.Count -gt 0)
  if(Test-Path -LiteralPath $dfPath){$report.DeepFreezeVersion=((Get-Item -LiteralPath $dfPath).VersionInfo.FileVersion -replace ',\s*','.').Trim()}
  $existing=Join-Path $env:ProgramData 'Compton-SecureBootMaintenance\state.json'
  # Audit mode never takes ownership, changes protection, starts tasks, or writes update flags.
  if($SecureBootAuditOnly){
   if(Test-Path -LiteralPath $existing){$report.Status='Audit_PendingState';$report.Message='An earlier maintenance transaction exists; inspect normal coordinator output.'}
   elseif(!$snapshot.FirmwareQueryVerified){
    if($snapshot.RegistryDeploymentComplete -and $snapshot.RegistrySecureBootEnabled -eq $true){$report.Status='Updated_FirmwareQueryRestricted'}else{$report.Status='FirmwareQueryRestricted'}
    $report.Message='Registry evidence is reported separately. Firmware verification is unavailable: '+$snapshot.FirmwareQueryError+'. Remediation and readiness require a successful firmware query.'
   }
   elseif($snapshot.Complete){$report.Status='AlreadyUpdated';$report.Ready=(!$encrypted -or [string]$volume.ProtectionStatus -eq 'On')}
   else{$report.Status='NeedsUpdate'}
   return $report
  }
  $mutex=New-Object Threading.Mutex($false,'Global\Compton_SecureBootMaintenance')
  try{$held=$mutex.WaitOne(0)}catch [Threading.AbandonedMutexException]{$held=$true}
  if(!$held){throw 'Another Secure Boot coordinator is running.'}
  $stateDir=Initialize-SecureBootStateDirectory;$statePath=Join-Path $stateDir 'state.json'
  if(Test-Path -LiteralPath $statePath){
   if((Get-Item -LiteralPath $statePath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'State file must not be a reparse point.'}
   $state=Get-Content -LiteralPath $statePath -Raw|ConvertFrom-Json
   if($state.SchemaVersion -ne 1 -or $state.ComputerName -ine $computer -or $state.UUID -ne $uuid -or $state.VolumeId -ne $volumeIdentity){$state=$null;throw 'Secure Boot state identity mismatch; manual review required.'}
  }
  if($SecureBootStartup -and !$state){$report.Status='NoActiveTransaction';$report.Ready=$snapshot.Complete;return $report}
  if($state -and $state.Phase -eq 'Failed') { Resume-OwnedSecureBootProtection $state $statePath;throw 'Previous Secure Boot transaction failed; review evidence before removing state.json to permit a new attempt.' }
  # Protection cleanup is allowed outside Sunday, but certificate writes/reboots are not.
  if($state -and ((Get-Date).DayOfWeek -ne 'Sunday' -or (Get-Date).Date -ne [datetime]::Parse($state.Started).Date)){
   Resume-OwnedSecureBootProtection $state $statePath
   $state.Phase='Failed';Save-SecureBootState $state $statePath
   throw 'Maintenance transaction expired; owned suspension was resumed. Manual verification required.'
  }
  if(!$snapshot.FirmwareQueryVerified){throw ('Firmware verification is unavailable; no certificate update permitted: '+$snapshot.FirmwareQueryError)}
  if(!$state -and $snapshot.Complete){
   $report.Status='AlreadyUpdated';$report.Ready=(!$encrypted -or [string]$volume.ProtectionStatus -eq 'On')
   if(!$report.Ready){$report.Message='Certificate deployment completed, but BitLocker was already suspended by another process; left unchanged.'}
   return $report
  }
  $policyPath=Join-Path $ScriptsRoot 'Maintenance.Policy.json'
  $policy=(Get-Content -LiteralPath $policyPath -Raw|ConvertFrom-Json).SecureBootMaintenance
  if(!$policy -or !$policy.Enabled){throw 'Secure Boot remediation is disabled or policy is missing.'}
  if($state){$policy=$state.Policy} # Preserve the transaction policy through updater/policy changes.
  if(!(Test-SecureBootPolicyHost $policy 'PilotComputers' $computer)){throw 'Computer is not enrolled in the Secure Boot remediation pilot.'}
  if($dfDetected -and !(Test-SecureBootPolicyHost $policy 'ThawedThroughVerificationComputers' $computer)){throw 'No confirmation that Deep Freeze will remain thawed through verification.'}
  if($encrypted -and $dfDetected){
   if(!$report.DeepFreezeVersion -or [version]$report.DeepFreezeVersion -lt [version]$policy.MinimumDeepFreezeVersion){throw 'Deep Freeze upgrade required before encrypted certificate remediation (minimum 10.20.220.5844).'}
  }
  Assert-SecureBootWindow $policy -NewUpdate:(!$state)
  if(!$snapshot.Enabled){throw 'Secure Boot is disabled; no automatic firmware setting change is permitted.'}
  if($snapshot.ErrorCode -ne 0){throw ('Secure Boot servicing reports error '+$snapshot.ErrorCode+'; investigate before retrying.')}
  if(!$state){
   if($encrypted){
    if([string]$volume.ProtectionStatus -ne 'On'){throw 'BitLocker is already suspended by another process; no ownership assumed.'}
    if(!(Test-SecureBootPolicyHost $policy 'RecoveryKeyVerifiedComputers' $computer)){throw 'Recovery-key availability has not been confirmed for this computer.'}
    if(!@($volume.KeyProtector|Where-Object KeyProtectorType -eq 'RecoveryPassword').Count){throw 'A recovery-password protector is required; no protector will be created automatically.'}
   }
   $task=Get-ScheduledTask -TaskPath '\Microsoft\Windows\PI\' -TaskName 'Secure-Boot-Update' -ErrorAction Stop
   if($task.State -eq 'Disabled'){throw 'Microsoft Secure-Boot-Update task is disabled.'}
   # Do not overwrite actions controlled by domain/Intune policy or unrelated update bits.
   $base=Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot'
   if($base.AvailableUpdatesPolicy){throw 'Secure Boot deployment is managed by Group Policy/Intune; coordinate that policy instead.'}
   if(($snapshot.AvailableUpdates -band (-bnot 0x5944)) -ne 0){throw 'AvailableUpdates contains other servicing flags; manual review required.'}
   $state=[pscustomobject]@{SchemaVersion=1;ComputerName=$computer;UUID=$uuid;VolumeId=$volumeIdentity;Started=(Get-Date).ToString('o');Phase='Updating';OwnsSuspension=$false;ResumeBoot=$null;RebootsRequested=0;LastRequestedBoot=$null;Policy=$policy}
   Save-SecureBootState $state $statePath
   Register-SecureBootResumeTask
   if($encrypted){
    # Persist intent first so failure between suspend and state write can still be cleaned up.
    $state.OwnsSuspension=$true;Save-SecureBootState $state $statePath
    Assert-SecureBootWindow $policy -NewUpdate
    Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount 2 -ErrorAction Stop|Out-Null
    if([string](Get-BitLockerVolume -MountPoint $env:SystemDrive).ProtectionStatus -ne 'Off'){throw 'BitLocker suspension could not be verified.'}
   }
   Assert-SecureBootWindow $policy -NewUpdate
   if($snapshot.Status -ne 'InProgress'){
    New-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot' -Name AvailableUpdates -PropertyType DWord -Value 0x5944 -Force -ErrorAction Stop|Out-Null
   }
  }
  $currentBoot=(Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().Ticks.ToString()
  if($state.LastRequestedBoot -eq $currentBoot){$report.Status='RebootAlreadyRequested';return $report}
  if($state.Phase -eq 'ProtectionResumed'){
   $boot=(Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().Ticks.ToString()
   if(!$snapshot.Complete){throw 'Certificate status is no longer complete after resuming protection.'}
   if($encrypted -and [string]$volume.ProtectionStatus -ne 'On'){throw 'BitLocker protection is not on after resume.'}
   if($state.ResumeBoot -eq $boot){$report.Status=Request-SecureBootReboot $state $policy $statePath;return $report}
   try{$recovery=@(Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-BitLocker-Driver';Id=24658,24705,24652;StartTime=(Get-CimInstance Win32_OperatingSystem).LastBootUpTime} -ErrorAction Stop)}catch{if($_.FullyQualifiedErrorId -match '^NoMatchingEventsFound'){$recovery=@()}else{throw 'Unable to read BitLocker events for the protected verification boot.'}}
   if($recovery.Count){throw 'BitLocker recovery/mismatch events occurred on the verification boot; manual review required.'}
   $report.Status='Updated_ProtectedBootVerified';$report.Ready=$true
   Remove-SecureBootResumeTask
   Move-Item -LiteralPath $statePath -Destination (Join-Path $stateDir 'last-completed.json') -Force
   return $report
  }
  Assert-SecureBootWindow $policy
  Start-ScheduledTask -TaskPath '\Microsoft\Windows\PI\' -TaskName 'Secure-Boot-Update' -ErrorAction Stop
  # Bounded polling; do not reset flags on each pass or claim a task launch is success.
  for($i=0;$i -lt 9;$i++){
   Start-Sleep -Seconds 5
   $snapshot=Get-SecureBootSnapshot;$report.Snapshot=$snapshot
   if($snapshot.Complete -or $snapshot.ErrorCode -ne 0 -or $snapshot.AvailableUpdates -eq 0x4100){break}
  }
  if($snapshot.ErrorCode -ne 0){throw ('Secure Boot update failed with code '+$snapshot.ErrorCode)}
  if($snapshot.Complete){
   Resume-OwnedSecureBootProtection $state $statePath
   # Even unencrypted systems must verify a subsequent boot before closing the transaction.
   if($state.Phase -ne 'ProtectionResumed'){$state.Phase='ProtectionResumed';$state.ResumeBoot=(Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().Ticks.ToString();Save-SecureBootState $state $statePath}
   $report.Status=Request-SecureBootReboot $state $policy $statePath
  }elseif($snapshot.AvailableUpdates -eq 0x4100){
   if($encrypted -and [string](Get-BitLockerVolume -MountPoint $env:SystemDrive).ProtectionStatus -ne 'Off'){throw 'Suspension expired before certificate deployment completed; manual review required.'}
   $report.Status=Request-SecureBootReboot $state $policy $statePath
  }else{$report.Status='UpdateInProgress';$report.Message='No reboot was requested: waiting for documented reboot phase or completion. Next Sunday pass will recheck.'}
  return $report
 }catch{
  $report.Status='BlockedOrFailed';$report.Message=$_.Exception.Message
  if($state -and $statePath){
   try{Resume-OwnedSecureBootProtection $state $statePath;$state.Phase='Failed';Save-SecureBootState $state $statePath}catch{$report.Message+='; PROTECTION CLEANUP FAILED: '+$_.Exception.Message}
  }
  return $report
 }finally{
  if($state){$report.Phase=$state.Phase;$report.RebootsRequested=$state.RebootsRequested}
  $report['EndTime']=(Get-Date).ToUniversalTime().ToString('o')
  try{
   try{$report['TelemetryDelivery']=Write-CloseoutTelemetry -Event $report -Directory $LogDirectory}catch{$report['TelemetryDelivery']='Failed: '+$_.Exception.Message;Write-Warning ('Telemetry write failed: '+$_.Exception.Message)}
   try{$report|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $LogDirectory 'SecureBoot-Maintenance.latest.json') -Encoding UTF8}catch{Write-Warning ('Local report write failed: '+$_.Exception.Message)}
  }finally{if($held){$mutex.ReleaseMutex()};if($mutex){$mutex.Dispose()}}
 }
}
if($SecureBootOnly -or $SecureBootAuditOnly -or $SecureBootStartup){
 $secureBootResult=Invoke-SecureBootMaintenance
 $secureBootResult|ConvertTo-Json -Depth 10|Write-Host
 if($secureBootResult.Ready){exit 0}
 if($secureBootResult.Status -in @('RebootRequested','RebootAlreadyRequested')){exit 3010}
 exit 2
}

try{
 Test-Window
 # Early passes install; the final pass performs an independent scan and closeout.
 if(!$VerifyOnly){
  $enable=Run-Script '01_Enable_Windows_Update_Services.ps1';if($enable -eq 3010){$result.Status='RebootStillRequired';throw 'Startup maintenance requested a reboot; update installation deferred.'};if($enable -ne 0){throw 'Unable to open update services.'}
  $update=Run-Script '06_Weekend_Windows_Updates.ps1';if($update -notin @(0,3010)){throw 'Windows update pass did not complete successfully; review 06 telemetry.'};if($update -eq 3010){$result.RebootRequired=$true;$result.Status='RebootStillRequired';throw 'Script 06 completed its pass and reports a reboot requirement/request. Closeout stops here; no duplicate Script 07 reboot is requested.'}
 }
 Test-Window
 $secureBootResult = Invoke-SecureBootMaintenance
 $result['SecureBoot'] = $secureBootResult
 if ($secureBootResult.Status -in @('RebootRequested','RebootAlreadyRequested')) { $result.Status='RebootStillRequired'; throw 'Secure Boot reboot queued; further update work deferred.' }
 $result.RebootRequired=Servicing-RebootRequired
 if($result.RebootRequired){
  $result.Status='RebootStillRequired';Log 'Windows servicing requires a reboot; requesting Sunday servicing-only reboot.'
  $reboot=Run-Script '07_Force_Reboot_Install_Updates.ps1' @('-WindowsServicingOnly')
  if($reboot -notin @(0,3010)){throw 'Reboot script failed.'}
 }else{
  $session=New-Object -ComObject Microsoft.Update.Session
  $searcher=$session.CreateUpdateSearcher()
  $scan=$searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
  if([int]$scan.ResultCode -ne 2){throw "Final update scan was not successful: ResultCode=$($scan.ResultCode)"}
  $result.RemainingUpdates=$scan.Updates.Count
  Test-Window
  $result.RebootRequired=Servicing-RebootRequired
  if($result.RemainingUpdates -gt 0){$result.Status='UpdatesRemaining'}
  elseif($result.RebootRequired){$result.Status='RebootStillRequired'}
  elseif(@(Get-Process -Name @('TiWorker','TrustedInstaller','MoUsoCoreWorker') -ErrorAction SilentlyContinue).Count -gt 0){$result.Status='ServicingProcessesPresent_Review'}
  elseif(!$VerifyOnly){$result.Status='UpdatePassVerified_NoPendingReboot'}
  else{
   $disable=Run-Script '09_Disable_Windows_Update_Services.ps1'
   if($disable -notin @(0,3010)){throw '09 closeout failed.'}
   $wrong=@(Get-CimInstance Win32_Service -ErrorAction Stop|Where-Object {$_.Name -in @('wuauserv','UsoSvc','BITS') -and ($_.StartMode -ne 'Disabled' -or $_.State -ne 'Stopped')})
   if($wrong.Count){throw 'Update-service closeout verification failed.'}
   Test-Window
   if(Servicing-RebootRequired){$result.Status='RebootStillRequired'}elseif(!$secureBootResult.Ready){throw ('Secure Boot verification incomplete: '+$secureBootResult.Status+'; '+$secureBootResult.Message)}elseif(!(Get-SecureBootSnapshot).Complete){throw 'Secure Boot completion changed during closeout.'}else{$result.Status='ReadyToFreeze';$result.ReadyToFreeze=$true}
  }
 }
 Log "Status=$($result.Status); ReadyToFreeze=$($result.ReadyToFreeze). External freeze schedule is not controlled."
}catch{if($result.Status -ne 'RebootStillRequired'){$result.Status='VerificationFailed'};$result.FailureMessage=$_.Exception.Message;Log $_.Exception.Message}
finally{
 $result['EndTime']=(Get-Date).ToString('o')
 $result|ConvertTo-Json -Depth 6|Set-Content (Join-Path $LogDirectory '15_Sunday_Update_Closeout.latest.json') -Encoding UTF8
 $event=[ordered]@{EventType='maintenance.sunday_update_closeout';Timestamp=(Get-Date).ToUniversalTime().ToString('o');ComputerName=$env:COMPUTERNAME;ScriptName='15_Sunday_Update_Closeout.ps1';Details=$result}
 try{$null=Write-CloseoutTelemetry -Event $event -Directory $LogDirectory}catch{Write-Warning ('Closeout telemetry write failed: '+$_.Exception.Message)}
}
if($result.ReadyToFreeze -or $result.Status -eq 'UpdatePassVerified_NoPendingReboot'){exit 0};if($result.Status -eq 'RebootStillRequired'){exit 3010};exit 2
