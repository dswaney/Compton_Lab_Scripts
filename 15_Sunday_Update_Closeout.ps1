#requires -Version 5.1
#requires -RunAsAdministrator
# ScriptVersion: 1.0.0
# LastUpdated: 2026-10-13
[CmdletBinding()]
param([string]$ScriptsRoot='C:\Scripts',[switch]$VerifyOnly,[string]$LogDirectory='C:\Logs')
$ErrorActionPreference='Stop'
New-Item -Path $LogDirectory -ItemType Directory -Force|Out-Null
$start=Get-Date;$result=[ordered]@{ComputerName=$env:COMPUTERNAME;ScriptVersion='1.0.0';StartTime=$start.ToString('o');Status='Running';ReadyToFreeze=$false;RemainingUpdates=$null;RebootRequired=$null;FailureMessage=$null;FreezeScheduleControlled=$false}
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
try{
 Test-Window
 # Early passes install; the final pass performs an independent scan and closeout.
 if(!$VerifyOnly){
  $enable=Run-Script '01_Enable_Windows_Update_Services.ps1';if($enable -notin @(0,3010)){throw 'Unable to open update services.'}
  $update=Run-Script '06_Weekend_Windows_Updates.ps1';if($update -notin @(0,3010)){throw 'Windows update pass did not complete successfully; review 06 telemetry.'}
 }
 Test-Window
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
  elseif(@(Get-Process -Name TiWorker,TrustedInstaller,MoUsoCoreWorker -ErrorAction SilentlyContinue).Count -gt 0){$result.Status='ServicingProcessesPresent_Review'}
  elseif(!$VerifyOnly){$result.Status='UpdatePassVerified_NoPendingReboot'}
  else{
   $disable=Run-Script '09_Disable_Windows_Update_Services.ps1'
   if($disable -notin @(0,3010)){throw '09 closeout failed.'}
   $wrong=@(Get-CimInstance Win32_Service -ErrorAction Stop|Where-Object {$_.Name -in @('wuauserv','UsoSvc','BITS') -and ($_.StartMode -ne 'Disabled' -or $_.State -ne 'Stopped')})
   if($wrong.Count){throw 'Update-service closeout verification failed.'}
   Test-Window
   if(Servicing-RebootRequired){$result.Status='RebootStillRequired'}else{$result.Status='ReadyToFreeze';$result.ReadyToFreeze=$true}
  }
 }
 Log "Status=$($result.Status); ReadyToFreeze=$($result.ReadyToFreeze). External freeze schedule is not controlled."
}catch{$result.Status='VerificationFailed';$result.FailureMessage=$_.Exception.Message;Log $_.Exception.Message}
finally{
 $result['EndTime']=(Get-Date).ToString('o')
 $result|ConvertTo-Json -Depth 6|Set-Content (Join-Path $LogDirectory '15_Sunday_Update_Closeout.latest.json') -Encoding UTF8
 $event=[ordered]@{EventType='maintenance.sunday_update_closeout';Timestamp=(Get-Date).ToUniversalTime().ToString('o');ComputerName=$env:COMPUTERNAME;ScriptName='15_Sunday_Update_Closeout.ps1';Details=$result}
 $event|ConvertTo-Json -Depth 8 -Compress|Add-Content (Join-Path $LogDirectory 'Maintenance-Telemetry.ndjson') -Encoding UTF8
}
if($result.ReadyToFreeze -or $result.Status -eq 'UpdatePassVerified_NoPendingReboot'){exit 0};if($result.Status -eq 'RebootStillRequired'){exit 3010};exit 2
