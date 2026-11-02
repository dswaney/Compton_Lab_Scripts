#requires -Version 5.1
#requires -RunAsAdministrator
<#
.SYNOPSIS
    Verifies and configures Dell Sunday-midnight BIOS power-on and wired Wake-on-LAN.
.DESCRIPTION
    Version 2.3.0 uses Dell's direct agentless CIM interface. No DellBIOSProvider,
    Visual C++ installer, module replacement or network dependency is required.
    Designed for unattended elevated/SYSTEM Insight execution on each computer.
    Only incorrect settings are written. All eleven settings must pass final read-back.
    Does not prompt, reboot, update firmware, or change BitLocker/Secure Boot settings.
    A BIOS password can be supplied as a SecureString from a trusted local caller.
    No password is embedded, requested interactively, changed, or logged.
    Exit 0: all desired values verified (or non-Dell skipped); 1: error/incomplete;
    2: audit found differences; 20: another instance is active.
    Reports: C:\Logs\<computer>-Dell-BIOS-Power-WOL-<run>.json/.log,
    and C:\Logs\Set-Dell-BIOS-Settings.latest.json.
    AuditOnly makes no BIOS changes but still writes local reports.
.NOTES
    ScriptVersion: 2.3.0
    Source: https://downloads.dell.com/manuals/common/dell-agentless-client-manageability.pdf
#>
[CmdletBinding()]
param(
    [switch]$AuditOnly,
    [Security.SecureString]$BiosAdminPassword,
    [string]$LogRoot='C:\Logs',
    [string]$StagingLogRoot='C:\Temp\Compton-DellBIOS-Logs'
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$script:RunId=[guid]::NewGuid().ToString('N')
$computer=$env:COMPUTERNAME
$RunTimestamp=(Get-Date -Format 'yyyy-MM-dd_HHmmss_fff')+'-'+$script:RunId.Substring(0,8)
$PublishedLogPath=Join-Path $LogRoot ($computer+'-Dell-BIOS-Power-WOL-'+$RunTimestamp+'.log')
$script:DellBiosLogPath=Join-Path $StagingLogRoot ([IO.Path]::GetFileName($PublishedLogPath))
$script:DellBiosLogLines=New-Object 'System.Collections.Generic.List[string]'
$script:DellBiosLogWarningShown=$false
$script:DellBiosNamespace='root\dcim\sysman\biosattributes'
$script:DesiredBiosValues=[ordered]@{
    AutoOnHr='0'; AutoOnMn='0';
    AutoOnMon='Disabled'; AutoOnTue='Disabled'; AutoOnWed='Disabled';
    AutoOnThur='Disabled'; AutoOnFri='Disabled'; AutoOnSat='Disabled';
    AutoOnSun='Enabled'; AutoOn='SelectDays'; WakeOnLan='LanOnly'
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $line = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    $script:DellBiosLogLines.Add($line)
    $bytes=(New-Object System.Text.UTF8Encoding($false)).GetBytes($line+[Environment]::NewLine)
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $stream=$null
        try {
            $stream=[IO.File]::Open($script:DellBiosLogPath,[IO.FileMode]::Append,[IO.FileAccess]::Write,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
            $stream.Write($bytes,0,$bytes.Length)
            $stream.Flush()
            return
        }
        catch {
            if ($attempt -lt 5) { Start-Sleep -Milliseconds (150 * $attempt) }
        }
        finally { if ($null -ne $stream) { $stream.Dispose() } }
    }

    # Keep the complete log in memory, even if every disk write is unavailable.
    # A fresh staged file avoids a persistent exclusive lock on the previous path.
    try {
        $fallback=Join-Path $StagingLogRoot ('Dell-BIOS-recovered-'+[guid]::NewGuid().ToString('N')+'.log')
        [IO.File]::WriteAllLines($fallback,$script:DellBiosLogLines.ToArray(),(New-Object System.Text.UTF8Encoding($false)))
        $script:DellBiosLogPath=$fallback
    }
    catch {
        if (-not $script:DellBiosLogWarningShown) {
            Write-Warning 'Disk logging is unavailable. BIOS processing continues; log messages remain visible in this console.'
            $script:DellBiosLogWarningShown=$true
        }
    }
}

function Publish-CompletedLog {
    [CmdletBinding()]
    param()

    if ($script:DellBiosLogLines.Count -eq 0) { return }
    $completed=$null
    try {
        # Rebuild from memory so failed appends cannot omit earlier messages.
        $completed=Join-Path $StagingLogRoot ('Dell-BIOS-completed-'+[guid]::NewGuid().ToString('N')+'.tmp')
        [IO.File]::WriteAllLines($completed,$script:DellBiosLogLines.ToArray(),(New-Object System.Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $completed -Destination $PublishedLogPath -ErrorAction Stop
        Write-Host ("Published completed Dell BIOS log: {0}" -f $PublishedLogPath)
    }
    catch {
        Write-Warning ("Unable to publish completed Dell BIOS log to {0}: {1}" -f $PublishedLogPath, $_.Exception.Message)
        if ($null -ne $completed -and (Test-Path -LiteralPath $completed -PathType Leaf)) {
            Write-Warning ("Complete log retained at: {0}" -f $completed)
        } else {
            Write-Warning ("Staged log path: {0}. Console output remains available." -f $script:DellBiosLogPath)
        }
    }
}


function Get-DellCimAttributes {
    $map=@{}
    foreach($class in @('EnumerationAttribute','IntegerAttribute')) {
        foreach($item in @(Get-CimInstance -Namespace $script:DellBiosNamespace -ClassName $class -OperationTimeoutSec 20 -ErrorAction Stop)) {
            if($script:DesiredBiosValues.Contains([string]$item.AttributeName)) {
                if($map.ContainsKey([string]$item.AttributeName)){throw ('Duplicate BIOS attribute: '+$item.AttributeName)}
                $map[[string]$item.AttributeName]=$item
            }
        }
    }
    return $map
}
function New-DellBiosPlan {
    param([hashtable]$Attributes)
    $plan=New-Object 'System.Collections.Generic.List[object]'
    foreach($name in $script:DesiredBiosValues.Keys) {
        $desired=[string]$script:DesiredBiosValues[$name]
        $entry=[pscustomobject]@{Name=$name;Desired=$desired;Before=$null;After=$null;Status='NotAttempted';MethodStatus=$null;Error=$null}
        if(!$Attributes.ContainsKey($name)) {$entry.Status='Unsupported';$entry.Error='Attribute not exposed by this BIOS.'}
        else {
            $item=$Attributes[$name];$entry.Before=[string]$item.CurrentValue
            if($entry.Before -ieq $desired){$entry.Status='AlreadyCompliant'}
            elseif([bool]$item.ReadOnly){$entry.Status='ReadOnly';$entry.Error='Incorrect value is read-only.'}
            elseif($name -in @('AutoOnHr','AutoOnMn')) {
                if([long]$item.LowerBound -gt 0 -or [long]$item.UpperBound -lt 0){$entry.Status='Unsupported';$entry.Error='Midnight value is outside supported bounds.'}
                else{$entry.Status='NeedsChange'}
            } elseif($desired -notin @($item.PossibleValue)){$entry.Status='Unsupported';$entry.Error='Desired enumeration value is not supported.'}
            else{$entry.Status='NeedsChange'}
        }
        $plan.Add($entry)
    }
    return $plan.ToArray()
}
function Set-DellCimAttribute {
    param([object]$Interface,[string]$Name,[string]$Value,[byte[]]$PasswordBytes=@())
    if(!$script:DesiredBiosValues.Contains($Name) -or [string]$script:DesiredBiosValues[$Name] -ine $Value){throw 'BIOS write is outside the fixed power-on/WOL allowlist.'}
    $securityType=[uint32]0
    if($PasswordBytes.Length -gt 0){$securityType=[uint32]1}
    $arguments=@{SecType=$securityType;SecHndCount=[uint32]$PasswordBytes.Length;SecHandle=[byte[]]$PasswordBytes;AttributeName=$Name;AttributeValue=$Value}
    $response=Invoke-CimMethod -InputObject $Interface -MethodName SetAttribute -Arguments $arguments -OperationTimeoutSec 20 -ErrorAction Stop
    if($null -eq $response -or $null -eq $response.PSObject.Properties['Status']){throw 'Dell SetAttribute returned no Status field.'}
    if($null -eq $response.Status){throw 'Dell SetAttribute returned a null Status.'}
    return [int]$response.Status
}
function Invoke-DellPowerConfiguration {
    param([switch]$CheckOnly,[Security.SecureString]$Password)
    $result=[ordered]@{ComputerName=$env:COMPUTERNAME;RunId=$script:RunId;ScriptVersion='2.3.0';StartTime=(Get-Date).ToUniversalTime().ToString('o');EndTime=$null;Model=$null;BIOSVersion=$null;Backend='DirectCIM';AuditOnly=[bool]$CheckOnly;Status='Running';Compliant=$false;SettingsChanged=0;FailureMessage=$null;Settings=@()}
    [byte[]]$passwordBytes=@()
    try {
        $system=Get-CimInstance -ClassName Win32_ComputerSystem -OperationTimeoutSec 20 -ErrorAction Stop
        $result.ComputerName=[string]$system.Name;$result.Model=[string]$system.Model
        if([string]$system.Manufacturer -notmatch 'Dell'){$result.Status='SkippedNonDell';return $result}
        $result.BIOSVersion=[string](Get-CimInstance -ClassName Win32_BIOS -OperationTimeoutSec 20 -ErrorAction Stop).SMBIOSBIOSVersion
        $attributes=Get-DellCimAttributes
        $plan=@(New-DellBiosPlan -Attributes $attributes);$result.Settings=$plan
        foreach($entry in $plan){Write-Log -Message ("{0}: before={1}; desired={2}; status={3}" -f $entry.Name,$entry.Before,$entry.Desired,$entry.Status)}
        if(@($plan|Where-Object Status -in @('Unsupported','ReadOnly')).Count){throw 'Preflight failed. One or more required settings cannot be configured; no BIOS writes attempted.'}
        $changes=@($plan|Where-Object Status -eq 'NeedsChange')
        if(!$CheckOnly -and $changes.Count -gt 0){
            $interfaces=@(Get-CimInstance -Namespace $script:DellBiosNamespace -ClassName BIOSAttributeInterface -OperationTimeoutSec 20 -ErrorAction Stop)
            if($interfaces.Count -ne 1){throw 'Expected exactly one Dell BIOSAttributeInterface instance; no changes attempted.'}
            if($null -ne $Password){
                $pointer=[IntPtr]::Zero;$plain=$null
                try{$pointer=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password);$plain=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer);$passwordBytes=[Text.Encoding]::UTF8.GetBytes($plain)}
                finally{$plain=$null;if($pointer -ne [IntPtr]::Zero){[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)}}
            }
            foreach($entry in $changes){
                try {
                    $code=Set-DellCimAttribute -Interface $interfaces[0] -Name $entry.Name -Value $entry.Desired -PasswordBytes $passwordBytes
                    $entry.MethodStatus=$code
                    if($code -ne 0){
                        $meanings=@{1='Failed';2='Invalid parameter';3='Access denied; a BIOS administrator password may be required';4='Not supported';5='Memory error';6='Protocol error'}
                        $detail='Unknown Dell status';if($meanings.ContainsKey($code)){$detail=$meanings[$code]}
                        throw ("Dell method status {0}: {1}." -f $code,$detail)
                    }
                    $fresh=Get-DellCimAttributes
                    if(!$fresh.ContainsKey($entry.Name)){throw 'Attribute disappeared during verification.'}
                    $entry.After=[string]$fresh[$entry.Name].CurrentValue
                    if($entry.After -ine $entry.Desired){throw ('Read-back mismatch: requested '+$entry.Desired+'; actual '+$entry.After)}
                    $entry.Status='AppliedAndVerified';$result.SettingsChanged++
                    Write-Log -Level OK -Message ('Verified '+$entry.Name+'='+$entry.After)
                }catch{
                    $entry.Status='Failed';$entry.Error=$_.Exception.Message
                    Write-Log -Level ERROR -Message ($entry.Name+': '+$entry.Error)
                    # Avoid repeated authentication attempts after a password/access failure.
                    if($entry.MethodStatus -eq 3){break}
                }
            }
        }
        # Independent final snapshot verifies unchanged as well as written attributes.
        $final=Get-DellCimAttributes
        foreach($entry in $plan){
            if($final.ContainsKey($entry.Name)){$entry.After=[string]$final[$entry.Name].CurrentValue}
            if($entry.After -ieq $entry.Desired -and $entry.Status -eq 'AlreadyCompliant'){Write-Log -Level OK -Message ('Already correct: '+$entry.Name+'='+$entry.After)}
        }
        $bad=@($plan|Where-Object {$_.After -ine $_.Desired -or $_.Status -eq 'Failed'})
        $result.Compliant=($bad.Count -eq 0)
        if($result.Compliant){$result.Status='Verified';Write-Log -Level OK -Message 'All BIOS settings verified: Sunday 00:00 only; WakeOnLan=LanOnly.'}
        elseif($CheckOnly){$result.Status='NeedsChanges';Write-Log -Level WARN -Message 'Audit found BIOS settings requiring correction.'}
        else{$result.Status='Incomplete';$result.FailureMessage='One or more BIOS settings failed or remain incorrect.'}
    }catch{$result.Status='Failed';$result.FailureMessage=$_.Exception.Message;Write-Log -Level ERROR -Message $result.FailureMessage}
    finally{if($passwordBytes.Length){[Array]::Clear($passwordBytes,0,$passwordBytes.Length)};$result.EndTime=(Get-Date).ToUniversalTime().ToString('o')}
    return $result
}

$mutex=$null;$held=$false;$exitCode=1
try {
    foreach($directory in @($LogRoot,$StagingLogRoot)){if(!(Test-Path -LiteralPath $directory -PathType Container)){New-Item -Path $directory -ItemType Directory -Force|Out-Null}}
    Write-Log -Message ('Starting Dell BIOS power configuration. ScriptVersion=2.3.0; Backend=DirectCIM; ScriptPath='+$PSCommandPath)
    if(![Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem){throw 'Use native 64-bit Windows PowerShell for this script.'}
    $mutex=New-Object Threading.Mutex($false,'Global\Compton_DellBiosPowerConfiguration')
    try{$held=$mutex.WaitOne(0)}catch [Threading.AbandonedMutexException]{$held=$true}
    if(!$held){$exitCode=20;throw 'Another Dell BIOS power configuration process is running.'}
    $result=Invoke-DellPowerConfiguration -CheckOnly:$AuditOnly -Password $BiosAdminPassword
    $json=$result|ConvertTo-Json -Depth 8
    foreach($destination in @([IO.Path]::ChangeExtension($PublishedLogPath,'.json'),(Join-Path $LogRoot 'Set-Dell-BIOS-Settings.latest.json'))){
        try {
            $temp=Join-Path $StagingLogRoot ([guid]::NewGuid().ToString('N')+'.json.tmp')
            [IO.File]::WriteAllText($temp,$json,(New-Object Text.UTF8Encoding($false)))
            Move-Item -LiteralPath $temp -Destination $destination -Force -ErrorAction Stop
        }catch{Write-Log -Level WARN -Message ('Could not save report '+$destination+': '+$_.Exception.Message)}
    }
    $result.Settings|Format-Table Name,Before,Desired,After,Status -AutoSize|Out-Host
    Write-Host ('Computer='+$result.ComputerName+'; Status='+$result.Status+'; Compliant='+$result.Compliant+'; SettingsChanged='+$result.SettingsChanged)
    if($result.Compliant -or $result.Status -eq 'SkippedNonDell'){$exitCode=0}
    elseif($result.Status -eq 'NeedsChanges'){$exitCode=2}
}catch{Write-Host ('Dell BIOS configuration failed: '+$_.Exception.Message) -ForegroundColor Red}
finally{
    try{Publish-CompletedLog}catch{Write-Warning ('Log publication failed: '+$_.Exception.Message)}
    if($held){$mutex.ReleaseMutex()};if($null -ne $mutex){$mutex.Dispose()}
}
exit $exitCode
