#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
Collects local evidence for Deep Freeze upgrade and console communication failures.
.DESCRIPTION
Read-only diagnostics except creating the output folder and ZIP. Does not thaw,
freeze, restart services, enable logging, install software, or reboot. Run in
64-bit Windows PowerShell 5.1 as administrator or SYSTEM. Copy output off a frozen
volume BEFORE reboot. Vendor logs and Windows events can contain internal details.
Console-side history must be exported separately. No Deep Freeze password needed.
.PARAMETER ConsoleServer
Optional actual console/communication server DNS name or IP (no URL).
.PARAMETER ConsolePort
Optional configured TCP port. No default is assumed. TCP success does not verify
UDP discovery, Deep Freeze protocol, or console authentication.
.PARAMETER AdditionalLogPath
Optional additional vendor log directory or individual log file.
.EXAMPLE
.\Collect-DeepFreezeDiagnostics.ps1
.EXAMPLE
.\Collect-DeepFreezeDiagnostics.ps1 -ConsoleServer 'actual-server' -ConsolePort 7725
#>
[CmdletBinding()]
param(
    [ValidateRange(1,90)][int]$Days = 14,
    [string]$OutputRoot = 'C:\Temp\DeepFreeze-Diagnostics',
    [string]$ConsoleServer,
    [ValidateRange(1,65535)][int]$ConsolePort,
    [string[]]$AdditionalLogPath = @(),
    [ValidateRange(1,1024)][int]$MaxLogMB = 100,
    [ValidateRange(10,4096)][int]$MaxTotalLogMB = 500,
    [string]$ObservedState = 'Not recorded; record Frozen/Thawed from console or workstation UI',
    [string]$AttemptTime = 'Not supplied'
)
$ErrorActionPreference = 'Stop'
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    throw 'Run in 64-bit PowerShell (SYSTEM agents can use C:\Windows\Sysnative\WindowsPowerShell\v1.0\powershell.exe).'
}
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$bundle = Join-Path $OutputRoot ("{0}-DeepFreeze-{1}-{2}" -f $env:COMPUTERNAME,$stamp,([guid]::NewGuid().ToString('N').Substring(0,6)))
New-Item -ItemType Directory -Path $bundle -Force | Out-Null
$issues = New-Object 'System.Collections.Generic.List[string]'
function Save-Check {
    param([string]$Name,[scriptblock]$Action)
    try { & $Action | Out-String -Width 300 | Out-File (Join-Path $bundle ($Name+'.txt')) -Encoding utf8 }
    catch { $issues.Add("${Name}: $($_.Exception.Message)") }
}
function Native-Check {
    param([string]$Name,[string]$Exe,[string[]]$Arguments)
    Save-Check $Name {
        & $Exe @Arguments 2>&1 | Out-String -Width 300
        $rc = $LASTEXITCODE
        "ExitCode: $rc"
        if ($rc -ne 0) { $issues.Add("$Name returned exit code $rc") }
    }
}
$start = (Get-Date).AddDays(-$Days)
Save-Check 'Overview' {
    [pscustomobject]@{
        Computer=$env:COMPUTERNAME; CollectionLocal=(Get-Date).ToString('o'); CollectionUTC=[datetime]::UtcNow.ToString('o')
        TimeZone=(Get-TimeZone).Id; EventWindowStart=$start.ToString('o'); AttemptTime=$AttemptTime
        ObservedDeepFreezeState=$ObservedState; ReportedOldBuild='5788'; IntendedBuild='5844'
        ConsoleServer=$ConsoleServer; ConfiguredTCPPort=$ConsolePort; Identity=[Security.Principal.WindowsIdentity]::GetCurrent().Name
        PowerShell=$PSVersionTable.PSVersion.ToString(); Output=$bundle
    } | Format-List
    'Build labels above are supplied incident context; see BinaryVersions.csv for measured versions.'
}
Save-Check 'OS' { Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version,BuildNumber,LastBootUpTime,LocalDateTime,FreePhysicalMemory | Format-List }
Save-Check 'Hardware' { Get-CimInstance Win32_ComputerSystem | Select-Object Manufacturer,Model,Domain,PartOfDomain | Format-List; Get-CimInstance Win32_BIOS | Select-Object SMBIOSBIOSVersion,ReleaseDate | Format-List; Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | Select-Object DeviceID,Size,FreeSpace | Format-Table }
$vendorPattern = 'Faronics|Deep\s*Freeze|DFServ|DFFrz|DFDisk|DFNet|DFMon|DFEngine'
$services = @(); $drivers = @(); $processes = @()
Save-Check 'Services' {
    $script:services = @(Get-CimInstance Win32_Service | Where-Object { ($_.Name+' '+$_.DisplayName+' '+$_.PathName) -match $vendorPattern })
    $services | Select-Object Name,DisplayName,State,StartMode,ProcessId,PathName,ExitCode,ServiceSpecificExitCode | Format-List
}
Save-Check 'Drivers' {
    $script:drivers = @(Get-CimInstance Win32_SystemDriver | Where-Object { ($_.Name+' '+$_.DisplayName+' '+$_.PathName) -match $vendorPattern })
    $drivers | Select-Object Name,DisplayName,State,StartMode,PathName,ExitCode | Format-List
}
Save-Check 'Processes' {
    $script:processes = @(Get-Process | Where-Object { $_.ProcessName -match $vendorPattern -or $_.Path -match 'Faronics|Deep Freeze' })
    $processes | Select-Object Id,ProcessName,Path,StartTime,CPU,WorkingSet64 | Format-Table -AutoSize
}
Save-Check 'InstalledProducts' {
    # Registry only: deliberately avoids Win32_Product and its MSI repair side effects.
    Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'Deep Freeze|Faronics' } |
        Select-Object DisplayName,DisplayVersion,Publisher,InstallDate,InstallLocation | Format-List
}
Save-Check 'PendingRestart' {
    [pscustomobject]@{
        CBS=Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        WindowsUpdate=Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
        PendingFileRenameOperations=[bool](Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue)
    } | Format-List
    'Restart flags alone do not establish the cause of a failed upgrade.'
}
Native-Check 'IPConfig' 'ipconfig.exe' @('/all')
Native-Check 'Routes' 'route.exe' @('print')
Native-Check 'Netstat' 'netstat.exe' @('-ano')
Native-Check 'TimeStatus' 'w32tm.exe' @('/query','/status')
Save-Check 'NetworkProfiles' { Get-NetConnectionProfile | Format-List; Get-NetAdapter | Select-Object Name,InterfaceDescription,Status,LinkSpeed | Format-Table }
Save-Check 'FirewallProfiles' { Get-NetFirewallProfile | Select-Object Name,Enabled,DefaultInboundAction,DefaultOutboundAction,LogFileName,LogBlocked,LogAllowed | Format-List }
Save-Check 'VendorFirewallRules' {
    $appFilters = @(Get-NetFirewallApplicationFilter | Where-Object Program -match 'Faronics|Deep Freeze|DFServ')
    $rules = @(Get-NetFirewallRule | Where-Object DisplayName -match 'Faronics|Deep Freeze')
    foreach ($filter in $appFilters) { $rules += @($filter | Get-NetFirewallRule) }
    foreach ($rule in ($rules | Sort-Object Name -Unique)) {
        $rule | Select-Object Name,DisplayName,Enabled,Direction,Action,Profile | Format-List
        $rule | Get-NetFirewallPortFilter | Format-List Protocol,LocalPort,RemotePort
        $rule | Get-NetFirewallAddressFilter | Format-List LocalAddress,RemoteAddress
        $rule | Get-NetFirewallApplicationFilter | Format-List Program
    }
}
Save-Check 'VendorSockets' {
    $ids = @($services | Select-Object -ExpandProperty ProcessId) + @($processes | Select-Object -ExpandProperty Id)
    $ids = @($ids | Where-Object { $_ -gt 0 } | Sort-Object -Unique)
    Get-NetTCPConnection | Where-Object { $_.OwningProcess -in $ids } | Format-Table -AutoSize
    Get-NetUDPEndpoint | Where-Object { $_.OwningProcess -in $ids } | Format-Table -AutoSize
}
if ($ConsoleServer) {
    Save-Check 'ConsoleDNS' { Resolve-DnsName -Name $ConsoleServer | Format-Table -AutoSize }
    if ($ConsolePort) {
        Save-Check 'ConsoleTCPTest' {
            $client = New-Object System.Net.Sockets.TcpClient
            try {
                $task = $client.ConnectAsync($ConsoleServer,$ConsolePort)
                if (-not $task.Wait(5000)) { throw 'TCP connection timed out after 5 seconds.' }
                if (-not $client.Connected) { throw 'TCP connection not established.' }
                "TCP connection succeeded to ${ConsoleServer}:$ConsolePort. Does not validate UDP or the Deep Freeze protocol."
            } finally { $client.Dispose() }
        }
    }
}
# Preserve native Windows events for support; exports are limited to the selected window.
$eventDir = Join-Path $bundle 'Events'
New-Item -ItemType Directory -Path $eventDir | Out-Null
$channels = @('System','Application','Setup','Microsoft-Windows-CodeIntegrity/Operational','Microsoft-Windows-Windows Defender/Operational')
try { $channels += @(Get-WinEvent -ListLog * -ErrorAction SilentlyContinue | Where-Object LogName -match 'Faronics|Deep.?Freeze' | Select-Object -ExpandProperty LogName) } catch { $issues.Add("Event channel discovery: $($_.Exception.Message)") }
$milliseconds = [long]$Days * 86400000
foreach ($channel in ($channels | Sort-Object -Unique)) {
    $safeName = $channel -replace '[^a-zA-Z0-9._-]','_'
    try {
        $meta = Get-WinEvent -ListLog $channel
        if (-not $meta.IsEnabled) { $issues.Add("Event channel disabled (left unchanged): $channel"); continue }
        $dest = Join-Path $eventDir ($safeName+'.evtx')
        $result = & wevtutil.exe epl $channel $dest "/q:*[System[TimeCreated[timediff(@SystemTime) <= $milliseconds]]]" /ow:true 2>&1
        $rc = $LASTEXITCODE
        if ($rc -ne 0) { throw "wevtutil exit $rc : $result" }
    } catch { $issues.Add("Event export ${channel}: $($_.Exception.Message)") }
}
Save-Check 'RelevantEventTimeline' {
    foreach ($channel in @('Application','System')) {
        try {
            Get-WinEvent -FilterHashtable @{LogName=$channel;StartTime=$start} -MaxEvents 30000 |
                Where-Object { $_.ProviderName -match 'Faronics|Deep.?Freeze|MsiInstaller|Service Control Manager|Application Error|Windows Error Reporting|Kernel-Power|EventLog|USER32|Tcpip|DNS|Dhcp|Netlogon' -or $_.Message -match 'Faronics|Deep Freeze|DFServ|DFInstall' } |
                Select-Object TimeCreated,LogName,ProviderName,Id,LevelDisplayName,Message | Format-List
        } catch { $issues.Add("Timeline ${channel}: $($_.Exception.Message)") }
    }
    'Timeline is capped at 30,000 newest events per channel; EVTX exports contain the full requested time window.'
}
# Search vendor trees, plus matching top-level installer logs in Windows/Temp.
# Never collect raw registry configuration, RDX files, passwords, or memory dumps.
$roots = @()
foreach ($base in @($env:ProgramFiles,${env:ProgramFiles(x86)},$env:ProgramData)) {
    if ($base) { $roots += Join-Path $base 'Faronics' }
}
$roots += $AdditionalLogPath
$files = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
$scanErrors = @()
foreach ($root in ($roots | Sort-Object -Unique)) {
    if (Test-Path -LiteralPath $root) {
        Get-ChildItem -LiteralPath $root -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +scanErrors | ForEach-Object { $files.Add($_) }
    }
}
foreach ($root in @($env:windir,(Join-Path $env:windir 'Temp'),$env:TEMP)) {
    if (Test-Path -LiteralPath $root) {
        Get-ChildItem -LiteralPath $root -File -ErrorAction SilentlyContinue -ErrorVariable +scanErrors |
            Where-Object Name -match '^(DF|Deep.?Freeze|Faronics).*(\.log|\.txt|\.etl|\.bak|\.old|\.zip)(\.\d+)?$' | ForEach-Object { $files.Add($_) }
    }
}
foreach ($err in $scanErrors) { $issues.Add("File discovery: $err") }
$uniqueFiles = @($files | Sort-Object FullName -Unique)
Save-Check 'BinaryVersionCollection' {
    $binaries = @($uniqueFiles | Where-Object Extension -in '.exe','.dll','.sys' | Select-Object -ExpandProperty FullName)
    foreach ($item in @($services)+@($drivers)) {
        $p = [string]$item.PathName
        if ($p -match '^"([^"]+)"') { $p = $matches[1] }
        elseif ($p -match '^(.+?\.(exe|sys))(?=\s|$)') { $p = $matches[1] }
        $p = $p -replace '^\\\?\?\\',''
        $p = $p -replace '^\\SystemRoot', $env:windir
        $p = [Environment]::ExpandEnvironmentVariables($p)
        if (Test-Path -LiteralPath $p -PathType Leaf) { $binaries += $p }
    }
    $rows = foreach ($p in ($binaries | Sort-Object -Unique)) {
        try {
            $f = Get-Item -LiteralPath $p
            [pscustomobject]@{Path=$p;FileVersion=$f.VersionInfo.FileVersion;ProductVersion=$f.VersionInfo.ProductVersion;Length=$f.Length;LastWriteTimeUTC=$f.LastWriteTimeUtc;SHA256=(Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash}
        } catch { $issues.Add("Binary ${p}: $($_.Exception.Message)") }
    }
    $rows | Export-Csv (Join-Path $bundle 'BinaryVersions.csv') -NoTypeInformation -Encoding utf8
    "Measured $(@($rows).Count) binary versions. Binaries themselves are not copied."
}
$logDir = Join-Path $bundle 'VendorLogs'
New-Item -ItemType Directory -Path $logDir | Out-Null
$totalBytes = 0L
$index = 0
$manifest = foreach ($file in ($uniqueFiles | Where-Object { $_.Name -match '\.(log|txt|etl|evtx|bak|old)(\.\d+)?$' -or $_.Name -match '^(DFInstall|dfupdate|DFServ).*\.zip$' } | Sort-Object LastWriteTimeUtc -Descending)) {
    $index++
    $status = 'Copied'; $relative = ''; $hash = ''
    try {
        if ($file.Length -gt ($MaxLogMB * 1MB)) { $status = 'Skipped: per-file size limit' }
        elseif (($totalBytes + $file.Length) -gt ($MaxTotalLogMB * 1MB)) { $status = 'Skipped: total log size limit' }
        else {
            $relative = 'VendorLogs\{0:D4}-{1}' -f $index,$file.Name
            $target = Join-Path $bundle $relative
            Copy-Item -LiteralPath $file.FullName -Destination $target
            $totalBytes += (Get-Item -LiteralPath $target).Length
            $hash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
        }
    } catch { $status = 'Failed: '+$_.Exception.Message; $issues.Add("Copy $($file.FullName): $status") }
    [pscustomobject]@{Source=$file.FullName;Length=$file.Length;LastWriteTimeUTC=$file.LastWriteTimeUtc;Status=$status;BundlePath=$relative;CopiedSHA256=$hash}
}
$manifest | Export-Csv (Join-Path $bundle 'LogManifest.csv') -NoTypeInformation -Encoding utf8
@"
Deep Freeze upgrade/connection evidence - collector v1.0
Computer: $env:COMPUTERNAME
Attempt time supplied: $AttemptTime
State supplied (not independently detected): $ObservedState

Start with Overview.txt, BinaryVersions.csv, RelevantEventTimeline.txt,
VendorLogs, LogManifest.csv, Services.txt, Drivers.txt, VendorSockets.txt,
and CollectionIssues.txt. EVTX files open in Windows Event Viewer.
Vendor logs are copied regardless of age, subject to size limits; event window is $Days days.
Missing/locked logs and disabled channels are recorded. No debug logging is enabled.
Existing evidence may not identify a failure if the vendor did not log it.
This collector does not determine frozen state or diagnose a root cause automatically.

Also attach the Enterprise Console's Show Log export for this hostname, exact
console version, full client version before/after, update attempt time/timezone,
and when the endpoint disappeared/reappeared. Export history before it ages out.
Run on one failed endpoint and one successful endpoint for comparison.
If possible capture while disconnected, then again after reconnecting.
Copy ZIP OFF any frozen volume before reboot; otherwise collection may be lost.

Privacy: contains hostnames, addresses, domain/user references, local paths,
Windows events and raw vendor logs. Review before external sharing. Logs are
not automatically redacted. No raw registry exports, password/configuration files,
BitLocker recovery keys, packet captures or memory dumps are requested.
No files are sent anywhere automatically. No services/settings are changed.
TCP test (if requested) cannot prove UDP or Deep Freeze protocol connectivity.
"@ | Set-Content (Join-Path $bundle 'READ-ME-FIRST.txt') -Encoding utf8
if ($issues.Count -eq 0) { 'No collection errors recorded.' | Set-Content (Join-Path $bundle 'CollectionIssues.txt') -Encoding utf8 }
else { $issues | Set-Content (Join-Path $bundle 'CollectionIssues.txt') -Encoding utf8 }
$zipPath = $bundle+'.zip'
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::CreateFromDirectory($bundle,$zipPath,[IO.Compression.CompressionLevel]::Optimal,$false)
    Write-Output "ZIP: $zipPath"
    Write-Output "Review collection issues: $(Join-Path $bundle 'CollectionIssues.txt')"
    Write-Output 'Copy the ZIP off this computer before reboot if the output volume is frozen.'
} catch {
    Write-Warning "ZIP creation failed: $($_.Exception.Message). Uncompressed evidence remains at $bundle"
    throw
}
