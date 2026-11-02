#Requires -RunAsAdministrator
<#
.SYNOPSIS
Injects a Dell RAID-to-AHCI WinPE preflight into Windows Setup boot.wim index 2.
Services install.wim using explicitly ordered offline update packages.
.DESCRIPTION
Dell machines in RAID mode request AHCI, tell the technician to boot USB again,
and wait for a keypress before reboot. HP and already-AHCI machines continue.
Validates that disk 0 is the only internal drive before launching the existing
unattended installer. Removes the Compton domain-join payload added by the
previous builder, if present; other answer-file settings are preserved.
Sets an automatic computer name and removes unattend domain-join components.
Use current compatible ADK DISM and matching architecture/version update packages.
Allow space for two extra install.wim copies, boot.wim backup, and mounted images.
This prepares destructive clean-install media, not in-place repair media.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateScript({Test-Path $_ -PathType Leaf})][string]$BootWim,
    [Parameter(Mandatory)][ValidateScript({Test-Path $_ -PathType Leaf})][string]$AnswerFile,
    [Parameter(Mandatory)][ValidateScript({Test-Path (Join-Path $_ 'cctk.exe') -PathType Leaf})][string]$DellConfigureFolder,
    [Parameter(Mandatory)][ValidateScript({Test-Path $_ -PathType Container})][string]$WinPEOCFolder,
    [Parameter(Mandatory)][ValidateScript({Test-Path $_ -PathType Leaf})][string]$InstallWim,
    # Explicit prerequisite order: checkpoint/SSU first, target cumulative update last.
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$UpdatePackages,
    # Omit to service every edition/index in install.wim.
    [int[]]$InstallIndexes,
    [string]$WorkDir = (Join-Path $env:TEMP 'StoragePreflightBuild')
)
$ErrorActionPreference = 'Stop'
$BootWim = (Resolve-Path -LiteralPath $BootWim).Path
$AnswerFile = (Resolve-Path -LiteralPath $AnswerFile).Path
$DellConfigureFolder = (Resolve-Path -LiteralPath $DellConfigureFolder).Path
$WinPEOCFolder = (Resolve-Path -LiteralPath $WinPEOCFolder).Path
$InstallWim = (Resolve-Path -LiteralPath $InstallWim).Path
if ([IO.Path]::GetExtension($InstallWim) -ine '.wim') { throw 'InstallWim must be an unsplit install.wim, not ESD or SWM.' }
$UpdatePackages = @($UpdatePackages | ForEach-Object {
    $package = (Resolve-Path -LiteralPath $_ -ErrorAction Stop).Path
    if ([IO.Path]::GetExtension($package) -notin @('.msu','.cab')) { throw "Expected MSU or CAB: $package" }
    $package
})
Import-Module Dism -ErrorAction Stop
$availableIndexes = @(Get-WindowsImage -ImagePath $InstallWim -ErrorAction Stop | Select-Object -ExpandProperty ImageIndex)
if (-not $InstallIndexes) { $InstallIndexes = $availableIndexes }
$InstallIndexes = @($InstallIndexes | Select-Object -Unique)
foreach ($index in $InstallIndexes) {
    if ($index -notin $availableIndexes) { throw "Install index $index does not exist." }
}
$usbRoot = Split-Path (Split-Path $BootWim -Parent) -Parent
if ((Split-Path $BootWim -Parent | Split-Path -Leaf) -ine 'sources' -or
    (Split-Path $AnswerFile -Parent) -ine $usbRoot) {
    throw 'BootWim must be in sources directly under the AnswerFile folder.'
}
$xml = [xml](Get-Content -LiteralPath $AnswerFile -Raw)
# Use Windows-generated names; do not request a technician-supplied name.
$shell = $xml.SelectSingleNode('//*[local-name()="settings" and @pass="specialize"]/*[local-name()="component" and @name="Microsoft-Windows-Shell-Setup"]')
if (-not $shell) { throw 'Missing specialize Shell-Setup component.' }
$computerName = $shell.SelectSingleNode('*[local-name()="ComputerName"]')
if (-not $computerName) {
    $computerName = $xml.CreateElement('ComputerName', $xml.DocumentElement.NamespaceURI)
    [void]$shell.AppendChild($computerName)
}
$computerName.InnerText = '*'
foreach ($join in @($xml.SelectNodes('//*[local-name()="component" and @name="Microsoft-Windows-UnattendedJoin"]'))) {
    [void]$join.ParentNode.RemoveChild($join)
}
$setup = @($xml.unattend.settings | Where-Object pass -eq 'windowsPE').component |
    Where-Object name -eq 'Microsoft-Windows-Setup'
$paths = @($setup.RunSynchronous.RunSynchronousCommand | ForEach-Object Path) -join ' '
if ($setup.ImageInstall.OSImage.InstallTo.DiskID -ne '0' -or
    $setup.ImageInstall.OSImage.InstallTo.PartitionID -ne '3' -or
    $paths -notmatch '(?i)SELECT DISK=0' -or $paths -notmatch '(?i)\bCLEAN\b') {
    throw 'Answer file must use the existing disk 0, partition 3 layout and diskpart CLEAN. Review it before building.'
}
$wimInfo = & dism.exe /English '/Get-WimInfo' "/WimFile:$BootWim" '/Index:2' 2>&1
if ($LASTEXITCODE -ne 0 -or ($wimInfo -join ' ') -notmatch '(?i)setup') {
    throw 'Index 2 must be the Windows Setup image.'
}
# These are the exact entries introduced by the older domain/inventory build.
# The original FirstLogon.ps1 and all other unattend content are left intact.
foreach ($command in @($xml.SelectNodes('//*[local-name()="settings" and @pass="oobeSystem"]/*[local-name()="component" and @name="Microsoft-Windows-Shell-Setup"]/*[local-name()="FirstLogonCommands"]/*[local-name()="SynchronousCommand"]'))) {
    if ($command.InnerText -match '(?i)Initialize-ComptonDeployment\.ps1') {
        [void]$command.ParentNode.RemoveChild($command)
    }
}
$extensions = $xml.DocumentElement.SelectSingleNode('*[local-name()="Extensions"]')
if ($extensions) {
    foreach ($file in @($extensions.ChildNodes | Where-Object {
        $_.LocalName -eq 'File' -and $_.GetAttribute('path') -in @(
            'C:\Windows\Setup\Scripts\Initialize-ComptonDeployment.ps1',
            'C:\ProgramData\ComptonDeployment\Post-Deployment.ps1',
            'C:\ProgramData\ComptonDeployment\JoinPlan.json')
    })) { [void]$extensions.RemoveChild($file) }
}
$runDir = Join-Path $WorkDir ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $runDir -Force | Out-Null
$mount = Join-Path $runDir 'Mount'
New-Item -ItemType Directory -Path $mount | Out-Null
$backup = Join-Path $runDir 'boot.wim.original'
$answerBackup = Join-Path $runDir 'autounattend.xml.original'
$mounted = $false
$installMounted = $false
$installBackupReady = $false
$recoveryComplete = $true
$installMount = Join-Path $runDir 'InstallMount'
$stagedInstall = Join-Path $runDir 'install.serviced.wim'
$installBackup = Join-Path $runDir 'install.wim.original'
New-Item -ItemType Directory -Path $installMount | Out-Null
function Invoke-Dism([string[]]$Arguments) {
    & dism.exe @Arguments | Out-Host
    if ($LASTEXITCODE -notin @(0,3010)) { throw "DISM failed (exit $LASTEXITCODE): $($Arguments -join ' ')" }
}
try {
    Copy-Item -LiteralPath $BootWim -Destination $backup -Force
    Copy-Item -LiteralPath $AnswerFile -Destination $answerBackup -Force
    # Service a staging copy so a failed package cannot leave original media half updated.
    Copy-Item -LiteralPath $InstallWim -Destination $stagedInstall -Force
    foreach ($index in $InstallIndexes) {
        Write-Host "Servicing install.wim index $index"
        Invoke-Dism -Arguments @('/Mount-Image', "/ImageFile:$stagedInstall", "/Index:$index", "/MountDir:$installMount")
        $installMounted = $true
        foreach ($package in $UpdatePackages) {
            Write-Host "Applying $package"
            Invoke-Dism -Arguments @("/Image:$installMount", '/Add-Package', "/PackagePath:$package", '/NoRestart')
        }
        Invoke-Dism -Arguments @("/Image:$installMount", '/Get-Packages')
        Invoke-Dism -Arguments @('/Unmount-Image', "/MountDir:$installMount", '/Commit')
        $installMounted = $false
    }
    Invoke-Dism -Arguments @('/Mount-Image', "/ImageFile:$BootWim", '/Index:2', "/MountDir:$mount")
    $mounted = $true
    $sys = Join-Path $mount 'Windows\System32'
    $launcher = Join-Path $sys 'winpeshl.ini'
    if (Test-Path -LiteralPath $launcher) {
        $prior = Get-Content -LiteralPath $launcher -Raw
        if ($prior -notmatch '(?i)StoragePreflight|setup\.exe') {
            throw 'An unrecognized WinPE launcher is already present. Review winpeshl.ini manually.'
        }
    }
    $needsComponents = -not ((Test-Path -LiteralPath $launcher) -and
        ((Get-Content -LiteralPath $launcher -Raw) -match '(?i)StoragePreflight'))
    if ($needsComponents) {
        foreach ($name in 'WinPE-WMI','WinPE-NetFX','WinPE-Scripting','WinPE-PowerShell','WinPE-StorageWMI') {
            $cab = Join-Path $WinPEOCFolder "$name.cab"
            if (-not (Test-Path -LiteralPath $cab)) { throw "Missing ADK component: $cab" }
            Invoke-Dism -Arguments @("/Image:$mount", '/Add-Package', "/PackagePath:$cab")
            $lang = Join-Path $WinPEOCFolder "en-us\${name}_en-us.cab"
            if (Test-Path -LiteralPath $lang) { Invoke-Dism -Arguments @("/Image:$mount", '/Add-Package', "/PackagePath:$lang") }
        }
    }
    $dellDest = Join-Path $sys 'DellConfigure'
    if (Test-Path -LiteralPath $dellDest) { Remove-Item -LiteralPath $dellDest -Recurse -Force }
    Copy-Item -LiteralPath $DellConfigureFolder -Destination $dellDest -Recurse -Force
    $runtime = @'
$ErrorActionPreference = 'Stop'
$log = 'X:\Windows\Temp\StoragePreflight.log'
function Log([string]$message) {
    $line = '{0} {1}' -f (Get-Date -Format s),$message
    try { $line | Out-File -LiteralPath $log -Append -Encoding UTF8 -ErrorAction Stop }
    catch { $line | Out-File -LiteralPath 'X:\Windows\Temp\StoragePreflight.log' -Append -Encoding UTF8 }
    Write-Host $message
}
function Block([string]$message) {
    Log "STOPPED: $message"
    Write-Host "Installation stopped. See $log" -ForegroundColor Red
    while ($true) { Start-Sleep -Seconds 3600 }
}
try {
    $media = @()
    foreach ($vol in @(Get-Volume | Where-Object DriveLetter)) {
        $root = '{0}:\' -f $vol.DriveLetter
        if ((Test-Path ($root + 'autounattend.xml')) -and (Test-Path ($root + 'sources\boot.wim'))) {
            $partition = Get-Partition -DriveLetter $vol.DriveLetter
            $disk = Get-Disk -Number $partition.DiskNumber
            $media += [pscustomobject]@{ Root=$root; Disk=$disk }
        }
    }
    if ($media.Count -ne 1 -or $media[0].Disk.BusType -ne 'USB') {
        Block 'Exactly one USB installer with boot.wim and autounattend.xml is required.'
    }
    $root = $media[0].Root
    try {
        $folder = $root + 'ComptonDeploy\Logs'
        New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
        $serial = ([string](Get-CimInstance Win32_BIOS).SerialNumber) -replace '[^A-Za-z0-9-]','_'
        $log = Join-Path $folder ('{0}-StoragePreflight-{1}.log' -f $serial,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    } catch { Log "USB logging unavailable: $($_.Exception.Message)" }
    Log "Installer USB detected at $root."
    $answer = [xml](Get-Content -LiteralPath ($root + 'autounattend.xml') -Raw)
    $setup = @($answer.unattend.settings | Where-Object pass -eq 'windowsPE').component |
        Where-Object name -eq 'Microsoft-Windows-Setup'
    $commands = @($setup.RunSynchronous.RunSynchronousCommand | ForEach-Object Path) -join ' '
    if ($setup.ImageInstall.OSImage.InstallTo.DiskID -ne '0' -or
        $setup.ImageInstall.OSImage.InstallTo.PartitionID -ne '3' -or
        $commands -notmatch '(?i)SELECT DISK=0' -or $commands -notmatch '(?i)\bCLEAN\b') {
        Block 'USB answer file no longer matches disk 0 safety gate.'
    }
    $manufacturer = [string](Get-CimInstance Win32_ComputerSystem).Manufacturer
    Log "Manufacturer: $manufacturer"
    if ($manufacturer -match '(?i)Dell') {
        $cctk = 'X:\Windows\System32\DellConfigure\cctk.exe'
        if (-not (Test-Path -LiteralPath $cctk)) { Block 'Dell Command Configure is missing.' }
        $query = (& $cctk '--EmbSataRaid' 2>&1 | Out-String)
        $queryCode = $LASTEXITCODE
        Log "Dell storage setting: exit=$queryCode; $query"
        if ($queryCode -ne 0) { Block 'Cannot query Dell storage mode; check model and BIOS permissions.' }
        if ($query -match '(?im)\b(raid|raidon|smartresponse)\b') {
            $marker = $root + 'DellAHCI-reboot-attempt.txt'
            if (Test-Path -LiteralPath $marker) { Block 'Dell still reports RAID after an AHCI reboot attempt.' }
            $output = (& $cctk '--EmbSataRaid=Ahci' 2>&1 | Out-String)
            $exitCode = $LASTEXITCODE
            Log "Dell AHCI request: exit=$exitCode; $output"
            if ($exitCode -ne 0) { Block 'Dell AHCI change failed. Check BIOS support or setup password.' }
            Set-Content -LiteralPath $marker -Value 'AHCI requested; do not reboot in a loop.' -Encoding ASCII
            Write-Host ''
            Write-Host 'RAID storage was detected and the change to AHCI was accepted.' -ForegroundColor Yellow
            Write-Host 'A reboot is required before Windows installation.' -ForegroundColor Yellow
            Write-Host 'After reboot, select this USB in the one-time boot menu again.' -ForegroundColor Yellow
            Write-Host 'Press any key when ready to reboot.' -ForegroundColor Cyan
            & cmd.exe /c pause
            if ($LASTEXITCODE -ne 0) { Block 'No keypress was read; reboot cancelled.' }
            & wpeutil.exe Reboot
            Block 'The reboot command returned unexpectedly.'
        }
        if ($query -notmatch '(?im)\bahci\b') { Block 'Unrecognized Dell storage mode.' }
        $marker = $root + 'DellAHCI-reboot-attempt.txt'
        if (Test-Path -LiteralPath $marker) { Remove-Item -LiteralPath $marker -Force }
    }
    $disks = @(Get-Disk)
    foreach ($disk in $disks) { Log "Disk $($disk.Number): $($disk.FriendlyName), bus=$($disk.BusType), size=$($disk.Size)" }
    $internal = @($disks | Where-Object { $_.Number -ne $media[0].Disk.Number -and $_.BusType -notin @('USB','SD','MMC','Virtual','File Backed Virtual') })
    $disk0 = @($disks | Where-Object Number -eq 0)
    if ($disks.Count -ne 2 -or $internal.Count -ne 1 -or $disk0.Count -ne 1 -or
        $disk0[0].Number -ne $internal[0].Number -or $disk0[0].IsOffline -or
        $disk0[0].IsReadOnly -or $disk0[0].Size -lt 64GB) {
        Block 'Disk 0 is not the sole internal disk, or USB/disk layout is ambiguous. No install attempted.'
    }
    Log 'Disk 0 verified as sole internal disk; starting Windows Setup.'
    & 'X:\sources\setup.exe' ('/unattend:' + $root + 'autounattend.xml')
    if ($LASTEXITCODE -ne 0) { Block "Windows Setup launcher returned exit code $LASTEXITCODE." }
    Log 'Setup launcher returned successfully; waiting while installation continues.'
} catch { Block $_.Exception.Message }
'@
    Set-Content -LiteralPath (Join-Path $sys 'StoragePreflight.ps1') -Value $runtime -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $sys 'StoragePreflight.cmd') -Value @'
@echo off
wpeinit
powershell.exe -NoProfile -ExecutionPolicy Bypass -File X:\Windows\System32\StoragePreflight.ps1
:hold
timeout /t 3600 /nobreak >nul
goto hold
'@ -Encoding ASCII
    Set-Content -LiteralPath $launcher -Value @'
[LaunchApps]
%SYSTEMDRIVE%\Windows\System32\cmd.exe,"/c %SYSTEMDRIVE%\Windows\System32\StoragePreflight.cmd"
'@ -Encoding ASCII
    Invoke-Dism -Arguments @('/Unmount-Image', "/MountDir:$mount", '/Commit')
    $mounted = $false
    Copy-Item -LiteralPath $InstallWim -Destination $installBackup -Force
    $installBackupReady = $true
    Copy-Item -LiteralPath $stagedInstall -Destination $InstallWim -Force
    $xml.Save($AnswerFile)
    Write-Host "Updated $BootWim. Cleaned prior domain-join additions from $AnswerFile."
    Write-Host 'HP and Dell already in AHCI continue to Setup after disk checks.'
    Write-Host "Serviced install.wim indexes: $($InstallIndexes -join ', '). Computer name is automatic; domain join is removed."
} catch {
    $recoveryComplete = $false
    if ($installMounted) {
        & dism.exe /Unmount-Image "/MountDir:$installMount" /Discard | Out-Host
        if ($LASTEXITCODE -eq 0) { $installMounted = $false }
    }
    if ($installBackupReady) {
        Copy-Item -LiteralPath $installBackup -Destination $InstallWim -Force
    }
    if ($mounted) {
        & dism.exe /Unmount-Image "/MountDir:$mount" /Discard | Out-Host
        if ($LASTEXITCODE -eq 0) { $mounted = $false }
    }
    if (-not $mounted -and (Test-Path -LiteralPath $backup)) {
        Copy-Item -LiteralPath $backup -Destination $BootWim -Force
    }
    if (Test-Path -LiteralPath $answerBackup) {
        Copy-Item -LiteralPath $answerBackup -Destination $AnswerFile -Force
    }
    $recoveryComplete = $true
    throw
} finally {
    if ($recoveryComplete -and -not $mounted -and -not $installMounted -and (Test-Path -LiteralPath $runDir)) {
        Remove-Item -LiteralPath $runDir -Recurse -Force
    }
    elseif (-not $recoveryComplete -or $mounted -or $installMounted) { Write-Warning "Recovery incomplete; retained recovery files at $runDir. Do not delete them until recovery is complete." }
}
