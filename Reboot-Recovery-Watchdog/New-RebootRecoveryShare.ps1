#requires -Version 5.1
<#
.SYNOPSIS
    Provision a dedicated recovery share on a Windows file server.
.NOTES
    Version 1.0.0. Run elevated on the file server, not a lab endpoint.
    Existing unrelated folders/shares are refused. No maintenance tasks are run.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]]$ComputerNames,
    [string]$Domain = 'COMPTON',
    [string]$ShareName = 'ComptonRebootRecovery$',
    [string]$Path = 'C:\ProgramData\ComptonRebootRecoveryShare'
)
$ErrorActionPreference='Stop'
$principal=New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {throw 'Run elevated on the Windows file server.'}
if ($ShareName -notmatch '^[A-Za-z0-9_-]+\$$') {throw 'Use a dedicated hidden share name ending with $.'}
$root=[IO.Path]::GetFullPath($Path).TrimEnd('\')
if ($root -eq [IO.Path]::GetPathRoot($root).TrimEnd('\') -or $root.StartsWith('\\')) {throw 'Use a dedicated local subdirectory on the file server.'}
$marker=Join-Path $root 'Compton-RebootRecovery-Share.marker'
$existingShare=Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if ($null -ne $existingShare -and [IO.Path]::GetFullPath($existingShare.Path).TrimEnd('\') -ine $root) {throw 'The share name already points to a different path.'}
if ((Test-Path -LiteralPath $root) -and -not (Test-Path -LiteralPath $marker -PathType Leaf)) {throw 'The target folder already exists and is not marked as this dedicated recovery share. Choose a new folder.'}
$accounts=@()
foreach ($computer in @($ComputerNames | Sort-Object -Unique)) {
    if ($computer -notmatch '^[A-Za-z0-9][A-Za-z0-9-]{0,14}$') {throw ('Invalid computer name: '+$computer)}
    $name=$computer.ToUpperInvariant()
    $account=New-Object Security.Principal.NTAccount($Domain,($name+'$'))
    $sid=$account.Translate([Security.Principal.SecurityIdentifier])
    $accounts += [pscustomobject]@{Name=$name;SID=$sid}
}
$domainComputers=(New-Object Security.Principal.NTAccount($Domain,'Domain Computers')).Translate([Security.Principal.SecurityIdentifier])
$admins=New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')
$system=New-Object Security.Principal.SecurityIdentifier('S-1-5-18')
if (-not (Test-Path -LiteralPath $root)) {
    New-Item -Path $root -ItemType Directory -Force | Out-Null
    $acl=New-Object Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true,$false);$acl.SetOwner($admins)
    foreach ($sid in @($admins,$system)) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sid,'FullControl','ContainerInherit,ObjectInherit','None','Allow')))
    }
    # Root only: computers can traverse the share, but cannot read other machines' data.
    $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($domainComputers,'ReadAndExecute','Allow')))
    Set-Acl -LiteralPath $root -AclObject $acl
    'Compton-RebootRecovery-v1' | Set-Content -LiteralPath $marker -Encoding ASCII
}
foreach ($item in $accounts) {
    $folder=Join-Path $root $item.Name
    if (Test-Path -LiteralPath $folder) {
        Write-Host ('Existing computer folder retained without changing permissions: '+$folder)
        continue
    }
    New-Item -Path $folder -ItemType Directory | Out-Null
    $acl=New-Object Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true,$false);$acl.SetOwner($admins)
    foreach ($sid in @($admins,$system)) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sid,'FullControl','ContainerInherit,ObjectInherit','None','Allow')))
    }
    $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($item.SID,'Modify','ContainerInherit,ObjectInherit','None','Allow')))
    Set-Acl -LiteralPath $folder -AclObject $acl
    Write-Host ('Provisioned private folder: '+$item.Name)
}
if ($null -eq $existingShare) {
    $adminName=$admins.Translate([Security.Principal.NTAccount]).Value
    $computersName=$domainComputers.Translate([Security.Principal.NTAccount]).Value
    New-SmbShare -Name $ShareName -Path $root -FullAccess $adminName -ChangeAccess $computersName -FolderEnumerationMode AccessBased -CachingMode None | Out-Null
}
Write-Host ('PersistentStateRoot = ''\\'+$env:COMPUTERNAME+'\'+$ShareName+'''') -ForegroundColor Green
Write-Host 'Verify access under the endpoint SYSTEM account using the diagnostic-only watchdog task.'
