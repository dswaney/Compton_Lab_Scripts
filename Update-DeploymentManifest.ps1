#requires -Version 5.1

<#
.SYNOPSIS
    Updates DeploymentManifest.json from the current files on \\filesvr\Labscripts.

.DESCRIPTION
    Reconciles DeploymentManifest.json against the approved active-file list,
    adding newly approved files and removing retired or unapproved entries. It
    extracts every file's embedded version and always recalculates SHA-256 so a
    content change cannot be missed merely because its version was not changed.

    The manifest is backed up before any write and the updated JSON is validated
    before replacing the original.

.NOTES
    ScriptVersion: 1.1.0
#>

[CmdletBinding()]
param(
    [string]$SourceRoot = '\\filesvr\Labscripts',
    [string]$ManifestName = 'DeploymentManifest.json',
    [switch]$WhatIfOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Keep this list aligned with $ApprovedMaintenanceFiles in
# 00_Update-Scripts-FromShare.ps1. Order here becomes order in the manifest.
$ApprovedFileDefinitions = @(
    [pscustomobject]@{ Name='00_Update-Scripts-FromShare.ps1';                         Role='Updater' },
    [pscustomobject]@{ Name='01_Enable_Windows_Update_Services.ps1';                  Role='ManagedFile' },
    [pscustomobject]@{ Name='02_Remove_User_Profiles.ps1';                            Role='ManagedFile' },
    [pscustomobject]@{ Name='03_Weekend_Apps_Update.ps1';                             Role='ManagedFile' },
    [pscustomobject]@{ Name='04_Sunday_Lab_Application_Maintenance.ps1';              Role='ManagedFile' },
    [pscustomobject]@{ Name='05_Weekend_HP_Drivers_Update.ps1';                       Role='ManagedFile' },
    [pscustomobject]@{ Name='06_Weekend_Windows_Updates.ps1';                         Role='ManagedFile' },
    [pscustomobject]@{ Name='07_Force_Reboot_Install_Updates.ps1';                    Role='ManagedFile' },
    [pscustomobject]@{ Name='08_System_Repair.ps1';                                   Role='ManagedFile' },
    [pscustomobject]@{ Name='09_Disable_Windows_Update_Services.ps1';                 Role='ManagedFile' },
    [pscustomobject]@{ Name='10_Sync_System_Time.ps1';                                Role='ManagedFile' },
    [pscustomobject]@{ Name='12_Enable-SystemRestore-And-Create-RestorePoint.ps1';     Role='ManagedFile' },
    [pscustomobject]@{ Name='14_Endpoint_Health_Inventory.ps1';                       Role='ManagedFile' },
    [pscustomobject]@{ Name='16_Check_Deep_Freeze_Status.ps1';                        Role='ManagedFile' },
    [pscustomobject]@{ Name='Get-MaintenanceFleetStatus.ps1';                         Role='ManagedFile' },
    [pscustomobject]@{ Name='Invoke-MaintenanceScript.ps1';                           Role='Launcher' },
    [pscustomobject]@{ Name='Maintenance.Framework.psm1';                             Role='Framework' },
    [pscustomobject]@{ Name='Maintenance.Policy.json';                                Role='Policy' },
    [pscustomobject]@{ Name='Register-Tasks_SYSTEM.ps1';                              Role='TaskRegistration' }
)

function Write-Status {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','CHANGE')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host ('[{0}] [{1,-6}] {2}' -f $timestamp, $Level, $Message)
}

function Get-EmbeddedVersion {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $extension = [IO.Path]::GetExtension($Path)

    if ($extension -ieq '.json') {
        try {
            $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
            foreach ($propertyName in @('Version','ScriptVersion','ModuleVersion','PolicyVersion')) {
                if ($null -ne $json.PSObject.Properties[$propertyName]) {
                    $value = [string]$json.$propertyName
                    if (-not [string]::IsNullOrWhiteSpace($value)) {
                        return $value.Trim()
                    }
                }
            }
        }
        catch { }
    }

    if ($extension -in @('.ps1','.psm1','.psd1','.json')) {
        $patterns = @(
            '^\s*#\s*ScriptVersion\s*:\s*(?<Version>[0-9]+(?:\.[0-9]+)+)\s*$',
            '^\s*#\s*ModuleVersion\s*:\s*(?<Version>[0-9]+(?:\.[0-9]+)+)\s*$',
            '^\s*#\s*Version\s*:\s*(?<Version>[0-9]+(?:\.[0-9]+)+)\s*$',
            '^\s*ScriptVersion\s*:\s*(?<Version>[0-9]+(?:\.[0-9]+)+)\s*$',
            '^\s*ModuleVersion\s*:\s*(?<Version>[0-9]+(?:\.[0-9]+)+)\s*$',
            '^\s*Version\s*:\s*(?<Version>[0-9]+(?:\.[0-9]+)+)\s*$',
            '^\s*\$script:ScriptVersion\s*=\s*[''"](?<Version>[0-9]+(?:\.[0-9]+)+)[''"]',
            '^\s*\$ScriptVersion\s*=\s*[''"](?<Version>[0-9]+(?:\.[0-9]+)+)[''"]',
            '^\s*\$script:ModuleVersion\s*=\s*[''"](?<Version>[0-9]+(?:\.[0-9]+)+)[''"]',
            '^\s*\$ModuleVersion\s*=\s*[''"](?<Version>[0-9]+(?:\.[0-9]+)+)[''"]',
            'ModuleVersion\s*=\s*[''"](?<Version>[0-9]+(?:\.[0-9]+)+)[''"]'
        )

        foreach ($line in Get-Content -LiteralPath $Path) {
            foreach ($pattern in $patterns) {
                if ($line -match $pattern) {
                    return $matches.Version
                }
            }
        }
    }

    try {
        $item = Get-Item -LiteralPath $Path
        $fileVersion = [string]$item.VersionInfo.FileVersion
        if (-not [string]::IsNullOrWhiteSpace($fileVersion)) {
            return $fileVersion.Trim()
        }
    }
    catch { }

    return $null
}

function Write-ManifestSafely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Manifest,
        [Parameter(Mandatory)][string]$ManifestPath
    )

    $dir = Split-Path -Parent $ManifestPath
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backup = Join-Path $dir ("DeploymentManifest_{0}.bak.json" -f $stamp)
    $temp = "$ManifestPath.tmp"

    Copy-Item -LiteralPath $ManifestPath -Destination $backup -Force
    Write-Status -Level OK -Message "Backup created: $backup"

    $Manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $temp -Encoding UTF8

    $null = Get-Content -LiteralPath $temp -Raw | ConvertFrom-Json

    Move-Item -LiteralPath $temp -Destination $ManifestPath -Force
    Write-Status -Level OK -Message "Updated manifest: $ManifestPath"
}

$manifestPath = Join-Path $SourceRoot $ManifestName

Write-Status -Message "Source root: $SourceRoot"
Write-Status -Message "Manifest: $manifestPath"

if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
    throw "Source root is unavailable: $SourceRoot"
}

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Manifest not found: $manifestPath"
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

if ($null -eq $manifest.PSObject.Properties['Files']) {
    throw 'Manifest does not contain a Files array.'
}

$changes = [System.Collections.Generic.List[object]]::new()
$existingByName = @{}

foreach ($entry in @($manifest.Files)) {
    $name = [string]$entry.Name
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        $existingByName[$name.ToLowerInvariant()] = $entry
    }
}

$approvedNames = @($ApprovedFileDefinitions | ForEach-Object { [string]$_.Name })
$retiredEntries = @(
    $manifest.Files |
    Where-Object { [string]$_.Name -notin $approvedNames }
)

foreach ($entry in $retiredEntries) {
    Write-Status -Level CHANGE -Message "Removing retired or unapproved manifest entry: $($entry.Name)"
    [void]$changes.Add([pscustomobject]@{
        Name   = [string]$entry.Name
        Change = 'Removed'
    })
}

$rebuiltEntries = [System.Collections.Generic.List[object]]::new()

foreach ($definition in $ApprovedFileDefinitions) {
    $name = [string]$definition.Name
    $role = [string]$definition.Role
    $sourcePath = Join-Path $SourceRoot $name

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required approved file is missing from the source package: $sourcePath"
    }

    $currentVersion = Get-EmbeddedVersion -Path $sourcePath
    if ([string]::IsNullOrWhiteSpace($currentVersion)) {
        throw "Could not determine a version for required approved file: $name"
    }

    $currentHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToUpperInvariant()
    $key = $name.ToLowerInvariant()
    $existing = if ($existingByName.ContainsKey($key)) { $existingByName[$key] } else { $null }

    $oldVersion = if ($null -ne $existing) { [string]$existing.Version } else { $null }
    $oldHash = if ($null -ne $existing) { [string]$existing.SHA256 } else { $null }
    $oldRole = if ($null -ne $existing) { [string]$existing.Role } else { $null }

    $changeTypes = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $existing) { [void]$changeTypes.Add('Added') }
    if ($oldVersion -cne $currentVersion) { [void]$changeTypes.Add('Version') }
    if ($oldHash -cne $currentHash) { [void]$changeTypes.Add('SHA256') }
    if ($oldRole -cne $role) { [void]$changeTypes.Add('Role') }

    if ($changeTypes.Count -gt 0) {
        $description = $changeTypes -join ', '
        Write-Status -Level CHANGE -Message "$name requires manifest update: $description"
        [void]$changes.Add([pscustomobject]@{
            Name       = $name
            Change     = $description
            OldVersion = $oldVersion
            NewVersion = $currentVersion
            OldSHA256  = $oldHash
            NewSHA256  = $currentHash
        })
    }
    else {
        Write-Status -Level OK -Message "$name is current. Version=$currentVersion"
    }

    [void]$rebuiltEntries.Add([pscustomobject][ordered]@{
        Name    = $name
        Role    = $role
        Version = $currentVersion
        SHA256  = $currentHash
    })
}

if ($changes.Count -eq 0) {
    Write-Status -Level OK -Message 'Manifest contents and hashes are current. No changes required.'
}
else {
    $manifest.Files = @($rebuiltEntries)
    if ($null -ne $manifest.PSObject.Properties['GeneratedUtc']) {
        $manifest.GeneratedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    Write-Status -Message "Manifest changes required: $($changes.Count)"

    if ($WhatIfOnly) {
        Write-Status -Level WARN -Message 'WhatIfOnly specified. No changes written.'
    }
    else {
        Write-ManifestSafely -Manifest $manifest -ManifestPath $manifestPath
    }
}

Write-Status -Level OK -Message 'Completed successfully.'
