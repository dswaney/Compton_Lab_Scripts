#requires -Version 5.1
# ScriptName:    Update-DeploymentManifest.ps1
# ScriptVersion: 2.2.0
# LastUpdated:   2026-09-18

<#
.SYNOPSIS
    Updates DeploymentManifest.json from the current files on \\SERVER\DeploymentShare.

.DESCRIPTION
    Synchronizes the manifest Files array with the approved maintenance-file
    catalog, adds missing approved files, removes retired or consolidated files, and
    updates Version and SHA256 whenever either value differs from the file on
    the deployment share.

    The manifest is backed up before any write and the updated JSON is validated
    before replacing the original.

.NOTES
    ScriptName:    Update-DeploymentManifest.ps1
    ScriptVersion: 2.2.0
    LastUpdated:   2026-09-18
    Changes:       v2.2.0 adds Maintenance.Copilot.psm1 to the approved catalog.
                   v2.1.0 removes former script 12 from the approved catalog,
                   adds it to the script 04 retirement mapping, and validates
                   that no retired file can also remain approved.
                   v2.0.0 adds consolidated script 04 and retires standalone
                   scripts 11, 12, 13, 15, 17, 18, and Stellarium script 19.
                   Legacy names 19_Set_Browser_Homepage and
                   20_Install_Honorlock are also retired into script 04.
    Changes:       v1.2.0 adds Script 19 for Stellarium Location Services.
    Changes:       v1.1.0 adds Scripts 17 and 18, retires their former Script 19/20
                   names, and detects hash-only changes.
#>

[CmdletBinding()]
param(
    [string]$SourceRoot = '\\SERVER\DeploymentShare',
    [string]$ManifestName = 'DeploymentManifest.json',
    [switch]$WhatIfOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '2.2.0'

# Keep this catalog synchronized with $ApprovedMaintenanceFiles in
# 00_Update-Scripts-FromShare.ps1. Files absent from the manifest are added.
$ApprovedFileCatalog = @(
    [pscustomobject]@{ Name='00_Update-Scripts-FromShare.ps1';                         Role='Updater' },
    [pscustomobject]@{ Name='01_Enable_Windows_Update_Services.ps1';                  Role='MaintenanceScript' },
    [pscustomobject]@{ Name='02_Remove_User_Profiles.ps1';                             Role='MaintenanceScript' },
    [pscustomobject]@{ Name='03_Weekend_Apps_Update.ps1';                              Role='MaintenanceScript' },
    [pscustomobject]@{ Name='04_Sunday_Lab_Application_Maintenance.ps1';               Role='MaintenanceScript' },
    [pscustomobject]@{ Name='05_Weekend_HP_Drivers_Update.ps1';                        Role='MaintenanceScript' },
    [pscustomobject]@{ Name='06_Weekend_Windows_Updates.ps1';                          Role='MaintenanceScript' },
    [pscustomobject]@{ Name='07_Force_Reboot_Install_Updates.ps1';                     Role='MaintenanceScript' },
    [pscustomobject]@{ Name='08_System_Repair.ps1';                                    Role='MaintenanceScript' },
    [pscustomobject]@{ Name='09_Disable_Windows_Update_Services.ps1';                  Role='MaintenanceScript' },
    [pscustomobject]@{ Name='10_Sync_System_Time.ps1';                                 Role='MaintenanceScript' },
    [pscustomobject]@{ Name='14_Endpoint_Health_Inventory.ps1';                        Role='MaintenanceScript' },
    [pscustomobject]@{ Name='16_Check_Deep_Freeze_Status.ps1';                         Role='MaintenanceScript' },
    [pscustomobject]@{ Name='Get-MaintenanceFleetStatus.ps1';                          Role='Utility' },
    [pscustomobject]@{ Name='Invoke-MaintenanceScript.ps1';                            Role='Launcher' },
    [pscustomobject]@{ Name='Maintenance.Copilot.psm1';                                Role='Framework' },
    [pscustomobject]@{ Name='Maintenance.Framework.psm1';                              Role='Framework' },
    [pscustomobject]@{ Name='Maintenance.Policy.json';                                 Role='Policy' },
    [pscustomobject]@{ Name='Register-Tasks_SYSTEM.ps1';                               Role='TaskRegistration' }
)

$RetiredFileReplacements = @(
    [pscustomobject]@{ OldName='11_Install_SharpDriver_And_PaperCut.ps1';          NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='12_Enable-SystemRestore-And-Create-RestorePoint.ps1'; NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='13_Configure_Autologon_And_Edge.ps1';             NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='15_Install_Elastic_Agent.ps1';                    NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='17_Set_Browser_Homepage.ps1';                     NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='18_Install_Honorlock_Chrome_Extension.ps1';        NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='19_Stellarium_Location_Services.ps1';             NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='19_Set_Browser_Homepage.ps1';                     NewName='04_Sunday_Lab_Application_Maintenance.ps1' },
    [pscustomobject]@{ OldName='20_Install_Honorlock_Chrome_Extension.ps1';        NewName='04_Sunday_Lab_Application_Maintenance.ps1' }
)

# Prevent a retired file from being reintroduced into the manifest if the two
# catalogs are edited independently in a future version.
$retiredNameLookup = @{}
foreach ($retiredFile in $RetiredFileReplacements) {
    $retiredNameLookup[[string]$retiredFile.OldName] = $true
}

$catalogConflicts = @(
    $ApprovedFileCatalog |
    Where-Object { $retiredNameLookup.ContainsKey([string]$_.Name) } |
    ForEach-Object { [string]$_.Name }
)

if ($catalogConflicts.Count -gt 0) {
    throw "Internal catalog error: retired files are also listed as approved: $($catalogConflicts -join ', ')"
}

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

function Set-ObjectProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$InputObject,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$Value
    )

    if ($null -ne $InputObject.PSObject.Properties[$Name]) {
        $InputObject.$Name = $Value
    }
    else {
        $InputObject | Add-Member -MemberType NoteProperty -Name $Name -Value $Value
    }
}

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
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

$changes = @()
$warnings = @()

$duplicateNames = @(
    @($manifest.Files) |
    Group-Object -Property Name |
    Where-Object { $_.Count -gt 1 }
)

if ($duplicateNames.Count -gt 0) {
    throw "Manifest contains duplicate file entries: $($duplicateNames.Name -join ', ')"
}

# Add approved files that are present on the share but missing from the manifest.
foreach ($approvedFile in $ApprovedFileCatalog) {
    $name = [string]$approvedFile.Name
    $existingEntry = @($manifest.Files | Where-Object { [string]$_.Name -ieq $name }) | Select-Object -First 1

    if ($null -ne $existingEntry) {
        continue
    }

    $sourcePath = Join-Path $SourceRoot $name
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        Write-Status -Level WARN -Message "Approved file is missing from the source share and cannot be added: $sourcePath"
        $warnings += "Missing approved file: $name"
        continue
    }

    $currentVersion = Get-EmbeddedVersion -Path $sourcePath
    if ([string]::IsNullOrWhiteSpace($currentVersion)) {
        Write-Status -Level WARN -Message "Could not determine version for approved file $name. Entry was not added."
        $warnings += "Version unreadable: $name"
        continue
    }

    $sha256 = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToUpperInvariant()
    $newEntry = [pscustomobject][ordered]@{
        Name    = $name
        Role    = [string]$approvedFile.Role
        Version = $currentVersion
        SHA256  = $sha256
    }

    $manifest.Files = @($manifest.Files) + $newEntry
    $changes += [pscustomobject]@{
        ChangeType = 'Added'
        Name       = $name
        OldVersion = $null
        NewVersion = $currentVersion
        OldSHA256  = $null
        NewSHA256  = $sha256
    }
    Write-Status -Level CHANGE -Message "Added approved manifest entry: $name (v$currentVersion)"
}

# Remove an obsolete name only after its replacement exists on the share and
# is represented in the manifest.
foreach ($replacement in $RetiredFileReplacements) {
    $retiredName = [string]$replacement.OldName
    $replacementName = [string]$replacement.NewName
    $retiredEntry = @($manifest.Files | Where-Object { [string]$_.Name -ieq $retiredName }) | Select-Object -First 1

    if ($null -eq $retiredEntry) {
        continue
    }

    $replacementSourcePath = Join-Path $SourceRoot $replacementName
    $replacementEntry = @($manifest.Files | Where-Object { [string]$_.Name -ieq $replacementName }) | Select-Object -First 1

    if (-not (Test-Path -LiteralPath $replacementSourcePath -PathType Leaf) -or $null -eq $replacementEntry) {
        Write-Status -Level WARN -Message "Retired entry was retained because its replacement is not ready: $retiredName -> $replacementName"
        $warnings += "Replacement not ready: $replacementName"
        continue
    }

    $manifest.Files = @($manifest.Files | Where-Object { [string]$_.Name -ine $retiredName })
    $changes += [pscustomobject]@{
        ChangeType = 'Removed'
        Name       = $retiredName
        OldVersion = [string](Get-ObjectPropertyValue -InputObject $retiredEntry -Name 'Version')
        NewVersion = $null
        OldSHA256  = [string](Get-ObjectPropertyValue -InputObject $retiredEntry -Name 'SHA256')
        NewSHA256  = $null
    }
    Write-Status -Level CHANGE -Message "Removed retired manifest entry: $retiredName"
}

foreach ($entry in @($manifest.Files)) {
    $name = [string]$entry.Name

    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Status -Level WARN -Message 'Skipping entry with no Name.'
        $warnings += 'Unnamed manifest entry'
        continue
    }

    $sourcePath = Join-Path $SourceRoot $name

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        Write-Status -Level WARN -Message "Missing referenced file: $sourcePath"
        $warnings += "Missing: $name"
        continue
    }

    $currentVersion = Get-EmbeddedVersion -Path $sourcePath
    $manifestVersion = [string](Get-ObjectPropertyValue -InputObject $entry -Name 'Version')
    $manifestHash = [string](Get-ObjectPropertyValue -InputObject $entry -Name 'SHA256')

    if ([string]::IsNullOrWhiteSpace($currentVersion)) {
        Write-Status -Level WARN -Message "Could not determine version for $name. Entry left unchanged."
        $warnings += "Version unreadable: $name"
        continue
    }

    $sha256 = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToUpperInvariant()

    if ($currentVersion -eq $manifestVersion -and $sha256 -eq $manifestHash) {
        Write-Status -Level OK -Message "$name is current. Version=$currentVersion; SHA256 verified."
        continue
    }

    Write-Status -Level CHANGE -Message "$name requires a manifest update. Version: $manifestVersion -> $currentVersion"
    Write-Status -Message "SHA256: $sha256"

    Set-ObjectProperty -InputObject $entry -Name 'Version' -Value $currentVersion
    Set-ObjectProperty -InputObject $entry -Name 'SHA256' -Value $sha256

    $changes += [pscustomobject]@{
        ChangeType = 'Updated'
        Name       = $name
        OldVersion = $manifestVersion
        NewVersion = $currentVersion
        OldSHA256  = $manifestHash
        NewSHA256  = $sha256
    }
}

if ($changes.Count -eq 0) {
    Write-Status -Level OK -Message 'No manifest differences found. Manifest was not changed.'
}
else {
    if ($null -ne $manifest.PSObject.Properties['GeneratedUtc']) {
        $manifest.GeneratedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    Write-Status -Message "Entries requiring update: $($changes.Count)"

    foreach ($change in $changes) {
        switch ($change.ChangeType) {
            'Added'   { Write-Status -Level CHANGE -Message ("Added {0} at v{1}" -f $change.Name, $change.NewVersion) }
            'Removed' { Write-Status -Level CHANGE -Message ("Removed retired entry {0}" -f $change.Name) }
            default   { Write-Status -Level CHANGE -Message ("Updated {0}: v{1} -> v{2}" -f $change.Name, $change.OldVersion, $change.NewVersion) }
        }
    }

    if ($WhatIfOnly) {
        Write-Status -Level WARN -Message 'WhatIfOnly specified. No changes written.'
    }
    else {
        Write-ManifestSafely -Manifest $manifest -ManifestPath $manifestPath
    }
}

if ($warnings.Count -gt 0) {
    Write-Status -Level WARN -Message "Completed with $($warnings.Count) warning(s)."
}
else {
    Write-Status -Level OK -Message 'Completed successfully.'
}
