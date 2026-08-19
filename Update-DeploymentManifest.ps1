#requires -Version 5.1

<#
.SYNOPSIS
    Updates DeploymentManifest.json from the current files on \\filesvr\Labscripts.

.DESCRIPTION
    Reads the manifest Files array, checks each referenced source file, extracts
    its embedded version, and compares it with the manifest version. When the
    version differs, the script calculates SHA-256 and updates both Version and
    SHA256 in DeploymentManifest.json.

    The manifest is backed up before any write and the updated JSON is validated
    before replacing the original.

.NOTES
    ScriptVersion: 1.0.0
#>

[CmdletBinding()]
param(
    [string]$SourceRoot = '\\filesvr\Labscripts',
    [string]$ManifestName = 'DeploymentManifest.json',
    [switch]$WhatIfOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

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

$changes = @()
$warnings = @()

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
    $manifestVersion = [string]$entry.Version

    if ([string]::IsNullOrWhiteSpace($currentVersion)) {
        Write-Status -Level WARN -Message "Could not determine version for $name. Entry left unchanged."
        $warnings += "Version unreadable: $name"
        continue
    }

    if ($currentVersion -eq $manifestVersion) {
        Write-Status -Level OK -Message "$name is current. Version=$currentVersion"
        continue
    }

    $sha256 = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToUpperInvariant()

    Write-Status -Level CHANGE -Message "$name version changed: $manifestVersion -> $currentVersion"
    Write-Status -Message "SHA256: $sha256"

    $oldHash = [string]$entry.SHA256
    $entry.Version = $currentVersion
    $entry.SHA256 = $sha256

    $changes += [pscustomobject]@{
        Name       = $name
        OldVersion = $manifestVersion
        NewVersion = $currentVersion
        OldSHA256  = $oldHash
        NewSHA256  = $sha256
    }
}

if ($changes.Count -eq 0) {
    Write-Status -Level OK -Message 'No version differences found. Manifest was not changed.'
}
else {
    if ($null -ne $manifest.PSObject.Properties['GeneratedUtc']) {
        $manifest.GeneratedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    Write-Status -Message "Entries requiring update: $($changes.Count)"

    foreach ($change in $changes) {
        Write-Status -Level CHANGE -Message ("{0}: v{1} -> v{2}" -f $change.Name, $change.OldVersion, $change.NewVersion)
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
