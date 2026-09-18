#requires -version 5.1
<#
.SYNOPSIS
    Shared Microsoft Copilot removal and disablement for Compton College endpoints.
.DESCRIPTION
    Provides one canonical routine used by post-deployment, profile cleanup,
    and optional system repair. The routine removes installed/provisioned
    Copilot packages, stops Copilot processes, disables Copilot tasks, applies
    machine and user policies, cleans shortcuts, and verifies the result.
#>

Set-StrictMode -Version 2.0

$script:CopilotModuleVersion = '1.0.0'

function Write-CopilotMessage {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO',
        [scriptblock]$Logger
    )

    if ($null -ne $Logger) {
        & $Logger $Message $Level
        return
    }

    Write-Verbose ("[{0}] {1}" -f $Level,$Message)
}

function Set-CopilotDword {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
    }
    New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force -ErrorAction Stop | Out-Null
}

function Set-CopilotUserPolicies {
    param([Parameter(Mandatory)][string]$HiveRoot)

    Set-CopilotDword -Path "$HiveRoot\Software\Policies\Microsoft\Windows\WindowsCopilot" -Name 'TurnOffWindowsCopilot' -Value 1
    Set-CopilotDword -Path "$HiveRoot\Software\Policies\Microsoft\Windows\WindowsAI" -Name 'RemoveMicrosoftCopilotApp' -Value 1
    Set-CopilotDword -Path "$HiveRoot\Software\Policies\Microsoft\Windows\Explorer" -Name 'HideCopilotButton' -Value 1
    Set-CopilotDword -Path "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name 'ShowCopilotButton' -Value 0
}

function Invoke-ComprehensiveCopilotRemoval {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param(
        [scriptblock]$Logger,
        [switch]$SkipOfflineUserHives,
        [switch]$SkipExplorerRestart
    )

    $result = [ordered]@{
        ModuleVersion                = $script:CopilotModuleVersion
        ProcessesStopped             = 0
        TasksDisabled                = 0
        InstalledPackagesFound       = 0
        InstalledPackagesRemoved     = 0
        ProvisionedPackagesFound     = 0
        ProvisionedPackagesRemoved   = 0
        UserHivesUpdated              = 0
        OfflineUserHivesUpdated       = 0
        ShortcutsRemoved              = 0
        RemainingInstalledPackages   = 0
        RemainingProvisionedPackages = 0
        Failures                      = 0
        RebootRequired                = $false
        Status                        = 'Running'
    }

    $packagePatterns = @(
        '*Copilot*',
        'Microsoft.Windows.Ai.Copilot.Provider',
        'Microsoft.MicrosoftOfficeHub'
    )

    Write-CopilotMessage -Logger $Logger -Message "Starting comprehensive Microsoft Copilot removal (module $($script:CopilotModuleVersion))."

    $machinePolicies = @(
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot'; Name='TurnOffWindowsCopilot'; Value=1 },
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI'; Name='RemoveMicrosoftCopilotApp'; Value=1 },
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer'; Name='HideCopilotButton'; Value=1 },
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Name='HubsSidebarEnabled'; Value=0 },
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Name='EdgeCopilotEnabled'; Value=0 },
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Name='CopilotPageContext'; Value=0 },
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Name='EdgeEntraCopilotPageContext'; Value=0 },
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Name='Microsoft365CopilotChatIconEnabled'; Value=0 }
    )
    foreach ($policy in $machinePolicies) {
        try {
            if ($PSCmdlet.ShouldProcess("$($policy.Path)\$($policy.Name)",'Set Copilot disable policy')) {
                Set-CopilotDword -Path $policy.Path -Name $policy.Name -Value $policy.Value
            }
        }
        catch {
            $result.Failures++
            Write-CopilotMessage -Logger $Logger -Message "Failed setting policy $($policy.Path)\$($policy.Name): $($_.Exception.Message)" -Level ERROR
        }
    }

    foreach ($process in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like '*Copilot*' })) {
        try {
            if ($PSCmdlet.ShouldProcess("$($process.ProcessName) [$($process.Id)]",'Stop Copilot process')) {
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
                $result.ProcessesStopped++
            }
        }
        catch {
            $result.Failures++
            Write-CopilotMessage -Logger $Logger -Message "Failed stopping Copilot process $($process.ProcessName): $($_.Exception.Message)" -Level WARN
        }
    }

    foreach ($task in @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like '*Copilot*' -or $_.TaskPath -like '*Copilot*' })) {
        try {
            if ($task.State -ne 'Disabled' -and $PSCmdlet.ShouldProcess("$($task.TaskPath)$($task.TaskName)",'Disable Copilot scheduled task')) {
                Disable-ScheduledTask -InputObject $task -ErrorAction Stop | Out-Null
                $result.TasksDisabled++
            }
        }
        catch {
            $result.Failures++
            Write-CopilotMessage -Logger $Logger -Message "Failed disabling Copilot task $($task.TaskPath)$($task.TaskName): $($_.Exception.Message)" -Level WARN
        }
    }

    $loadedUserSids = @(Get-ChildItem Registry::HKEY_USERS -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21-(?:\d+-){3}\d+$' } |
        Select-Object -ExpandProperty PSChildName)
    foreach ($sid in $loadedUserSids) {
        try {
            if ($PSCmdlet.ShouldProcess($sid,'Apply Copilot user policies')) {
                Set-CopilotUserPolicies -HiveRoot "Registry::HKEY_USERS\$sid"
                $result.UserHivesUpdated++
            }
        }
        catch {
            $result.Failures++
            Write-CopilotMessage -Logger $Logger -Message "Failed applying Copilot policies to loaded user $sid: $($_.Exception.Message)" -Level WARN
        }
    }

    if (-not $SkipOfflineUserHives) {
        $profileEntries = @(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^S-1-5-21-(?:\d+-){3}\d+$' })
        foreach ($profile in $profileEntries) {
            $sid = [string]$profile.PSChildName
            if ($sid -in $loadedUserSids) { continue }
            $profilePath = [Environment]::ExpandEnvironmentVariables([string]$profile.ProfileImagePath)
            $hiveFile = Join-Path $profilePath 'NTUSER.DAT'
            if (-not (Test-Path -LiteralPath $hiveFile -PathType Leaf)) { continue }
            $mountName = 'ComptonCopilot_{0}' -f ([guid]::NewGuid().ToString('N'))
            try {
                if (-not $PSCmdlet.ShouldProcess($profilePath,'Load offline hive and apply Copilot policies')) { continue }
                & reg.exe load "HKU\$mountName" $hiveFile 2>&1 | Out-Null
                if ($LASTEXITCODE -ne 0) { throw "reg.exe load returned $LASTEXITCODE" }
                try {
                    Set-CopilotUserPolicies -HiveRoot "Registry::HKEY_USERS\$mountName"
                    $result.OfflineUserHivesUpdated++
                }
                finally {
                    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
                    & reg.exe unload "HKU\$mountName" 2>&1 | Out-Null
                }
            }
            catch {
                $result.Failures++
                Write-CopilotMessage -Logger $Logger -Message "Failed applying Copilot policies to offline profile $profilePath: $($_.Exception.Message)" -Level WARN
            }
        }

        $defaultHive = Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT'
        if (Test-Path -LiteralPath $defaultHive -PathType Leaf) {
            $mountName = 'ComptonCopilotDefault'
            try {
                if ($PSCmdlet.ShouldProcess($defaultHive,'Apply Copilot policies to Default User')) {
                    & reg.exe load "HKU\$mountName" $defaultHive 2>&1 | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "reg.exe load returned $LASTEXITCODE" }
                    try {
                        Set-CopilotUserPolicies -HiveRoot "Registry::HKEY_USERS\$mountName"
                        $result.OfflineUserHivesUpdated++
                    }
                    finally {
                        [gc]::Collect(); [gc]::WaitForPendingFinalizers()
                        & reg.exe unload "HKU\$mountName" 2>&1 | Out-Null
                    }
                }
            }
            catch {
                $result.Failures++
                Write-CopilotMessage -Logger $Logger -Message "Failed applying Copilot policies to Default User: $($_.Exception.Message)" -Level WARN
            }
        }
    }

    $installed = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
        $name = [string]$_.Name; $fullName = [string]$_.PackageFullName
        @($packagePatterns | Where-Object { $name -like $_ -or $fullName -like $_ }).Count -gt 0
    } | Sort-Object PackageFullName -Unique)
    $result.InstalledPackagesFound = $installed.Count
    foreach ($package in $installed) {
        try {
            if ($PSCmdlet.ShouldProcess($package.PackageFullName,'Remove Copilot Appx package for all users')) {
                Remove-AppxPackage -Package $package.PackageFullName -AllUsers -ErrorAction Stop
                $result.InstalledPackagesRemoved++
                $result.RebootRequired = $true
            }
        }
        catch {
            $result.Failures++
            Write-CopilotMessage -Logger $Logger -Message "Failed removing installed package $($package.PackageFullName): $($_.Exception.Message)" -Level ERROR
        }
    }

    $provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object {
        $name = [string]$_.DisplayName; $fullName = [string]$_.PackageName
        @($packagePatterns | Where-Object { $name -like $_ -or $fullName -like $_ }).Count -gt 0
    } | Sort-Object PackageName -Unique)
    $result.ProvisionedPackagesFound = $provisioned.Count
    foreach ($package in $provisioned) {
        try {
            if ($PSCmdlet.ShouldProcess($package.PackageName,'Remove provisioned Copilot Appx package')) {
                Remove-AppxProvisionedPackage -Online -PackageName $package.PackageName -AllUsers -ErrorAction Stop | Out-Null
                $result.ProvisionedPackagesRemoved++
                $result.RebootRequired = $true
            }
        }
        catch {
            $result.Failures++
            Write-CopilotMessage -Logger $Logger -Message "Failed removing provisioned package $($package.PackageName): $($_.Exception.Message)" -Level ERROR
        }
    }

    $shortcutRoots = @(
        (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'),
        (Join-Path $env:PUBLIC 'Desktop')
    )
    foreach ($profile in @(Get-ChildItem (Join-Path $env:SystemDrive 'Users') -Directory -Force -ErrorAction SilentlyContinue)) {
        $shortcutRoots += Join-Path $profile.FullName 'Desktop'
        $shortcutRoots += Join-Path $profile.FullName 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs'
    }
    foreach ($root in $shortcutRoots | Select-Object -Unique) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        foreach ($shortcut in @(Get-ChildItem -LiteralPath $root -Recurse -Force -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '(?i)copilot' -and $_.Extension -in @('.lnk','.url') })) {
            try {
                if ($PSCmdlet.ShouldProcess($shortcut.FullName,'Remove Copilot shortcut')) {
                    Remove-Item -LiteralPath $shortcut.FullName -Force -ErrorAction Stop
                    $result.ShortcutsRemoved++
                }
            }
            catch {
                $result.Failures++
                Write-CopilotMessage -Logger $Logger -Message "Failed removing shortcut $($shortcut.FullName): $($_.Exception.Message)" -Level WARN
            }
        }
    }

    if (-not $SkipExplorerRestart -and $result.RebootRequired) {
        foreach ($explorer in @(Get-Process explorer -ErrorAction SilentlyContinue)) {
            try { Stop-Process -Id $explorer.Id -Force -ErrorAction Stop } catch {}
        }
    }

    $result.RemainingInstalledPackages = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
        $name = [string]$_.Name; $fullName = [string]$_.PackageFullName
        @($packagePatterns | Where-Object { $name -like $_ -or $fullName -like $_ }).Count -gt 0
    }).Count
    $result.RemainingProvisionedPackages = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object {
        $name = [string]$_.DisplayName; $fullName = [string]$_.PackageName
        @($packagePatterns | Where-Object { $name -like $_ -or $fullName -like $_ }).Count -gt 0
    }).Count

    if ($WhatIfPreference) {
        $result.Status = 'WhatIf'
    }
    elseif (($result.RemainingInstalledPackages + $result.RemainingProvisionedPackages) -gt 0) {
        $result.Status = 'FailedVerification'
        $result.Failures += $result.RemainingInstalledPackages + $result.RemainingProvisionedPackages
    }
    elseif ($result.Failures -gt 0) {
        $result.Status = 'CompletedWithErrors'
    }
    else {
        $result.Status = 'RemovedOrNotPresent'
    }

    $level = if ($result.Failures -gt 0) { 'ERROR' } else { 'OK' }
    Write-CopilotMessage -Logger $Logger -Level $level -Message ("Copilot removal completed. Status={0}; InstalledRemoved={1}; ProvisionedRemoved={2}; UserHives={3}; OfflineHives={4}; Failures={5}" -f $result.Status,$result.InstalledPackagesRemoved,$result.ProvisionedPackagesRemoved,$result.UserHivesUpdated,$result.OfflineUserHivesUpdated,$result.Failures)
    return [pscustomobject]$result
}

Export-ModuleMember -Function 'Invoke-ComprehensiveCopilotRemoval'
