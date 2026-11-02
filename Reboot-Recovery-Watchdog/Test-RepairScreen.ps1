#requires -Version 5.1
# Offline display protocol tests. The registry backend is replaced with an
# in-memory store BEFORE restore is called. No HKLM/display/task/WTS operations.
[CmdletBinding()]
param([string]$ModulePath=(Join-Path $PSScriptRoot 'RepairScreen.psm1'))
$ErrorActionPreference='Stop'
$tokens=$null;$errors=$null
[void][Management.Automation.Language.Parser]::ParseFile($ModulePath,[ref]$tokens,[ref]$errors)
if ($errors.Count -gt 0) {throw ($errors.Message -join '; ')}
$module=Import-Module -Name $ModulePath -Force -PassThru -DisableNameChecking
function Assert-True {param([bool]$Condition,[string]$Name);if (-not $Condition) {throw ('FAILED: '+$Name)};Write-Host ('PASS: '+$Name)}
Assert-True ((Select-ComptonRepairMethod -ProductSku 48) -eq 'NoticeOnly') 'Pro selects notice-only'
Assert-True ((Select-ComptonRepairMethod -ProductSku 164) -eq 'NoticeOnly') 'Pro Education is not confused with Education'
Assert-True ((Select-ComptonRepairMethod -ProductSku 4) -eq 'LockImageAndNotice') 'Enterprise selects lock image and notice'
Assert-True ((Select-ComptonRepairMethod -ProductSku 140) -eq 'LockImageAndNotice') 'Enterprise Subscription is recognized'
Assert-True ((Select-ComptonRepairMethod -ProductSku 121) -eq 'LockImageAndNotice') 'Education selects lock image and notice'
Assert-True ((Select-ComptonRepairMethod -ProductSku 188) -eq 'LockImageAndNotice') 'IoT Enterprise is recognized'
Assert-True ((Select-ComptonRepairMethod -ProductSku 999) -eq 'NoticeOnly') 'Unknown SKU uses conservative fallback'
Assert-True ((Select-ComptonRepairMethod -ProductSku 4 -Confident $false) -eq 'NoticeOnly') 'Uncertain/conflicting edition uses conservative fallback'
Assert-True ((Select-ComptonRepairMethod -ProductSku 4 -ProductType 3) -eq 'NoticeOnly') 'Server is not treated as a supported client image target'
$text=Join-ComptonRepairNotice -Message 'Repair message' -OriginalCaption 'Corporate policy' -OriginalText 'Existing legal notice'
Assert-True ($text.Contains('Repair message') -and $text.Contains('Corporate policy') -and $text.Contains('Existing legal notice')) 'Original organization notice is retained in repair text'

# Replace module-local private registry functions with a verified in-memory backend.
& $module {
    $script:MockValues=@{}
    function script:Get-RepairRegistryValue {
        param([string]$Path,[string]$Name)
        $key=$Path+'|'+$Name
        if ($script:MockValues.ContainsKey($key)) {return $script:MockValues[$key]}
        return [pscustomobject]@{Exists=$false;Kind=$null;Data=$null}
    }
    function script:Set-RepairRegistryValue {
        param([string]$Path,[string]$Name,[string]$Kind,$Data)
        $script:MockValues[$Path+'|'+$Name]=[pscustomobject]@{Exists=$true;Kind=$Kind;Data=$Data}
    }
    function script:Remove-RepairRegistryValue {
        param([string]$Path,[string]$Name)
        $script:MockValues.Remove($Path+'|'+$Name)
    }
}
$mockConfirmed=& $module {(Get-Command Get-RepairRegistryValue).Definition.Contains('MockValues') -and (Get-Command Set-RepairRegistryValue).Definition.Contains('MockValues') -and (Get-Command Remove-RepairRegistryValue).Definition.Contains('MockValues')}
if (-not $mockConfirmed) {throw 'Registry mocks did not install; refusing to run restore tests.'}
$temp=Join-Path ([IO.Path]::GetTempPath()) ('Compton-RepairScreen-Test-'+[guid]::NewGuid().ToString())
New-Item -Path $temp -ItemType Directory | Out-Null
try {
    $incident=[guid]::NewGuid().ToString()
    $system='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    $personalization='HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization'
    $values=@(
        [pscustomobject]@{Path=$system;Name='legalnoticecaption';Original=[pscustomobject]@{Exists=$true;Kind='String';Data='Original policy'};Applied=[pscustomobject]@{Exists=$true;Kind='String';Data='IT REPAIR IN PROGRESS'}},
        [pscustomobject]@{Path=$system;Name='legalnoticetext';Original=[pscustomobject]@{Exists=$false;Kind=$null;Data=$null};Applied=[pscustomobject]@{Exists=$true;Kind='String';Data='Repair message'}},
        [pscustomobject]@{Path=$personalization;Name='NoChangingLockScreen';Original=[pscustomobject]@{Exists=$true;Kind='DWord';Data=0};Applied=[pscustomobject]@{Exists=$true;Kind='DWord';Data=1}}
    )
    $snapshot=[pscustomobject]@{SchemaVersion=1;IncidentId=$incident;ComputerName='TEST';Phase='Applied';Method='NoticeOnly';Title='IT REPAIR IN PROGRESS';RepairMessage='Repair message';Values=$values;Restoration=@()}
    $file=Join-Path $temp 'RepairScreenState.json'
    $snapshot | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file -Encoding UTF8
    & $module {param($items) foreach ($v in $items) {$script:MockValues[$v.Path+'|'+$v.Name]=$v.Applied}} $values
    $result=Restore-ComptonRepairScreen -IncidentFolder $temp -IncidentId $incident -ComputerName 'TEST'
    Assert-True $result.Restored 'Restoration succeeds'
    $observed=& $module {param($a,$b) @((Get-RepairRegistryValue $a 'legalnoticecaption'),(Get-RepairRegistryValue $a 'legalnoticetext'),(Get-RepairRegistryValue $b 'NoChangingLockScreen'))} $system $personalization
    Assert-True ($observed[0].Data -eq 'Original policy') 'Original string restored'
    Assert-True (-not $observed[1].Exists) 'Originally absent value removed'
    Assert-True ($observed[2].Kind -eq 'DWord' -and $observed[2].Data -eq 0) 'Original DWORD kind/value restored'

    # A newer management setting must survive restoration.
    $snapshot.Phase='Applied'
    $snapshot | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file -Encoding UTF8
    & $module {param($a) $script:MockValues[$a+'|legalnoticecaption']=[pscustomobject]@{Exists=$true;Kind='String';Data='New organization policy'}} $system
    $result=Restore-ComptonRepairScreen -IncidentFolder $temp -IncidentId $incident -ComputerName 'TEST'
    $current=& $module {param($a) Get-RepairRegistryValue $a 'legalnoticecaption'} $system
    Assert-True ($current.Data -eq 'New organization policy') 'Newer external organization policy preserved'

    # A modified message that still contains repair content must not pass readiness.
    $snapshot.Phase='Applied'
    $snapshot | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file -Encoding UTF8
    & $module {param($a) $script:MockValues[$a+'|legalnoticetext']=[pscustomobject]@{Exists=$true;Kind='String';Data='Repair message plus external changes'}} $system
    $blocked=$false
    try {[void](Restore-ComptonRepairScreen -IncidentFolder $temp -IncidentId $incident -ComputerName 'TEST')} catch {$blocked=$true}
    Assert-True $blocked 'Unresolved repair content blocks readiness'
    Write-Host 'Offline display tests passed. No real registry/display settings were changed.' -ForegroundColor Green
}
finally {Remove-Module $module.Name -Force;Remove-Item -LiteralPath $temp -Recurse -Force}
