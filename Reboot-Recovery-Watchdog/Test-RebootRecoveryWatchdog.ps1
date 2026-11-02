#requires -Version 5.1
[CmdletBinding()]
param([string]$WatchdogPath = (Join-Path $PSScriptRoot 'Watch-RebootRecovery.ps1'))
$ErrorActionPreference='Stop'
$tokens=$null;$parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($WatchdogPath,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count -gt 0) {throw ($parseErrors.Message -join '; ')}
# Load only these pure helper definitions. NEVER dot-source the full watchdog.
foreach ($name in @('Get-LoopDecision','Test-NormalMaintenance','Quote-NativeArgument','Save-Json','Get-TaskArgumentText','Read-PlainScriptHeader','Start-RecoveryProcess')) {
    $definitions=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true))
    if ($definitions.Count -ne 1) {throw ('Expected exactly one function: '+$name)}
    . ([scriptblock]::Create($definitions[0].Extent.Text))
}
function Assert-True {param([bool]$Condition,[string]$Name); if (-not $Condition) {throw ('FAILED: '+$Name)};Write-Host ('PASS: '+$Name)}
$execTask=[pscustomobject]@{Actions=@([pscustomobject]@{Arguments='-File test.ps1'})}
$comTask=[pscustomobject]@{Actions=@([pscustomobject]@{ClassId='COM-TEST'})}
Assert-True ((Get-TaskArgumentText -Task $execTask) -eq '-File test.ps1') 'Executable task arguments are read'
Assert-True ((Get-TaskArgumentText -Task $comTask) -eq '') 'COM-handler tasks without Arguments do not fail task discovery'
$now=[datetime]'2026-10-28 14:30:00'
$comment='Pending reboot flags detected during maintenance closeout.'
$requests=@(foreach ($minutes in @(25,16,7)) {[pscustomobject]@{TimeCreated=$now.AddMinutes(-$minutes);Process='C:\Windows\System32\shutdown.exe (TEST)';Comment=$comment}})
$boots=@(foreach ($minutes in @(24,15,6)) {$now.AddMinutes(-$minutes)})
$d=Get-LoopDecision -Requests $requests -Boots $boots -Now $now
Assert-True ($d.Detected -and $d.CompletedRestarts -eq 3) 'Three completed maintenance restarts detected'
$d=Get-LoopDecision -Requests $requests -Boots @() -Now $now
Assert-True (-not $d.Detected) 'Requests without completed boots do not trigger recovery'
$d=Get-LoopDecision -Requests $requests -Boots @($boots[0],$boots[0],$boots[0]) -Now $now
Assert-True ($d.CompletedRestarts -eq 1) 'Duplicate boot events count once'
$manual=@($requests | ForEach-Object {[pscustomobject]@{TimeCreated=$_.TimeCreated;Process=$_.Process;Comment='Manual one-reboot repair: finish existing pending updates; verify once after startup.'}})
Assert-True (-not (Get-LoopDecision -Requests $manual -Boots $boots -Now $now).Detected) 'Manual Force reboot comments are excluded'
$other=@($requests | ForEach-Object {[pscustomobject]@{TimeCreated=$_.TimeCreated;Process='C:\Program Files\Faronics\DFServ.exe';Comment=$comment}})
Assert-True (-not (Get-LoopDecision -Requests $other -Boots $boots -Now $now).Detected) 'Deep Freeze and other processes are excluded'
$intervening=$requests+@([pscustomobject]@{TimeCreated=$boots[2].AddSeconds(-5);Process='C:\Windows\System32\shutdown.exe';Comment='User requested restart'})
Assert-True (-not (Get-LoopDecision -Requests $intervening -Boots $boots -Now $now).Detected) 'Intervening user restart prevents false attribution'
Assert-True (-not (Get-LoopDecision -Requests $requests -Boots $boots -Now $now.AddHours(2)).Detected) 'Old restart history expires'
Assert-True (Test-NormalMaintenance -Now ([datetime]'2026-11-01 01:15:00')) 'Sunday maintenance is excluded'
Assert-True (-not (Test-NormalMaintenance -Now ([datetime]'2026-10-28 14:30:00'))) 'Weekday is outside the normal window'
Assert-True (Test-NormalMaintenance -Now ([datetime]'2026-11-01 21:44:59')) 'Sunday exclusion includes the new 9:45 PM end boundary'
Assert-True (-not (Test-NormalMaintenance -Now ([datetime]'2026-11-01 21:45:00'))) 'Sunday exclusion ends at 9:45 PM'
$quoted=Quote-NativeArgument 'a b"c\'
Assert-True ($quoted -eq '"a b\"c\\"') 'Native argument quoting handles spaces, quotes, trailing backslash'
$temp=Join-Path ([IO.Path]::GetTempPath()) ('Compton-Watchdog-Test-'+[guid]::NewGuid().ToString())
New-Item -Path $temp -ItemType Directory | Out-Null
try {
    $file=Join-Path $temp 'state.json'
    Save-Json -Path $file -Value @{Attempts=0}
    Save-Json -Path $file -Value @{Attempts=1}
    $read=Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
    Assert-True ($read.Attempts -eq 1) 'State write and atomic replacement read back correctly'
}
finally {Remove-Item -LiteralPath $temp -Recurse -Force}
Write-Host 'All offline tests passed. No DFC, task, service, share or reboot operations were performed.' -ForegroundColor Green

# Plain header regression: no provider metadata, correct limit, small JSON.
$headerTestPath=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString()+'.txt')
try {
    [IO.File]::WriteAllLines($headerTestPath,[string[]]@('first','second','third'))
    $plain=@(Read-PlainScriptHeader -Path $headerTestPath -MaximumLines 2)
    Assert-True ($plain.Count -eq 2 -and $plain[0] -eq 'first' -and $plain[1] -eq 'second') 'Plain header limits lines correctly'
    Assert-True ($null -eq $plain[0].PSObject.Properties['PSPath']) 'Plain header has no file-provider metadata'
    $headerJson=ConvertTo-Json -InputObject @{Header=$plain} -Depth 12 -Compress
    Assert-True ($headerJson -eq '{"Header":["first","second"]}') 'Plain header serializes as a compact string array'
}
finally {Remove-Item -LiteralPath $headerTestPath -Force -ErrorAction SilentlyContinue}

# Native process regression; launches only PowerShell child processes that exit.
$nativePowerShell=Join-Path $PSHOME 'powershell.exe'
foreach ($expectedCode in @(0,7)) {
    $probeProcess=Start-RecoveryProcess -FilePath $nativePowerShell -Arguments ('-NoLogo -NoProfile -NonInteractive -Command "exit '+$expectedCode+'"')
    try {
        if (-not $probeProcess.WaitForExit(15000)) {$probeProcess.Kill();throw 'Exit-code regression child timed out.'}
        Assert-True ($null -ne $probeProcess.ExitCode -and $probeProcess.ExitCode -eq $expectedCode) ('Owned process reports exit '+$expectedCode)
    }
    finally {$probeProcess.Dispose()}
}

$redirectProbe=Start-RecoveryProcess -FilePath $nativePowerShell -Arguments '-NoLogo -NoProfile -NonInteractive -Command "Write-Output probe; exit 7"' -OutputPath (Join-Path ([IO.Path]::GetTempPath()) 'Compton-ExitCode-Probe.out')
try {
    if (-not $redirectProbe.WaitForExit(15000)) {$redirectProbe.Kill();throw 'Redirected exit-code child timed out.'}
    Assert-True ($null -ne $redirectProbe.ExitCode -and $redirectProbe.ExitCode -eq 7) 'Redirected owned process reports nonzero exit code'
    Assert-True ($redirectProbe.RecoveryOutputTask.Wait(5000) -and $redirectProbe.RecoveryOutputTask.Result.Trim() -eq 'probe') 'Redirected output is drained without losing exit code'
}
finally {$redirectProbe.Dispose()}
