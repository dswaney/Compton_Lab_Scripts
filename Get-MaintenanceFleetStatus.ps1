#requires -Version 5.1
[CmdletBinding()]
param(
 [string]$StatusRoot='\\filesvr\MaintenanceStatus',
 [string]$OutputPath='C:\Logs\Fleet-Maintenance-Status.html',
 [int]$WarningDays=3,[int]$CriticalDays=7
)
Set-StrictMode -Version 2.0
$now=Get-Date
$rows=@()
Get-ChildItem -LiteralPath $StatusRoot -Filter '*.json' -File -ErrorAction Stop | ForEach-Object {
 try {
  $r=Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
  $last=[datetime]$r.LastCheckIn; $age=($now.ToUniversalTime()-$last.ToUniversalTime()).TotalDays
  $stale=if($age -ge $CriticalDays){'Critical'}elseif($age -ge $WarningDays){'Warning'}else{'Current'}
  $rows+=[pscustomobject]@{ComputerName=$r.ComputerName;LastCheckIn=$last;AgeDays=[math]::Round($age,1);StaleStatus=$stale;ScriptName=$r.ScriptName;Status=$r.Status;ExitCode=$r.ExitCode;FrameworkVersion=$r.FrameworkVersion;PolicyVersion=$r.PolicyVersion;CorrelationId=$r.CorrelationId}
 } catch {}
}
$style='<style>body{font-family:Segoe UI,Arial}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px}th{background:#eee}.Critical{background:#f8d7da}.Warning{background:#fff3cd}</style>'
$html=$rows | Sort-Object StaleStatus,ComputerName | ConvertTo-Html -Title 'Compton Maintenance Fleet Status' -Head $style -PreContent ("<h1>Compton Maintenance Fleet Status</h1><p>Generated {0}. Warning: {1} days; Critical: {2} days.</p>" -f $now,$WarningDays,$CriticalDays)
$html=$html -replace '<tr><td>([^<]+)</td><td>([^<]+)</td><td>([^<]+)</td><td>(Critical|Warning|Current)</td>', '<tr class="$4"><td>$1</td><td>$2</td><td>$3</td><td>$4</td>'
$dir=Split-Path -Parent $OutputPath;if(-not(Test-Path $dir)){New-Item $dir -ItemType Directory -Force|Out-Null}
$html | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$rows | Export-Csv -LiteralPath ([IO.Path]::ChangeExtension($OutputPath,'.csv')) -NoTypeInformation -Encoding UTF8
$rows
