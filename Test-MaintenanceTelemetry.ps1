#requires -version 5.1
[CmdletBinding()]
param(
    [int]$ProcessCount = 6,
    [int]$RecordsPerProcess = 100
)

$module = 'C:\Scripts\Maintenance.Framework.psm1'
$testPath = 'C:\Logs\Maintenance-Telemetry-Test.ndjson'
$testMutex = 'Global\Compton_MaintenanceTelemetry_TEST'
Remove-Item -LiteralPath $testPath -Force -ErrorAction SilentlyContinue

$jobs = for ($process = 1; $process -le $ProcessCount; $process++) {
    Start-Job -ArgumentList $module,$testPath,$testMutex,$process,$RecordsPerProcess -ScriptBlock {
        param($Module,$Path,$Mutex,$ProcessNumber,$Count)
        Import-Module $Module -Force
        1..$Count | ForEach-Object {
            $event = [ordered]@{
                Process = $ProcessNumber
                Record = $_
                Timestamp = (Get-Date).ToUniversalTime().ToString('o')
                Payload = ('X' * 1024)
            }
            Write-MaintenanceTelemetryLine -Path $Path -JsonLine ($event | ConvertTo-Json -Compress) -MaxSizeMB 1 -RetentionDays 1 -MutexName $Mutex
        }
    }
}

$jobs | Wait-Job | Receive-Job
$jobs | Remove-Job -Force

$lines = @(Get-Content -LiteralPath $testPath -ErrorAction SilentlyContinue)
$invalid = @($lines | Where-Object { try { $_ | ConvertFrom-Json -ErrorAction Stop | Out-Null; $false } catch { $true } })

[pscustomobject]@{
    ActiveFileRecords = $lines.Count
    InvalidJsonLines = $invalid.Count
    RotatedArchives = @(Get-ChildItem 'C:\Logs\Maintenance-Telemetry-Test_*.zip' -ErrorAction SilentlyContinue).Count
    Result = if ($invalid.Count -eq 0) { 'PASS' } else { 'FAIL' }
} | Format-List
