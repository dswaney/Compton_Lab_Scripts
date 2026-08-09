#requires -version 5.1
[CmdletBinding()]
param([string]$SourceRoot = 'C:\Scripts')
$ErrorActionPreference = 'Stop'
$files = Get-ChildItem -LiteralPath $SourceRoot -File | Where-Object { $_.Extension -in '.ps1','.psm1' } | Sort-Object Name
$results = foreach ($file in $files) {
    $tokens=$null; $errors=$null
    [System.Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors) | Out-Null
    if ($errors.Count -eq 0) {
        [pscustomobject]@{File=$file.Name;ErrorCount=0;Line=$null;Column=$null;ErrorId=$null;Message='No parser errors';Code=$null}
    } else {
        foreach ($e in $errors) {
            [pscustomobject]@{File=$file.Name;ErrorCount=$errors.Count;Line=$e.Extent.StartLineNumber;Column=$e.Extent.StartColumnNumber;ErrorId=$e.ErrorId;Message=$e.Message;Code=$e.Extent.Text}
        }
    }
}
$results | Format-List *
if (($results | Where-Object ErrorCount -gt 0).Count -gt 0) { exit 1 }
exit 0
