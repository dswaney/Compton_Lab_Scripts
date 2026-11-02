<# 
Check-ADPasswordExpiry.ps1
Prompts for an AD username and shows password expiry status/date.
#>

try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Host "ERROR: ActiveDirectory module not found. Install RSAT (AD PowerShell) and try again." -ForegroundColor Red
    exit 1
}

$userInput = Read-Host "Enter AD username (samAccountName like jdoe OR UPN like jdoe@domain.com)"
if ([string]::IsNullOrWhiteSpace($userInput)) {
    Write-Host "No username entered. Exiting." -ForegroundColor Yellow
    exit 1
}

try {
    $u = Get-ADUser -Identity $userInput -Properties `
        Enabled, PasswordExpired, PasswordNeverExpires, PasswordLastSet, LockedOut, `
        msDS-UserPasswordExpiryTimeComputed -ErrorAction Stop
} catch {
    try {
        $u = Get-ADUser -Filter "UserPrincipalName -eq '$userInput'" -Properties `
            Enabled, PasswordExpired, PasswordNeverExpires, PasswordLastSet, LockedOut, `
            msDS-UserPasswordExpiryTimeComputed -ErrorAction Stop
    } catch {
        Write-Host "ERROR: Could not find user '$userInput' in AD." -ForegroundColor Red
        exit 1
    }
}

# Compute expiry date (safe)
$expiry = $null
$rawExpiry = $u.'msDS-UserPasswordExpiryTimeComputed'

try { $rawExpiry = [Int64]$rawExpiry } catch { $rawExpiry = $null }

if (-not $u.PasswordNeverExpires -and $rawExpiry -and $rawExpiry -gt 0) {
    try { $expiry = [DateTime]::FromFileTime($rawExpiry) } catch { $expiry = $null }
}

Write-Host ""
Write-Host "Results:" -ForegroundColor Cyan
[pscustomobject]@{
    SamAccountName       = $u.SamAccountName
    Enabled              = $u.Enabled
    LockedOut            = $u.LockedOut
    PasswordExpired      = $u.PasswordExpired
    PasswordNeverExpires = $u.PasswordNeverExpires
    PasswordLastSet      = $u.PasswordLastSet
    PasswordExpiryDate   = $expiry
} | Format-List

if ($u.PasswordNeverExpires) {
    Write-Host "Summary: Password is set to NEVER expire." -ForegroundColor Yellow
} elseif ($u.PasswordExpired) {
    Write-Host "Summary: Password IS EXPIRED." -ForegroundColor Red
} elseif ($expiry) {
    Write-Host ("Summary: Password is NOT expired. Expires: {0}" -f $expiry) -ForegroundColor Green
} else {
    Write-Host "Summary: Password is NOT marked expired. Expiry date could not be computed." -ForegroundColor Green
}
