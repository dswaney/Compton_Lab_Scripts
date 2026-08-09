COMPTON COLLEGE MAINTENANCE FRAMEWORK v2.1.1

CANONICAL PATHS
  Scripts: C:\Scripts
  Logs:    C:\Logs

This release intentionally contains no Scripts path or workstation-specific fallback.
All scheduled tasks, module imports, update operations, tests, and build tools use C:\Scripts.

FRAMEWORK
  Maintenance.Framework.psm1 provides:
  - Central configuration
  - Standard exit codes
  - Administrative checks
  - Run context creation
  - Common status/logging helpers
  - File resolution
  - NDJSON telemetry writing
  - Global telemetry mutex
  - 50 MB rotation
  - ZIP compression
  - 60-day archive retention

STANDARD EXIT CODES
  0 Success
  1 Success with warnings
  2 Recoverable failure
  3 Critical failure
  4 Telemetry failure

VALIDATION
  Run elevated Windows PowerShell 5.1:
    C:\Scripts\Build\Validate-AllScripts.ps1

TELEMETRY TEST
    C:\Scripts\Test-MaintenanceTelemetry.ps1

CREATE A RELEASE
    C:\Scripts\Build\Create-Release.ps1 -Version 2.1.1

DEPLOYMENT
  Copy the top-level files and Build folder to the central Labscripts share.
  Script 00 synchronizes approved production files to C:\Scripts.

IMPORTANT
  Existing scripts retain script-specific functions where their behavior differs.
  Shared infrastructure is centralized in Maintenance.Framework.psm1; future script
  revisions should call the exported framework helpers instead of adding new copies.

MANIFEST-DRIVEN UPDATER (SCRIPT 00 VERSION 4.0.0)
--------------------------------------------------
The updater uses DeploymentManifest.json from \\filesvr\Labscripts, with
\\10.2.3.30\Labscripts as the fallback source.

Startup behavior:
1. Reach the source share and read DeploymentManifest.json.
2. Verify the manifest structure and required entries.
3. Bootstrap Maintenance.Framework.psm1 first if it is missing or changed.
4. Verify every source file against the manifest SHA-256 value.
5. Copy each changed file into C:\Scripts\.staging.
6. Run the Windows PowerShell parser against staged .ps1/.psm1/.psd1 files.
7. Back up the current local file under C:\Scripts\Rollback before replacement.
8. Install and verify the destination hash and parser result.
9. Roll back automatically if replacement or post-install validation fails.
10. Update Script 00 last and relaunch the verified replacement.
11. Run Register-Tasks_SYSTEM.ps1 after synchronization succeeds.

Deploy DeploymentManifest.json to the root of \\filesvr\Labscripts together
with every file listed in that manifest. A file not listed in the deployment
manifest is not installed by Script 00.

SHA-256 validation protects against incomplete copies, corruption, and
accidental modification. It does not authenticate a compromised file share;
Authenticode signing would be required for publisher authenticity.

Framework command cleanup:
- Initialize-MaintenanceDirectory uses an approved PowerShell verb and replaces Ensure-MaintenanceDirectory.


V2.2 ENHANCEMENTS
- Invoke-MaintenanceScript.ps1 is the scheduled-task launcher.
- Maintenance.Policy.json controls windows, dependencies, lock timeout, log rotation, and fleet-status paths.
- Fleet status is published as <ComputerName>.json to \\filesvr\MaintenanceStatus with IP fallback.
- Get-MaintenanceFleetStatus.ps1 creates HTML and CSV stale-client reports.
- Scheduled tasks are compliance-checked and reconciled only when drift is found.
- Windows Event Log: Compton Maintenance / Compton-Maintenance.
- Correlation IDs connect launcher logs, telemetry, status, and events.
- All launcher logs rotate at 20 MB and retain 60 days by default.

REQUIRED SHARE PERMISSIONS
Grant Domain Computers Modify/Create Files on \\filesvr\MaintenanceStatus.
Grant Domain Computers Read & Execute on \\filesvr\Labscripts.
