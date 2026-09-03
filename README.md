# Compton College Lab Maintenance Scripts

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2011-0078D4?logo=windows11&logoColor=white)](https://www.microsoft.com/windows/windows-11)
[![Execution](https://img.shields.io/badge/Execution-SYSTEM-success)](#scheduled-task-deployment)

PowerShell-based maintenance automation for Compton College Windows lab computers. The project keeps recurring endpoint maintenance consistent across multiple labs while providing readable logs, structured telemetry, rollback protection, computer-name targeting, and centralized Task Scheduler management.

The scripts are designed primarily for 64-bit Windows PowerShell 5.1 and normally run as `NT AUTHORITY\SYSTEM` with highest privileges.

> [!IMPORTANT]
> Review all paths, computer-name patterns, service settings, credentials, enrollment information, and maintenance times before deploying these scripts in another environment.

## Contents

- [Project goals](#project-goals)
- [Active maintenance scripts](#active-maintenance-scripts)
- [Detailed script descriptions](#detailed-script-descriptions)
- [Supporting files](#supporting-files)
- [Scheduled task deployment](#scheduled-task-deployment)
- [Sunday maintenance schedule](#sunday-maintenance-schedule)
- [Deployment workflow](#deployment-workflow)
- [Logging and telemetry](#logging-and-telemetry)
- [Security considerations](#security-considerations)
- [Retired scripts](#retired-scripts)

## Project goals

- Automate routine Sunday maintenance on Windows lab computers.
- Keep the scripts on each endpoint synchronized with a central file share.
- Run maintenance under SYSTEM without requiring an interactive administrator session.
- Verify changes instead of assuming that a command succeeded.
- Preserve human-readable operational logs and structured Elastic-compatible telemetry.
- Target lab-specific settings using computer-name prefixes and wildcard patterns.
- Reduce the number of separate scheduled tasks by consolidating related application and lab configuration work.
- Keep rollback copies before replacing or retiring locally deployed scripts.

## Active maintenance scripts

| File | Purpose |
|---|---|
| [`00_Update-Scripts-FromShare.ps1`](./00_Update-Scripts-FromShare.ps1) | Synchronizes the approved maintenance package from the primary or fallback share, validates PowerShell files, preserves rollback copies, retires replaced scripts, and reconciles scheduled tasks. |
| [`01_Enable_Windows_Update_Services.ps1`](./01_Enable_Windows_Update_Services.ps1) | Restores Windows Update services, scheduled tasks, policy settings, and required Windows configuration before the update stages begin. |
| [`02_Remove_User_Profiles.ps1`](./02_Remove_User_Profiles.ps1) | Performs scheduled cleanup of eligible local user profiles while preserving required and protected profiles. |
| [`03_Weekend_Apps_Update.ps1`](./03_Weekend_Apps_Update.ps1) | Runs the weekend application update stage before driver and operating-system maintenance. |
| [`04_Sunday_Lab_Application_Maintenance.ps1`](./04_Sunday_Lab_Application_Maintenance.ps1) | Runs the consolidated printer, PaperCut, Autologon, Edge, Elastic Agent, browser homepage, Honorlock, and Stellarium Location Services maintenance. |
| [`05_Weekend_HP_Drivers_Update.ps1`](./05_Weekend_HP_Drivers_Update.ps1) | Performs HP driver and firmware maintenance with safeguards around sensitive storage-related driver categories. |
| [`06_Weekend_Windows_Updates.ps1`](./06_Weekend_Windows_Updates.ps1) | Installs Windows Updates and is scheduled for two passes during the Sunday maintenance window. |
| [`07_Force_Reboot_Install_Updates.ps1`](./07_Force_Reboot_Install_Updates.ps1) | Forces the planned maintenance reboot, supports update completion, and provides startup-resume verification. |
| [`08_System_Repair.ps1`](./08_System_Repair.ps1) | Runs Windows health and repair operations after the update and reboot stages. |
| [`09_Disable_Windows_Update_Services.ps1`](./09_Disable_Windows_Update_Services.ps1) | Returns Windows Update services and related settings to the college's post-maintenance state. |
| [`10_Sync_System_Time.ps1`](./10_Sync_System_Time.ps1) | Synchronizes system time and runs independently every four hours. |
| [`12_Enable-SystemRestore-And-Create-RestorePoint.ps1`](./12_Enable-SystemRestore-And-Create-RestorePoint.ps1) | Enables System Restore where required and creates a weekly restore point before maintenance changes. |
| [`14_Endpoint_Health_Inventory.ps1`](./14_Endpoint_Health_Inventory.ps1) | Captures the endpoint's final weekly health and compliance inventory after the other Sunday stages finish. |
| [`16_Check_Deep_Freeze_Status.ps1`](./16_Check_Deep_Freeze_Status.ps1) | Checks Deep Freeze status at startup and exits safely on systems where Deep Freeze is not installed. |

## Detailed script descriptions

### `00_Update-Scripts-FromShare.ps1`

The manifest-driven updater maintains `C:\Scripts` from the college deployment share. It prefers `\\filesvr\Labscripts` and uses `\\10.2.3.30\Labscripts` as a fallback.

Major functions include:

- Loads and validates `DeploymentManifest.json`.
- Limits deployment to an explicit approved-file list.
- Deploys selected supplemental files even while the manifest is being refreshed.
- Compares actual source and local SHA-256 hashes.
- Parses PowerShell files before installation and validates them again afterward.
- Creates rollback copies before replacing existing files.
- Updates itself last and relaunches the new version safely.
- Runs `Register-Tasks_SYSTEM.ps1` after synchronization.
- Cleans old staging directories and rollback folders according to retention rules.
- Deploys `04_Sunday_Lab_Application_Maintenance.ps1` and removes the six standalone scripts it replaces.

Retired scripts are removed from `C:\Scripts` only after script 04 exists locally and passes PowerShell parser validation. Removed files are moved into the updater's rollback structure and retained for 30 days.

### `01_Enable_Windows_Update_Services.ps1`

Prepares the endpoint for the Sunday update window by restoring and validating the services, tasks, and settings required by Windows Update.

The script includes:

- Bootstrap synchronization through the latest script 00 before loading shared framework components.
- Service recovery and retry handling for Windows Update-related services.
- Validation of services such as Windows Update, BITS, Delivery Optimization, Update Orchestrator, Cryptographic Services, and Windows Installer.
- Recovery of required Microsoft update scheduled tasks.
- Cleanup of conflicting Windows Update policy values.
- Windows 11 configuration enforcement used by the maintenance environment.
- Controlled reboot handling if critical update services cannot be recovered.
- Scheduled-task reconciliation through `Register-Tasks_SYSTEM.ps1`.
- A consistent Sunday 1:00 AM validation time for the updater task.

### `02_Remove_User_Profiles.ps1`

Performs the weekly local-profile cleanup stage. The script is intended to reduce stale profile accumulation on shared lab systems while protecting accounts and profiles that must remain available.

Review the configured age threshold, exclusions, and account protections before expanding deployment to new computer groups.

### `03_Weekend_Apps_Update.ps1`

Runs the general application-update phase of the maintenance window. It is placed before the combined lab configuration, device-driver, and Windows Update stages so application servicing can finish before later reboots.

### `04_Sunday_Lab_Application_Maintenance.ps1`

Consolidates six former standalone scripts into one scheduled maintenance runner. Each internal section runs in an isolated 64-bit Windows PowerShell process so duplicate helper functions, strict-mode settings, and a section's final `exit` statement cannot interfere with later work.

The combined sections are:

1. SHARP printer driver, PaperCut Print Deploy, and `StudentSecurePrint` connection maintenance.
2. Autologon configuration and Microsoft Edge InPrivate startup.
3. Elastic Agent installation, enrollment, health checking, and package fallback handling.
4. Chrome, Edge, and Firefox homepage/startup policy configuration.
5. Honorlock Chrome extension force-install policy configuration.
6. Windows Location Services configuration for Stellarium.

The easy-to-edit configuration area near the top contains:

- Autologon computer wildcard patterns.
- Elastic Agent computer-name prefixes.
- Honorlock computer wildcard patterns.
- Stellarium computer wildcard patterns.
- A `$true` or `$false` switch for each internal section.
- The shared browser homepage URL.

The runner continues to the next section when one section fails and produces its own combined JSON summary in addition to the preserved per-section logs and telemetry.

> [!CAUTION]
> This script contains the embedded configuration from the former Autologon and Elastic Agent scripts. Treat access to the file and deployment share as sensitive.

### `05_Weekend_HP_Drivers_Update.ps1`

Runs the device-driver and firmware maintenance stage for supported HP systems. Storage, chipset, Intel RST, VMD, and NVMe-related updates should remain subject to the script's safety controls because unattended installation of those categories can affect boot or storage availability.

### `06_Weekend_Windows_Updates.ps1`

Performs Windows Update installation during two scheduled passes:

- The first pass installs the initially available updates.
- The second pass runs after the planned reboot to find updates that became applicable only after the first pass or reboot.

### `07_Force_Reboot_Install_Updates.ps1`

Coordinates the planned maintenance reboot and update-completion stage. A companion startup task runs the script with `-StartupResume` so verification can continue after Windows starts again.

### `08_System_Repair.ps1`

Runs system integrity and repair checks after the Windows Update stages. It is given a one-hour window before the final health inventory so repair operations have time to complete.

### `09_Disable_Windows_Update_Services.ps1`

Applies the college's post-maintenance Windows Update service configuration after both update passes have finished.

### `10_Sync_System_Time.ps1`

Maintains reliable system time independently of the weekly Sunday chain. `Register-Tasks_SYSTEM.ps1` creates daily triggers at:

- 12:00 AM
- 4:00 AM
- 8:00 AM
- 12:00 PM
- 4:00 PM
- 8:00 PM

### `12_Enable-SystemRestore-And-Create-RestorePoint.ps1`

Ensures System Restore is enabled as required and creates a restore point early in the maintenance window, before application, driver, update, and configuration changes.

### `14_Endpoint_Health_Inventory.ps1`

Collects a final weekly endpoint snapshot after the other maintenance stages. The inventory is designed to support proactive troubleshooting, compliance reporting, and Elastic dashboards.

Collected areas include:

- CPU, memory, disk utilization, and disk health.
- Windows edition, version, build, uptime, and last boot.
- Pending-reboot state.
- Device Manager problems.
- Microsoft Defender and Windows Firewall status.
- BitLocker, TPM, and Secure Boot state.
- Network adapters, IP addressing, gateways, and DNS configuration.
- Recent update state.
- Critical system and crash events.
- Required service health.
- Management and monitoring agent status.

### `16_Check_Deep_Freeze_Status.ps1`

Runs at startup to capture Deep Freeze status. Systems without Deep Freeze can exit without creating unnecessary maintenance-launcher telemetry.

## Supporting files

### `Maintenance.Framework.psm1`

Shared PowerShell module used by the maintenance scripts. It centralizes recurring functionality such as:

- Maintenance environment initialization.
- Staged text-log creation and publication.
- Log archival and retention.
- Structured NDJSON telemetry writes.
- Windows Event Log entries.
- Fleet-status publication.
- Common configuration and policy handling.

### `Maintenance.Policy.json`

Central maintenance policy consumed by the framework. It provides shared configuration such as fleet-status locations and policy version information.

### `Invoke-MaintenanceScript.ps1`

Standard launcher used by managed scheduled tasks. It provides consistent invocation behavior and launcher-level telemetry around maintenance scripts.

### `Get-MaintenanceFleetStatus.ps1`

Reads and summarizes the latest maintenance status information published by managed endpoints.

### `DeploymentManifest.json`

Lists files managed by the deployment package, including version and hash metadata. The source share remains authoritative, while the manifest provides package structure and traceability.

### `Update-DeploymentManifest.ps1`

Regenerates or refreshes `DeploymentManifest.json` after maintenance files are added, changed, renamed, or removed.

## Scheduled task deployment

### `Register-Tasks_SYSTEM.ps1`

Creates, updates, validates, and removes Compton College maintenance tasks. It is idempotent: running it again leaves correct tasks alone and repairs only managed properties that differ.

It manages:

- Weekly Sunday task actions and start times.
- SYSTEM principals with highest privileges.
- Task execution settings and time limits.
- The system-time synchronization task.
- The post-reboot startup-resume task.
- The Deep Freeze startup check.
- Removal of obsolete task names from earlier schedules.
- Removal of any managed task whose action references one of the retired standalone scripts.
- Structured task-reconciliation telemetry and verification results.

The script creates the following weekly schedule:

| Order | Sunday time | Scheduled task | Script |
|---:|---:|---|---|
| 1 | 1:00 AM | Check for Updated Scripts | `00_Update-Scripts-FromShare.ps1` |
| 2 | 1:15 AM | Create Weekly System Restore Point | `12_Enable-SystemRestore-And-Create-RestorePoint.ps1` |
| 3 | 1:30 AM | Enable Windows Update Services | `01_Enable_Windows_Update_Services.ps1` |
| 4 | 1:45 AM | Remove User Profiles Weekly | `02_Remove_User_Profiles.ps1` |
| 5 | 2:15 AM | Weekend Apps Update | `03_Weekend_Apps_Update.ps1` |
| 6 | 3:15 AM | Sunday Lab Application Maintenance | `04_Sunday_Lab_Application_Maintenance.ps1` |
| 7 | 4:15 AM | Weekend HP Drivers Update | `05_Weekend_HP_Drivers_Update.ps1` |
| 8 | 5:15 AM | Weekend Windows Updates—First Pass | `06_Weekend_Windows_Updates.ps1` |
| 9 | 6:15 AM | Force Reboot and Install Updates | `07_Force_Reboot_Install_Updates.ps1` |
| 10 | 6:45 AM | Weekend Windows Updates—Second Pass | `06_Weekend_Windows_Updates.ps1` |
| 11 | 7:45 AM | Disable Windows Update Services | `09_Disable_Windows_Update_Services.ps1` |
| 12 | 8:00 AM | System Repair | `08_System_Repair.ps1` |
| 13 | 9:00 AM | Weekly Endpoint Health Inventory | `14_Endpoint_Health_Inventory.ps1` |

Additional managed triggers:

| Trigger | Task | Script and arguments |
|---|---|---|
| Every four hours | Sync System Time | `10_Sync_System_Time.ps1` |
| At system startup | Resume Reboot Verification | `07_Force_Reboot_Install_Updates.ps1 -StartupResume` |
| At system startup | Check Deep Freeze Status | `16_Check_Deep_Freeze_Status.ps1` |

The schedule deliberately leaves larger windows around application maintenance, drivers, Windows Updates, and system repair. Task start times are fixed; they do not guarantee that an earlier task has finished, so execution duration should continue to be monitored through telemetry.

## Deployment workflow

1. Place the approved scripts and supporting files in `\\filesvr\Labscripts`.
2. Keep the fallback share at `\\10.2.3.30\Labscripts` synchronized as required.
3. Remove retired standalone scripts from the active deployment-share folder or archive them outside the managed folder.
4. Run `Update-DeploymentManifest.ps1` after files are added, changed, renamed, or removed.
5. Test `00_Update-Scripts-FromShare.ps1` on a pilot endpoint.
6. Confirm that `04_Sunday_Lab_Application_Maintenance.ps1` was deployed to `C:\Scripts`.
7. Confirm that retired files and retired scheduled tasks were removed.
8. Review `Register-Tasks_SYSTEM.latest.json` and Task Scheduler for the expected task count and times.
9. Expand deployment after the pilot endpoint completes successfully.

## Logging and telemetry

Operational logs are normally written beneath:

```text
C:\Logs
```

Shared structured telemetry is written to:

```text
C:\Logs\Maintenance-Telemetry.ndjson
```

Most scripts also maintain a script-specific `*.latest.json` document containing the most recent execution state. Framework-enabled scripts stage their text logs and publish the completed log only after telemetry finalization so log ingestion receives a stable file.

The combined script adds:

```text
C:\Logs\04_Sunday_Lab_Application_Maintenance.log
C:\Logs\04_Sunday_Lab_Application_Maintenance.latest.json
```

## Security considerations

- Run scripts only from a trusted and access-controlled deployment share.
- Restrict modification rights on `C:\Scripts`, the central share, manifests, framework files, and scheduled tasks.
- Protect Autologon credentials and Elastic Agent enrollment information contained in script 04.
- Remember that Base64 encoding is obfuscation, not encryption.
- Use SYSTEM only where required and keep task actions limited to approved scripts.
- Review log and telemetry output to prevent accidental exposure of passwords, tokens, or other secrets.
- Pilot changes before broad deployment.
- Retain rollback copies only as long as operationally necessary and protect their permissions.
- Treat computer-name targeting as deployment scope control, not as a security boundary.

## Retired scripts

The following standalone scripts were consolidated into [`04_Sunday_Lab_Application_Maintenance.ps1`](./04_Sunday_Lab_Application_Maintenance.ps1). They should no longer remain in `C:\Scripts`, the active deployment-share folder, `DeploymentManifest.json`, or active scheduled tasks.

| Retired file | Replacement section in script 04 |
|---|---|
| `11_Install_SharpDriver_And_PaperCut.ps1` | SHARP printer driver, PaperCut Print Deploy, and shared printer maintenance |
| `13_Configure_Autologon_And_Edge.ps1` | Autologon and Edge InPrivate startup configuration |
| `15_Install_Elastic_Agent.ps1` | Elastic Agent installation, enrollment, and health verification |
| `17_Set_Browser_Homepage.ps1` | Chrome, Edge, and Firefox homepage/startup policies |
| `18_Install_Honorlock_Chrome_Extension.ps1` | Honorlock Chrome extension force-install policy |
| `19_Stellarium_Location_Services.ps1` | Windows Location Services configuration for Stellarium |
