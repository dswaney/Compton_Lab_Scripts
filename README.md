# 🖥️ Compton Lab Maintenance Automation

PowerShell-based maintenance and automation framework designed to manage Windows computer lab workstations in a consistent, repeatable, and observable manner.

The project automates common workstation maintenance tasks including:

- 🔄 Windows Updates
- 📦 Application Updates
- 🧹 User Profile Cleanup
- 🛠️ Windows System Repair
- 💾 Driver and Firmware Maintenance
- 🖨️ Printer / PaperCut Deployment
- 🕐 Time Synchronization
- ♻️ System Restore Management
- 🔐 Lab-Specific Configuration
- 📊 Endpoint Health Collection
- 📡 Elastic Agent Deployment
- ❄️ Deep Freeze Status Monitoring
- 🌐 Browser Homepage and Sign-In Policy
- 🧩 Honorlock Chrome Extension Deployment
- 📝 Centralized Maintenance Logging
- 📈 Maintenance Telemetry
- 🔎 Deployment Validation

The scripts are designed to operate primarily through **Windows Task Scheduler under the SYSTEM account**, allowing routine maintenance to occur without requiring an administrator to manually visit each workstation.

---

## 📋 Project Goals

The goal of this project is to move computer lab maintenance from a collection of individual scripts toward a **centrally managed endpoint maintenance framework**.

The system is designed around several principles:

**Automation**  
Routine maintenance should occur automatically and consistently.

**Centralized Management**  
Scripts can be maintained centrally and distributed to endpoints.

**Reliability**  
Maintenance tasks include logging, error handling, validation, and fallback mechanisms.

**Observability**  
Maintenance results can be collected and forwarded to centralized logging platforms for reporting and troubleshooting.

**Controlled Deployment**  
Configuration and deployment policies allow changes to be introduced gradually instead of immediately affecting every workstation.

**Maintainability**  
Common functionality is moved into shared framework components instead of being independently recreated in every script.

---

# 🔄 Maintenance Workflow

A typical workstation maintenance cycle follows approximately this process:

```text
                 Central Script Repository
                          │
                          ▼
             00_Update-Scripts-FromShare
                          │
                          ▼
                  C:\Scripts on Endpoint
                          │
                          ▼
                Windows Task Scheduler
                          │
          ┌───────────────┼────────────────┐
          ▼               ▼                ▼
     Maintenance       Updates        Health Checks
       Scripts
          │               │                │
          └───────────────┼────────────────┘
                          ▼
                       C:\Logs
                          │
                          ▼
                Maintenance Telemetry
                          │
                          ▼
               Central Logging / Elastic
```

This architecture allows the endpoint to continue performing its scheduled maintenance while also producing structured information that can be used for centralized monitoring and reporting.

---

# 📂 Maintenance Scripts

## `00_Update-Scripts-FromShare.ps1`

### Central Script Synchronization

Maintains the local maintenance script set on each workstation.

The script synchronizes the contents of the centralized maintenance repository with the workstation's local script directory, allowing updated versions of maintenance scripts to be deployed without manually copying files to individual computers.

Key responsibilities include:

- Updating locally installed maintenance scripts.
- Retrieving files from the configured central source.
- Supporting alternate/fallback sources when the primary source is unavailable.
- Validating deployment files.
- Maintaining the local maintenance framework.
- Providing logging for synchronization operations.
- Helping ensure endpoints are running the expected versions of maintenance scripts.

This script acts as the **deployment mechanism for the rest of the maintenance environment**.

---

## `01_Enable_Windows_Update_Services.ps1`

### Prepare Windows Update

Prepares the workstation for its scheduled Windows Update maintenance window.

The script ensures that services required by Windows Update are available before the update process begins.

This allows Windows Update services to remain restricted outside the maintenance period while still allowing the automated update process to operate when scheduled.

Typical responsibilities include:

- Checking required Windows Update services.
- Correcting service startup configuration where appropriate.
- Starting required update services.
- Logging service state and maintenance results.
- Preparing the computer for later Windows Update operations.

This script works together with:

`06_Weekend_Windows_Updates.ps1`

and

`09_Disable_Windows_Update_Services.ps1`

to create a controlled Windows Update maintenance window.

---

## `02_Remove_User_Profiles.ps1`

### User Profile Cleanup

Performs automated cleanup of unnecessary local user profiles on shared lab workstations.

Computer labs can accumulate large numbers of profiles as students and other users sign into machines over time. These profiles consume disk space and can eventually contribute to storage and performance problems.

The script is designed to:

- Identify profiles eligible for removal.
- Protect required/system profiles.
- Remove qualifying user profiles.
- Recover disk space.
- Record cleanup activity.
- Report errors encountered during profile removal.

Automating this process reduces the need for technicians to manually clean profiles from individual computers.

---

## `03_Weekend_Apps_Update.ps1`

### Application Maintenance

Performs scheduled application updates during the maintenance window.

The goal is to keep commonly installed applications current without requiring technicians to manually update software on individual lab computers.

The script handles the application maintenance process while generating logs that can later be reviewed to determine whether updates completed successfully.

This helps reduce:

- Outdated software.
- Application vulnerabilities.
- Version inconsistencies between lab computers.
- Manual technician workload.

---

## `04_Update_Edge_Silent.ps1`

### Microsoft Edge Update

Legacy/dedicated Microsoft Edge update script.

This script was originally used to perform silent Microsoft Edge maintenance independently from the broader application update process.

> **Note:** This functionality may overlap with newer application maintenance processes and is retained primarily for compatibility or historical deployment requirements.

---

## `05_Weekend_HP_Drivers_Update.ps1`

### HP Driver and Firmware Maintenance

Automates supported HP workstation maintenance using HP management/update tooling.

The script is designed to keep appropriate device drivers and firmware current while avoiding driver categories that could introduce unnecessary risk to unattended lab systems.

Responsibilities include:

- Detecting applicable HP hardware.
- Running HP update tooling.
- Evaluating available updates.
- Installing approved update categories.
- Restricting selected storage-related or otherwise sensitive unattended updates.
- Logging update results and failures.

Driver maintenance is intentionally more controlled than general application updating because certain storage, chipset, firmware, or controller changes can affect system bootability.

---

## `06_Weekend_Windows_Updates.ps1`

### Windows Update Installation

Performs the primary scheduled Windows Update maintenance operation.

After Windows Update services have been prepared by script `01`, this script handles installation of available Windows updates during the scheduled maintenance window.

The process is designed for unattended operation and includes logging so update activity can later be reviewed centrally.

Typical responsibilities include:

- Checking for available Windows updates.
- Installing approved updates.
- Recording update results.
- Detecting update failures.
- Identifying conditions requiring a reboot.

This is one of the primary scripts responsible for keeping lab computers patched.

---

## `07_Force_Reboot_Install_Updates.ps1`

### Maintenance Reboot

Handles scheduled reboot operations required to complete maintenance.

Some Windows updates and system changes cannot fully complete until the computer restarts. This script ensures machines do not remain indefinitely in a pending-reboot state.

Responsibilities include:

- Detecting maintenance/reboot conditions.
- Allowing pending Windows maintenance to complete.
- Performing scheduled system restarts.
- Logging reboot-related activity.

This helps ensure that updates installed earlier in the maintenance cycle actually become active.

---

## `08_System_Repair.ps1`

### Windows Health and Repair

Performs automated Windows system integrity checks and repair operations.

The script is intended to detect and repair common Windows component and operating system corruption before it develops into a larger support issue.

Maintenance can include Windows servicing and system-file integrity operations such as:

```powershell
DISM
SFC
```

The script records the results of repair operations so recurring integrity problems can be identified instead of silently occurring on individual computers.

---

## `09_Disable_Windows_Update_Services.ps1`

### Close Windows Update Maintenance Window

Returns Windows Update-related services to their expected post-maintenance configuration.

This script runs after the scheduled update cycle and complements:

`01_Enable_Windows_Update_Services.ps1`

Together, these scripts provide a controlled update window:

```text
Enable Update Services
        │
        ▼
Install Windows Updates
        │
        ▼
Reboot / Complete Updates
        │
        ▼
Disable Update Services
```

The script also accounts for Windows services that may be protected or controlled directly by the operating system.

---

## `10_Sync_System_Time.ps1`

### System Time Synchronization

Checks and corrects Windows time synchronization.

Accurate system time is important for:

- Active Directory authentication.
- Kerberos.
- Event log correlation.
- Security investigations.
- Scheduled tasks.
- Certificate validation.
- Centralized logging.

The script helps ensure lab computers remain synchronized with the expected time source and records synchronization results.

---

## `11_Install_SharpDriver_And_PaperCut.ps1`

### Printing Environment Deployment

Automates installation and configuration of required Sharp printing components and PaperCut software.

The script reduces the amount of manual printer configuration required when deploying or repairing lab computers.

Responsibilities include installation and validation of the required printing environment while logging deployment results.

---

## `12_Enable-SystemRestore-And-Create-RestorePoint.ps1`

### System Restore Protection

Ensures Windows System Restore is available and creates a known restore point.

This provides an additional recovery mechanism before or during maintenance operations.

The script can:

- Enable System Restore.
- Verify restore configuration.
- Create a restore point.
- Record whether the operation succeeded.

This provides another recovery option when maintenance or software changes create an unexpected workstation problem.

---

## `13_Configure_Autologon_And_Edge.ps1`

### Lab-Specific Workstation Configuration

Applies configuration settings required by selected computer labs.

Unlike scripts intended to run identically across every workstation, this script uses workstation naming or lab-selection logic to determine where configuration should be applied.

Configuration responsibilities can include:

- Lab autologon configuration.
- Default user configuration.
- Domain configuration.
- Microsoft Edge configuration.
- Removal or correction of obsolete settings.
- Lab-specific workstation behavior.

This allows specialized lab requirements to remain automated without applying those settings to unrelated systems.

> ⚠️ Autologon configuration should always be treated as security-sensitive and access to configuration data should be appropriately restricted.

---

## `14_Endpoint_Health_Inventory.ps1`

### Endpoint Health and Inventory Collection

Collects a broad health snapshot from each workstation.

Rather than making configuration changes, this script provides visibility into the current condition of the endpoint.

Collected information can include:

- Computer identification.
- Hardware information.
- CPU information.
- Memory utilization.
- Disk utilization and health.
- Windows version/build.
- System uptime.
- Last boot time.
- Pending reboot state.
- Device Manager problems.
- Microsoft Defender status.
- Windows Firewall status.
- BitLocker status.
- TPM status.
- Secure Boot status.
- Network adapter information.
- IP configuration.
- DNS configuration.
- Windows Update information.
- Recent crash or critical events.
- Important Windows service status.
- Management/monitoring agent status.

The resulting data can be used to identify trends and eventually support centralized health dashboards and proactive maintenance.

---

## `15_Install_Elastic_Agent.ps1`

### Elastic Agent Deployment

Automates installation of the Elastic Agent on selected lab computers.

The script supports phased deployment rather than immediately installing the agent across the entire workstation environment.

Its purpose is to prepare endpoints for centralized collection of maintenance telemetry and other approved log sources.

The deployment process is designed to support:

- Lab-based targeting.
- Controlled rollout.
- Local installation packages.
- Alternate package sources.
- Download fallback when necessary.
- Installation validation.
- Deployment logging.

Once configured, Elastic Agent provides the connection between workstation telemetry and the centralized Elastic logging environment.

---

## `16_Check_Deep_Freeze_Status.ps1`

### Deep Freeze Monitoring

Checks the current status of Faronics Deep Freeze on applicable lab workstations.

Deep Freeze state is important during maintenance because updates or configuration changes performed while a system is frozen may be lost after restart.

The script provides visibility into whether the workstation is in the expected Deep Freeze state and records that information for maintenance reporting.

This information can eventually be correlated with update failures or maintenance anomalies.

---

## `17_Set_Browser_Homepage.ps1`

### Browser Homepage and Sign-In Policy

Configures a consistent machine-wide homepage and startup page for Mozilla Firefox, Google Chrome, and Microsoft Edge.

The default homepage is:

`https://www.compton.edu`

The script applies browser policies under the computer-level registry so the configuration is available to every user who signs in to the workstation.

Responsibilities include:

- Setting the Firefox, Chrome, and Edge homepages.
- Opening the configured homepage when each browser starts.
- Enabling the browser Home button where supported.
- Disabling Chrome browser/profile sign-in prompts and synchronization.
- Preventing Chrome sign-in interception and promotional sign-in tabs.
- Preserving the ability to sign in normally to websites.
- Verifying every applied policy value.
- Recording changes, verification results, logs, and maintenance telemetry.

The browser policies take effect after the affected browser is closed and reopened or after its enterprise policies refresh.

---

## `18_Install_Honorlock_Chrome_Extension.ps1`

### Honorlock Chrome Extension Deployment

Force-installs the Honorlock extension in Google Chrome through machine-wide Chrome enterprise policy.

The script is designed for controlled lab deployment and currently targets computer names matching:

- `SSC-216*`
- `AHB-146*`

Computers outside the configured pattern list record a successful `NotTargeted` result and make no policy changes.

Responsibilities include:

- Evaluating computer-name wildcard patterns before deployment.
- Creating Chrome's `ExtensionInstallForcelist` policy when required.
- Preserving other extensions already present in the force-install list.
- Updating an existing Honorlock entry when its policy value is incorrect.
- Selecting the next available numeric policy entry for a new installation.
- Verifying that the Honorlock policy exists after configuration.
- Producing local logs and structured maintenance telemetry.

Chrome installs or updates the extension when browser policy refreshes or when Chrome next starts. Additional labs can be introduced by adding their computer-name patterns to the script configuration.

---

# 🧰 Maintenance Framework

The repository contains several additional files that provide the underlying deployment, logging, validation, and orchestration framework.

---

## `Maintenance.Framework.psm1`

### Shared PowerShell Framework Module

Contains reusable PowerShell functions used throughout the maintenance environment.

Instead of duplicating common functionality in every maintenance script, shared operations can be maintained in this module.

The framework helps standardize areas such as:

- Logging.
- Error handling.
- Script initialization.
- Maintenance result reporting.
- Telemetry generation.
- Policy handling.
- Common validation operations.

Centralizing these functions makes the maintenance environment easier to update and helps scripts produce consistent output.

---

## `Maintenance.Policy.json`

### Maintenance Policy Configuration

Provides centralized configuration used by the maintenance framework.

Separating policy from script logic makes it possible to modify operational behavior without rewriting every maintenance script.

Policy configuration can control how framework components behave and provides a common configuration source for maintenance operations.

---

## `Invoke-MaintenanceScript.ps1`

### Maintenance Script Wrapper

Provides a standardized method for launching maintenance scripts through the framework.

Rather than every scheduled task independently implementing initialization, validation, execution, logging, and error handling, this wrapper provides a common execution path.

Conceptually:

```text
Task Scheduler
      │
      ▼
Invoke-MaintenanceScript
      │
      ▼
Maintenance Framework
      │
      ▼
Selected Maintenance Script
      │
      ▼
Logging / Telemetry
```

This makes execution behavior more predictable across the maintenance environment.

---

# ⏱️ Scheduled Task Management

## `Register-Tasks_SYSTEM.ps1`

Creates and maintains the Windows Scheduled Tasks used to execute the maintenance scripts.

Tasks are configured to run under the Windows **SYSTEM** account so maintenance can occur without requiring an interactive administrator login.

The script centralizes task configuration including:

- Script execution.
- Maintenance schedules.
- SYSTEM execution.
- PowerShell execution parameters.
- Task replacement/update behavior.
- Scheduled maintenance sequencing.

This allows scheduled-task changes to be deployed consistently instead of manually modifying Task Scheduler on every computer.

The current rotation includes the browser-homepage policy at **08:40 Sunday** and the Honorlock policy at **08:45 Sunday**. The Honorlock task can exist on every endpoint because Script 18 performs its own computer-name targeting before making changes.

---

# 📊 Fleet Status and Monitoring

## `Get-MaintenanceFleetStatus.ps1`

Provides a maintenance status view intended to help determine the condition of deployed endpoints.

The script can be used to aggregate or evaluate maintenance information and identify computers that may require additional attention.

This supports the larger goal of moving from:

> **Reactive workstation support**

toward:

> **Proactive endpoint monitoring and maintenance**

---

# 🧪 Testing and Validation

## `Test-AllMaintenanceScripts.ps1`

Performs validation of the maintenance PowerShell scripts.

This provides a way to identify script problems before updated scripts are broadly deployed to lab computers.

Testing the complete script collection helps reduce the possibility that a syntax or framework problem will interrupt an unattended maintenance cycle.

---

## `Test-MaintenanceTelemetry.ps1`

Tests maintenance telemetry generation and processing.

This utility is useful when validating the logging pipeline without needing to wait for an actual scheduled maintenance operation.

It helps verify that expected telemetry can be generated before connecting that data to centralized monitoring and dashboards.

---

# 📦 Deployment Management

## `DeploymentManifest.json`

### Deployment Manifest

Defines information about the files that make up the expected maintenance deployment.

The manifest provides a structured source that can be used to determine what files belong on an endpoint and assist with deployment validation.

---

## `Update-DeploymentManifest.ps1`

### Manifest Generator / Updater

Synchronizes the deployment manifest with the approved maintenance-file catalog.

The utility:

- Adds approved files that are present on the deployment share but missing from the manifest.
- Includes Scripts 17 and 18 in the approved deployment set.
- Removes retired Script 19/20 names only after their Script 17/18 replacements are ready.
- Extracts embedded file versions.
- Recalculates SHA-256 hashes, including hash-only changes where the version number did not change.
- Detects duplicate manifest entries.
- Creates a backup before writing.
- Validates the generated JSON before replacing the active manifest.
- Supports `-WhatIfOnly` for a read-only preview.

This reduces manual deployment-metadata maintenance and helps ensure the manifest accurately represents the files currently approved for endpoint deployment.

---

## `SHA256SUMS.txt`

### File Integrity Information

Contains SHA-256 hashes for deployment files.

Cryptographic hashes provide a method of verifying that a deployed file matches the expected version and has not been accidentally modified or corrupted.

Example concept:

```text
Repository File
      │
      ▼
Calculate SHA-256
      │
      ▼
Compare Expected Hash
      │
      ├── MATCH ──► File Valid
      │
      └── FAIL ───► Investigate / Replace
```

---

## `BUILD-VALIDATION.json`

### Build Validation Results

Stores structured information associated with validation of the maintenance script build.

This provides a machine-readable record that can be used as part of deployment validation and testing.

---

## `README-Maintenance-Framework.txt`

Contains additional technical documentation specifically related to the maintenance framework.

This file is useful when troubleshooting or modifying the underlying framework rather than simply operating the individual maintenance scripts.

---

# 📁 Repository Layout

```text
Compton_Lab_Scripts/
│
├── 00_Update-Scripts-FromShare.ps1
├── 01_Enable_Windows_Update_Services.ps1
├── 02_Remove_User_Profiles.ps1
├── 03_Weekend_Apps_Update.ps1
├── 04_Update_Edge_Silent.ps1
├── 05_Weekend_HP_Drivers_Update.ps1
├── 06_Weekend_Windows_Updates.ps1
├── 07_Force_Reboot_Install_Updates.ps1
├── 08_System_Repair.ps1
├── 09_Disable_Windows_Update_Services.ps1
├── 10_Sync_System_Time.ps1
├── 11_Install_SharpDriver_And_PaperCut.ps1
├── 12_Enable-SystemRestore-And-Create-RestorePoint.ps1
├── 13_Configure_Autologon_And_Edge.ps1
├── 14_Endpoint_Health_Inventory.ps1
├── 15_Install_Elastic_Agent.ps1
├── 16_Check_Deep_Freeze_Status.ps1
├── 17_Set_Browser_Homepage.ps1
├── 18_Install_Honorlock_Chrome_Extension.ps1
│
├── Maintenance.Framework.psm1
├── Maintenance.Policy.json
├── Invoke-MaintenanceScript.ps1
├── Get-MaintenanceFleetStatus.ps1
├── Register-Tasks_SYSTEM.ps1
│
├── Test-AllMaintenanceScripts.ps1
├── Test-MaintenanceTelemetry.ps1
│
├── DeploymentManifest.json
├── Update-DeploymentManifest.ps1
├── BUILD-VALIDATION.json
├── SHA256SUMS.txt
│
├── README-Maintenance-Framework.txt
└── README.md
```

---

# 📝 Logging and Telemetry

Maintenance operations are designed to generate logs rather than silently executing.

This is important because unattended automation without visibility can make troubleshooting more difficult.

The logging architecture is intended to provide information such as:

```text
Computer
   │
   ├── Script Executed
   ├── Start Time
   ├── Completion Time
   ├── Result
   ├── Warnings
   ├── Errors
   └── Maintenance Details
            │
            ▼
         Local Logs
            │
            ▼
        Elastic Agent
            │
            ▼
       Elastic Stack
            │
            ▼
     Kibana Dashboards
```

Centralized telemetry makes it possible to eventually answer questions such as:

- Which computers failed maintenance?
- Which machines have not reported recently?
- Which computers are low on disk space?
- Which systems require a reboot?
- Which computers have Windows Update failures?
- Which endpoints have hardware or Device Manager problems?
- Which systems have unhealthy security services?
- Which systems have incorrect Deep Freeze states?
- Which maintenance scripts are failing most frequently?
- Which lab is experiencing the most maintenance problems?

---

# 🔐 Security Considerations

Many scripts in this repository perform privileged system operations and are intended to execute as **SYSTEM**.

Before deploying these scripts in another environment:

1. Review all configuration values.
2. Review network paths and deployment sources.
3. Review lab/computer naming rules.
4. Remove organization-specific credentials or secrets.
5. Restrict modification rights to deployment locations.
6. Validate SHA-256 hashes before deployment where appropriate.
7. Test changes on a limited group of computers before broad deployment.
8. Protect log data that may contain system or configuration information.

> **Never assume these scripts are appropriate for another environment without reviewing and testing them first.**

---

# ⚠️ Environment Specificity

This project was developed for a specific Windows computer-lab environment.

Several scripts may contain assumptions regarding:

- Computer naming conventions
- Active Directory/domain configuration
- Network paths
- Printer configuration
- Installed applications
- HP hardware
- Deep Freeze
- Elastic infrastructure
- Scheduled maintenance windows

Anyone adapting the project for another environment should review these dependencies before deployment.

---

# 🚧 Project Status

This project is under active development.

The maintenance framework continues to evolve toward increased:

- Centralized logging
- Deployment validation
- Endpoint health visibility
- Automated failure detection
- Fleet-level reporting
- Elastic/Kibana integration
- Proactive alerting
- Maintenance dashboards

The long-term objective is to provide a centralized view of workstation health and maintenance status while reducing repetitive hands-on support work.

---

# 🛠️ Technologies

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-blue?logo=windows)
![Elastic](https://img.shields.io/badge/Elastic-Logging-005571?logo=elastic)
![Kibana](https://img.shields.io/badge/Kibana-Dashboards-005571?logo=kibana)
![License](https://img.shields.io/badge/Status-Active%20Development-orange)

Primary technologies include:

- PowerShell
- Windows Task Scheduler
- Windows Update
- Windows Event Logs
- Elastic Agent
- Elasticsearch
- Kibana
- JSON configuration
- SHA-256 integrity validation

---

# 👤 Author

**Daniel Swaney**

PowerShell automation, endpoint maintenance, infrastructure monitoring, and security-focused systems administration.

---

## 📌 Disclaimer

These scripts are provided as examples of automation developed for a specific computer lab environment.

Administrative scripts can make significant changes to Windows systems. Review and test all scripts before using them in a production environment.

Use at your own risk.
