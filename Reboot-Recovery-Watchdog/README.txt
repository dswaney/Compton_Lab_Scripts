CURRENT PUBLIC PACKAGE NOTES — 2026-11-02
Runtime Watch-RebootRecovery.ps1: 1.5.4. Read the root README first.
Configured credentials are blank in this public package. Configure only in a
restricted private deployment copy; never commit a DFC password/enrollment token.
AutomaticFreezeAfterVerifiedRepair is TRUE in the supplied runtime: verified
recovery may request one final frozen restart, in addition to the bounded thaw
and optional repair restart. Older text below describing a mandatory thawed end
state or a two-restart total is superseded by this note and the current code.
Review production script hash pins explicitly; public sanitization changes bytes.
Configure-RecoveryShare.py is referenced in supplied notes but was NOT included.
New-RebootRecoveryShare.ps1 is the supplied Windows share setup helper.
The root RepairScreen.psm1 carries newer version metadata than this folder copy.
Offline tests and old validation text do not establish live Windows validation.

--- Supplied historical instructions follow ---

COMPTON REBOOT RECOVERY WATCHDOG v1.5.0
Prepared 2026-10-28. Windows PowerShell 5.1, Administrator/SYSTEM.

This is a separate watchdog. It does not replace Script 07, the task registrar,
Script 00, or your existing Sunday schedules. The password section is near the
TOP of Watch-RebootRecovery.ps1, immediately after param().

WHAT IT DOES
- Diagnostic-only unless explicitly installed with -EnableRecovery or run with -Recover.
- Startup trigger (15-second delay) plus a check every five minutes, as SYSTEM.
- Matches three distinct COMPLETED restarts within 30 minutes: known Script 07
  maintenance shutdown comments, shutdown.exe, and subsequent Kernel-General 12
  boot events. A request alone does not count. The last request before a boot
  must match; intervening user/Deep Freeze requests invalidate that match.
- Excludes Sunday 00:15-21:45 local time and defers recovery transitions about
  to enter the Sunday window. Manual -Force comments and watchdog restarts do
  not feed the detection counter. BSODs/power failures alone do not trigger thaw.
- Creates a durable per-computer incident record BEFORE consuming recovery actions.
- Stops/disables only recognized non-manual Script 07 tasks and associated
  powershell.exe invocations. Records original enabled states. Cancels an active
  countdown only when its latest request is a recognized maintenance comment.
- While frozen, requests THAWNEXTBOOT and one 60-second restart. Verifies the next
  boot really is THAWED. Does not call /BOOTTHAWED in addition to its own restart.
- While thawed, replays pre-reboot evidence, runs the approved updater once, and
  requires the reviewed Script 07/09 hashes before any repair execution.
- Lets the boot settle, then checks pending work. If none remains, runs Script 09
  and verifies background controls without adding an unnecessary reboot.
- If work remains, invokes approved Script 07 -Force ONCE. Script 07 handles its
  existing one-reboot/verification-only continuation. The watchdog never requests
  a second -Force pass for that incident.
- Verifies clean flags and stopped/disabled wuauserv/BITS/UsoSvc; checks the
  UpdateOrchestrator Reboot task, restores prior Script 07 task enabled states,
  and logs ReadyForManualFreeze. It leaves the machine THAWED for technician review.
- One thaw restart plus, only if needed, one Script 07 repair restart: at most TWO
  requested restarts per incident. This cannot limit reboots initiated by Windows,
  Deep Freeze schedules, other administrators, firmware or unrelated software.
- Successful incidents have a 24-hour cooldown. Failed incidents remain latched
  NeedsAttention until a technician fixes and explicitly acknowledges them.
- Missing storage, unknown freeze state, unapproved scripts, interrupted phases,
  failed writes or unavailable telemetry stop automatic actions. Pending reboot
  registry values are never cleared by this watchdog.

FILES
Watch-RebootRecovery.ps1        Complete watchdog; edit configuration at the top.
Test-RebootRecoveryWatchdog.ps1 Offline unit tests. No DFC/service/task/reboot calls.
RepairScreen.psm1               Edition detection, temporary display, verified restore.
Test-RepairScreen.ps1           Offline tests with an in-memory registry backend.
Preview-RepairScreen.ps1        Edition report or reversible visual pilot test.
New-RebootRecoveryShare.ps1     Optional dedicated Windows file-server share setup.
VALIDATION.txt                 Exact validation performed and remaining live checks.

PASSWORD / STORAGE
The requested password field is plaintext in the script. Do not upload/email the
edited script or put it on the broadly readable Labscripts share. Installation
copies it to an Administrators/SYSTEM-only directory under:
C:\ProgramData\Compton\RebootRecovery
DFC requires the password as a command-line argument. The watchdog does not write
that argument to its logs, but privileged process-command-line auditing or an
administrator inspecting processes may still observe it.

Keep the unedited package for sharing. Protect or remove your edited setup copy
once installation is verified. The installer does not delete your original copy.

PersistentStateRoot is also REQUIRED for automatic recovery. It MUST be a real UNC
folder on persistent server storage. A frozen C: state file would reset attempts
and could create another loop. Do not substitute a C: path. Diagnostic mode works
without the share; recovery refuses to run without writable durable state.

QUICK START: FIRST PILOT ON IB1-103-03 WHILE THAWED

1. Extract the ZIP on the endpoint using elevated PowerShell:

Expand-Archive -LiteralPath "$env:USERPROFILE\Downloads\Compton-Reboot-Recovery-v1.5.0.zip" -DestinationPath 'C:\Scripts\RebootRecoverySetup' -Force
Set-Location 'C:\Scripts\RebootRecoverySetup\Reboot-Recovery-Watchdog'

2. Native syntax and offline checks (BEFORE editing or installing):

.\Watch-RebootRecovery.ps1 -ValidateOnly
.\Test-RebootRecoveryWatchdog.ps1
.\Test-RepairScreen.ps1

These do not execute the operational watchdog. The tests write/delete only a
unique test folder under your Windows TEMP directory.
If anything fails, stop there and share the error, not a password-bearing file.

3. Provision persistent storage, if you do not already have a suitable share.
Copy New-RebootRecoveryShare.ps1 to the WINDOWS FILE SERVER (filesvr).
Run this there in elevated PowerShell, not on the lab computer:

.\New-RebootRecoveryShare.ps1 -ComputerNames 'IB1-103-03' -Domain 'COMPTON'

Default new share: \\filesvr\ComptonRebootRecovery$
It prints the actual UNC name based on the server's computer name. Use that output.
The helper creates a dedicated folder/share only, gives Domain Computers share
Change access, and restricts each NTFS computer folder to that computer account,
Administrators and SYSTEM. It refuses unrelated existing locations. Existing
computer folders are retained without silently changing their permissions.
For more endpoints, rerun on the file server with their computer names.
No AD module, new account, package or endpoint software installation is required.
If filesvr is not Windows, provision an equivalent persistent SMB share manually.

4. On the ENDPOINT, edit the configuration at the top:

notepad .\Watch-RebootRecovery.ps1

Set:
$DeepFreezeCommandLinePassword = ''
$PersistentStateRoot = '\\filesvr\ComptonRebootRecovery$'

Use the actual UNC path printed by step 3. If the password contains an apostrophe,
double it inside single quotes: 'abc''def'. Do not put the password in a shell command.
Retain the other defaults for the first pilot, including approved hashes.

5. Install DIAGNOSTIC-ONLY first while THAWED:

.\Watch-RebootRecovery.ps1 -Install
Start-ScheduledTask -TaskName 'Compton. Reboot Loop Recovery Watchdog'

After about 20 seconds:

Get-Content 'C:\Logs\RebootRecovery.latest.json' -Raw | ConvertFrom-Json | Format-List *

DetailsJson should report:
- ExecutionIdentity NT AUTHORITY\SYSTEM
- FreezeState Thawed
- PersistentStorage AvailableInThisExecutionContext
- TelemetryReady true
This scheduled test verifies the COMPUTER ACCOUNT can use the share; an interactive
administrator's successful share test alone does not prove that.
No forced reboots, thaw configuration changes or service/task quarantine occur in audit mode.
Audit mode may write a transient share probe and a diagnostic bundle if a loop is detected.

6. Verify telemetry reaches Elastic before enabling recovery.
In Kibana Discover, use:

maintenance_event.EventType : "endpoint.reboot_recovery"

For a just-installed audit task, look for Status AuditNoLoop or LoopDetected,
ComputerName IB1-103-03, and this version. If no events arrive, inspect your existing
Elastic Agent filestream/NDJSON ingestion rules. Local writing is not proof of delivery.
No direct Elasticsearch credentials or new n8n workflow are required by this package.

7. Enable bounded automatic recovery while still THAWED:

.\Watch-RebootRecovery.ps1 -Install -EnableRecovery

This replaces ONLY the watchdog task definition with the -Recover argument.
It does not immediately force a reboot or change the ordinary maintenance schedule.
The periodic/startup task will act only when the configured loop conditions are met.
Do not deliberately reproduce a classroom reboot loop as a deployment test.

8. Once installed/audited, verify the task and protected installed file remain
present after a normal thawed restart before returning the computer to its normal
frozen configuration. A reboot is not initiated by this instruction/package installer.

ELASTIC LOGGING AND DIAGNOSTICS
Uses the existing framework's Write-MaintenanceTelemetryLine to append to:
C:\Logs\Maintenance-Telemetry.ndjson

Latest local watchdog status:
C:\Logs\RebootRecovery.latest.json

Event types:
endpoint.reboot_recovery
endpoint.reboot_recovery.evidence

Search both:
maintenance_event.EventType : "endpoint.reboot_recovery" or maintenance_event.EventType : "endpoint.reboot_recovery.evidence"

Filter by maintenance_event.ComputerName and maintenance_event.IncidentId.
Detection/failure events use Severity Critical. Recovery transitions are logged.
Unchanged idle/cooldown heartbeats are coalesced to about once per 30 minutes.

On detection, evidence includes:
- Matching restart requests and verified boot history.
- Current Deep Freeze state/DFC file version, boot time and configured Sunday window.
- Current pending flags including Edge/file rename entries; Windows Update COM status.
- Installed 00/01/06/07/09/15/registrar headers, hashes and file timestamps.
- Available latest JSON reports, newest text logs and available last-Sunday text logs.
- Matching last-Sunday execution summaries from the newest 4 MiB / 20000 NDJSON lines.
- Relevant task actions, triggers, enabled settings, SYSTEM identity, last-run results.
- Retained System, WindowsUpdateClient/Operational and TaskScheduler/Operational
  events for last Sunday and the preceding two hours.
- Coverage notes identifying missing logs, unavailable channels and truncation.

Evidence is stored BEFORE thaw/reboot under:
<PersistentStateRoot>\<ComputerName>\<IncidentId>\Evidence.ndjson
<PersistentStateRoot>\<ComputerName>\<IncidentId>\Diagnostics.zip

It is appended locally before the thaw reboot and replayed after verified thaw.
EvidenceFingerprint allows backend deduplication of those at-least-once copies;
this package does not change Elastic pipelines/document IDs or guarantee deduplication.
The ZIP stays on the share; Elasticsearch receives structured evidence/log excerpts,
not an uploaded ZIP attachment. Evidence publication success still needs agent/server
confirmation. The agent must remain functional and monitor the existing NDJSON file.

Limits: 2000 evidence records / approximately 2 MiB; each evidence value about 12000
characters; at most 500 lines per selected text file; 100 newest events per channel
per time window; copied originals capped around 8 MiB and 1 MiB per file. Tail-only
copies and truncation are marked. Archived ZIP logs, complete WindowsUpdate ETLs,
and complete CBS logs are not collected. Missing Sunday data cannot be reconstructed.
No edited watchdog source/password configuration is included in the diagnostic bundle.
Display events include detected Windows edition/SKU, selected method, policy conflicts
and restoration results. Deep Freeze edition labels are not used for display selection.
Retain/rotate incident folders server-side according to your organization's policy;
this version does not delete incident history automatically.

APPROVED SCRIPT HASHES
The updater and 07/09 hashes match the supplied files we reviewed. After legitimate
future changes, an unknown hash stops repair. Review the new scripts' behavior,
then add the approved hashes at the top and reinstall the watchdog while thawed.
Do not simply add an old/unreviewed hash to silence the check.
A pinned starting updater still uses your administrator-controlled production share.

FAILED INCIDENT / MANUAL ACKNOWLEDGMENT
NeedsAttention blocks automatic retries, leaves suspended reboot tasks quarantined,
and requires investigation. The machine is not automatically refrozen on failure.
Inspect the incident ZIP, latest report, and 07/09 reports. After fixing the cause,
with the computer thawed, current approved scripts, no pending reboot flags and
verified background controls, acknowledge using the protected installed copy:

& 'C:\ProgramData\Compton\RebootRecovery\Watch-RebootRecovery.ps1' -AcknowledgeIncident

This verifies readiness, restores prior enabled states of recognized reboot tasks,
and marks the incident complete. It does not request a reboot or automatically freeze.

UNINSTALL (WHILE THAWED)
& 'C:\ProgramData\Compton\RebootRecovery\Watch-RebootRecovery.ps1' -Uninstall

Removes only the watchdog task. Does not delete incident evidence, remove the share,
restore quarantined tasks, or erase the credential-bearing installed file. Review an
active incident before uninstalling and protect/delete that file separately if desired.

LIMITATIONS
Requires functioning EventLog, DFC, SMB access, shared maintenance framework, and
existing Elastic Agent ingestion. No detection from lost/rolled-back event history.
Cannot guarantee Sunday maintenance occurred, cannot reconstruct unavailable events,
and cannot prevent all restarts from unrelated software. Its own requested recovery
is bounded. No automatic refreeze, download/install scan, registry flag cleanup or
firmware/BitLocker changes are performed by the watchdog. Script 00 and the reviewed
07/09 retain their own existing behavior.


NEW IN v1.2.0: REPAIR SCREEN
The password/share configuration remains near the top of Watch-RebootRecovery.ps1.
Also configurable there: EnableRepairScreen, RepairScreenTitle, RepairScreenMessage.
The display module is copied beside the protected installed watchdog. Extract the
WHOLE package; do not copy only the .ps1 file. Re-enter your existing password/share
settings in the new setup copy before reinstalling while thawed. Do not upload it.
The new package retains v1.0.0 incident receipts; no retry budgets are reset.

Detection:
- Windows GetProductInfo is primary, with Win32_OperatingSystem and EditionID evidence.
- It logs the source, chosen SKU, WMI SKU, registry EditionID and uncertainty.
- Confirmed client Enterprise/Education/IoT Enterprise: repair lock/sign-in image
  plus a temporary sign-in notice.
- Pro/Pro Education/unknown or conflicting fallback evidence: sign-in notice only.
- No edition upgrade, activation, Intune/CSP configuration or Shared PC enablement.

Notice text defaults to:
IT REPAIR IN PROGRESS
This computer is temporarily unavailable. Repairs may take several minutes.
Please do not turn it off. For assistance, call ext. 2181.

Application/restoration:
- Applied only after confirmed thaw. No attempt to persist a new registry policy
  written while frozen into the next boot.
- Original value existence, types and data are backed up in the incident's
  RepairScreenState.json BEFORE any sign-in policy is changed.
- Existing organization legal notice text is appended, then restored afterward.
- Eligible editions get a locally generated JPEG under the separate public-image
  folder C:\ProgramData\Compton\RepairScreenAssets. No password is in that folder.
- Active user sessions get a best-effort, dismissible message (120-second timeout)
  through Windows WTS. This avoids trying to show a Session 0 service window.
- Lock/sign-in policy refresh/caching is controlled by Windows; it may require a
  subsequent lock/sign-in. Visual display must be verified on the pilot.
- It is informational, not a sign-in block. It does not disable keyboard/mouse,
  kill Explorer/LogonUI, sign users out, replace boot logos, or replace Windows
  update progress screens.
- Partial display setup is reported; bounded repair can continue. Any saved display
  state must be successfully restored before the watchdog reports ReadyForManualFreeze.
- Restoration removes only originally absent values, restores original types/data,
  and preserves newer external management settings if they no longer contain repair
  content. A conflicting value still containing repair content blocks readiness.
- The generated JPEG is retained for Windows cache stability; original images are
  never overwritten. Empty registry keys created for policies may remain harmlessly.
- Failure keeps repair settings in place for investigation. Complete recovery or the
  existing -AcknowledgeIncident path restores them before readiness.
- Use the WATCHDOG's ReadyForManualFreeze status, not Script 07's earlier report,
  when this feature is active. The watchdog still does not automatically freeze.
- It cannot intercept an unrelated administrator/Deep Freeze freeze command. Do not
  externally freeze before display restoration/readiness is confirmed.

VISUAL PILOT (NO THAW COMMAND OR REBOOT REQUEST)
Keep the pilot THAWED and the watchdog AUDIT-ONLY for this test.
First report the chosen method without changing settings:

.\Preview-RepairScreen.ps1 -Mode Report

After the persistent private computer folder is provisioned, apply a preview:

.\Preview-RepairScreen.ps1 -Mode Apply -PersistentStateRoot '\\filesvr\ComptonRebootRecovery$'

Use YOUR actual persistent UNC root. The helper refuses an active automatic recovery
or a task already configured with -Recover. It preserves a separate preview backup.
Check the next lock/sign-in screen on the pilot using normal Windows controls.
You may see the dismissible message immediately if already signed in. Existing
corporate policy can override image/notice settings; the watchdog records conflicts
rather than repeatedly fighting domain policy.

Restore the preview before enabling recovery or freezing:

.\Preview-RepairScreen.ps1 -Mode Restore -PersistentStateRoot '\\filesvr\ComptonRebootRecovery$'

An active/unrestored preview blocks automatic recovery/enablement. After verifying
normal appearance is restored, proceed with the existing -Install -EnableRecovery step.

The display offline tests mock registry access and use a private TEMP folder; they
never change actual sign-in policies or send a message. Visual/API behavior remains
untested in the authoring environment. v1.2.0 also fixes initialization scope so
framework/display helper commands remain available after their loader returns.

PRIMARY REFERENCES
https://learn.microsoft.com/windows/configuration/background/
https://learn.microsoft.com/windows/win32/api/sysinfoapi/nf-sysinfoapi-getproductinfo
https://github.com/microsoft/win32metadata/blob/main/generation/WinSDK/RecompiledIdlHeaders/um/winnt.h
https://learn.microsoft.com/windows/win32/api/wtsapi32/nf-wtsapi32-wtssendmessagew
https://learn.microsoft.com/previous-versions/windows/it-pro/windows-10/security/threat-protection/security-policy-settings/interactive-logon-message-text-for-users-attempting-to-log-on

MANUAL RECOVERY TEST (-Force), v1.2.0
Run elevated after configuring and installing this same version:
  .\Watch-RebootRecovery.ps1 -Install -EnableRecovery
  & "$env:ProgramData\Compton\RebootRecovery\Watch-RebootRecovery.ps1" -Force

-Force implies -Recover and bypasses ONLY the repeated-reboot threshold. It can
thaw/restart a frozen computer, refresh scripts, close update services, and run
Script 07 -Force once if pending flags require it. Save work before testing.
A clean thawed machine need not reboot. Existing active incidents are resumed;
NeedsAttention and the 24-hour completed-incident cooldown are NOT reset.
Sunday exclusion, share access, script hashes, telemetry, and budgets still apply.
ManualTestStarted is logged separately from LoopDetected. Do not put -Force in
a scheduled task. The recovery task must already be enabled and use -Recover.

REGISTRATION
Install while THAWED, in elevated Windows PowerShell 5.1, with the complete
package extracted together. Set DeepFreezeCommandLinePassword and
PersistentStateRoot in Watch-RebootRecovery.ps1 first. Provision the private
computer folder on that share as described above. Then:
  .\Watch-RebootRecovery.ps1 -Install -EnableRecovery
This registers 'Compton. Reboot Loop Recovery Watchdog' as SYSTEM, at startup
with a 15-second delay and every five minutes. Reinstall replaces that same task.
The installed protected script and module are under:
  C:\ProgramData\Compton\RebootRecovery
For audit-only registration, use -Install without -EnableRecovery.
Verify registration:
  Get-ScheduledTask -TaskName 'Compton. Reboot Loop Recovery Watchdog' | Format-List TaskName,State,Actions,Triggers
Check latest status:
  Get-Content C:\Logs\RebootRecovery.latest.json -Raw

VERSION 1.3.0 CORRECTIONS
- Correct null-string arguments for durable JSON replacement in both files.
- XML event fields read InnerText, including empty elements.
- NoLoopDetected is informational.
- UTC receipt times use DateTimeOffset.UtcDateTime for timeout/boot/cooldown checks.
- Pre-thaw recovery evidence runs in a separate evidence-only process, with a
  120-second parent-enforced timeout and stage progress in latest JSON/telemetry.
  A timeout fails closed to NeedsAttention with no thaw/reboot, retains partial
  evidence and suspended-task settings for review. It does not silently skip
  diagnostics or automatically retry. Collector output is protected under
  C:\ProgramData\Compton\RebootRecovery (EvidenceCollector-<incident>.out/.err).
The specific cause of the pilot collection stall is not established. Progress
will identify the stage if it recurs; a timeout is not a successful repair.
The restored pilot incident remains latched; install does not clear it.
Do not resume the pilot until that incident is explicitly reviewed/archived.

v1.3.1: Script headers use StreamReader plain strings, avoiding Get-Content
provider metadata in Windows PowerShell 5.1 JSON serialization. Pilot confirmed
hash/header reads finish and plain-header JSON takes 1 ms; decorated serialization
was the likely blocked operation. Progress now separates hashing, header reading
and serialization. Scheduled-task definitions are plain XML strings instead of
raw CIM action/trigger objects. The 120-second timeout remains in place.

v1.3.2: native processes are started/owned directly through Diagnostics.Process,
not Start-Process. Redirected streams are drained asynchronously. Explicit exit
code checks refuse unknown codes for DFC, updater/maintenance children and the
evidence collector. A failed parent reloads the child's persisted evidence fields
before latching NeedsAttention. Includes Windows exit-code regression checks.
Pilot v1.3.1 completed evidence collection, but parent reported a blank exit code
and overwrote its evidence references; actual diagnostics remain on the share.

VERSION 1.4.0 CLASSROOM RECOVERY FLOW (supersedes older manual-freeze descriptions)
- Startup trigger has no added delay. Recurring checks run every ONE minute.
- Watchdog boot-settle minimum is ONE minute. The configured three-restarts
  detection threshold and 30-minute evidence window remain unchanged.
- Repair notice is informational. No keyboard/input lock or sign-in restrictions
  are added; users can dismiss the notice. Repair may reboot an occupied session.
- After thaw verification and Script 00 refresh, always invoke the approved
  Script 07 -Force once, even if preliminary pending flags are empty.
- Script 07's own servicing/verification waits are unchanged (its startup check
  currently waits two minutes). No fresh update scan/download/install pass added.
  Flags are verified by Script 07; they are not blindly deleted by the watchdog.
- After Script 07 confirms the repair, recheck pending indicators and service
  closeout, restore original display/task settings, issue FREEZENEXTBOOT, and
  request a final forced restart with a ten-second countdown. A later watchdog
  check verifies the machine actually booted frozen and logs CompletedFrozen.
- Maximum successful flow: one thaw restart + one Script 07 repair restart +
  one final frozen restart. Final-freeze receipt/budget is committed BEFORE DFC.
- On verification/DFC failure, latch NeedsAttention instead of reporting success
  or blindly freezing unresolved work. Such a failure can still leave a machine
  thawed and requires IT intervention. Normal Sunday exclusion still applies.
- Existing in-memory runs retain the old behavior: install only while thawed and
  when the current recovery pass is idle. Do not kill a running Script 07 pass
  or replace/delete its active incident receipt.

READABILITY (v1.4.0)
Custom supported-edition background fonts: title 80 pt, body 44 pt, footer 26 pt.
Built-in Windows sign-in legal notice font is controlled by Windows; this update
shortens the message but does not claim to enlarge that system dialog's text.
On Pro using notice-only mode, the font remains Windows-controlled.

v1.4.1 EARLIER NOTICE
Active recovery receipt is loaded before reboot-history queries. On a thawed
repair boot the display is requested immediately after that receipt is available;
active incidents skip unnecessary redetection queries. Startup delay remains zero.
Network availability and Windows sign-in screen refresh still affect visibility.
No immediate pre-Windows/boot-screen popup is promised. Install only after the
current recovery completes; changing files does not update an in-memory run.

VERSION 1.5.0: DEDICATED STARTUP NOTICE
Install registers two SYSTEM tasks:
  Compton. Reboot Loop Recovery Watchdog (startup + every one minute)
  Compton. Recovery Notice at Startup (startup only, no added delay)
The second task runs Show-RecoveryNotice.ps1 with a protected, non-secret
RecoveryNotice.config.json generated at install. Do not edit that generated
configuration; configure the main watchdog and reinstall instead.
The helper queries DFC state without a password; frozen computers are untouched.
When thawed, it reads the persistent UNC receipt, validates computer/incident
identity and active phase, and calls the shared display module. It does not run
detection, evidence collection, updater, service changes, DFC state commands or
reboots. It retries unavailable state storage for up to 45 seconds and has a
two-minute task execution limit for blocked operations. Main-watchdog display
logic remains a fallback. Network and Windows UI refresh can still delay it.
Display backup/apply/restore share a named mutex across both processes and the
preview helper. A restored incident snapshot is never reapplied.
Boot helper status: C:\Logs\RecoveryNotice.latest.json
Telemetry EventType: endpoint.reboot_recovery.notice (existing NDJSON transport).
Display is requested BEFORE loading the maintenance telemetry framework.
Final closeout now restores maintenance-task settings and obtains DFC acceptance
of FREEZENEXTBOOT BEFORE restoring the original display. The normal ten-second
final reboot countdown remains. There is no guarantee the notice remains visible
until the exact instant Windows shuts down.
Audit-only installation removes the dedicated notice task; uninstall removes
both task registrations and retains incident evidence/configuration for review.
Installing v1.5.0 does not clear a completed incident's 24-hour cooldown.
