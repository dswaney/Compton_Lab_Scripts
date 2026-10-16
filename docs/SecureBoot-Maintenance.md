# Compton Secure Boot maintenance pilot — source pilot package 3.0.19

Prepared from the five files supplied October 16, 2026. This is a pilot release, not a campus-wide deployment recommendation. No code has been run against a Windows computer during preparation.

## Source pilot files changed

- 01_Enable_Windows_Update_Services.ps1: 2.8.8. Calls the Script 15 coordinator after task reconciliation. An update reboot request returns 3010 instead of reporting no reboot needed.
- 15_Sunday_Update_Closeout.ps1: 1.1.2. Adds certificate status checks, gated remediation, startup continuation, protection verification and standard NDJSON telemetry. Ordinary service closeout is preserved even if certificate verification is blocked.
- Maintenance.Policy.json: 1.3.0. Adds exact-host pilot prerequisites. All three computer lists are empty initially, so no computer will start certificate remediation by default.
- DeploymentManifest.json: 3.0.19. Updates only the versions/hashes of the three changed managed files; retains other entries.
- Test-SecureBootOffline.ps1: offline development tests. Do not copy this test script into the production share or run it on Windows. It mocks Windows APIs and changes process environment variables in its disposable test process.

Script 07 was reviewed and remains unchanged. The Secure Boot coordinator handles its own bounded Sunday reboot sequence because its reboot reason is distinct from Windows servicing flags. The existing Sunday tasks remain unchanged.

## First step: read-only audit

Extract the package into C:\Temp\Compton-SecureBoot-Pilot. In an elevated, native 64-bit Windows PowerShell window:

```powershell
& 'C:\Temp\Compton-SecureBoot-Pilot\15_Sunday_Update_Closeout.ps1' `
    -ScriptsRoot 'C:\Temp\Compton-SecureBoot-Pilot' -SecureBootAuditOnly
```

This writes a report to C:\Logs\SecureBoot-Maintenance.latest.json and appends a maintenance.secure_boot event to C:\Logs\Maintenance-Telemetry.ndjson. Audit mode does not modify Secure Boot flags, BitLocker protection, scheduled tasks, or freeze state. It can run any day while frozen or thawed. An Updated system with no servicing error and AvailableUpdates 0 or 0x4000 needs no flag rewrite. Capable is informational only.

The existing Elastic filestream can collect the NDJSON event. Verify that an event with maintenance_event.EventType = maintenance.secure_boot is arriving. No new Elastic pipeline or dashboard has been installed by this package. Dynamic mapping limits may affect searchable new fields; the JSON report is also local evidence.

## Before enrolling a remediation pilot

For an encrypted Deep Freeze pilot, upgrade to a supported Deep Freeze release at least 10.20.220.5844. Faronics documents fix 51348 for BitLocker recovery after a Windows Update Workstation Task. This version gate is a local precaution, not a claim that Faronics guarantees every recovery cause is fixed.

Confirm a usable recovery password is safely available. Do not paste it into policy or telemetry. Confirm the computer will remain THAWED through the entire transaction and verification boots, even if an update fails. A live DFC state check proves current state, not future thaw duration.

The scripts do not override the external Deep Freeze schedule and cannot hold a failed computer thawed. Arrange that separately before adding its name to policy. A failed ReadyToFreeze check alone will not stop the external freeze schedule.

After the read-only audit, prepare a controlled Sunday pilot by deploying the two changed scripts, policy and manifest together. Back up the existing share files first. Preserve any newer/local policy changes before applying this package; the supplied policy contains existing FleetStatusRoots placeholder values unchanged.

On the staged policy, add the pilot's exact hostname to:

- SecureBootMaintenance.PilotComputers
- SecureBootMaintenance.RecoveryKeyVerifiedComputers (encrypted pilots)
- SecureBootMaintenance.ThawedThroughVerificationComputers (Deep Freeze pilots)

These are operator confirmations, not automated key escrow or freeze-schedule verification. No wildcard enrollment is supported. No pilot is pre-enrolled, including SSC-216-007, which already reported Updated.

If editing policy after packaging, recompute its DeploymentManifest SHA256 to match the staged policy. Keep policy and manifest coherent before deployment. Do not deploy a modified manifest blindly over a newer share package.

## Sunday behavior

1. Script 01 runs the coordinator; Script 15 passes also recheck it.
2. Already completed systems are reported without repeating the update. BitLocker suspension owned by another process is left unchanged and reported as not ready.
3. New updates require local Sunday, live thaw verification, policy enrollment, no current servicing error, an enabled Microsoft Secure-Boot-Update task, and a compatible encrypted Deep Freeze version. Missing/unknown state fails closed.
4. New transactions may start before 18:00. Reboot requests stop before 20:15. A transaction records computer UUID, hostname and OS volume identity in an administrator/SYSTEM-only ProgramData directory.
5. On encrypted pilots, the coordinator records ownership and suspends protection with RebootCount 2. It does not decrypt the disk. It sets Microsoft's AvailableUpdates 0x5944 only when beginning a deployment that is not already InProgress; subsequent passes do not reset the flags. Existing unrelated flags and Group Policy/Intune-controlled deployment are deferred for review.
6. The Microsoft task is started and polled for up to 45 seconds. Reboots are requested only at the documented 0x4100 stage or for a completed deployment's protected verification boot. Other in-progress states wait for another pass. Each transaction has a maximum of three reboot requests; protection is not repeatedly suspended to extend the budget.
7. The temporary SYSTEM task “Compton Secure Boot Update Resume” runs two minutes after startup and at 20:15. It does not start a new transaction without state. The 20:15 run is a cleanup deadline, not a certificate-update opportunity. Sleep/offline devices cannot be forced to execute cleanup; the two-reboot suspension bound remains relevant.
8. After deployment is complete, owned suspension is resumed and protection On is verified. A subsequent thawed boot is required; recovery/mismatch events or unreadable System events fail verification. A successful verified transaction is archived as last-completed.json and its temporary task removed.
9. Failure, expiry or deadline triggers an attempt to resume owned protection and a failure event. It never reports ReadyToFreeze on failed verification. An incomplete firmware transition may still need manual recovery; restoring protection does not prove that the next boot is safe.
10. Sunday service shutdown remains in place independently of certificate readiness. The external Deep Freeze schedule remains external.

In-progress state: C:\ProgramData\Compton-SecureBootMaintenance\state.json. A failed transaction is not automatically retried from scratch. Inspect its evidence and correct the cause before an administrator archives/removes that state file to permit another attempt. Never delete state while the coordinator owns suspended protection.

## Decryption as a fallback

No automatic decryption is enabled or implemented in this pilot. Update failure alone cannot predict whether the next boot will request recovery. Suspending protection and decrypting a volume are different operations.

A separately approved decryption fallback would require the computer to remain thawed until FullyDecrypted is verified. It must report DecryptionInProgress separately from FullyDecrypted, retain a clear security exception for follow-up, and later explicitly re-enable encryption and confirm recovery-key escrow. This package does not claim a request to decrypt is completion and does not auto-encrypt a previously decrypted disk.

## Verification and limitations

PowerShell 7 parser checks passed for both changed scripts and the offline test. Twenty-three offline tests cover updated/audit behavior, frozen and version blocks, recovery-key prerequisite, servicing errors, normal update/reboot/resume/protected-boot flow, repeated pass idempotency, unrelated suspension, orphan startup, recovery evidence, unreadable events, weekdays, late start, deadline cleanup, transaction expiry and reboot budget exhaustion.

Windows PowerShell 5.1 syntax is used, but native Windows integration is not yet tested. DFC semantics follow the existing Script 15 implementation. A real encrypted, thawed pilot must validate DFC, task execution, firmware servicing, BitLocker handling, telemetry ingestion, and a final frozen boot before expanding rollout. The first pilot should be attended with the recovery key available.

This targets the Secure Boot 2023 certificate/boot manager transition. It cannot eliminate recovery caused by unrelated BIOS/TPM/boot configuration changes, disk faults, or freeze rollback. No TPM clear, PCR validation removal, Secure Boot disabling, hand-installed certificate blob or DBX revocation is performed.

Sources:
- https://support.microsoft.com/en-us/servicing/os/secure-boot/2025/09/registry-key-updates-for-secure-boot-windows-devices-with-it-managed-updates
- https://support.microsoft.com/en-US/servicing/os/secure-boot/2026/03/secure-boot-troubleshooting-guide
- https://www.faronics.com/webhelp/DFE/ReleaseNotes/DFE_ReleaseNotes.htm

## Telemetry sharing correction (1.1.1)

Replaces Add-Content telemetry writes with explicit FileStream append sharing for readers, five bounded attempts, and a uniquely named local pending NDJSON file if an exclusive lock remains. Telemetry failure cannot suppress the audit result or bypass mutex release. TelemetryDelivery in the latest report distinguishes Appended, Pending, and Failed. Pending files require explicit forwarding/replay; this update does not change the Elastic input paths.

## Restricted firmware query reporting (1.1.2)

If Confirm-SecureBootUEFI fails, audit mode retains certificate deployment evidence, the separately labeled Windows registry Secure Boot state, BitLocker protection and Deep Freeze version. Updated_FirmwareQueryRestricted means registry evidence indicates completion/enabled state, but the live firmware query was denied. Ready and Complete remain false. Enabled is null, not false, when the firmware query cannot be made. FirmwareQueryError records the restriction. RegistryDeploymentComplete and RegistrySecureBootEnabled are observations, not substitutes for successful live verification. No privilege grants, security-policy changes or firmware-query bypasses are performed. Both remediation and final readiness still require a successful live query.

## Repository synchronization

The repository package is 3.0.20. Its public paths and hashes differ from production bytes. This guide describes the source pilot package; use the repository manifest for repository files. Script 02 and reporting/registrar/launcher synchronization are documented in the main README.
