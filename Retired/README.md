# Retired Maintenance Scripts

These scripts are retained for historical reference only. They are not part of
the active Compton College lab-maintenance package and must not be copied to
`C:\Scripts`, placed in the active deployment-share folder, added to
`DeploymentManifest.json`, or registered as scheduled tasks.

The former Autologon password and Elastic Agent enrollment token are blank in
these public archive copies. Operational credentials belong only in the
access-controlled deployment-share copy and must never be committed to GitHub.

| Retired script | Current replacement |
|---|---|
| `04_Update_Edge_Silent.ps1` | Microsoft Edge application servicing and health visibility from script 14 |
| `11_Install_SharpDriver_And_PaperCut.ps1` | Printer and PaperCut section of `04_Sunday_Lab_Application_Maintenance.ps1` |
| `13_Configure_Autologon_And_Edge.ps1` | Autologon and Edge section of script 04 |
| `15_Install_Elastic_Agent.ps1` | Elastic Agent section of script 04 |
| `17_Set_Browser_Homepage.ps1` | Browser homepage section of script 04 |
| `18_Install_Honorlock_Chrome_Extension.ps1` | Honorlock section of script 04 |
| `19_Stellarium_Location_Services.ps1` | Stellarium Location Services section of script 04 |

Because Git retains earlier revisions, deleting or sanitizing a current file is
not a substitute for rotating any credential that was previously committed.
