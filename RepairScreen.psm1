#requires -Version 5.1
# ScriptName: RepairScreen.psm1
# ScriptVersion: 1.5.3
# LastUpdated: 2026-10-30
# Changes: Adds manifest-readable version metadata; display behavior unchanged.
# Compton repair display helpers. Importing this module makes no changes.
Set-StrictMode -Version 2.0

function Initialize-RepairNative {
    if ('ComptonRepairScreen.NativeV11' -as [type]) {return}
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ComptonRepairScreen {
 public static class NativeV11 {
  [DllImport("kernel32.dll", SetLastError=true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool GetProductInfo(uint major, uint minor, uint spMajor, uint spMinor, out uint sku);
  [StructLayout(LayoutKind.Sequential)]
  struct Session {public int Id; public IntPtr Name; public int State;}
  [DllImport("wtsapi32.dll", EntryPoint="WTSEnumerateSessionsW", SetLastError=true)]
  static extern bool Enumerate(IntPtr server, int reserved, int version, out IntPtr buffer, out int count);
  [DllImport("wtsapi32.dll")] static extern void WTSFreeMemory(IntPtr memory);
  [DllImport("wtsapi32.dll", EntryPoint="WTSSendMessageW", CharSet=CharSet.Unicode, SetLastError=true)]
  static extern bool Send(IntPtr server,int session,string title,int titleBytes,string text,int textBytes,int style,int timeout,out int response,bool wait);
  public static int Notify(string title,string text) {
   IntPtr buffer=IntPtr.Zero; int count=0; int sent=0;
   if (!Enumerate(IntPtr.Zero,0,1,out buffer,out count)) return -1;
   try {
    int size=Marshal.SizeOf(typeof(Session));
    for(int i=0;i<count;i++) {
     Session session=(Session)Marshal.PtrToStructure(IntPtr.Add(buffer,i*size),typeof(Session));
     if(session.Id<=0 || session.State!=0) continue;
     int response;
     if(Send(IntPtr.Zero,session.Id,title,title.Length*2,text,text.Length*2,0x40,120,out response,false))sent++;
    }
   } finally {WTSFreeMemory(buffer);}
   return sent;
  }
 }
}
'@ -ErrorAction Stop | Out-Null
}

function Select-ComptonRepairMethod {
    param([int]$ProductSku,[int]$ProductType=1,[bool]$Confident=$true)
    $supported=@(4,27,70,72,84,121,122,125,126,129,130,140,141,175,188,191)
    if ($Confident -and $ProductType -eq 1 -and $ProductSku -in $supported) {return 'LockImageAndNotice'}
    return 'NoticeOnly'
}

function Get-ComptonRepairEdition {
    $caption='Unknown';$cimSku=0;$productType=1;$osVersion='10.0';$editionId='Unknown'
    try {
        $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $caption=[string]$os.Caption;$cimSku=[int]$os.OperatingSystemSKU;$productType=[int]$os.ProductType;$osVersion=[string]$os.Version
    } catch { }
    try {$editionId=[string](Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name EditionID -ErrorAction Stop).EditionID} catch { }
    $sku=$cimSku;$source='Win32_OperatingSystem';$confident=($cimSku -gt 0)
    try {
        Initialize-RepairNative
        [uint32]$nativeSku=0;$version=[version]$osVersion
        if ([ComptonRepairScreen.NativeV11]::GetProductInfo([uint32]$version.Major,[uint32]$version.Minor,0,0,[ref]$nativeSku)) {
            $sku=[int]$nativeSku;$source='Windows GetProductInfo';$confident=($sku -gt 0)
        }
    } catch { }
    if ($source -ne 'Windows GetProductInfo' -and $editionId -match '^Professional' -and $cimSku -in @(4,27,70,72,84,121,122,125,126,129,130,140,141,175,188,191)) {$confident=$false}
    [pscustomobject]@{Caption=$caption;SelectedSku=$sku;CimSku=$cimSku;EditionID=$editionId;ProductType=$productType;Source=$source;Confident=$confident;Method=(Select-ComptonRepairMethod -ProductSku $sku -ProductType $productType -Confident $confident)}
}

function Get-RepairRegistryValue {
    param([string]$Path,[string]$Name)
    if (-not (Test-Path -LiteralPath $Path -ErrorAction Stop)) {return [pscustomobject]@{Exists=$false;Kind=$null;Data=$null}}
    $key=Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($Name -notin @($key.GetValueNames())) {return [pscustomobject]@{Exists=$false;Kind=$null;Data=$null}}
    [pscustomobject]@{Exists=$true;Kind=$key.GetValueKind($Name).ToString();Data=$key.GetValue($Name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)}
}

function Set-RepairRegistryValue {
    param([string]$Path,[string]$Name,[string]$Kind,$Data)
    if (-not (Test-Path -LiteralPath $Path)) {New-Item -Path $Path -Force -ErrorAction Stop | Out-Null}
    $value=$Data
    switch ($Kind) {
        'DWord' {$value=[int]$Data}
        'QWord' {$value=[long]$Data}
        'Binary' {$value=[byte[]]$Data}
        'MultiString' {$value=[string[]]$Data}
        default {$value=[string]$Data}
    }
    New-ItemProperty -LiteralPath $Path -Name $Name -PropertyType $Kind -Value $value -Force -ErrorAction Stop | Out-Null
}

function Remove-RepairRegistryValue {
    param([string]$Path,[string]$Name)
    if (Test-Path -LiteralPath $Path) {
        $key=Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($Name -in @($key.GetValueNames())) {Remove-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop}
    }
}

function Test-RepairValueEqual {
    param($Left,$Right)
    if ([bool]$Left.Exists -ne [bool]$Right.Exists) {return $false}
    if (-not $Left.Exists) {return $true}
    if ([string]$Left.Kind -ne [string]$Right.Kind) {return $false}
    return ((ConvertTo-Json -InputObject $Left.Data -Compress -Depth 5) -ceq (ConvertTo-Json -InputObject $Right.Data -Compress -Depth 5))
}

function Save-RepairScreenState {
    param([string]$Path,$Value)
    $temp=$Path+'.'+[guid]::NewGuid().ToString()+'.tmp'
    $text=ConvertTo-Json -InputObject $Value -Depth 12
    try {
        $stream=New-Object IO.FileStream($temp,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try {$bytes=[Text.Encoding]::UTF8.GetBytes($text);$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)} finally {$stream.Dispose()}
        if (Test-Path -LiteralPath $Path) {[IO.File]::Replace($temp,$Path,[System.Management.Automation.Language.NullString]::Value)}
        if (-not (Test-Path -LiteralPath $Path)) {[IO.File]::Move($temp,$Path)}
        if ([IO.File]::ReadAllText($Path) -cne $text) {throw 'Repair screen backup verification failed.'}
    } finally {Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue}
}

function New-RepairValuePlan {
    param([string]$Path,[string]$Name,[string]$Kind,$Data)
    [pscustomobject]@{Path=$Path;Name=$Name;Original=(Get-RepairRegistryValue -Path $Path -Name $Name);Applied=[pscustomobject]@{Exists=$true;Kind=$Kind;Data=$Data}}
}

function New-RepairBackground {
    param([string]$ImagePath,[string]$Title,[string]$Message)
    # Public image ONLY. Credentials and backups stay in the protected/persistent folders.
    $folder=Split-Path -Parent $ImagePath
    $current=$folder
    while ($current) {
        if ((Test-Path -LiteralPath $current) -and ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {throw 'Repair image path contains a reparse point.'}
        $parent=Split-Path -Parent $current
        if ($parent -eq $current) {break};$current=$parent
    }
    New-Item -Path $folder -ItemType Directory -Force | Out-Null
    $acl=New-Object Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true,$false)
    foreach ($sidText in @('S-1-5-18','S-1-5-32-544')) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier($sidText)),'FullControl','ContainerInherit,ObjectInherit','None','Allow')))
    }
    foreach ($sidText in @('S-1-5-11','S-1-15-2-1')) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier($sidText)),'ReadAndExecute','ContainerInherit,ObjectInherit','None','Allow')))
    }
    Set-Acl -LiteralPath $folder -AclObject $acl
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $bitmap=New-Object Drawing.Bitmap(1920,1080)
    $graphics=[Drawing.Graphics]::FromImage($bitmap)
    $titleFont=New-Object Drawing.Font('Segoe UI',80,[Drawing.FontStyle]::Bold)
    $bodyFont=New-Object Drawing.Font('Segoe UI',44,[Drawing.FontStyle]::Regular)
    $footerFont=New-Object Drawing.Font('Segoe UI',26,[Drawing.FontStyle]::Regular)
    $brush=New-Object Drawing.SolidBrush([Drawing.Color]::White)
    try {
        $graphics.Clear([Drawing.Color]::FromArgb(12,29,52))
        $graphics.TextRenderingHint=[Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $graphics.DrawString('COMPTON COLLEGE IT',$footerFont,$brush,[single]180,[single]220)
        $graphics.DrawString($Title,$titleFont,$brush,(New-Object Drawing.RectangleF(180,330,1560,170)))
        $graphics.DrawString($Message,$bodyFont,$brush,(New-Object Drawing.RectangleF(180,525,1500,310)))
        $graphics.DrawString('Temporary repair mode',$footerFont,$brush,[single]180,[single]900)
        $bitmap.Save($ImagePath,[Drawing.Imaging.ImageFormat]::Jpeg)
    } finally {$brush.Dispose();$footerFont.Dispose();$bodyFont.Dispose();$titleFont.Dispose();$graphics.Dispose();$bitmap.Dispose()}
}

function Get-RepairSnapshotPath {
    param([string]$IncidentFolder)
    Join-Path $IncidentFolder 'RepairScreenState.json'
}

function Get-ValidatedRepairSnapshot {
    param([string]$IncidentFolder,[string]$IncidentId,[string]$ComputerName)
    $path=Get-RepairSnapshotPath $IncidentFolder
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {return $null}
    $state=Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($state.SchemaVersion -ne 1 -or $state.IncidentId -ne $IncidentId -or $state.ComputerName -ne $ComputerName -or
        $state.Phase -notin @('Saved','Applying','Applied','Restored')) {throw 'Repair screen backup is invalid or belongs to another incident.'}
    $permitted=@(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|legalnoticecaption',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|legalnoticetext',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|DisableAcrylicBackgroundOnLogon',
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization|LockScreenImage',
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization|NoChangingLockScreen'
    )
    if (@($state.Values).Count -lt 2 -or @($state.Values).Count -gt 5) {throw 'Repair display backup has an invalid value count.'}
    $seen=@{}
    foreach ($value in @($state.Values)) {
        $key=$value.Path+'|'+$value.Name
        if ($seen.ContainsKey($key)) {throw 'Duplicate target in repair display backup.'}
        $seen[$key]=$true
        if (($value.Path+'|'+$value.Name) -notin $permitted) {throw 'Unrecognized registry target in repair screen backup.'}
    }
    return $state
}

function Join-ComptonRepairNotice {
    param([string]$Message,[string]$OriginalCaption,[string]$OriginalText)
    if ([string]::IsNullOrWhiteSpace($OriginalText)) {return $Message}
    return $Message+"`r`n`r`n--- Existing organization sign-in notice ---`r`n"+$OriginalCaption+"`r`n"+$OriginalText
}

function Enable-ComptonRepairScreenCore {
    param([string]$IncidentFolder,[string]$IncidentId,[string]$ComputerName,[string]$Title,[string]$Message)
    $path=Get-RepairSnapshotPath $IncidentFolder
    $state=Get-ValidatedRepairSnapshot -IncidentFolder $IncidentFolder -IncidentId $IncidentId -ComputerName $ComputerName
    if ($null -ne $state -and $state.Phase -in @('Applied','Restored')) {
        $matches=@($state.Values | Where-Object {-not (Test-RepairValueEqual (Get-RepairRegistryValue $_.Path $_.Name) $_.Applied)}).Count -eq 0
        return [pscustomobject]@{Method=$state.Method;Edition=$state.Edition;Phase=$state.Phase;AlreadyConfigured=$true;CurrentSettingsMatch=$matches;Snapshot=$path;Note='Policies are not repeatedly forced against domain management.'}
    }
    if ($null -eq $state) {
        New-Item -Path $IncidentFolder -ItemType Directory -Force | Out-Null
        $edition=Get-ComptonRepairEdition
        $method=$edition.Method;$image=$null;$imageFailure=$null
        if ($method -eq 'LockImageAndNotice') {
            $image=Join-Path $env:ProgramData ('Compton\RepairScreenAssets\Repair-'+$IncidentId+'.jpg')
            try {New-RepairBackground -ImagePath $image -Title $Title -Message $Message}
            catch {$imageFailure='Image creation unavailable; using the sign-in notice.';$method='NoticeOnly';$image=$null}
        }
        $system='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        $priorCaption=Get-RepairRegistryValue -Path $system -Name 'legalnoticecaption'
        $priorText=Get-RepairRegistryValue -Path $system -Name 'legalnoticetext'
        $text=Join-ComptonRepairNotice -Message $Message -OriginalCaption ([string]$priorCaption.Data) -OriginalText ([string]$priorText.Data)
        $values=@(
            New-RepairValuePlan -Path $system -Name 'legalnoticecaption' -Kind 'String' -Data $Title
            New-RepairValuePlan -Path $system -Name 'legalnoticetext' -Kind 'String' -Data $text
        )
        if ($method -eq 'LockImageAndNotice') {
            $personalization='HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization'
            $values += New-RepairValuePlan -Path $personalization -Name 'LockScreenImage' -Kind 'String' -Data $image
            $values += New-RepairValuePlan -Path $personalization -Name 'NoChangingLockScreen' -Kind 'DWord' -Data 1
            $values += New-RepairValuePlan -Path $system -Name 'DisableAcrylicBackgroundOnLogon' -Kind 'DWord' -Data 1
        }
        $state=[pscustomobject]@{SchemaVersion=1;IncidentId=$IncidentId;ComputerName=$ComputerName;Phase='Saved';Method=$method;Edition=$edition;Title=$Title;RepairMessage=$Message;Values=$values;Image=$image;ImageFailure=$imageFailure;NotifyCount=0;Restoration=@()}
        # Original existence, registry types and values are durable BEFORE changes.
        Save-RepairScreenState -Path $path -Value $state
    }
    $state.Phase='Applying';Save-RepairScreenState -Path $path -Value $state
    foreach ($value in @($state.Values)) {
        $current=Get-RepairRegistryValue -Path $value.Path -Name $value.Name
        if (-not (Test-RepairValueEqual $current $value.Original) -and -not (Test-RepairValueEqual $current $value.Applied)) {
            throw 'An administrator/domain policy changed a display setting during repair. Automatic overwrite refused.'
        }
        Set-RepairRegistryValue -Path $value.Path -Name $value.Name -Kind $value.Applied.Kind -Data $value.Applied.Data
        if (-not (Test-RepairValueEqual (Get-RepairRegistryValue $value.Path $value.Name) $value.Applied)) {throw 'Repair display setting did not verify.'}
    }
    $state.Phase='Applied'
    try {Initialize-RepairNative;$state.NotifyCount=[ComptonRepairScreen.NativeV11]::Notify($Title,$Message)} catch {$state.NotifyCount=-1}
    Save-RepairScreenState -Path $path -Value $state
    [pscustomobject]@{Method=$state.Method;Edition=$state.Edition;Phase=$state.Phase;CurrentSettingsMatch=$true;UserNotificationsRequested=$state.NotifyCount;Snapshot=$path;Image=$state.Image;ImageFailure=$state.ImageFailure;Note='Sign-in notice and image appear when Windows next refreshes the lock/sign-in UI. Neither blocks sign-in.'}
}

function Restore-ComptonRepairScreenCore {
    param([string]$IncidentFolder,[string]$IncidentId,[string]$ComputerName)
    $path=Get-RepairSnapshotPath $IncidentFolder
    $state=Get-ValidatedRepairSnapshot -IncidentFolder $IncidentFolder -IncidentId $IncidentId -ComputerName $ComputerName
    if ($null -eq $state) {return [pscustomobject]@{Restored=$true;Method='NoSnapshot';Changes=@()}}
    if ($state.Phase -eq 'Restored') {return [pscustomobject]@{Restored=$true;Method=$state.Method;Changes=$state.Restoration}}
    $results=@()
    foreach ($value in @($state.Values)) {
        $current=Get-RepairRegistryValue -Path $value.Path -Name $value.Name
        if (Test-RepairValueEqual $current $value.Original) {$results += $value.Name+': already original';continue}
        if (Test-RepairValueEqual $current $value.Applied) {
            if ($value.Original.Exists) {Set-RepairRegistryValue -Path $value.Path -Name $value.Name -Kind $value.Original.Kind -Data $value.Original.Data}
            if (-not $value.Original.Exists) {Remove-RepairRegistryValue -Path $value.Path -Name $value.Name}
            if (-not (Test-RepairValueEqual (Get-RepairRegistryValue $value.Path $value.Name) $value.Original)) {throw 'Original display setting restoration did not verify.'}
            $results += $value.Name+': restored'
            continue
        }
        # Preserve a newer external policy, provided it is not still our image/notice.
        $data=[string]$current.Data
        $containsRepair=$data.IndexOf($IncidentId,[StringComparison]::OrdinalIgnoreCase) -ge 0
        if ($value.Name -in @('legalnoticecaption','legalnoticetext')) {
            if ($state.Title -and $data.IndexOf([string]$state.Title,[StringComparison]::OrdinalIgnoreCase) -ge 0) {$containsRepair=$true}
            if ($state.RepairMessage -and $data.IndexOf([string]$state.RepairMessage,[StringComparison]::OrdinalIgnoreCase) -ge 0) {$containsRepair=$true}
        }
        if ($containsRepair) {
            throw 'A repair display value was changed externally and still contains repair content; manual restoration is required.'
        }
        $results += $value.Name+': newer external setting preserved'
    }
    $state.Restoration=$results;$state.Phase='Restored'
    Save-RepairScreenState -Path $path -Value $state
    # The generated picture is harmless and is retained; deleting a still-cached
    # logon resource immediately can cause a blank display. No original image is changed.
    [pscustomobject]@{Restored=$true;Method=$state.Method;Changes=$results;Snapshot=$path;Note='Registry restoration verified; Windows UI refresh timing is controlled by Windows.'}
}

# A single display mutex protects backup, application, and restoration across
# the watchdog, preview helper and dedicated startup notice process.
function Enable-ComptonRepairScreen {
    param([string]$IncidentFolder,[string]$IncidentId,[string]$ComputerName,[string]$Title,[string]$Message)
    $mutex=New-Object Threading.Mutex($false,'Global\Compton_RepairDisplay')
    $held=$false
    try {
        try {$held=$mutex.WaitOne(30000)} catch [Threading.AbandonedMutexException] {$held=$true}
        if (-not $held) {throw 'Repair display is busy; no settings were changed by this caller.'}
        Enable-ComptonRepairScreenCore @PSBoundParameters
    }
    finally {if ($held) {$mutex.ReleaseMutex()};$mutex.Dispose()}
}

function Restore-ComptonRepairScreen {
    param([string]$IncidentFolder,[string]$IncidentId,[string]$ComputerName)
    $mutex=New-Object Threading.Mutex($false,'Global\Compton_RepairDisplay')
    $held=$false
    try {
        try {$held=$mutex.WaitOne(30000)} catch [Threading.AbandonedMutexException] {$held=$true}
        if (-not $held) {throw 'Repair display restoration is busy; readiness is blocked.'}
        Restore-ComptonRepairScreenCore @PSBoundParameters
    }
    finally {if ($held) {$mutex.ReleaseMutex()};$mutex.Dispose()}
}

Export-ModuleMember -Function @('Get-ComptonRepairEdition','Select-ComptonRepairMethod','Enable-ComptonRepairScreen','Restore-ComptonRepairScreen','Join-ComptonRepairNotice')
