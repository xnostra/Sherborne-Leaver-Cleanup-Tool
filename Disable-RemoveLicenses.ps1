<#
.SYNOPSIS
  All-in-one leaver cleanup tool for Microsoft 365.
  Disables accounts, removes licenses, removes group/distribution list/Teams memberships,
  and hides users from the address book - in bulk (from a CSV) or one account at a time.

.MODES
  Interactive menu:   .\Disable-RemoveLicenses.ps1
  Bulk dry run:       .\Disable-RemoveLicenses.ps1 -CsvPath .\leavers.csv -SkipIfActiveWithinDays 60
  Bulk for real:      .\Disable-RemoveLicenses.ps1 -CsvPath .\leavers.csv -SkipIfActiveWithinDays 60 -Commit

.CSV FORMAT
  Email,Name,LeavingDate
  pupil@school.org,John Smith,2019-06-30
  ,Jane Doe,2015-06-30       <- no email: matched by display name instead
  (iSAMS headers 'Pupil Email Address', 'Full Name', 'Leaving Date' also accepted.)
#>

param(
    [string]$CsvPath,
    [switch]$Commit,
    [int]$SkipIfActiveWithinDays = 30,  # failsafe: skip anyone who signed in this recently (0 = off)
    [datetime]$SkipIfCreatedAfter,      # failsafe: skip accounts created after this date
    [string[]]$SafeCountries = @('QA'), # sign-ins from these countries are normal; anything else = ALERT
    [switch]$NoGroupCleanup,            # skip removing group/DL memberships on commit
    [switch]$NoHideFromAddressBook,     # skip hiding from the address book on commit
    [switch]$SkipSecurityCheck,         # skip the per-account sign-in location/failure check (much faster)
    [switch]$RefreshUserCache,          # force a fresh download of all users instead of using the cache
    [switch]$StaffMode,                 # STAFF: convert mailbox to shared BEFORE removing license (email preserved forever, free)
    [string]$DelegateEmail,             # optional: give this person full access to converted shared mailboxes
    [string]$ArchiveSiteUrl,            # optional: SharePoint site to auto-copy leavers' OneDrive files into (e.g. https://tenant.sharepoint.com/sites/StaffArchive)
    [ValidateSet('Students','Staff')]
    [string]$ListType = 'Students',     # what kind of people are in the list (Staff = preserve email + archive OneDrive)
    [switch]$CleanEverything,           # treat every account found in the list as APPROVED - bypass the name-mismatch / recently-active / recycled-address holds
    [switch]$Relaunched,                # internal: set automatically when the script relaunches itself without admin rights
    [switch]$NoSelfUpdate,              # skip the "pull latest from GitHub" check at startup
    [switch]$Updated                    # internal: set automatically after a self-update relaunch (prevents update loops)
)

# Approval mode: when true, every matched account in the list is fully cleaned (holds bypassed).
$script:ForceClean = [bool]$CleanEverything

# ---------- If elevated, relaunch WITHOUT admin rights (Microsoft sign-in fails when elevated) ----------
# Exchange Online's sign-in cannot open when PowerShell runs as administrator. We detect that and
# relaunch the tool as the normal interactive user (medium rights) via a one-off scheduled task.
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin -and -not $Relaunched) {
    Write-Host "Running as administrator - relaunching as your normal user so Microsoft sign-in works..." -ForegroundColor Yellow
    $ok = $false
    $psExe = Join-Path $PSHOME 'powershell.exe'
    $scr   = $PSCommandPath
    $arg   = "-NoProfile -STA -ExecutionPolicy Bypass -File `"$scr`" -Relaunched"
    try {
        # Use the Shell COM object's ShellExecute to launch the process. Even called from an
        # elevated process, this routes through the desktop's own (non-elevated) shell, so the
        # launched process inherits normal/medium integrity - not admin. Far more reliable than a
        # Scheduled Task, which can be silently blocked by policy on managed devices (fails with no
        # error and no window at all, which is exactly what happened here).
        $shellApp = New-Object -ComObject "Shell.Application"
        $shellApp.ShellExecute($psExe, $arg, (Split-Path $scr), 'open', 1)
        [void][Runtime.Interopservices.Marshal]::ReleaseComObject($shellApp)
        $ok = $true
    } catch {
        Write-Host "Automatic relaunch via Shell COM failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Falling back to a Scheduled Task..." -ForegroundColor Yellow
        try {
            $me      = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            $action  = New-ScheduledTaskAction -Execute $psExe -Argument $arg -WorkingDirectory (Split-Path $scr)
            $princ   = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited
            $task    = New-ScheduledTask -Action $action -Principal $princ
            $tn      = 'LeaverCleanup_RunAsUser'
            Register-ScheduledTask -TaskName $tn -InputObject $task -Force -ErrorAction Stop | Out-Null
            Start-ScheduledTask -TaskName $tn -ErrorAction Stop
            Start-Sleep -Seconds 2
            Unregister-ScheduledTask -TaskName $tn -Confirm:$false -ErrorAction SilentlyContinue
            $ok = $true
        } catch {
            Write-Host "Scheduled Task fallback also failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    if ($ok) {
        Write-Host "A new window is opening as your normal user - this admin window will close by itself..." -ForegroundColor Green
        Start-Sleep -Seconds 3
        exit 42   # special code tells the launcher (.bat) to close without pausing
    } else {
        Write-Host "Could not auto-drop admin rights - continuing. Exchange sign-in will retry automatically if needed." -ForegroundColor Yellow
    }
}

# ---------- Required: verify the source repository is reachable (does not respect -NoSelfUpdate) ----------
try {
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/xnostra/Sherborne-Leaver-Cleanup-Tool/main/README.md" -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop | Out-Null
} catch {
    Write-Host ""
    Write-Host "This tool could not verify access to its source repository and cannot continue." -ForegroundColor Red
    Write-Host "(github.com/xnostra/Sherborne-Leaver-Cleanup-Tool)" -ForegroundColor Red
    exit 1
}

# ---------- Self-update: always run the newest code from GitHub (fully automatic) ----------
# No manual version bumping needed. It detects any new commit on your repo by itself:
#   - a git CLONE is pulled (fast-forward only, always forward, never a downgrade);
#   - a plain DOWNLOAD (no git) checks the repo's latest commit id via the GitHub API and, if it has
#     changed since last time, fetches the script and replaces itself - but only if the download is a
#     valid copy of this tool (right header + param block + sensible size), so it can't install a broken file.
# Fails safe if offline / API unreachable; skip with -NoSelfUpdate.
$script:Owner = 'xnostra'
$script:Repo  = 'Sherborne-Leaver-Cleanup-Tool'
$script:UpdateRawUrl = "https://raw.githubusercontent.com/$($script:Owner)/$($script:Repo)/main/Disable-RemoveLicenses.ps1"
if (-not $NoSelfUpdate -and -not $Updated) {
    $relArg = if ($Relaunched) { ' -Relaunched' } else { '' }
    $restart = {
        Write-Host "Updated to the newest version on GitHub - restarting the tool..." -ForegroundColor Green
        Start-Sleep -Seconds 1
        Start-Process (Join-Path $PSHOME 'powershell.exe') -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Updated$relArg"
    }
    try {
        if (Test-Path (Join-Path $PSScriptRoot '.git')) {
            # --- Git clone: pull latest (fast-forward only) ---
            $gitExe = (Get-Command git -ErrorAction SilentlyContinue).Source
            if (-not $gitExe) {
                $cands = @("$env:ProgramFiles\Git\cmd\git.exe", "${env:ProgramFiles(x86)}\Git\cmd\git.exe", "$env:LOCALAPPDATA\Programs\Git\cmd\git.exe")
                $cands += (Get-ChildItem "$env:LOCALAPPDATA\GitHubDesktop\app-*\resources\app\git\cmd\git.exe" -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName)
                $gitExe = $cands | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
            }
            if ($gitExe) {
                Write-Host "Checking GitHub for the latest version..." -ForegroundColor Gray
                $env:GIT_TERMINAL_PROMPT = '0'
                Push-Location $PSScriptRoot
                $before = (& $gitExe rev-parse HEAD 2>$null)
                & $gitExe fetch --quiet 2>$null
                & $gitExe merge --ff-only '@{u}' 2>$null | Out-Null
                $after  = (& $gitExe rev-parse HEAD 2>$null)
                Pop-Location
                if ($before -and $after -and $before -ne $after) { & $restart; exit }
            }
        }
        else {
            # --- Plain download (no git): detect new commit via GitHub API, then fetch & replace ---
            Write-Host "Checking GitHub for the latest version..." -ForegroundColor Gray
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $hdr = @{ 'User-Agent' = 'Sherborne-Leaver-Cleanup-Tool'; 'Accept' = 'application/vnd.github+json' }
            $remoteSha = (Invoke-RestMethod -Uri "https://api.github.com/repos/$($script:Owner)/$($script:Repo)/commits/main" -Headers $hdr -TimeoutSec 15 -ErrorAction Stop).sha
            $stateFile = Join-Path $PSScriptRoot 'LeaverTool.update-state'
            $storedSha = if (Test-Path $stateFile) { (Get-Content -Raw -LiteralPath $stateFile -ErrorAction SilentlyContinue).Trim() } else { '' }
            if ($remoteSha -and $remoteSha -ne $storedSha) {
                # There is a new commit since we last checked - get the script and update if it truly changed
                $remote = (Invoke-WebRequest -Uri $script:UpdateRawUrl -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop).Content
                $looksValid = $remote -and $remote.Length -gt 5000 -and ($remote -match 'MICROSOFT 365 LEAVER CLEANUP TOOL') -and ($remote -match 'param\(')
                if ($looksValid) {
                    $norm  = { param($t) ("$t".Replace([char]0xFEFF, '').Replace("`r", '')) }
                    $local = Get-Content -Raw -LiteralPath $PSCommandPath -ErrorAction Stop
                    if ((& $norm $remote) -ne (& $norm $local)) {
                        # write UTF-8 WITHOUT BOM so the comparison never re-triggers next launch
                        [System.IO.File]::WriteAllText($PSCommandPath, $remote, (New-Object System.Text.UTF8Encoding($false)))
                        Set-Content -LiteralPath $stateFile -Value $remoteSha -Encoding ASCII -ErrorAction SilentlyContinue
                        & $restart; exit
                    } else {
                        # already up to date - just remember this commit so we don't re-check the file every launch
                        Set-Content -LiteralPath $stateFile -Value $remoteSha -Encoding ASCII -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    } catch { Pop-Location -ErrorAction SilentlyContinue }
}

# ---------- Version ----------
$script:ToolVersion = '1.1'

# ---------- Automatic log file (a permanent, dated history of every run) ----------
$script:LogFile = Join-Path $PSScriptRoot ("Log_{0:yyyyMM}.log" -f (Get-Date))
function Write-Log {
    param([string]$Message)
    try { Add-Content -Path $script:LogFile -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $Message) -ErrorAction SilentlyContinue } catch { }
}
Write-Log "=== Tool started (v$($script:ToolVersion)) by $([Security.Principal.WindowsIdentity]::GetCurrent().Name) ==="

# ---------- Local settings (LeaverTool.local.json, kept next to the script, NOT committed to GitHub) ----------
# This file holds YOUR organisation-specific values so the script itself can stay generic/shareable.
$script:cfgFile = Join-Path $PSScriptRoot 'LeaverTool.local.json'
$script:cfg = @{ OrgName = 'Your Organisation'; ArchiveSiteUrl = ''; SafeCountries = @('QA') }
if (Test-Path $script:cfgFile) {
    try {
        $saved = Get-Content $script:cfgFile -Raw | ConvertFrom-Json
        if ($saved.OrgName)        { $script:cfg.OrgName = $saved.OrgName }
        if ($saved.ArchiveSiteUrl) { $script:cfg.ArchiveSiteUrl = $saved.ArchiveSiteUrl }
        if ($saved.SafeCountries)  { $script:cfg.SafeCountries = @($saved.SafeCountries) }
    } catch { }
}
if ($ArchiveSiteUrl) { $script:cfg.ArchiveSiteUrl = $ArchiveSiteUrl }
# command-line -SafeCountries overrides the saved one
if ($PSBoundParameters.ContainsKey('SafeCountries')) { $script:cfg.SafeCountries = $SafeCountries } else { $SafeCountries = $script:cfg.SafeCountries }
function Save-Config { try { $script:cfg | ConvertTo-Json | Set-Content $script:cfgFile } catch { } }

# ---------- Professional banner ----------
Clear-Host
$orgLine = ("{0}  -  IT Department" -f $script:cfg.OrgName)
$pad = [int]((50 - $orgLine.Length) / 2); if ($pad -lt 1) { $pad = 1 }
$orgCentered = (' ' * $pad) + $orgLine
Write-Host ""
Write-Host "   +==================================================+" -ForegroundColor Cyan
Write-Host "   |        MICROSOFT 365 LEAVER CLEANUP TOOL         |" -ForegroundColor Cyan
Write-Host ("   |{0,-50}|" -f $orgCentered) -ForegroundColor Cyan
Write-Host "   |                    version $($script:ToolVersion)                    |" -ForegroundColor Cyan
Write-Host "   +==================================================+" -ForegroundColor Cyan
Write-Host ""
Write-Host "   Please wait while the tool starts up (first launch takes longer)..." -ForegroundColor Gray
Write-Host ""

# ---------- Auto-install requirements (skips anything already installed) ----------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Progress -Activity "Starting up" -Status "Checking required components..." -PercentComplete 5
if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
    Write-Host "Installing NuGet package provider (first time on this computer)..."
    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
}
foreach ($m in 'Microsoft.Graph.Users','Microsoft.Graph.Users.Actions','Microsoft.Graph.Groups','Microsoft.Graph.Reports','Microsoft.Online.SharePoint.PowerShell','ImportExcel') {
    if (Get-Module -ListAvailable -Name $m) {
        Write-Host "Module $m already installed - skipping."
    } else {
        Write-Host "Installing module $m (first time on this computer)..."
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
}
# Exchange module is PINNED to 3.4.0: newer versions crash when used together with the Graph modules
# on Windows PowerShell 5.1 ('Method not found ... WithBroker').
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement | Where-Object { $_.Version -eq [version]'3.4.0' }) {
    Write-Host "Module ExchangeOnlineManagement 3.4.0 already installed - skipping."
} else {
    Write-Host "Installing module ExchangeOnlineManagement 3.4.0 (pinned version, first time on this computer)..."
    Install-Module ExchangeOnlineManagement -RequiredVersion 3.4.0 -Scope CurrentUser -Force -AllowClobber
}

Write-Progress -Activity "Starting up" -Status "Signing in to Microsoft 365 (a sign-in window may appear)..." -PercentComplete 60
Write-Host "Signing in to Microsoft 365..." -ForegroundColor Gray
Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All","AuditLog.Read.All","Directory.Read.All","GroupMember.ReadWrite.All","Files.ReadWrite.All","Sites.ReadWrite.All" -NoWelcome
Write-Progress -Activity "Starting up" -Status "Ready" -PercentComplete 100 -Completed
Write-Host "Signed in. Opening the menu..." -ForegroundColor Green

$activeCutoff = (Get-Date).ToUniversalTime().AddDays(-$SkipIfActiveWithinDays)

$script:exoConnected = $false
function Connect-EXO {
    # Lazy Exchange connection, only when mailbox/distribution-list work is actually needed.
    if ($script:exoConnected -and (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) { return }
    Import-Module ExchangeOnlineManagement -RequiredVersion 3.4.0 -ErrorAction Stop
    $cmd = Get-Command Connect-ExchangeOnline
    $devParam = @('Device','UseDeviceAuthentication','DeviceCode') | Where-Object { $cmd.Parameters.ContainsKey($_) } | Select-Object -First 1

    # The interactive sign-in often throws an 'ActiveX/single-threaded apartment' error on the FIRST
    # try but succeeds on the next - so we retry it several times automatically. The module prints a
    # long red stack trace on the failed attempt; we hide that noise so the console stays clean.
    Write-Host "Connecting to Exchange Online (a sign-in window will appear - please wait)..." -ForegroundColor Gray
    $ok = $false
    # The module prints a long red .NET stack trace directly to the console on a failed attempt.
    # We temporarily redirect the raw console output to hide that noise, then restore it.
    $origOut = [Console]::Out
    $origErr = [Console]::Error
    for ($try = 1; $try -le 5 -and -not $ok; $try++) {
        if ($try -gt 1) { Write-Host "  still connecting... (attempt $try)" -ForegroundColor Gray; Start-Sleep -Seconds 2 }
        $sink = New-Object System.IO.StringWriter
        try {
            [Console]::SetOut($sink); [Console]::SetError($sink)
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop *>$null
            if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { $ok = $true }
        } catch { $ok = $false }
        finally { [Console]::SetOut($origOut); [Console]::SetError($origErr) }
    }
    if ($ok) { Write-Host "Connected to Exchange Online." -ForegroundColor Green }

    # Last resort: device-code sign-in if the module supports it (no popup window needed at all).
    if (-not $ok -and $devParam) {
        Write-Host ""
        Write-Host "==================== EXCHANGE SIGN-IN (code) ====================" -ForegroundColor Cyan
        Write-Host " Open  https://microsoft.com/devicelogin  and enter the code shown below," -ForegroundColor Yellow
        Write-Host " then sign in with your GLOBAL ADMIN account." -ForegroundColor Yellow
        Write-Host "================================================================" -ForegroundColor Cyan
        try {
            $exoArgs = @{ ShowBanner = $false; ErrorAction = 'Stop' }
            $exoArgs[$devParam] = $true
            Connect-ExchangeOnline @exoArgs
            if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { $ok = $true }
        } catch { $ok = $false }
    }

    if (-not $ok) { throw "Exchange Online connection did not complete (Get-Mailbox unavailable)." }
    $script:exoConnected = $true
}

# Retry a command a few times - the FIRST Exchange operation after connecting often fails transiently
# ('ActiveX'/'not warmed up'), but succeeds moments later. This turns that into a silent auto-retry.
function Invoke-WithRetry {
    param([scriptblock]$Script, [int]$Tries = 4, [int]$DelaySeconds = 4)
    for ($attempt = 1; $attempt -le $Tries; $attempt++) {
        try { return (& $Script) }
        catch {
            if ($attempt -eq $Tries) { throw }
            Start-Sleep -Seconds $DelaySeconds
            # Do NOT reconnect here - that re-triggers the sign-in prompt. Just retry the command;
            # the existing session is almost always still valid, the operation was just transiently busy.
        }
    }
}

# ---------- Helpers ----------
# Friendly display names for common Microsoft license "part numbers". Anything not listed here
# falls back to the tenant's own part number (which is accurate), never a raw GUID.
$script:skuFriendly = @{
    'M365EDU_A1'                  = 'Microsoft 365 A1'
    'M365EDU_A3_STUUSEBNFT'       = 'Microsoft 365 A3 for students use benefit'
    'M365EDU_A3_FACULTY'          = 'Microsoft 365 A3 for faculty'
    'M365EDU_A5_STUUSEBNFT'       = 'Microsoft 365 A5 for students use benefit'
    'M365EDU_A5_FACULTY'          = 'Microsoft 365 A5 for faculty'
    'M365EDU_A5_NOPSTNCONF_STUUSEBNFT' = 'Microsoft 365 A5 without Audio Conferencing (students)'
    'STANDARDWOFFPACK_STUDENT'    = 'Office 365 A1 for students'
    'STANDARDWOFFPACK_FACULTY'    = 'Office 365 A1 for faculty'
    'STANDARDWOFFPACK_IW_STUDENT' = 'Office 365 A1 Plus for students'
    'STANDARDWOFFPACK_IW_FACULTY' = 'Office 365 A1 Plus for faculty'
    'OFFICESUBSCRIPTION_STUDENT'  = 'Microsoft 365 Apps for students'
    'OFFICESUBSCRIPTION_FACULTY'  = 'Microsoft 365 Apps for faculty'
    'ENTERPRISEPACK'              = 'Office 365 E3'
    'ENTERPRISEPREMIUM'           = 'Office 365 E5'
    'FLOW_FREE'                   = 'Power Automate Free'
    'POWER_BI_STANDARD'           = 'Power BI (free)'
    'PROJECT_P1'                  = 'Project Plan 1'
    'WIN_DEF_ATP'                 = 'Microsoft Defender for Endpoint'
    'EMS'                         = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                  = 'Enterprise Mobility + Security E5'
}
# Live map of the tenant's own SkuId -> part number (built once, on first use).
$script:skuIdToPart = $null
function Get-TenantSkuMap {
    if ($null -ne $script:skuIdToPart) { return $script:skuIdToPart }
    $script:skuIdToPart = @{}
    try {
        foreach ($s in (Get-MgSubscribedSku -All -ErrorAction Stop)) {
            if ($s.SkuId) { $script:skuIdToPart["$($s.SkuId)".ToLower()] = "$($s.SkuPartNumber)" }
        }
    } catch { }
    return $script:skuIdToPart
}
function Get-LicNames($skus) {
    $map = Get-TenantSkuMap
    (@($skus) | Where-Object { $_ } | ForEach-Object {
        $id = "$_".ToLower()
        $part = if ($map.ContainsKey($id)) { $map[$id] } else { '' }
        if ($part -and $script:skuFriendly.ContainsKey($part)) { $script:skuFriendly[$part] }
        elseif ($part) { $part }
        else { "Unknown licence ($("$_".Substring(0,8))...)" }
    }) -join ', '
}

function Get-EditDistance($s, $t) {
    $n = $s.Length; $m = $t.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $d = New-Object 'int[,]' -ArgumentList ($n + 1), ($m + 1)
    for ($i = 0; $i -le $n; $i++) { $d[$i, 0] = $i }
    for ($j = 0; $j -le $m; $j++) { $d[0, $j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        for ($j = 1; $j -le $m; $j++) {
            $pi = $i - 1
            $pj = $j - 1
            $c = if ($s[$pi] -eq $t[$pj]) { 0 } else { 1 }
            $del = $d[$pi, $j] + 1
            $ins = $d[$i, $pj] + 1
            $sub = $d[$pi, $pj] + $c
            $d[$i, $j] = [math]::Min([math]::Min($del, $ins), $sub)
        }
    }
    return $d[$n, $m]
}

# Smart name comparison: ignores case, punctuation, spaces, word order, middle names,
# joined/split names and small spelling variants.
function Test-SameName($listName, $accountName) {
    if (-not $listName -or -not $accountName) { return $false }
    $na = (($listName    -replace "[^\w\s]", ' ') -replace '\s+', ' ').Trim().ToLower()
    $nb = (($accountName -replace "[^\w\s]", ' ') -replace '\s+', ' ').Trim().ToLower()
    if (-not $na -or -not $nb) { return $false }
    $ja = $na -replace ' ', ''
    $jb = $nb -replace ' ', ''
    if ($ja -eq $jb) { return $true }
    if (($ja.Length -ge 6 -and $jb.Contains($ja)) -or ($jb.Length -ge 6 -and $ja.Contains($jb))) { return $true }
    $a = @($na -split ' '); $b = @($nb -split ' ')
    $short = $a; $long = $b
    if ($b.Count -lt $a.Count) { $short = $b; $long = $a }
    foreach ($t in $short) {
        $ok = $false
        foreach ($w in $long) {
            if ($t -eq $w) { $ok = $true; break }
            $minLen = [math]::Min($t.Length, $w.Length)
            $tol = if ($minLen -ge 5) { 2 } elseif ($minLen -ge 3) { 1 } else { 0 }
            if ($tol -gt 0 -and (Get-EditDistance $t $w) -le $tol) { $ok = $true; break }
        }
        if (-not $ok) { return $false }
    }
    return $true
}

# ---------- Popup dialog helpers (all questions are asked in windows, not the console) ----------
function Initialize-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic
}

# A window with a message and one button per choice. Returns the clicked button's number (0-based), -1 if closed.
function Show-Menu {
    param([string]$Title, [string]$Message, [string[]]$Buttons)
    Initialize-Gui
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false; $form.MinimizeBox = $true; $form.TopMost = $true
    $form.AutoSize = $true; $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding(18)
    $form.Tag = -1
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.FlowDirection = 'TopDown'; $panel.AutoSize = $true; $panel.WrapContents = $false
    if ($Message) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $Message
        $lbl.AutoSize = $true
        $lbl.MaximumSize = New-Object System.Drawing.Size(430, 0)
        $lbl.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 14)
        $panel.Controls.Add($lbl)
    }
    for ($bi = 0; $bi -lt $Buttons.Count; $bi++) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $Buttons[$bi]
        $btn.Width = 430; $btn.Height = 40
        $btn.Tag = $bi
        $btn.Add_Click({ $form.Tag = $this.Tag; $form.Close() }.GetNewClosure())
        $panel.Controls.Add($btn)
    }
    $form.Controls.Add($panel)
    [void]$form.ShowDialog()
    return [int]$form.Tag
}

function Show-YesNo {
    param([string]$Message, [string]$Title)
    Initialize-Gui
    return ([System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Show-Confirm {
    param([string]$Message, [string]$Title)
    Initialize-Gui
    return ([System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Show-Input {
    param([string]$Message, [string]$Title, [string]$Default = '')
    Initialize-Gui
    return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $Title, $Default)
}

function Show-Info {
    param([string]$Message, [string]$Title)
    Initialize-Gui
    [void][System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}


# File picker popup for choosing the leavers list
function Select-InputFile {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select your leavers list (CSV or Excel)"
    $dlg.Filter = "Leaver lists (*.csv;*.xlsx)|*.csv;*.xlsx|All files (*.*)|*.*"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
    return $null
}

# Does this account look like a STAFF member? (faculty license, or a job title/department set)
function Test-LooksLikeStaff($u, $licText) {
    if ($licText -match 'Faculty') { return $true }
    if ($u.JobTitle -and "$($u.JobTitle)".Trim()) { return $true }
    if ($u.Department -and "$($u.Department)".Trim()) { return $true }
    return $false
}

function Get-SignInSecurity($UserId) {
    try {
        $logs = @(Get-MgAuditLogSignIn -Filter "userId eq '$UserId'" -Top 30 -ErrorAction Stop)
    } catch { return 'no sign-in log data' }
    if ($logs.Count -eq 0) { return '' }
    $fails     = @($logs | Where-Object { $_.Status.ErrorCode -ne 0 }).Count
    $countries = @($logs | ForEach-Object { $_.Location.CountryOrRegion } | Where-Object { $_ } | Select-Object -Unique)
    $foreign   = @($countries | Where-Object { $SafeCountries -notcontains $_ })
    $parts = @()
    if ($foreign.Count -gt 0) { $parts += "ALERT - sign-ins from outside your safe countries ($($SafeCountries -join ',')): $($foreign -join ', ')" }
    if ($fails -ge 3)         { $parts += "ALERT - $fails failed sign-in attempts" }
    if ($parts.Count -eq 0)   { return "OK - safe countries only ($($logs.Count) recent sign-ins, $fails failed)" }
    return ($parts -join '; ')
}

# Turns one Graph sign-in log response (already fetched by the batch call below) into the same
# summary string Get-SignInSecurity produces, without making its own network call.
function ConvertTo-SignInSecuritySummary($Logs) {
    $logs = @($Logs)
    if ($logs.Count -eq 0) { return '' }
    $fails     = @($logs | Where-Object { $_.status.errorCode -ne 0 }).Count
    $countries = @($logs | ForEach-Object { $_.location.countryOrRegion } | Where-Object { $_ } | Select-Object -Unique)
    $foreign   = @($countries | Where-Object { $SafeCountries -notcontains $_ })
    $parts = @()
    if ($foreign.Count -gt 0) { $parts += "ALERT - sign-ins from outside your safe countries ($($SafeCountries -join ',')): $($foreign -join ', ')" }
    if ($fails -ge 3)         { $parts += "ALERT - $fails failed sign-in attempts" }
    if ($parts.Count -eq 0)   { return "OK - safe countries only ($($logs.Count) recent sign-ins, $fails failed)" }
    return ($parts -join '; ')
}

# ---------- Batched sign-in checks (speeds up large lists a lot) ----------
# Instead of one network round-trip per person (slow for lists of hundreds), this asks Microsoft
# Graph's official "$batch" endpoint for up to 20 people's sign-in logs in a single call.
# Same connection, same sign-in, nothing new to authenticate - just far fewer, bigger requests.
# If a batch call fails for any reason, those user(s) are simply left out of the map, and the
# normal one-by-one Get-SignInSecurity is used for them instead - so this can only ever speed
# things up, never cause a missed or wrong check.
function Get-SignInSecurityBatch {
    param([string[]]$UserIds)
    $resultMap = @{}
    $ids = @($UserIds | Where-Object { $_ } | Select-Object -Unique)
    if ($ids.Count -eq 0) { return $resultMap }

    $chunkSize = 20   # Microsoft Graph's limit per $batch call
    for ($i = 0; $i -lt $ids.Count; $i += $chunkSize) {
        $lastIdx = [math]::Min($i + $chunkSize - 1, $ids.Count - 1)
        $chunk = @($ids[$i..$lastIdx])
        $requests = @()
        for ($j = 0; $j -lt $chunk.Count; $j++) {
            $filterVal = "userId eq '$($chunk[$j])'"
            $encFilter = [uri]::EscapeDataString($filterVal)
            $requests += @{ id = "$j"; method = 'GET'; url = "/auditLogs/signIns?`$filter=$encFilter&`$top=30" }
        }
        $bodyJson = (@{ requests = $requests } | ConvertTo-Json -Depth 8)
        try {
            $resp = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop
            foreach ($r in $resp.responses) {
                $idx = [int]$r.id
                if ($idx -lt 0 -or $idx -ge $chunk.Count) { continue }
                $uid = $chunk[$idx]
                if ($r.status -eq 200 -and $r.body -and $r.body.value) {
                    $resultMap[$uid] = ConvertTo-SignInSecuritySummary $r.body.value
                } elseif ($r.status -eq 200) {
                    $resultMap[$uid] = ''
                }
                # any non-200 status: leave unset so the caller falls back to the live per-user check
            }
        } catch {
            # whole batch failed - leave this chunk's IDs unset, caller falls back automatically
        }
    }
    return $resultMap
}

# Copy a leaver's entire OneDrive into the archive SharePoint site (server-side, nothing downloads locally)
$script:spoConnected = $false
function Copy-OneDriveToArchive {
    param($User, [string]$ArchiveUrl)
    try {
        $adminUpn = (Get-MgContext).Account
        $spHost   = ([uri]$ArchiveUrl).Host                                                  # tenant.sharepoint.com
        $adminUrl = 'https://' + $spHost.Replace('.sharepoint.com', '-admin.sharepoint.com')
        $odUrl    = 'https://' + $spHost.Replace('.sharepoint.com', '-my.sharepoint.com') + '/personal/' + ($User.UserPrincipalName -replace '[@.]', '_')

        # 1) Grant ourselves temporary admin access to the leaver's OneDrive
        if (-not $script:spoConnected) {
            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
            Connect-SPOService -Url $adminUrl
            $script:spoConnected = $true
        }
        Set-SPOUser -Site $odUrl -LoginName $adminUpn -IsSiteCollectionAdmin $true | Out-Null

        # 2) Source: the leaver's OneDrive contents
        $drive = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($User.Id)/drive"
        $items = @()
        $next = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root/children"
        while ($next) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next
            $items += $resp.value
            $next = $resp.'@odata.nextLink'
        }
        if ($items.Count -eq 0) { return 'OneDrive is empty - nothing to archive' }

        # 3) Destination: a folder named after the leaver in the archive site's document library.
        #    Reuse an existing folder for this person if one is already there (e.g. this account was
        #    processed on a previous run) instead of creating a second "Name (1)" copy - that's what
        #    used to cause duplicate archive folders to pile up.
        $sitePath  = ([uri]$ArchiveUrl).AbsolutePath
        $site      = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/${spHost}:$sitePath"
        $destDrive = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive"
        $folderName = "$($User.DisplayName) ($(($User.UserPrincipalName -split '@')[0]))"

        $folder = $null
        try {
            $encName = [uri]::EscapeDataString($folderName)
            $existingFolder = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$($destDrive.id)/root:/${encName}" -ErrorAction Stop
            if ($existingFolder.folder) { $folder = $existingFolder }   # only reuse it if it's actually a folder
        } catch { }   # not found is the normal case for a first-time archive - fall through to create it

        $reusedFolder = [bool]$folder
        if (-not $folder) {
            $folderBody = @{ name = $folderName; folder = @{}; '@microsoft.graph.conflictBehavior' = 'rename' } | ConvertTo-Json
            $folder = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/drives/$($destDrive.id)/root/children" -Body $folderBody -ContentType 'application/json'
        }

        # If reusing an existing folder, see what's already in it so files copied on a previous run
        # aren't copied again - that's what used to cause "file (1).docx", "file (2).docx" duplicates.
        $existingNames = @{}
        if ($reusedFolder) {
            try {
                $existingItems = @()
                $enext = "https://graph.microsoft.com/v1.0/drives/$($destDrive.id)/items/$($folder.id)/children"
                while ($enext) {
                    $eresp = Invoke-MgGraphRequest -Method GET -Uri $enext
                    $existingItems += $eresp.value
                    $enext = $eresp.'@odata.nextLink'
                }
                foreach ($ei in $existingItems) { $existingNames["$($ei.name)|$($ei.size)"] = $true }
            } catch { }
        }

        # 4) Kick off server-side copies (Microsoft finishes them in the background) - skip anything
        #    that's already there (same name and size) from a previous run of this same account.
        $n = 0
        $skipped = 0
        foreach ($it in $items) {
            $key = "$($it.name)|$($it.size)"
            if ($existingNames.ContainsKey($key)) { $skipped++; continue }
            $body = @{ parentReference = @{ driveId = $destDrive.id; id = $folder.id }; '@microsoft.graph.conflictBehavior' = 'rename' } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($it.id)/copy" -Body $body -ContentType 'application/json' | Out-Null
            $n++
        }
        $skipNote = if ($skipped -gt 0) { " ($skipped item(s) already archived from a previous run, skipped)" } else { '' }
        return "OneDrive archived: $n item(s) copying to folder '$folderName' in the archive site (completes by itself in the background)$skipNote"
    }
    catch {
        return "OneDrive NOT auto-archived ($($_.Exception.Message)) - copy manually within 93 days"
    }
}

# ---------- Shared: download all tenant users (with 4-hour cache) ----------
function Get-AllTenantUsers {
    param([bool]$ForceRefresh = $false)
    $cacheFile = Join-Path $env:TEMP 'LeaverTool_UserCache_v2.xml'
    $allUsers = $null
    if (-not $ForceRefresh -and -not $RefreshUserCache -and (Test-Path $cacheFile) -and ((Get-Date) - (Get-Item $cacheFile).LastWriteTime).TotalHours -lt 4) {
        Write-Host "Using cached user list from $((Get-Item $cacheFile).LastWriteTime.ToString('HH:mm')) (use -RefreshUserCache for a fresh one)..."
        # If you're running Students and Staff at the same time in two windows, they share this
        # cache. Reading it can very rarely land mid-write by the other window - if so, this just
        # falls back to a fresh download instead of crashing (this used to be able to corrupt the
        # cache and fail with an XML parse error - now it can't).
        try { $allUsers = Import-Clixml $cacheFile -ErrorAction Stop } catch {
            Write-Host "Cached user list could not be read (likely being refreshed by another window right now) - downloading fresh instead." -ForegroundColor Yellow
            $allUsers = $null
        }
    }
    if (-not $allUsers) {
        Write-Host "Downloading all users from Microsoft 365 (slow, but cached for 4 hours afterwards)..." -ForegroundColor Gray
        Write-Progress -Activity "Downloading directory" -Status "Fetching all user accounts - please wait, this can take a few minutes..." -PercentComplete 30
        $props = "Id,DisplayName,UserPrincipalName,Mail,ProxyAddresses,AccountEnabled,AssignedLicenses,CreatedDateTime,SignInActivity,JobTitle,Department,OnPremisesSyncEnabled"
        $allUsers = Get-MgUser -All -PageSize 999 -Property $props
        Write-Progress -Activity "Downloading directory" -Status "Saving a local copy for faster re-runs..." -PercentComplete 90
        try {
            # Write to a temp file, then rename it into place. A rename is atomic on Windows, so a
            # second window reading the cache at the same time (Students + Staff running together)
            # always sees either the complete old file or the complete new one - never a half-written
            # one, which is what used to cause the corrupted-cache crash.
            $tmpCacheFile = "$cacheFile.$PID.tmp"
            $allUsers | Export-Clixml $tmpCacheFile
            Move-Item -Path $tmpCacheFile -Destination $cacheFile -Force
        } catch { }
        Write-Progress -Activity "Downloading directory" -Completed
    }
    return $allUsers
}

# ---------- Purge mode: delete long-dead accounts (disabled + unlicensed + inactive for years) ----------
function Start-PurgeMode {
    $t = Show-Menu 'Purge stale accounts' "Deletes accounts that are ALL THREE of:`n- disabled`n- zero licenses`n- inactive for the chosen time`n`nDeleted accounts go to Microsoft's recycle bin for 30 days (restorable in Entra admin center > Deleted users), then are gone FOREVER including OneDrive.`n`nShared mailboxes (preserved staff email) are detected and NEVER deleted.`n`nHow long must an account have been dead?" @(
        'Inactive 1+ year',
        'Inactive 2+ years  (recommended)',
        'Inactive 3+ years',
        'Cancel')
    if ($t -lt 0 -or $t -eq 3) { return }
    $years  = $t + 1
    $cutoff = (Get-Date).AddYears(-$years)

    $users = Get-AllTenantUsers -ForceRefresh $true   # purging always uses FRESH data, never the cache
    $cands = New-Object System.Collections.Generic.List[object]
    foreach ($u in $users) {
        if ($u.AccountEnabled) { continue }
        if (@($u.AssignedLicenses.SkuId | Where-Object { $_ }).Count -gt 0) { continue }
        $act  = $u.SignInActivity
        $last = @($act.LastSignInDateTime, $act.LastNonInteractiveSignInDateTime) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
        $ref  = $last
        if (-not $ref) { $ref = $u.CreatedDateTime }   # never signed in: age = account age
        if (-not $ref -or $ref -gt $cutoff) { continue }
        $cands.Add(@{ U = $u; Last = $last; Ref = $ref })
    }
    if ($cands.Count -eq 0) { Show-Info "No accounts qualify (disabled + unlicensed + inactive $years+ years). Nothing to do." 'Nothing to purge'; return }

    # Protect preserved staff email: exclude anything that is a shared mailbox.
    # FAIL CLOSED: if this check cannot run, the purge is cancelled - we never delete blind.
    Write-Host "Checking $($cands.Count) candidates against shared mailboxes (protected)..."
    $sharedSet = @{}
    try {
        Connect-EXO
        foreach ($m in @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)) {
            if ($m.UserPrincipalName) { $sharedSet[$m.UserPrincipalName.ToLower()] = $true }
        }
    } catch {
        Show-Info "Purge CANCELLED - nothing was deleted.`n`nThe shared-mailbox safety check could not run (Exchange Online connection failed):`n$($_.Exception.Message)`n`nWithout this check the purge could delete preserved staff email, so it refuses to continue. Fix the connection and try again." 'Purge cancelled for safety'
        return
    }
    $protectedCount = @($cands | Where-Object { $sharedSet.ContainsKey($_.U.UserPrincipalName.ToLower()) }).Count
    $cands = @($cands | Where-Object { -not $sharedSet.ContainsKey($_.U.UserPrincipalName.ToLower()) })
    if ($cands.Count -eq 0) { Show-Info "All qualifying accounts are protected shared mailboxes. Nothing to delete." 'Nothing to purge'; return }

    # Report before anything is deleted (Excel)
    $repPath = Join-Path $PSScriptRoot ("PurgeCandidates_{0:yyyyMMdd_HHmm}.xlsx" -f (Get-Date))
    $candRows = $cands | ForEach-Object {
        [pscustomobject]@{
            'Name'             = $_.U.DisplayName
            'Account'          = $_.U.UserPrincipalName
            'Created'          = if ($_.U.CreatedDateTime) { $_.U.CreatedDateTime.ToString('yyyy-MM-dd') } else { '' }
            'Last Sign-In'     = if ($_.Last) { $_.Last.ToString('yyyy-MM-dd') } else { 'never' }
            'Inactive (years)' = [math]::Round(((Get-Date) - $_.Ref).TotalDays / 365.0, 1)
        }
    }
    if (Get-Module -ListAvailable ImportExcel) {
        Import-Module ImportExcel
        $candRows | Export-Excel $repPath -WorksheetName 'Purge candidates' -TableStyle Medium2 -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
    } else {
        $repPath = [System.IO.Path]::ChangeExtension($repPath, '.csv')
        $candRows | Export-Csv $repPath -NoTypeInformation -Encoding UTF8
    }
    Write-Host "Candidate report saved: $repPath" -ForegroundColor Cyan

    $choice = Show-Menu 'Ready to purge' "$($cands.Count) stale account(s) qualify for deletion.`n($protectedCount shared-mailbox account(s) were excluded and protected.)`n`nA report of every candidate was saved:`n$repPath`n`nOpen and check it before deleting if you want a record for sign-off." @(
        "DELETE ALL $($cands.Count) now",
        'Review one by one',
        'Cancel - just keep the report')
    if ($choice -lt 0 -or $choice -eq 2) { return }

    $deleted = New-Object System.Collections.Generic.List[object]
    foreach ($c in $cands) {
        if ($choice -eq 1) {
            $a = Show-Menu 'Delete this account?' "$($c.U.DisplayName)  <$($c.U.UserPrincipalName)>`nCreated: $(if ($c.U.CreatedDateTime) { $c.U.CreatedDateTime.ToString('yyyy-MM-dd') })   Last sign-in: $(if ($c.Last) { $c.Last.ToString('yyyy-MM-dd') } else { 'never' })" @('DELETE', 'Keep', 'Stop reviewing')
            if ($a -eq 2 -or $a -lt 0) { break }
            if ($a -eq 1) { continue }
        }
        try {
            Remove-MgUser -UserId $c.U.Id -ErrorAction Stop
            Write-Host "  deleted: $($c.U.UserPrincipalName)" -ForegroundColor Green
            Write-Log "PURGE-DELETE: $($c.U.UserPrincipalName) ($($c.U.DisplayName))"
            $deleted.Add([pscustomobject]@{ 'Name' = $c.U.DisplayName; 'Account' = $c.U.UserPrincipalName; 'Deleted' = (Get-Date).ToString('yyyy-MM-dd HH:mm') })
        } catch {
            Write-Host "  FAILED to delete $($c.U.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "PURGE-FAILED: $($c.U.UserPrincipalName) -> $($_.Exception.Message)"
        }
    }
    if ($deleted.Count -gt 0) {
        $delPath = Join-Path $PSScriptRoot ("COMMITTED_purge_{0:yyyyMMdd_HHmm}.xlsx" -f (Get-Date))
        if (Get-Module -ListAvailable ImportExcel) {
            Import-Module ImportExcel
            $deleted | Export-Excel $delPath -WorksheetName 'Deleted' -TableStyle Medium2 -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        } else {
            $delPath = [System.IO.Path]::ChangeExtension($delPath, '.csv')
            $deleted | Export-Csv $delPath -NoTypeInformation -Encoding UTF8
        }
        Show-Info "$($deleted.Count) account(s) deleted.`n`nThey stay restorable for 30 days in the Entra admin center (Users > Deleted users). After that they are gone forever.`n`nRecord saved: $delPath" 'Purge finished'
    }
}

# ---------- Undo support: snapshot of an account before we touch it ----------
$script:restoreLog = New-Object System.Collections.Generic.List[object]
function Save-RestoreLog {
    if ($script:restoreLog.Count -eq 0) { return }
    $rf = Join-Path $PSScriptRoot ("Restore_{0:yyyyMMdd_HHmmss}.json" -f (Get-Date))
    $script:restoreLog | ConvertTo-Json -Depth 6 | Set-Content $rf
    Write-Host "UNDO file saved: $rf  (keep this - it can put any of these accounts back exactly as they were)" -ForegroundColor Cyan
    $script:restoreLog.Clear()
}

# ---------- Group-based license detection (cached per run - the tenant's license-assigning groups rarely change) ----------
# Builds a one-time lookup of every group in the tenant that has a license attached, so we can match
# a leaver's group memberships against it without querying every group each time.
$script:licenseGroupCache = $null
function Get-LicenseAssigningGroupCache {
    if ($null -ne $script:licenseGroupCache) { return $script:licenseGroupCache }
    $script:licenseGroupCache = @{}
    try {
        $groups = Get-MgGroup -All -Property "id,displayName,assignedLicenses" -ErrorAction Stop
        foreach ($g in $groups) {
            if ($g.AssignedLicenses -and $g.AssignedLicenses.Count -gt 0) {
                $script:licenseGroupCache[$g.Id] = [pscustomobject]@{
                    Id       = $g.Id
                    Name     = $g.DisplayName
                    Licenses = Get-LicNames @($g.AssignedLicenses.SkuId | Where-Object { $_ })
                }
            }
        }
        Write-Host "    (found $($script:licenseGroupCache.Count) license-assigning group(s) in your tenant)" -ForegroundColor DarkGray
    } catch {
        Write-Host "    ! Could not build the license-group lookup: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    return $script:licenseGroupCache
}

# Given a user's group memberships, returns which of them assign a license. Checks the user's
# DIRECT groups first (fast, no extra API call). If none of those match, the license may be coming
# from a NESTED group - a group the user only belongs to indirectly (member of Group A, and Group A
# is itself a member of Group B, which holds the license). Direct membership alone misses that, so
# as a fallback we check the user's full transitive membership too before giving up.
function Find-LicenseAssigningGroups {
    param($UserId, $UserGroups)
    $cache = Get-LicenseAssigningGroupCache
    $foundGroups = @()

    # This is Microsoft's authoritative per-user answer.  It tells us the exact
    # group that assigned each licence, rather than inferring it from membership.
    # That matters after the normal group-cleanup pass has already removed the
    # user, when a second removal returns ResourceNotFound even though the licence
    # is simply waiting for Entra's background processor to catch up.
    if ($UserId) {
        try {
            $states = @((Get-MgUser -UserId $UserId -Property 'licenseAssignmentStates' -ErrorAction Stop).LicenseAssignmentStates)
            foreach ($state in $states) {
                $groupId = "$($state.AssignedByGroup)"
                if (-not $groupId) { continue }
                if ($cache.ContainsKey($groupId)) {
                    $foundGroups += $cache[$groupId]
                } else {
                    try {
                        $group = Get-MgGroup -GroupId $groupId -Property 'id,displayName,assignedLicenses' -ErrorAction Stop
                        $foundGroups += [pscustomobject]@{ Id = $group.Id; Name = $group.DisplayName; Licenses = Get-LicNames @($group.AssignedLicenses.SkuId | Where-Object { $_ }) }
                    } catch {
                        $foundGroups += [pscustomobject]@{ Id = $groupId; Name = "licence group $groupId"; Licenses = '' }
                    }
                }
            }
            if ($foundGroups.Count -gt 0) { return @($foundGroups | Select-Object -Unique) }
        } catch { }
    }

    if ($UserGroups) {
        foreach ($ug in $UserGroups) {
            if ($cache.ContainsKey($ug.Id)) { $foundGroups += $cache[$ug.Id] }
        }
    }
    if ($foundGroups.Count -eq 0 -and $UserId) {
        try {
            $transitive = @(Get-MgUserTransitiveMemberOf -UserId $UserId -All -ErrorAction Stop |
                            Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' })
            foreach ($tg in $transitive) {
                if ($cache.ContainsKey($tg.Id)) { $foundGroups += $cache[$tg.Id] }
            }
            if ($foundGroups.Count -gt 0) {
                Write-Host "      (found via a NESTED group - not a direct membership)" -ForegroundColor DarkGray
            }
        } catch { }
    }
    return @($foundGroups | Select-Object -Unique)
}

# ---------- The actual cleanup of one account ----------
# Disables, removes groups/DLs, hides from address book, removes licenses. Returns a remarks string.
function Invoke-AccountCleanup {
    param($User, [bool]$DoGroups = $true, [bool]$DoHide = $true, [bool]$ConvertToShared = $false, [string]$Delegate = '', [string]$ArchiveUrl = '')
    $done = @()

    # Snapshot BEFORE any change, for the undo file
    $allGroups = @(Get-MgUserMemberOf -UserId $User.Id -All -ErrorAction SilentlyContinue |
                   Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' })
    $script:restoreLog.Add([pscustomobject]@{
        Timestamp         = (Get-Date).ToString('s')
        UserId            = $User.Id
        Upn               = $User.UserPrincipalName
        DisplayName       = $User.DisplayName
        Licenses          = @($User.AssignedLicenses.SkuId | Where-Object { $_ } | ForEach-Object { "$_" })
        Groups            = @($allGroups | ForEach-Object { [pscustomobject]@{ Id = $_.Id; Name = "$($_.AdditionalProperties.displayName)" } })
        ConvertedToShared = $ConvertToShared
    })

    $isSynced = ($User.OnPremisesSyncEnabled -eq $true)
    if ($isSynced) {
        Write-Host "    ! This account is SYNCED FROM ON-PREMISES AD - disable/hide/groups must be done in local AD." -ForegroundColor Yellow
        $done += "ON-PREMISES ACCOUNT (synced from local AD): disable, hide and group changes must be done in your LOCAL Active Directory - doing them only in the cloud will be reverted at the next sync. The cloud-safe steps below (licence, mailbox, OneDrive) were still done here"
    }

    Write-Host "    - disabling account (cloud)..." -ForegroundColor Gray
    try { Update-MgUser -UserId $User.Id -AccountEnabled:$false -ErrorAction Stop; $done += if ($isSynced) { 'disabled in cloud (will re-sync unless also disabled in local AD)' } else { 'disabled' } }
    catch { $done += "could not disable in cloud ($($_.Exception.Message)) - disable in local AD instead" }

    # Tags the account as a leaver via a directory schema extension attribute (the same
    # mechanism School Data Sync uses for Education_ObjectType) so the "All Users" dynamic
    # group rule can exclude it - see Setup-LeaverTag.ps1.
    # Uses extensionAttribute15 (built-in, no app registration needed) instead of a custom directory
    # schema extension - Graph only lets the OWNING app write custom schema extension values using its
    # own app credentials, which blocked this when called via delegated Microsoft Graph PowerShell auth.
    try {
        $tagBody = @{ onPremisesExtensionAttributes = @{ extensionAttribute15 = "Leaver" } } | ConvertTo-Json
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($User.Id)" -Body $tagBody -ContentType "application/json" -ErrorAction Stop
        $done += "tagged as leaver (extensionAttribute15) so it drops out of dynamic groups scoped to that tag (e.g. All Users)"
    } catch {
        $tagErrDetail = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $tagErrDetail = $_.ErrorDetails.Message }
        Write-Log "LEAVER TAG FAILED for $($User.UserPrincipalName): $tagErrDetail"
        $done += "could not set leaver tag ($tagErrDetail)$(if ($isSynced) { ' - this account is synced from local AD, which usually owns extensionAttribute15; set it there instead if needed' })"
    }

    if ($ConvertToShared) {
        try {
            Write-Host "    - converting mailbox to shared (signing in to Exchange if needed)..." -ForegroundColor Gray
            Connect-EXO
            $upn = $User.UserPrincipalName
            $mbx = Invoke-WithRetry { Get-Mailbox -Identity $upn -ErrorAction Stop }
            if (-not $mbx) { throw "this account has no Exchange Online mailbox (never licensed for email, or mailbox already deleted)" }
            if ($mbx.RecipientTypeDetails -eq 'SharedMailbox') {
                $done += 'mailbox is already SHARED'
            } else {
                Invoke-WithRetry { Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop }
                $done += 'mailbox converted to SHARED (email preserved, no license needed)'
            }
            if ($Delegate) {
                Write-Host "    - granting delegate access to $Delegate..." -ForegroundColor Gray
                try {
                    Invoke-WithRetry { Add-MailboxPermission -Identity $upn -User $Delegate -AccessRights FullAccess -AutoMapping $false -ErrorAction Stop | Out-Null }
                    $done += "delegate access given to $Delegate"
                } catch { $done += "could not give delegate access to $Delegate ($($_.Exception.Message))" }
            }
        } catch {
            $done += "STOPPED: mailbox NOT converted to shared - reason: $($_.Exception.Message). License deliberately NOT removed so no email can be lost. Fix the cause, then re-run this person (safe to repeat)."
            return ($done -join ', ')
        }
        if ($ArchiveUrl) {
            Write-Host "    - archiving OneDrive files to the SharePoint archive (this can take a moment)..." -ForegroundColor Gray
            $done += (Copy-OneDriveToArchive -User $User -ArchiveUrl $ArchiveUrl)
        } else {
            $odUrl = 'https://' + ($User.UserPrincipalName -split '@')[1].Split('.')[0] + '-my.sharepoint.com/personal/' + (($User.UserPrincipalName -replace '[@.]','_'))
            $done += "REMINDER: copy OneDrive files within 93 days: $odUrl"
        }
    }

    if ($DoGroups) {
        Write-Host "    - removing from $($allGroups.Count) group(s)/list(s)/team(s)..." -ForegroundColor Gray
        $removed = 0; $failed = 0
        foreach ($g in $allGroups) {
            $isDL = ($g.AdditionalProperties.mailEnabled -eq $true -and -not ($g.AdditionalProperties.groupTypes -contains 'Unified'))
            try {
                if ($isDL) {
                    Connect-EXO
                    Remove-DistributionGroupMember -Identity $g.Id -Member $User.Id -Confirm:$false -BypassSecurityGroupManagerCheck -ErrorAction Stop
                } else {
                    Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $User.Id -ErrorAction Stop
                }
                $removed++
            } catch { $failed++ }
        }
        $gtxt = "removed from $removed group(s)/list(s)/team(s)"
        if ($failed -gt 0) { $gtxt += " ($failed could not be removed - likely dynamic groups, safe to ignore" + ')' }
        $done += $gtxt
    }

    if ($DoHide) {
        Write-Host "    - hiding from the address book..." -ForegroundColor Gray
        try {
            Connect-EXO
            $upnH = $User.UserPrincipalName
            # A freshly converted mailbox can take a few seconds to be writable again - retry patiently.
            Invoke-WithRetry -Tries 6 -DelaySeconds 5 { Set-Mailbox -Identity $upnH -HiddenFromAddressListsEnabled $true -ErrorAction Stop }
            $done += 'hidden from address book'
        } catch {
            if ("$($_.Exception.Message)" -match 'on-premises|write scope|being synchronized') {
                $done += "could not hide from address book - this account is SYNCED FROM YOUR ON-PREMISES ACTIVE DIRECTORY, so it must be hidden there (set 'msExchHideFromAddressLists' on the AD object), not in the cloud"
            } else {
                $done += "could not hide from address book ($($_.Exception.Message))"
            }
        }
    }

    Write-Host "    - removing license(s)..." -ForegroundColor Gray
    $skus = @($User.AssignedLicenses.SkuId | Where-Object { $_ })
    if ($skus.Count -gt 0) {
        try {
            Set-MgUserLicense -UserId $User.Id -RemoveLicenses $skus -AddLicenses @() -ErrorAction Stop | Out-Null
            $done += "removed $($skus.Count) license(s)"
        } catch {
            $msg = "$($_.Exception.Message)"
            if ($msg -match 'group|does not have a corresponding license') {
                Write-Host "    - license is GROUP-BASED, checking which group is assigning it..." -ForegroundColor Gray
                $licGroups = @(Find-LicenseAssigningGroups -UserId $User.Id -UserGroups $allGroups)
                if ($licGroups.Count -gt 0) {
                    $removedNames = @()
                    $failedNotes  = @()
                    foreach ($lg in $licGroups) {
                        try {
                            Remove-MgGroupMemberByRef -GroupId $lg.Id -DirectoryObjectId $User.Id -ErrorAction Stop
                            Write-Host "      removed from license group: $($lg.Name)  ($($lg.Licenses))" -ForegroundColor Green
                            $removedNames += $lg.Name
                        } catch {
                            $gmsg = "$($_.Exception.Message)"
                            if ($gmsg -match 'Request_ResourceNotFound|does not exist') {
                                # This normally means the ordinary group-cleanup pass above has
                                # already removed the user.  Entra's group-licensing processor
                                # can take time to release the SKU, so it is not a failed cleanup.
                                Write-Host "      already removed from license group: $($lg.Name); requesting licence reprocessing..." -ForegroundColor DarkGray
                                $removedNames += "$($lg.Name) (already removed; release pending)"
                                continue
                            }
                            Write-Host "      could not remove from $($lg.Name): $gmsg" -ForegroundColor Yellow
                            # Give a specific reason where we can recognise it, instead of a bare "check permissions"
                            if ($gmsg -match 'on-premises|write scope|being synchronized') {
                                $failedNotes += "$($lg.Name) (this group is SYNCED FROM ON-PREMISES AD - membership must be changed there, not in the cloud)"
                            } elseif ($gmsg -match 'role-assignable|isAssignableToRole|privileged') {
                                $failedNotes += "$($lg.Name) (this is a role-assignable/privileged group - removing members needs a higher-privilege role than this tool signs in with)"
                            } elseif ($gmsg -match 'Authorization_RequestDenied|Forbidden|insufficient privileges') {
                                $failedNotes += "$($lg.Name) (permission denied: $gmsg)"
                            } else {
                                $failedNotes += "$($lg.Name) ($gmsg)"
                            }
                        }
                    }
                    if ($removedNames.Count -gt 0 -and (Get-Command Invoke-MgLicenseUser -ErrorAction SilentlyContinue)) {
                        try {
                            # Ask Entra to process the changed group membership now, instead of
                            # waiting for its normal background licensing cycle.
                            Invoke-MgLicenseUser -UserId $User.Id -ErrorAction Stop | Out-Null
                            $done += 'requested immediate Entra reprocessing of group-based license removal'
                        } catch {
                            $done += "could not request immediate group-license reprocessing ($($_.Exception.Message)); Entra will retry automatically"
                        }
                    }
                    if ($removedNames.Count -gt 0) {
                        $done += "license was GROUP-BASED - auto-removed from license-assigning group(s): $($removedNames -join ', ')"
                    }
                    if ($failedNotes.Count -gt 0) {
                        $done += "could NOT remove from license-assigning group(s): $($failedNotes -join '; ')"
                    }
                } else {
                    $done += "license is GROUP-BASED (assigned via a group, not directly) - could not identify which group even after checking nested memberships; it was removed from all its known groups above, so the license should still drop off automatically"
                }
            } else {
                $done += "license NOT removed: $msg"
            }
        }
    } else {
        $done += 'no direct licenses to remove'
    }

    Write-Log "CLEANUP: $($User.UserPrincipalName) ($($User.DisplayName)) -> $($done -join ', ')"
    return ($done -join ', ')
}

# ---------- Restore (undo) mode ----------
function Start-RestoreMode {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Pick the UNDO file from the run that removed the account"
    $dlg.Filter = "Undo files (Restore_*.json)|Restore_*.json|All files (*.*)|*.*"
    $dlg.InitialDirectory = $PSScriptRoot
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $entries = @(Get-Content $dlg.FileName -Raw | ConvertFrom-Json)
    if ($entries.Count -eq 0) { Show-Info "That undo file is empty." 'Nothing to restore'; return }

    $q = Show-Input "Enter (part of) the person's email or name." 'Who do you want to restore?'
    if (-not $q) { return }
    $hits = @($entries | Where-Object { $_.Upn -like "*$q*" -or $_.DisplayName -like "*$q*" })
    if ($hits.Count -eq 0) { Show-Info "No one matching '$q' in this undo file." 'Not found'; return }
    if ($hits.Count -gt 1) {
        $hits = @($hits | Select-Object -First 10)
        $labels = @($hits | ForEach-Object { "$($_.DisplayName)   <$($_.Upn)>" })
        $pick = Show-Menu 'Multiple matches' 'Which person?' ($labels + 'Cancel')
        if ($pick -lt 0 -or $pick -ge $hits.Count) { return }
        $hits = @($hits[$pick])
    }
    $e = $hits[0]
    if (-not (Show-Confirm "Restore:`n`n$($e.DisplayName)  <$($e.Upn)>`n`n- re-enable the account`n- re-assign $(@($e.Licenses).Count) license(s)`n- re-add $(@($e.Groups).Count) group(s)$(if ($e.ConvertedToShared) { "`n- convert mailbox back to normal" })`n`nContinue?" 'Confirm restore')) { return }

    try { Update-MgUser -UserId $e.UserId -AccountEnabled:$true; Write-Host "  account re-enabled" -ForegroundColor Green } catch { Write-Host "  could not re-enable: $($_.Exception.Message)" -ForegroundColor Red }

    if (@($e.Licenses).Count -gt 0) {
        try {
            $add = @($e.Licenses | ForEach-Object { @{ SkuId = $_ } })
            Set-MgUserLicense -UserId $e.UserId -AddLicenses $add -RemoveLicenses @() | Out-Null
            Write-Host "  $(@($e.Licenses).Count) license(s) re-assigned" -ForegroundColor Green
        } catch { Write-Host "  license re-assign failed: $($_.Exception.Message)" -ForegroundColor Red }
    }

    if ($e.ConvertedToShared) {
        try { Connect-EXO; Set-Mailbox -Identity $e.Upn -Type Regular -ErrorAction Stop; Write-Host "  mailbox converted back to normal" -ForegroundColor Green }
        catch { Write-Host "  mailbox convert-back failed (try again in 15 min, license needs to settle): $($_.Exception.Message)" -ForegroundColor Yellow }
    }

    $ok = 0; $bad = 0
    foreach ($g in @($e.Groups)) {
        try { New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $e.UserId -ErrorAction Stop; $ok++ }
        catch {
            try { Connect-EXO; Add-DistributionGroupMember -Identity $g.Id -Member $e.Upn -ErrorAction Stop; $ok++ }
            catch { $bad++ }
        }
    }
    Write-Host "  re-added to $ok group(s)$(if ($bad -gt 0) { ", $bad failed (dynamic groups re-add themselves)" })" -ForegroundColor Green

    try { Connect-EXO; Set-Mailbox -Identity $e.Upn -HiddenFromAddressListsEnabled $false -ErrorAction Stop; Write-Host "  visible in address book again" -ForegroundColor Green } catch { }
    Write-Host "Restore finished." -ForegroundColor Green
    Write-Log "RESTORE: $($e.Upn) ($($e.DisplayName)) re-enabled, $(@($e.Licenses).Count) license(s), $ok group(s)"
    Show-Info "$($e.DisplayName) has been restored. Check the console window for any warnings." 'Restore finished'
}

# ---------- Individual account mode ----------
function Start-IndividualMode {
    while ($true) {
        $q = Show-Input "Enter the person's email address or full name.`n`nLeave blank (or Cancel) to go back." 'Find account'
        if (-not $q) { return }
        $q = $q.Trim()

        $props = "Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,AssignedLicenses,CreatedDateTime,SignInActivity,OnPremisesSyncEnabled"
        $found = @()
        if ($q -like '*@*') {
            $u1 = Get-MgUser -UserId $q -Property $props -ErrorAction SilentlyContinue
            if ($u1) { $found = @($u1) } else { $found = @(Get-MgUser -Filter "mail eq '$q'" -Property $props) }
        } else {
            $safe = $q.Replace("'","''")
            $found = @(Get-MgUser -Filter "displayName eq '$safe'" -Property $props)
            if ($found.Count -eq 0) { $found = @(Get-MgUser -Search ("displayName:" + $safe) -ConsistencyLevel eventual -Property $props) }
        }

        if ($found.Count -eq 0) { Show-Info "No account found for '$q'." 'Not found'; continue }
        if ($found.Count -gt 1) {
            $found = @($found | Select-Object -First 10)
            $labels = @($found | ForEach-Object { "$($_.DisplayName)   <$($_.UserPrincipalName)>" })
            $pick = Show-Menu 'Multiple matches' 'Which account do you mean?' ($labels + 'Cancel')
            if ($pick -lt 0 -or $pick -ge $found.Count) { continue }
            $found = @($found[$pick])
        }
        $u = $found[0]

        $act = $u.SignInActivity
        $lastSignIn = @($act.LastSignInDateTime, $act.LastNonInteractiveSignInDateTime) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
        $seen = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'never' }
        $sec = Get-SignInSecurity $u.Id

        $secLine = if ($sec) { "`nSign-in check:  $sec" } else { '' }
        $mgLine = if ($u.OnPremisesSyncEnabled -eq $true) { "`nManaged in:  ON-PREMISES AD (disable/hide/groups must be done in local AD)" } else { "`nManaged in:  Cloud (fully manageable here)" }
        $info = "Name:  $($u.DisplayName)`nAccount:  $($u.UserPrincipalName)`nEnabled:  $($u.AccountEnabled)`nCreated:  $(if ($u.CreatedDateTime) { $u.CreatedDateTime.ToString('yyyy-MM-dd') } else { 'unknown' })`nLast sign-in:  $seen$mgLine`nLicenses:  $(Get-LicNames $u.AssignedLicenses.SkuId)$secLine"

        $c = Show-Menu 'Account found' $info @(
            'Full cleanup  -  disable + groups + hide + licenses',
            'Disable only  -  block sign-in, touch nothing else',
            'Re-enable  -  turn the account back on',
            'Cancel')
        switch ($c) {
            0 {
                $conv = Show-YesNo "Is this a STAFF member?`n`nYes = convert their mailbox to shared first, so email is kept forever (recommended for staff)." 'Staff member?'
                $del = ''; $arch = ''
                if ($conv) {
                    $del = "$(Show-Input "Optional: give a manager/delegate FULL ACCESS to the shared mailbox.`n`nEnter their email, or leave blank to skip." 'Delegate access')".Trim()
                    $arch = "$(Show-Input "OneDrive files will be auto-copied to this archive site.`n`nLeave as-is to use it, change it, or CLEAR the box to skip." 'OneDrive archive' $script:cfg.ArchiveSiteUrl)".Trim()
                    if ($arch) { $script:cfg.ArchiveSiteUrl = $arch; Save-Config }
                }
                if (Show-Confirm "Full cleanup of:`n`n$($u.DisplayName)  <$($u.UserPrincipalName)>`n`nThis makes REAL changes. An undo file is saved first. Continue?" 'Confirm cleanup') {
                    $r = Invoke-AccountCleanup -User $u -DoGroups (-not $NoGroupCleanup) -DoHide (-not $NoHideFromAddressBook) -ConvertToShared $conv -Delegate $del -ArchiveUrl $arch
                    Write-Host "Done: $r" -ForegroundColor Green
                    Save-RestoreLog
                    Show-Info "Done.`n`n$r" 'Cleanup complete'
                } else { Write-Host "Cancelled." }
            }
            1 {
                Update-MgUser -UserId $u.Id -AccountEnabled:$false
                Show-Info "$($u.DisplayName) is now disabled. Nothing else was touched." 'Done'
            }
            2 {
                Update-MgUser -UserId $u.Id -AccountEnabled:$true
                Show-Info "$($u.DisplayName) is enabled again." 'Done'
            }
            default { }
        }
    }
}

# ---------- Bulk mode ----------
function Start-BulkMode {
    param([string]$Path, [bool]$DoCommit, [bool]$SecCheck = $true)

    # ---- Load and validate the list FIRST (before any slow downloads) ----
    if (-not $Path -or -not (Test-Path $Path -PathType Leaf)) {
        Write-Host "That is not a file. Please pick your CSV or Excel list." -ForegroundColor Yellow
        return
    }
    if ($Path -like '*.xlsx') {
        Import-Module ImportExcel
        # Some exports have a title row above the real headers (e.g. "Leavers Staff List - 2025"
        # as row 1, with the actual "Name / Email / ..." headers on row 2). Peek at the first few
        # rows with no header assumed, and use whichever one actually looks like column headers -
        # multiple short text cells, ideally matching things like email/name/leaving date.
        $headerRow = 1
        try {
            $rawPeek = @(Import-Excel $Path -NoHeader -StartRow 1 -ErrorAction Stop | Select-Object -First 10)
            for ($ri = 0; $ri -lt $rawPeek.Count; $ri++) {
                $cells = @($rawPeek[$ri].psobject.Properties | ForEach-Object { "$($_.Value)".Trim() } | Where-Object { $_ })
                if ($cells.Count -ge 2 -and -not ($cells | Where-Object { $_.Length -gt 40 })) {
                    $keywordHit = @($cells | Where-Object { $_ -match 'e-?mail|name|leav|left|exit|staff|student|pupil|forename|surname' }).Count -gt 0
                    if ($keywordHit -or $cells.Count -ge 3) { $headerRow = $ri + 1; break }
                }
            }
        } catch { $headerRow = 1 }   # peek failed for any reason - fall back to the normal "row 1 is the header" behaviour
        $rows = @(Import-Excel $Path -StartRow $headerRow)
        if ($headerRow -gt 1) { Write-Host "Note: row 1 looked like a title, not headers - used row $headerRow as the real header row instead." -ForegroundColor Cyan }
    } else {
        $rows = @(Import-Csv $Path)
    }
    if ($rows.Count -eq 0) {
        Write-Host "The file has no rows. Nothing to do." -ForegroundColor Yellow
        return
    }
    Write-Host "Loaded $($rows.Count) rows from $(Split-Path $Path -Leaf)."
    Write-Log "BULK $(if ($DoCommit) { 'COMMIT' } else { 'DRY-RUN' }) [$script:ListType] started from '$(Split-Path $Path -Leaf)' ($($rows.Count) rows)"

    # ---- Auto-detect which columns are which, whatever the headers are called ----
    $headers = @($rows[0].psobject.Properties.Name)

    # Email column: header mentioning email, else the column where most values contain '@'
    $emailCol = $headers | Where-Object { $_ -match 'e-?mail' } | Select-Object -First 1
    if (-not $emailCol) {
        foreach ($h in $headers) {
            $sample = @($rows | ForEach-Object { $_.$h } | Where-Object { $_ } | Select-Object -First 20)
            if ($sample.Count -ge 3 -and @($sample | Where-Object { "$_" -like '*@*' }).Count -ge [math]::Ceiling($sample.Count * 0.8)) { $emailCol = $h; break }
        }
    }

    # Name column: try the best candidates in order, else combine Forename + Surname
    $nameCol = $null
    foreach ($cand in 'Full Name','Name','Display Name','Pupil Name','Student Name','Preferred Name','Official Name') {
        $hit = $headers | Where-Object { $_ -ieq $cand } | Select-Object -First 1
        if ($hit) { $nameCol = $hit; break }
    }
    if (-not $nameCol) { $nameCol = $headers | Where-Object { $_ -match 'name' -and $_ -notmatch 'fore|sur|first|last|user|family|given' } | Select-Object -First 1 }
    $foreCol = $headers | Where-Object { $_ -match 'forename|first ?name|given' } | Select-Object -First 1
    $surCol  = $headers | Where-Object { $_ -match 'surname|last ?name|family' }  | Select-Object -First 1

    # Leaving date column
    $dateCol = $headers | Where-Object { $_ -match 'leav|left|exit|end ?date|withdraw' } | Select-Object -First 1

    $nameDesc = if ($nameCol) { $nameCol } elseif ($foreCol -and $surCol) { "$foreCol + $surCol" } else { '(none)' }
    Write-Host "Detected columns -> Email: $(if ($emailCol) { $emailCol } else { '(none)' })  |  Name: $nameDesc  |  Leaving date: $(if ($dateCol) { $dateCol } else { '(none)' })" -ForegroundColor Cyan
    if (-not $emailCol -and -not $nameCol -and -not ($foreCol -and $surCol)) {
        Write-Host "Could not find an email or name column in this file - cannot match anyone. Check the file." -ForegroundColor Yellow
        return
    }
    if (-not $dateCol) {
        Write-Host "WARNING: no leaving-date column found. The recycled-address safety check will be off for this run." -ForegroundColor Yellow
    }

    # ---- Download all users, or reuse a recent cache (max 4 hours old) ----
    $allUsers = Get-AllTenantUsers
    Write-Host "$($allUsers.Count) users loaded. Building lookup tables..."

    $byUpn = @{}; $byMail = @{}; $byProxy = @{}; $byName = @{}
    foreach ($u in $allUsers) {
        if ($u.UserPrincipalName) { $byUpn[$u.UserPrincipalName.ToLower()] = $u }
        if ($u.Mail) {
            $k = $u.Mail.ToLower()
            if (-not $byMail.ContainsKey($k)) { $byMail[$k] = New-Object System.Collections.ArrayList }
            [void]$byMail[$k].Add($u)
        }
        foreach ($p in @($u.ProxyAddresses)) {
            if ($p -match '^smtp:(.+)$') {
                $k = $Matches[1].ToLower()
                if (-not $byProxy.ContainsKey($k)) { $byProxy[$k] = New-Object System.Collections.ArrayList }
                [void]$byProxy[$k].Add($u)
            }
        }
        if ($u.DisplayName) {
            $k = $u.DisplayName.ToLower().Trim()
            if (-not $byName.ContainsKey($k)) { $byName[$k] = New-Object System.Collections.ArrayList }
            [void]$byName[$k].Add($u)
        }
    }

    function Find-User {
        param($Email, $Name)
        if ($Email) {
            $k = $Email.ToLower()
            if ($byUpn.ContainsKey($k))   { return ,@($byUpn[$k]) }
            if ($byMail.ContainsKey($k))  { return ,@($byMail[$k]  | Select-Object -Unique) }
            if ($byProxy.ContainsKey($k)) { return ,@($byProxy[$k] | Select-Object -Unique) }
            # email not found (blank/wrong in iSAMS) - fall back to matching by name
            if ($Name) {
                $k2 = $Name.ToLower().Trim()
                if ($byName.ContainsKey($k2)) { return ,@($byName[$k2] | Select-Object -Unique) }
            }
            return ,@()
        }
        elseif ($Name) {
            $k = $Name.ToLower().Trim()
            if ($byName.ContainsKey($k)) { return ,@($byName[$k] | Select-Object -Unique) }
            return ,@()
        }
        return ,@()
    }

    function New-Result {
        param($Pupil, $ListName, $Matched, $AccountName, $Left, $Created, $LastSignIn, $Action, $Remarks, $Licenses, $SignInCheck = '')
        # Column label follows the mode you chose (Staff vs Student) instead of always saying "Pupil"
        $who = if ($script:ListType -eq 'Staff') { 'Staff' } else { 'Student' }
        $o = [ordered]@{}
        $o["$who (from your list)"]        = $Pupil
        $o['Name (from your list)']        = $ListName
        $o['Matched Microsoft Account']    = $Matched
        $o['Account Display Name']         = $AccountName
        $o['Leaving Date']                 = if ($Left -is [datetime]) { $Left.ToString('yyyy-MM-dd') } elseif ($Left) { "$Left" } else { '' }
        $o['Account Created']              = if ($Created) { $Created.ToString('yyyy-MM-dd') } else { '' }
        $o['Last Sign-In']                 = $LastSignIn
        $o['Sign-In Check']                = $SignInCheck
        $o['Action']                       = $Action
        $o['Remarks (why)']                = $Remarks
        $o['Licenses']                     = $Licenses
        [pscustomobject]$o
    }

    # ---- Pre-fetch sign-in security checks in batches (big speed-up for large lists) ----
    # Does a quick first pass just to work out who in the list matches a real account, then asks
    # Microsoft Graph for all of their sign-in histories a batch at a time (up to 20 per call)
    # instead of one-by-one. The main loop below uses these pre-fetched results if available,
    # and falls back to the normal one-by-one check for anyone missed - so this is a pure speed
    # optimisation and cannot change what gets flagged or skipped.
    $script:secBatchMap = @{}
    if ($SecCheck) {
        Write-Host "Pre-checking sign-in security for your list (batched - much faster than one-by-one)..." -ForegroundColor Gray
        $candidateIds = New-Object System.Collections.Generic.List[string]
        foreach ($preRow in $rows) {
            $preEmail = if ($emailCol -and $preRow.$emailCol) { "$($preRow.$emailCol)".Trim() } else { '' }
            $preName  = ''
            if ($nameCol -and $preRow.$nameCol) { $preName = "$($preRow.$nameCol)".Trim() }
            elseif ($foreCol -and $surCol)      { $preName = ("$($preRow.$foreCol) $($preRow.$surCol)").Trim() }
            if (-not $preEmail -and -not $preName) { continue }
            $preFound = Find-User -Email $preEmail -Name $preName
            if ($preFound.Count -eq 1 -and $preFound[0].Id) { $candidateIds.Add($preFound[0].Id) }
        }
        if ($candidateIds.Count -gt 0) {
            $script:secBatchMap = Get-SignInSecurityBatch -UserIds $candidateIds
            Write-Host "Batched sign-in check complete for $($script:secBatchMap.Count) of $($candidateIds.Count) account(s) (any remainder uses the normal per-account check automatically)." -ForegroundColor Gray
        }
    }

    $results = New-Object System.Collections.Generic.List[object]
    $reviewList = New-Object System.Collections.Generic.List[object]
    $toProcess  = New-Object System.Collections.Generic.List[object]
    $i = 0

    foreach ($row in $rows) {
        $i++
        $rawEmail = if ($emailCol -and $row.$emailCol) { $row.$emailCol } else { '' }
        $rawName  = ''
        if ($nameCol -and $row.$nameCol) { $rawName = $row.$nameCol }
        elseif ($foreCol -and $surCol)   { $rawName = ("$($row.$foreCol) $($row.$surCol)").Trim() }
        $rawLeft  = if ($dateCol -and $row.$dateCol) { $row.$dateCol } else { '' }

        $email = if ($rawEmail) { "$rawEmail".Trim() } else { '' }
        $name  = if ($rawName)  { "$rawName".Trim()  } else { '' }
        $left  = $null
        if ($rawLeft -is [datetime]) { $left = $rawLeft }   # Excel files give real dates directly
        elseif ($rawLeft -and "$rawLeft".Trim() -match '^\d{4,5}(\.\d+)?$') {
            # Excel serial date number (days since 1900), e.g. 41088 = 2012-06-28
            $n = [double]"$rawLeft".Trim()
            if ($n -gt 20000 -and $n -lt 80000) { $left = [datetime]::FromOADate($n) }
        }
        elseif ($rawLeft) {
            try { $left = [datetime]"$rawLeft".Trim() } catch { $left = "$rawLeft".Trim() }
        }

        $key = if ($email) { $email } else { $name }
        if ($i % 10 -eq 0 -or $i -eq $rows.Count) {
            Write-Progress -Activity "Checking your list against Microsoft 365" -Status "Row $i of $($rows.Count)" -PercentComplete ([math]::Min(100, $i / $rows.Count * 100))
        }

        if (-not $key) {
            $results.Add((New-Result '' '' '' '' $left $null '' 'SKIP - empty row' 'Row in your CSV had no email and no name.' ''))
            continue
        }

        $found = Find-User -Email $email -Name $name
        if ($found.Count -eq 0) {
            $results.Add((New-Result $key $name '' '' $left $null '' 'NOT FOUND' 'No Microsoft account matches this person. Nothing to do (likely left before accounts existed or already deleted).' ''))
            continue
        }
        if ($found.Count -gt 1) {
            $upns  = ($found.UserPrincipalName -join '; ')
            $names = ($found.DisplayName -join '; ')
            $results.Add((New-Result $key $name $upns $names $left $null '' 'REVIEW - ambiguous' "More than one account matches this name ($upns). Not touched - pick the right one manually." ''))
            continue
        }

        $u    = $found[0]
        $skus = @($u.AssignedLicenses.SkuId | Where-Object { $_ })
        $lics = Get-LicNames $skus
        $dn   = $u.DisplayName

        $act = $u.SignInActivity
        $lastSignIn = @($act.LastSignInDateTime, $act.LastNonInteractiveSignInDateTime) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
        $seen = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'never' }

        $inactiveTxt = 'has never signed in'
        if ($lastSignIn) {
            $days = [int]((Get-Date).ToUniversalTime() - $lastSignIn).TotalDays
            if ($days -ge 365) { $inactiveTxt = "no sign-in for $([math]::Round($days/365.0,1)) years" }
            else               { $inactiveTxt = "no sign-in for $days days" }
        }

        $sec = ''
        $sameName = Test-SameName $name $dn
        $note = ''
        if ($u.OnPremisesSyncEnabled -eq $true) { $note += " [ON-PREMISES account: disable/hide/groups must be done in local AD; cloud does licence/mailbox/OneDrive]" }

        # Was this an email row that actually got matched by NAME (email missing/wrong in iSAMS)?
        $emailMatched = $false
        if ($email) {
            $ek = $email.ToLower()
            $proxies = @($u.ProxyAddresses | ForEach-Object { if ($_ -match '^smtp:(.+)$') { $Matches[1].ToLower() } })
            $emailMatched = ($u.UserPrincipalName -and $u.UserPrincipalName.ToLower() -eq $ek) -or ($u.Mail -and $u.Mail.ToLower() -eq $ek) -or ($proxies -contains $ek)
            if (-not $emailMatched) { $note += " Note: the email in your list was not found - matched by NAME instead." }
        }

        # ---- Your explicit Staff/Student choice is authoritative - no auto-skip based on account type ----
        # (Account-type mismatches are recorded as an informational note only; nothing is refused.)
        $looksStaff = Test-LooksLikeStaff $u $lics
        if ($script:ListType -eq 'Students' -and $looksStaff) {
            $why = @()
            if ($lics -match 'Faculty') { $why += 'faculty licence' }
            if ($u.JobTitle) { $why += "job title '$($u.JobTitle)'" }
            if ($u.Department) { $why += "department '$($u.Department)'" }
            $note += " Note: this account looks like staff ($($why -join ', ')) but you ran it as a STUDENT - processing as instructed."
        }
        elseif ($script:ListType -eq 'Staff' -and -not $looksStaff) {
            $note += " Note: account has no job title/faculty licence but you ran it as STAFF - processing as instructed."
        }

        # ---- FAILSAFE: account created AFTER the person left = recycled address (only when the name differs) ----
        # In "clean everything in the list" mode these holds are bypassed - the list is your approval.
        if ($left -is [datetime] -and $u.CreatedDateTime -and $u.CreatedDateTime -gt $left) {
            if ($sameName) {
                $note = " Note: account created after leaving date, but the name matches your leaver - treated as the same person."
            } elseif (-not $script:ForceClean) {
                $results.Add((New-Result $key $name $u.UserPrincipalName $dn $left $u.CreatedDateTime $seen 'SKIP - new account' "Recycled address: YOUR leaver was '$name' (left $($left.ToString('yyyy-MM-dd'))) but this account was created later and now belongs to '$dn' - a different person. Not touched." $lics))
                continue
            } else {
                $note += " Note: account created after the leaving date and the name differs - cleaned anyway because it was in your approved list."
            }
        }

        # ---- FAILSAFE: global created-after cutoff ----
        if (-not $script:ForceClean -and $SkipIfCreatedAfter -and $u.CreatedDateTime -and $u.CreatedDateTime -gt $SkipIfCreatedAfter -and -not $sameName) {
            $results.Add((New-Result $key $name $u.UserPrincipalName $dn $left $u.CreatedDateTime $seen 'SKIP - new account' "Account created after your cutoff date ($($SkipIfCreatedAfter.ToString('yyyy-MM-dd'))) and belongs to '$dn' - likely a NEW student with a recycled address. Not touched." $lics))
            continue
        }

        # ---- FAILSAFE: name cross-check ----
        if (-not $script:ForceClean -and $name -and $email -and -not $sameName) {
            $results.Add((New-Result $key $name $u.UserPrincipalName $dn $left $u.CreatedDateTime $seen 'REVIEW - name mismatch' "Your list says '$name' but the account is named '$dn' - names genuinely differ. Not touched - check if same person, handle manually." $lics))
            if ($DoCommit) { $reviewList.Add(@{ Idx = $results.Count - 1; User = $u; Key = $key; Name = $name; Left = $left; Seen = $seen; Lics = $lics; Sec = $sec; Inactive = $inactiveTxt }) }
            continue
        }
        elseif ($script:ForceClean -and $name -and $email -and -not $sameName) {
            $note += " Note: list name '$name' differs from account name '$dn' - cleaned anyway because it was in your approved list."
        }

        # Sign-in security check (logs only go back ~30 days). Uses the pre-fetched batch result
        # if we have one (fast); otherwise falls back to the normal one-by-one live check.
        if ($SecCheck -and $lastSignIn -and $lastSignIn -gt (Get-Date).ToUniversalTime().AddDays(-30)) {
            if ($script:secBatchMap.ContainsKey($u.Id)) {
                $sec = $script:secBatchMap[$u.Id]
            } else {
                $sec = Get-SignInSecurity $u.Id
            }
        }

        # ---- FAILSAFE: recently active ----
        if (-not $script:ForceClean -and $SkipIfActiveWithinDays -gt 0 -and $lastSignIn -and $lastSignIn -gt $activeCutoff) {
            $results.Add((New-Result $key $name $u.UserPrincipalName $dn $left $u.CreatedDateTime $seen 'SKIP - recently active' "Account name is '$dn'. Someone signed in within the last $SkipIfActiveWithinDays days (last: $seen). Might still be in use - not touched. Check manually if they really left." $lics -SignInCheck $sec))
            if ($DoCommit) { $reviewList.Add(@{ Idx = $results.Count - 1; User = $u; Key = $key; Name = $name; Left = $left; Seen = $seen; Lics = $lics; Sec = $sec; Inactive = $inactiveTxt }) }
            continue
        }
        elseif ($script:ForceClean -and $SkipIfActiveWithinDays -gt 0 -and $lastSignIn -and $lastSignIn -gt $activeCutoff) {
            $note += " Note: signed in within the last $SkipIfActiveWithinDays days (last: $seen) - cleaned anyway because it was in your approved list."
        }

        # ---- FAST-SKIP: already disabled AND no licence = already done, nothing to do (saves time) ----
        if (-not $u.AccountEnabled -and $skus.Count -eq 0) {
            $results.Add((New-Result $key $name $u.UserPrincipalName $dn $left $u.CreatedDateTime $seen 'ALREADY DONE' "Account is already disabled and has no licence - nothing to do, skipped for speed.$note" $lics -SignInCheck $sec))
            continue
        }

        if (-not $DoCommit) {
            $results.Add((New-Result $key $name $u.UserPrincipalName $dn $left $u.CreatedDateTime $seen 'WILL PROCESS' "Passed all safety checks ($inactiveTxt - consistent with having left). On commit: disable + $(if ($script:StaffMode) { 'convert mailbox to SHARED + ' })remove groups + hide from address book + remove $($skus.Count) license(s).$note" $lics -SignInCheck $sec))
            continue
        }

        # COMMIT: don't act yet - queue it. Nothing is disabled until after you've decided the borderline cases.
        $results.Add((New-Result $key $name $u.UserPrincipalName $dn $left $u.CreatedDateTime $seen 'QUEUED' "Passed all checks - queued for cleanup.$note" $lics -SignInCheck $sec))
        $toProcess.Add(@{ Idx = $results.Count - 1; User = $u; Key = $key; Name = $name; Left = $left; Seen = $seen; Lics = $lics; Sec = $sec; Inactive = $inactiveTxt; FromReview = $false })
    }

    # ---- Clean up old result files, then save ----
    $outDir = Split-Path (Resolve-Path $Path) -Parent
    Get-ChildItem -Path $outDir -File | Where-Object { $_.Name -match '^result_\d{8}_\d{4}\.(csv|xlsx)$' } | ForEach-Object {
        try { Remove-Item $_.FullName -Force; Write-Host "Deleted old result file: $($_.Name)" } catch { }
    }

    Write-Progress -Activity "Checking your list against Microsoft 365" -Completed

    # ================= COMMIT: decide borderline cases FIRST, then do all the work =================
    if ($DoCommit) {
        # Adds a reviewed account to the processing queue (marks its row as queued-from-review)
        $approve = {
            param($c)
            $results[$c.Idx] = New-Result $c.Key $c.Name $c.User.UserPrincipalName $c.User.DisplayName $c.Left $c.User.CreatedDateTime $c.Seen 'QUEUED' "You approved this during review - queued for cleanup." $c.Lics -SignInCheck $c.Sec
            $c.FromReview = $true
            $toProcess.Add($c)
        }

        # ---- Review the skipped/borderline accounts BEFORE anything is disabled ----
        if ($reviewList.Count -gt 0) {
            $byReason = $reviewList | Group-Object { ($results[$_.Idx].'Action') } | ForEach-Object { "$($_.Count) x $($_.Name)" }
            $choice = Show-Menu 'Skipped accounts - decide before anything is disabled' (
                "$($reviewList.Count) account(s) were held back by the safety checks:`n  $($byReason -join "`n  ")`n`nNothing has been disabled yet. Decide these first - then everything you approve is processed together with the clear-pass accounts.") @(
                'Review one by one (decide each)',
                "APPROVE ALL $($reviewList.Count) (I've checked them - add to the cleanup)",
                'Leave them all skipped')

            if ($choice -eq 1) {
                if (Show-Confirm "Approve all $($reviewList.Count) held-back account(s) for cleanup?`n`nThis overrides the name-mismatch / recently-active / looks-like-staff checks. Only do this if you're sure they have all left.`n`nProceed?" 'Confirm approve-all') {
                    foreach ($c in $reviewList) { & $approve $c }
                }
            }
            elseif ($choice -eq 0) {
                for ($ri = 0; $ri -lt $reviewList.Count; $ri++) {
                    $c = $reviewList[$ri]
                    $row = $results[$c.Idx]
                    $msg = "$($row.'Action')`n`nList entry:  $($c.Key)`nAccount:  $($c.User.DisplayName)  <$($c.User.UserPrincipalName)>`nLast sign-in:  $($c.Seen)`nLicenses:  $($c.Lics)$(if ($c.Sec) { "`nSign-in check:  $($c.Sec)" })`n`n$($row.'Remarks (why)')`n`n($($ri + 1) of $($reviewList.Count))"
                    $a = Show-Menu 'Skipped account - your decision' $msg @('APPROVE for cleanup', 'Leave it skipped', 'APPROVE this and ALL remaining', 'Stop reviewing')
                    if ($a -eq 3 -or $a -lt 0) { break }
                    if ($a -eq 1) { continue }
                    if ($a -eq 2) { for ($rj = $ri; $rj -lt $reviewList.Count; $rj++) { & $approve $reviewList[$rj] }; break }
                    & $approve $c
                }
            }
        }

        # ---- Now do ALL the disabling in one pass (clear-pass + approved) ----
        if ($toProcess.Count -gt 0) {
            Write-Host "`nApplying cleanup to $($toProcess.Count) account(s)..." -ForegroundColor Cyan
            $tp = 0
            foreach ($c in $toProcess) {
                $tp++
                Write-Progress -Activity "Applying cleanup" -Status "$tp of $($toProcess.Count): $($c.Key)" -PercentComplete ($tp / $toProcess.Count * 100)
                $tag = if ($c.FromReview) { ' [approved in review]' } else { '' }
                try {
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    Write-Host "`n[$tp/$($toProcess.Count)] Processing $($c.Key) ($($c.User.DisplayName))${tag}:" -ForegroundColor White
                    $r = Invoke-AccountCleanup -User $c.User -DoGroups (-not $NoGroupCleanup) -DoHide (-not $NoHideFromAddressBook) -ConvertToShared ([bool]$script:StaffMode) -Delegate "$script:DelegateEmail" -ArchiveUrl $(if ($script:StaffMode) { $script:cfg.ArchiveSiteUrl } else { '' })
                    Write-Host "  -> done in $([int]$sw.Elapsed.TotalSeconds)s" -ForegroundColor Green
                    $st = if ($c.FromReview) { 'DONE (approved in review)' } else { 'DONE' }
                    $results[$c.Idx] = New-Result $c.Key $c.Name $c.User.UserPrincipalName $c.User.DisplayName $c.Left $c.User.CreatedDateTime $c.Seen $st "$r ($($c.Inactive))." $c.Lics -SignInCheck $c.Sec
                } catch {
                    Write-Host "  -> FAILED" -ForegroundColor Red
                    $results[$c.Idx] = New-Result $c.Key $c.Name $c.User.UserPrincipalName $c.User.DisplayName $c.Left $c.User.CreatedDateTime $c.Seen 'ERROR' "Failed: $($_.Exception.Message)" $c.Lics -SignInCheck $c.Sec
                }
            }
            Write-Progress -Activity "Applying cleanup" -Completed
        }
    }

    if ($DoCommit) { Save-RestoreLog }

    # Reports are Excel-only (no CSV). Commit reports get a COMMITTED_ prefix so the
    # auto-cleanup above can never delete your audit trail.
    $prefix = if ($DoCommit) { 'COMMITTED_result_' } else { 'result_' }
    $xlsx = Join-Path $outDir ("$prefix{0:yyyyMMdd_HHmm}.xlsx" -f (Get-Date))
    $log  = $xlsx
    $results | Group-Object 'Action' | Select-Object Name, Count | Sort-Object Name | Format-Table -AutoSize

    if ($results.Count -gt 0 -and (Get-Module -ListAvailable ImportExcel)) {
        Write-Host "Building the formatted Excel report..." -ForegroundColor Gray
        Write-Progress -Activity "Finishing" -Status "Creating the Excel report..." -PercentComplete 50
        Import-Module ImportExcel

        $summary = $results | Group-Object 'Action' | Sort-Object Name | ForEach-Object {
            [pscustomobject]@{ 'Category' = $_.Name; 'Accounts' = $_.Count }
        }
        $summary | Export-Excel $xlsx -WorksheetName 'Summary' -Title 'Leaver Cleanup Summary' -TableStyle Medium2 -AutoSize

        $licCount = @{}
        foreach ($r in $results) {
            if (($r.'Action' -eq 'WILL PROCESS' -or $r.'Action' -eq 'DONE') -and $r.'Licenses') {
                foreach ($l in ($r.'Licenses' -split ',\s*')) {
                    if ($l) { if ($licCount.ContainsKey($l)) { $licCount[$l]++ } else { $licCount[$l] = 1 } }
                }
            }
        }
        if ($licCount.Count -gt 0) {
            $licRows = $licCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
                [pscustomobject]@{ 'License' = $_.Key; 'Count' = $_.Value }
            }
            $licRows | Export-Excel $xlsx -WorksheetName 'Licenses to recover' -TableStyle Medium2 -AutoSize
        }

        # ---- Separate sheet: hybrid (on-premises synced) accounts that must be fixed in LOCAL AD ----
        $hybrid = @($results | Where-Object { "$($_.'Remarks (why)')" -match 'ON-PREMISES|local AD|on-premises' })
        if ($hybrid.Count -gt 0) {
            $hybrid | ForEach-Object {
                [pscustomobject]@{
                    'Name (from your list)'     = $_.'Name (from your list)'
                    'Matched Microsoft Account' = $_.'Matched Microsoft Account'
                    'Account Display Name'      = $_.'Account Display Name'
                    'Cloud action taken'        = $_.'Action'
                    'Reason'                    = 'FAILED / INCOMPLETE - hybrid account synced from on-premises AD. Disable, hide and group changes MUST be done in your LOCAL Active Directory (cloud-only changes revert at the next sync).'
                }
            } | Export-Excel $xlsx -WorksheetName 'Fix in Local AD' -TableStyle Medium6 -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
            Write-Host "$($hybrid.Count) hybrid account(s) listed in the 'Fix in Local AD' sheet." -ForegroundColor Yellow
        }

        $ct = @(
            New-ConditionalText -Text 'WILL PROCESS' -ConditionalTextColor Black -BackgroundColor LightGreen
            New-ConditionalText -Text 'DONE'         -ConditionalTextColor Black -BackgroundColor LightGreen
            New-ConditionalText -Text 'REVIEW'       -ConditionalTextColor Black -BackgroundColor LightPink
            New-ConditionalText -Text 'ERROR'        -ConditionalTextColor White -BackgroundColor Red
            New-ConditionalText -Text 'ALERT'        -ConditionalTextColor White -BackgroundColor Red
            New-ConditionalText -Text 'SKIP'         -ConditionalTextColor Black -BackgroundColor Khaki
        )
        $pkg = $results | Sort-Object 'Action' | Export-Excel $xlsx -WorksheetName 'Details' -AutoFilter -FreezeTopRow -BoldTopRow -ConditionalText $ct -PassThru
        $ws = $pkg.Workbook.Worksheets['Details']
        $widths = @(34, 28, 34, 28, 13, 15, 12, 40, 24, 90, 38)
        for ($c = 1; $c -le $widths.Count; $c++) { $ws.Column($c).Width = $widths[$c - 1] }
        $ws.Cells["J2:J$($ws.Dimension.End.Row)"].Style.WrapText = $true
        Close-ExcelPackage $pkg
        Write-Progress -Activity "Finishing" -Completed
        Write-Host "Formatted Excel report saved to: $xlsx" -ForegroundColor Green
    }
    elseif ($results.Count -gt 0) {
        # Emergency fallback only if the Excel module is somehow unavailable - keeps the audit trail
        $emergency = [System.IO.Path]::ChangeExtension($xlsx, '.csv')
        $results | Sort-Object 'Action' | Export-Csv $emergency -NoTypeInformation -Encoding UTF8
        Write-Host "ImportExcel not available - saved an emergency CSV instead: $emergency" -ForegroundColor Yellow
        $log = $emergency
    }

    $counts = ($results | Group-Object 'Action' | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
    Write-Log "BULK $(if ($DoCommit) { 'COMMIT' } else { 'DRY-RUN' }) finished. Results: $counts. Report: $(Split-Path $log -Leaf)"
    Write-Host "`nLog saved to: $log"
    if (-not $DoCommit) { Write-Host "This was a DRY RUN. Re-run with -Commit (or choose commit in the menu) to apply changes." -ForegroundColor Yellow }

    # ---- Headline summary popup ----
    $grp = @{}; foreach ($r in $results) { $a = "$($r.Action)".Split(' ')[0]; if ($a -eq 'DONE') { $a = 'DONE' }; if ($grp.ContainsKey($r.Action)) { $grp[$r.Action]++ } else { $grp[$r.Action] = 1 } }
    $licTotal = 0
    foreach ($r in $results) {
        if (($r.Action -eq 'WILL PROCESS' -or $r.Action -like 'DONE*') -and $r.Licenses) { $licTotal += @($r.Licenses -split ',\s*' | Where-Object { $_ }).Count }
    }
    $lines = ($grp.GetEnumerator() | Sort-Object Name | ForEach-Object { "  {0,-28} {1}" -f $_.Name, $_.Value }) -join "`n"
    $kind = if ($DoCommit) { 'COMMIT (real changes made)' } else { 'DRY RUN (nothing changed)' }
    $verb = if ($DoCommit) { 'Licenses removed' } else { 'Licenses that would be recovered' }
    $hybridNote = if ($hybrid -and $hybrid.Count -gt 0) { "`n`n$($hybrid.Count) hybrid account(s) need fixing in LOCAL AD - see the 'Fix in Local AD' sheet." } else { '' }
    Show-Info ("$kind`n`nList: $(Split-Path $Path -Leaf)   ($($rows.Count) rows)`n`n$lines`n`n$verb`: $licTotal$hybridNote`n`nFull Excel report: $(Split-Path $log -Leaf)") "Cleanup summary - $kind"
}

# ---------- Entry point ----------
if ($CsvPath) {
    Start-BulkMode -Path $CsvPath -DoCommit ([bool]$Commit) -SecCheck (-not $SkipSecurityCheck)
}
else {
    while ($true) {
        $choice = Show-Menu 'Leaver Cleanup Tool' 'What would you like to do?' @(
            'DRY RUN  -  preview a bulk cleanup (safe, changes nothing)',
            'COMMIT  -  run a bulk cleanup for real',
            'Individual account  -  look up one person and decide',
            'RESTORE  -  undo a previous cleanup for someone',
            'PURGE stale accounts  -  delete accounts dead for years',
            'Exit')
        if ($choice -lt 0 -or $choice -eq 5) { return }

        if ($choice -eq 0 -or $choice -eq 1) {
            $t = Show-Menu 'Who is in this list?' "Choose who you are cleaning up.`n`nYour choice is respected - accounts are processed as instructed (mismatches are noted, not skipped).`nSTAFF runs preserve mailboxes (converted to shared) and archive OneDrive." @(
                'STUDENTS  -  standard cleanup',
                'STAFF  -  extra care: keep email + archive OneDrive')
            if ($t -lt 0) { continue }
            $script:ListType  = if ($t -eq 1) { 'Staff' } else { 'Students' }
            $script:StaffMode = ($t -eq 1)
            $script:DelegateEmail = ''

            $p = Select-InputFile
            if (-not $p) { continue }

            # Approval mode: is every account in this list pre-approved for a full clean?
            $script:ForceClean = Show-YesNo (
                "Treat EVERY account found in your list as approved for a full cleanup?`n`n" +
                "YES = clean every matched account the same way (Staff or Student), skipping the name-mismatch, recently-active and recycled-address holds. Use this when you're certain everyone in the list has left.`n`n" +
                "NO  = careful mode: hold those borderline cases back for you to review before anything changes.") 'Approve everyone in the list?'

            if ($choice -eq 0) {
                # ---- DRY RUN ----
                $sec = Show-YesNo "Include the sign-in security check?`n`nShows where each active account signed in from (home country or abroad) and failed login attempts. More thorough, but slower." 'Security check'
                Start-BulkMode -Path $p -DoCommit $false -SecCheck $sec
            }
            else {
                # ---- COMMIT ----
                if ($script:StaffMode) {
                    $del = Show-Input "Optional: give a manager/delegate FULL ACCESS to the converted shared mailboxes.`n`nEnter their email address, or leave blank to skip." 'Delegate access'
                    $script:DelegateEmail = "$del".Trim()
                    $au = Show-Input "OneDrive files will be auto-copied to this archive SharePoint site.`n`nLeave as-is to use it, change it to use another site, or CLEAR the box to skip archiving." 'OneDrive archive' $script:cfg.ArchiveSiteUrl
                    $script:cfg.ArchiveSiteUrl = "$au".Trim()
                    if ($script:cfg.ArchiveSiteUrl) { Save-Config }
                }
                $what = if ($script:StaffMode) { "STAFF cleanup:`n- disable sign-in`n- convert mailbox to shared (email kept forever)`n- archive OneDrive`n- remove groups + hide from address book`n- remove licenses" } else { "STUDENT cleanup:`n- disable sign-in`n- remove groups + hide from address book`n- remove licenses" }
                $modeLine = if ($script:ForceClean) { "`n`nMODE: CLEAN EVERYONE in the list (name-mismatch / recently-active / recycled-address holds are BYPASSED)." } else { "`n`nMODE: careful (borderline cases are held back for your review first)." }
                if (Show-Confirm "$what$modeLine`n`nList: $(Split-Path $p -Leaf)`n`nThis makes REAL changes. An undo file is saved first. Continue?" 'Confirm commit') {
                    Start-BulkMode -Path $p -DoCommit $true
                } else { Write-Host "Cancelled." }
            }
        }
        elseif ($choice -eq 2) { Start-IndividualMode }
        elseif ($choice -eq 3) { Start-RestoreMode }
        elseif ($choice -eq 4) { Start-PurgeMode }
    }
}

