<#
.SYNOPSIS
    One-liner launcher for the Microsoft 365 Leaver Cleanup Tool.

.DESCRIPTION
    Installs the tool into a PERMANENT folder (Desktop\Sherborne-Leaver-Cleanup-Tool) and
    launches it there. A permanent folder is used on purpose: this tool creates
    undo files (Restore_*.json), audit logs, and a config file that MUST persist
    between runs — running from a temporary folder would lose your safety net.

    Run with:
    irm https://raw.githubusercontent.com/xnostra/Sherborne-Leaver-Cleanup-Tool/main/invoke-leaver.ps1 | iex

.NOTES
    Version: 1.0
    LastModified: 2026-07-20

.LINK
    https://github.com/xnostra/Sherborne-Leaver-Cleanup-Tool
#>

$repoRaw    = "https://raw.githubusercontent.com/xnostra/Sherborne-Leaver-Cleanup-Tool/main"
$scriptUrl  = "$repoRaw/Disable-RemoveLicenses.ps1"

# Permanent install folder so undo files, logs, and config persist
$installDir = Join-Path ([Environment]::GetFolderPath('Desktop')) "Sherborne-Leaver-Cleanup-Tool"
$scriptPath = Join-Path $installDir "Disable-RemoveLicenses.ps1"

Write-Host "Microsoft 365 Leaver Cleanup Tool" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

try {
    if (-not (Test-Path $installDir)) {
        New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    }

    Write-Host "Installing/updating the tool in:" -ForegroundColor Yellow
    Write-Host "  $installDir" -ForegroundColor Gray
    Write-Host ""

    # Download without adding a second UTF-8 BOM. The GitHub source already has
    # one; piping Invoke-RestMethod into Out-File on Windows PowerShell adds
    # another and makes the opening <# documentation comment invalid syntax.
    $downloadedScript = (Invoke-WebRequest -Uri $scriptUrl -UseBasicParsing -ErrorAction Stop).Content
    $downloadedScript = $downloadedScript.TrimStart([char]0xFEFF)
    [System.IO.File]::WriteAllText($scriptPath, $downloadedScript, (New-Object System.Text.UTF8Encoding($false)))

    # Keep the child window open so startup errors (module installation, sign-in,
    # network, or permissions) remain visible instead of making the one-liner
    # appear to do nothing.
    $psArgs = "-NoProfile -STA -ExecutionPolicy Bypass -NoExit -File `"$scriptPath`""

    if ($isAdmin) {
        # PowerShell is elevated, but this tool's Microsoft sign-in breaks under admin on most
        # machines. Try to launch the tool as the NORMAL interactive user via a one-off scheduled
        # task (a child process would otherwise inherit admin - a scheduled task at "Limited" run
        # level is normally the reliable way to drop back to normal rights).
        #
        # Some Windows setups (UAC disabled, an account that's always elevated, etc.) let the
        # scheduled task "succeed" without ever actually producing a working window. So we don't
        # just trust Register/Start-ScheduledTask not throwing - we verify the relaunch really
        # started (a small marker file the relaunched process writes as its first action) before
        # believing it worked. If it didn't, we just run the tool right here instead of leaving
        # you with a closed window and nothing else happening.
        Write-Host "Detected an Administrator PowerShell." -ForegroundColor Yellow
        Write-Host "Trying to launch the tool as your normal user (so Microsoft sign-in works)..." -ForegroundColor Yellow
        Write-Host ""

        $markerPath = Join-Path $installDir '.launch_marker'
        Remove-Item $markerPath -Force -ErrorAction SilentlyContinue

        $innerCommand   = "try { '' | Out-File -FilePath '$markerPath' -Force -Encoding ascii } catch {}; & '$scriptPath'; Write-Host ''; Read-Host 'Press Enter to close'"
        $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($innerCommand))
        $relaunchArgs   = "-NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"

        $launched = $false
        try {
            $me     = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            $action = New-ScheduledTaskAction -Execute (Join-Path $PSHOME 'powershell.exe') -Argument $relaunchArgs -WorkingDirectory $installDir
            $princ  = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited
            $task   = New-ScheduledTask -Action $action -Principal $princ
            $tn     = 'LeaverCleanup_LaunchAsUser'
            Register-ScheduledTask -TaskName $tn -InputObject $task -Force -ErrorAction Stop | Out-Null
            Start-ScheduledTask -TaskName $tn -ErrorAction Stop

            # Give it up to 8 seconds to prove it actually started before we trust it
            $deadline = (Get-Date).AddSeconds(8)
            while ((Get-Date) -lt $deadline) {
                if (Test-Path $markerPath) { $launched = $true; break }
                Start-Sleep -Milliseconds 400
            }
            Unregister-ScheduledTask -TaskName $tn -Confirm:$false -ErrorAction SilentlyContinue
        } catch {
            Write-Host "Scheduled-task launch failed: $($_.Exception.Message)" -ForegroundColor Red
        }

        if ($launched) {
            Write-Host "The tool is opening in a new window (as your normal user)." -ForegroundColor Green
            Write-Host "All reports, undo files, and logs are saved in:" -ForegroundColor Gray
            Write-Host "  $installDir" -ForegroundColor Gray
        } else {
            Write-Host "Could not confirm a normal-user window actually opened - running the tool right here instead." -ForegroundColor Yellow
            Write-Host "(If Microsoft sign-in then fails because this window is elevated, close it and re-run from a non-administrator PowerShell window.)" -ForegroundColor Yellow
            Write-Host ""
            & $scriptPath -Relaunched
        }
    }
    else {
        # Already a normal user - launch directly in a fresh STA window
        Start-Process powershell.exe -ArgumentList $psArgs
        Write-Host "The tool is opening in a new window." -ForegroundColor Green
        Write-Host "All reports, undo files, and logs are saved in:" -ForegroundColor Gray
        Write-Host "  $installDir" -ForegroundColor Gray
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: Could not download or launch the tool." -ForegroundColor Red
    Write-Host "URL: $scriptUrl" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Read-Host "Press Enter to close"
}
