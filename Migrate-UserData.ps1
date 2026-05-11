#Requires -Version 5.1
<#
.SYNOPSIS
    Migrates user profile data to a new domain destination.

.DESCRIPTION
    Copies Chrome bookmarks, Chrome saved passwords, and common user folders
    (Desktop, Documents, Pictures, Downloads, AppData\Roaming, C:\Temp) to a
    specified destination path. Intended for use during domain migrations.

.PARAMETER SourceUser
    The username whose data is being migrated. Defaults to the current user.

.PARAMETER Destination
    Root path where migrated data will be written.
    E.g. "\\fileserver\migration$\jdoe" or "D:\Migration\jdoe"

.PARAMETER IncludeChromePasswords
    If specified, copies the Chrome Login Data (SQLite) file which contains
    saved passwords. Note: the file is encrypted with DPAPI tied to the source
    Windows profile; include it only so it can be imported on the same machine
    or decrypted with the appropriate tooling.

.EXAMPLE
    .\Migrate-UserData.ps1 -Destination "\\server\migration$\jdoe"

.EXAMPLE
    .\Migrate-UserData.ps1 -SourceUser "jdoe" -Destination "D:\Migration\jdoe" -IncludeChromePasswords
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$SourceUser    = $env:USERNAME,
    [Parameter(Mandatory)]
    [string]$Destination,
    [switch]$IncludeChromePasswords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'   # log errors but keep going

# ── helpers ──────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red    }
        default { Write-Host $line }
    }
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Copy-ItemSafe {
    <#
        Wraps Copy-Item so a failure on one file is logged without stopping the
        rest of the copy. Returns the number of bytes copied (approximate).
    #>
    param(
        [string]$Source,
        [string]$Dest,
        [switch]$Recurse
    )

    if (-not (Test-Path $Source)) {
        Write-Log "Source not found, skipping: $Source" WARN
        return
    }

    try {
        $destDir = if ((Get-Item $Source).PSIsContainer) { $Dest } else { Split-Path $Dest }
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        $params = @{ Path = $Source; Destination = $Dest; Force = $true; ErrorAction = 'Stop' }
        if ($Recurse) { $params['Recurse'] = $true }

        Copy-Item @params
        Write-Log "Copied: $Source  ->  $Dest"
    }
    catch {
        Write-Log "Failed to copy '$Source': $_" ERROR
    }
}

# ── resolve source profile root ───────────────────────────────────────────────

$profileRoot = if ($SourceUser -eq $env:USERNAME) {
    $env:USERPROFILE
} else {
    Join-Path 'C:\Users' $SourceUser
}

if (-not (Test-Path $profileRoot)) {
    Write-Error "Profile root not found: $profileRoot"
    exit 1
}

# ── set up destination & log ──────────────────────────────────────────────────

New-Item -ItemType Directory -Path $Destination -Force | Out-Null
$logFile = Join-Path $Destination "migration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
New-Item -ItemType File -Path $logFile -Force | Out-Null

Write-Log "=== Domain Migration: $SourceUser ==="
Write-Log "Profile root : $profileRoot"
Write-Log "Destination  : $Destination"

# ── define what to copy ───────────────────────────────────────────────────────

$chromeBase = Join-Path $profileRoot 'AppData\Local\Google\Chrome\User Data'

# Each entry: @{ Src = '...'; Dst = relative subfolder under $Destination; Recurse = $true/$false }
$copyTasks = @(

    # ── Chrome bookmarks (all profiles) ─────────────────────────────────────
    @{
        Label   = 'Chrome Bookmarks (Default)'
        Src     = Join-Path $chromeBase 'Default\Bookmarks'
        Dst     = 'Chrome\Default'
        Recurse = $false
    }
    @{
        Label   = 'Chrome Bookmarks Backup (Default)'
        Src     = Join-Path $chromeBase 'Default\Bookmarks.bak'
        Dst     = 'Chrome\Default'
        Recurse = $false
    }

    # ── User folders ─────────────────────────────────────────────────────────
    @{
        Label   = 'Desktop'
        Src     = Join-Path $profileRoot 'Desktop'
        Dst     = 'UserFolders\Desktop'
        Recurse = $true
    }
    @{
        Label   = 'Documents'
        Src     = Join-Path $profileRoot 'Documents'
        Dst     = 'UserFolders\Documents'
        Recurse = $true
    }
    @{
        Label   = 'Pictures'
        Src     = Join-Path $profileRoot 'Pictures'
        Dst     = 'UserFolders\Pictures'
        Recurse = $true
    }
    @{
        Label   = 'Downloads'
        Src     = Join-Path $profileRoot 'Downloads'
        Dst     = 'UserFolders\Downloads'
        Recurse = $true
    }
    @{
        Label   = 'AppData\Roaming'
        Src     = Join-Path $profileRoot 'AppData\Roaming'
        Dst     = 'UserFolders\AppData_Roaming'
        Recurse = $true
    }

    # ── C:\Temp (machine-level, not profile-scoped) ───────────────────────────
    @{
        Label   = 'C:\Temp'
        Src     = 'C:\Temp'
        Dst     = 'UserFolders\C_Temp'
        Recurse = $true
    }
)

# Optionally include Chrome saved-password database
if ($IncludeChromePasswords) {
    $copyTasks += @{
        Label   = 'Chrome Login Data (passwords – DPAPI encrypted)'
        Src     = Join-Path $chromeBase 'Default\Login Data'
        Dst     = 'Chrome\Default'
        Recurse = $false
    }
    $copyTasks += @{
        Label   = 'Chrome Login Data For Account'
        Src     = Join-Path $chromeBase 'Default\Login Data For Account'
        Dst     = 'Chrome\Default'
        Recurse = $false
    }
    Write-Log "Chrome password files included per -IncludeChromePasswords flag." WARN
    Write-Log "These files are DPAPI-encrypted and only decryptable under the originating Windows account." WARN
}

# ── also pick up any additional numbered Chrome profiles (Profile 1, 2 …) ─────

$extraProfiles = Get-ChildItem -Path $chromeBase -Directory -Filter 'Profile *' -ErrorAction SilentlyContinue
foreach ($ep in $extraProfiles) {
    $safeName = $ep.Name -replace '\s','_'
    $copyTasks += @{
        Label   = "Chrome Bookmarks ($($ep.Name))"
        Src     = Join-Path $ep.FullName 'Bookmarks'
        Dst     = "Chrome\$safeName"
        Recurse = $false
    }
    if ($IncludeChromePasswords) {
        $copyTasks += @{
            Label   = "Chrome Login Data ($($ep.Name))"
            Src     = Join-Path $ep.FullName 'Login Data'
            Dst     = "Chrome\$safeName"
            Recurse = $false
        }
    }
}

# ── execute copies ────────────────────────────────────────────────────────────

$total   = $copyTasks.Count
$current = 0

foreach ($task in $copyTasks) {
    $current++
    $pct = [int](($current / $total) * 100)
    Write-Progress -Activity 'Migrating user data' -Status $task.Label -PercentComplete $pct

    Write-Log "[$current/$total] $($task.Label)"
    $destPath = Join-Path $Destination $task.Dst

    if ($PSCmdlet.ShouldProcess($task.Src, "Copy to $destPath")) {
        Copy-ItemSafe -Source $task.Src -Dest $destPath -Recurse:$task.Recurse
    }
}

Write-Progress -Activity 'Migrating user data' -Completed

# ── summary ───────────────────────────────────────────────────────────────────

$errorCount = (Select-String -Path $logFile -Pattern '\[ERROR\]').Count
$warnCount  = (Select-String -Path $logFile -Pattern '\[WARN\]' ).Count

Write-Log "=== Migration complete. Errors: $errorCount  Warnings: $warnCount ==="
Write-Log "Log file: $logFile"

if ($errorCount -gt 0) {
    Write-Host "`nSome files could not be copied. Review the log for details:" -ForegroundColor Yellow
    Write-Host $logFile -ForegroundColor Cyan
}
