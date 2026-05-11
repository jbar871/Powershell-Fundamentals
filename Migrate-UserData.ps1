#Requires -Version 5.1
<#
.SYNOPSIS
    Migrates user profile data to a new domain destination.

.DESCRIPTION
    Copies Chrome/Edge bookmarks, IE Favorites, Chrome saved passwords, and
    common user folders (Desktop, Documents, Pictures, Downloads,
    AppData\Roaming, C:\Temp) to a specified destination path.
    Interactively asks whether to copy a user's OneDrive - ARS folder.
    Intended for use during domain migrations.

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
[CmdletBinding()]
param(
    [string]$SourceUser,
    [Parameter(Mandatory)]
    [string]$Destination,
    [switch]$IncludeChromePasswords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'   # log errors but keep going

# ---- helpers ------------------------------------------------------------------------------------------------------------------------------------

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
    param(
        [string]$Source,
        [string]$Dest,     # For FILES: the target folder to copy into.
        [switch]$Recurse   # Ignored for files; always recurses for directories via robocopy.
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        Write-Log "Source not found, skipping: $Source" WARN
        return
    }

    try {
        $item = Get-Item -LiteralPath $Source -Force -ErrorAction Stop

        if ($item.PSIsContainer) {
            # Use robocopy for directories:
            #   /E   - include subdirectories (including empty ones)
            #   /XJ  - skip junction points (avoids My Music / My Pictures loops in Documents)
            #   /256 - bypass the 260-character MAX_PATH limit
            #   /R:1 /W:1 - one retry, one-second wait between retries
            #   /NFL /NDL /NJH /NJS /NC /NS - suppress per-file noise; errors still print
            if (-not (Test-Path $Dest)) {
                New-Item -ItemType Directory -Path $Dest -Force | Out-Null
            }
            $roboArgs = @($Source, $Dest, '/E', '/XJ', '/256', '/R:1', '/W:1',
                          '/XD', $Dest,   # exclude destination from source scan (prevents self-copy when dest is inside source)
                          '/XA:H',        # skip hidden files (OneDrive lock files, OS temp files)
                          '/NFL', '/NDL', '/NJH', '/NJS', '/NC', '/NS')
            & robocopy @roboArgs | Out-Null
            # Robocopy exit codes 0-7 indicate success or partial success; 8+ are real errors.
            if ($LASTEXITCODE -ge 8) {
                Write-Log "Robocopy reported errors copying '$Source' (exit $LASTEXITCODE)" ERROR
            } else {
                Write-Log "Copied: $Source  ->  $Dest"
            }
        } else {
            # $Dest is the folder the file should land in.
            if (-not (Test-Path $Dest)) {
                New-Item -ItemType Directory -Path $Dest -Force | Out-Null
            }
            Copy-Item -LiteralPath $Source -Destination $Dest -Force -ErrorAction Stop
            Write-Log "Copied: $Source  ->  $Dest"
        }
    }
    catch {
        Write-Log "Failed to copy '$Source': $_" ERROR
    }
}

# helper: add bookmark tasks for any Chrome/Edge-family profile folder
function Add-BrowserProfileTasks {
    param(
        [System.Collections.ArrayList]$Tasks,
        [string]$ProfileDir,    # full path to a profile folder (Default, Profile 1, ...)
        [string]$DstBase,       # e.g. 'Chrome\Default' or 'Edge\Profile_1'
        [bool]$Passwords
    )

    $null = $Tasks.Add(@{
        Label   = "Bookmarks ($DstBase)"
        Src     = Join-Path $ProfileDir 'Bookmarks'
        Dst     = $DstBase
        Recurse = $false
    })
    $null = $Tasks.Add(@{
        Label   = "Bookmarks.bak ($DstBase)"
        Src     = Join-Path $ProfileDir 'Bookmarks.bak'
        Dst     = $DstBase
        Recurse = $false
    })

    if ($Passwords) {
        $null = $Tasks.Add(@{
            Label   = "Login Data ($DstBase)"
            Src     = Join-Path $ProfileDir 'Login Data'
            Dst     = $DstBase
            Recurse = $false
        })
        $null = $Tasks.Add(@{
            Label   = "Login Data For Account ($DstBase)"
            Src     = Join-Path $ProfileDir 'Login Data For Account'
            Dst     = $DstBase
            Recurse = $false
        })
    }
}

# ---- ask which user to migrate --------------------------------------------------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($SourceUser)) {
    Write-Host ''
    Write-Host '-----------------------------------------------' -ForegroundColor Cyan
    Write-Host '  Which user are you migrating?' -ForegroundColor Cyan
    Write-Host '-----------------------------------------------' -ForegroundColor Cyan
    Write-Host "Currently logged on as: $env:USERNAME"
    Write-Host 'Enter the username whose profile should be copied (e.g. jsmith, not the admin account).'
    $SourceUser = (Read-Host 'Username').Trim()
    if ([string]::IsNullOrWhiteSpace($SourceUser)) {
        Write-Error 'No username entered. Exiting.'
        exit 1
    }
}

# ---- resolve source profile root --------------------------------------------------------------------------------------------

$profileRoot = if ($SourceUser -eq $env:USERNAME) {
    $env:USERPROFILE
} else {
    Join-Path 'C:\Users' $SourceUser
}

if (-not (Test-Path $profileRoot)) {
    Write-Error "Profile root not found: $profileRoot"
    exit 1
}

# ---- set up destination & log ----------------------------------------------------------------------------------------------------

New-Item -ItemType Directory -Path $Destination -Force | Out-Null
$logFile = Join-Path $Destination "migration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
New-Item -ItemType File -Path $logFile -Force | Out-Null

Write-Log "=== Domain Migration: $SourceUser ==="
Write-Log "Profile root : $profileRoot"
Write-Log "Destination  : $Destination"

# ---- build copy task list ------------------------------------------------------------------------------------------------------------

$copyTasks = [System.Collections.ArrayList]@()

# ---- Chrome bookmarks (all install variants) --------------------------------------------------------------------

$chromeVariants = @(
    @{ Path = 'AppData\Local\Google\Chrome\User Data';      Label = 'Chrome'       }
    @{ Path = 'AppData\Local\Google\Chrome Beta\User Data'; Label = 'ChromeBeta'   }
    @{ Path = 'AppData\Local\Google\Chrome SxS\User Data';  Label = 'ChromeCanary' }
)

foreach ($variant in $chromeVariants) {
    $base = Join-Path $profileRoot $variant.Path

    if (-not (Test-Path $base)) { continue }

    # Default profile
    $defaultDir = Join-Path $base 'Default'
    if (Test-Path $defaultDir) {
        Add-BrowserProfileTasks -Tasks $copyTasks `
            -ProfileDir $defaultDir `
            -DstBase "$($variant.Label)\Default" `
            -Passwords $IncludeChromePasswords.IsPresent
    }

    # Numbered profiles (Profile 1, Profile 2, ...)
    Get-ChildItem -Path $base -Directory -Filter 'Profile *' -ErrorAction SilentlyContinue |
        ForEach-Object {
            $safeName = $_.Name -replace '\s', '_'
            Add-BrowserProfileTasks -Tasks $copyTasks `
                -ProfileDir $_.FullName `
                -DstBase "$($variant.Label)\$safeName" `
                -Passwords $IncludeChromePasswords.IsPresent
        }
}

if ($IncludeChromePasswords) {
    Write-Log "Chrome password files included per -IncludeChromePasswords flag." WARN
    Write-Log "These files are DPAPI-encrypted and only decryptable under the originating Windows account." WARN
}

# ---- Microsoft Edge (Chromium) bookmarks ----------------------------------------------------------------------------

$edgeBase = Join-Path $profileRoot 'AppData\Local\Microsoft\Edge\User Data'
if (Test-Path $edgeBase) {
    $edgeDefault = Join-Path $edgeBase 'Default'
    if (Test-Path $edgeDefault) {
        Add-BrowserProfileTasks -Tasks $copyTasks `
            -ProfileDir $edgeDefault `
            -DstBase 'Edge\Default' `
            -Passwords $false
    }

    Get-ChildItem -Path $edgeBase -Directory -Filter 'Profile *' -ErrorAction SilentlyContinue |
        ForEach-Object {
            $safeName = $_.Name -replace '\s', '_'
            Add-BrowserProfileTasks -Tasks $copyTasks `
                -ProfileDir $_.FullName `
                -DstBase "Edge\$safeName" `
                -Passwords $false
        }
}

# ---- IE / Edge Legacy Favorites ------------------------------------------------------------------------------------------------

$null = $copyTasks.Add(@{
    Label   = 'IE / Edge Legacy Favorites'
    Src     = Join-Path $profileRoot 'Favorites'
    Dst     = 'UserFolders\Favorites'
    Recurse = $true
})

# ---- Standard user folders --------------------------------------------------------------------------------------------------------

foreach ($folder in @('Desktop','Documents','Pictures','Downloads')) {
    $null = $copyTasks.Add(@{
        Label   = $folder
        Src     = Join-Path $profileRoot $folder
        Dst     = "UserFolders\$folder"
        Recurse = $true
    })
}

$null = $copyTasks.Add(@{
    Label   = 'AppData\Roaming'
    Src     = Join-Path $profileRoot 'AppData\Roaming'
    Dst     = 'UserFolders\AppData_Roaming'
    Recurse = $true
})

$null = $copyTasks.Add(@{
    Label   = 'C:\Temp'
    Src     = 'C:\Temp'
    Dst     = 'UserFolders\C_Temp'
    Recurse = $true
})

# ---- interactive: OneDrive - ARS --------------------------------------------------------------------------------------------

Write-Host ''
Write-Host '-----------------------------------------------' -ForegroundColor Cyan
Write-Host '  OneDrive - ARS' -ForegroundColor Cyan
Write-Host '-----------------------------------------------' -ForegroundColor Cyan

$odUser = Read-Host 'Enter the username for the OneDrive path  (C:\Users\<USERNAME>\OneDrive - ARS)  or press Enter to skip'

if ($odUser.Trim() -ne '') {
    $odPath = "C:\Users\$($odUser.Trim())\OneDrive - ARS"

    if (Test-Path $odPath) {
        Write-Host "Found: $odPath" -ForegroundColor Green

        do {
            $odChoice = (Read-Host 'Copy this OneDrive folder to the destination? (Y/N)').Trim().ToUpper()
        } until ($odChoice -in 'Y','N')

        if ($odChoice -eq 'Y') {
            $null = $copyTasks.Add(@{
                Label   = "OneDrive - ARS ($($odUser.Trim()))"
                Src     = $odPath
                Dst     = 'OneDrive_ARS'
                Recurse = $true
            })
            Write-Log "OneDrive - ARS path queued: $odPath"
        } else {
            Write-Log "OneDrive - ARS copy skipped by user." WARN
        }
    } else {
        Write-Host "Path not found: $odPath" -ForegroundColor Yellow
        Write-Log "OneDrive - ARS path not found: $odPath" WARN
    }
} else {
    Write-Log "OneDrive - ARS prompt skipped (no username entered)." WARN
}

Write-Host ''

# ---- execute copies ------------------------------------------------------------------------------------------------------------------------

$total   = $copyTasks.Count
$current = 0

foreach ($task in $copyTasks) {
    $current++
    $pct = [int](($current / $total) * 100)
    Write-Progress -Activity 'Migrating user data' -Status $task.Label -PercentComplete $pct

    Write-Log "[$current/$total] $($task.Label)"
    $destPath = Join-Path $Destination $task.Dst
    Copy-ItemSafe -Source $task.Src -Dest $destPath -Recurse:$task.Recurse
}

Write-Progress -Activity 'Migrating user data' -Completed

# ---- summary ------------------------------------------------------------------------------------------------------------------------------------

$errorCount = @(Select-String -Path $logFile -Pattern '\[ERROR\]').Count
$warnCount  = @(Select-String -Path $logFile -Pattern '\[WARN\]' ).Count

Write-Log "=== Migration complete. Errors: $errorCount  Warnings: $warnCount ==="
Write-Log "Log file: $logFile"

if ($errorCount -gt 0) {
    Write-Host "`nSome files could not be copied. Review the log for details:" -ForegroundColor Yellow
    Write-Host $logFile -ForegroundColor Cyan
}
