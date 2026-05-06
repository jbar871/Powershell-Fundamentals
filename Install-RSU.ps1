# Install-RSU.ps1
# Silent install wrapper for RSUSetupWSF26001.exe

$ExeName   = "RSUSetupWSF26001.exe"
$ExePath   = Join-Path $PSScriptRoot $ExeName
$LogDir    = "C:\Logs"
$LogFile   = "$LogDir\RSU_Install.log"
$ExtractDir = "$env:TEMP\RSU_Extract"

function Write-Log {
    param([string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Write-Output $entry
    Add-Content -Path $LogFile -Value $entry
}

# Ensure log dir exists
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

Write-Log "Starting RSU install wrapper"
Write-Log "Installer path: $ExePath"

if (-not (Test-Path $ExePath)) {
    Write-Log "ERROR: Installer not found at $ExePath"
    exit 1
}

# --- Step 1: Try to extract and find an MSI inside ---
Write-Log "Attempting extraction to find embedded MSI..."

if (-not (Test-Path $ExtractDir)) { New-Item -ItemType Directory -Path $ExtractDir | Out-Null }

$sevenZip = "${env:ProgramFiles}\7-Zip\7z.exe"
if (Test-Path $sevenZip) {
    & $sevenZip x $ExePath -o"$ExtractDir" -y | Out-Null
    $msi = Get-ChildItem -Path $ExtractDir -Filter "*.msi" -Recurse | Select-Object -First 1

    if ($msi) {
        Write-Log "Found embedded MSI: $($msi.FullName)"
        Write-Log "Installing via msiexec silently..."
        $msiLog = "$LogDir\RSU_MSI.log"
        $proc = Start-Process -FilePath "msiexec.exe" `
            -ArgumentList "/i `"$($msi.FullName)`" /qn /norestart /log `"$msiLog`"" `
            -Wait -PassThru
        Write-Log "msiexec exit code: $($proc.ExitCode)"
        exit $proc.ExitCode
    } else {
        Write-Log "No MSI found inside archive, falling back to direct exe switches..."
    }
} else {
    Write-Log "7-Zip not found, skipping extraction step..."
}

# --- Step 2: Try silent switches in order ---
$switches = @(
    "/S /v`"/qn /norestart`"",   # NSIS wrapping MSI
    "/s /v`"/qn /norestart`"",
    "/silent /norestart",
    "/quiet /norestart",
    "/S",
    "/s",
    "/verysilent /norestart",    # Inno Setup
    "/qn",
    "/QB-!"                       # InstallShield
)

foreach ($args in $switches) {
    Write-Log "Trying: $ExeName $args"
    $proc = Start-Process -FilePath $ExePath -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
    Write-Log "Exit code: $($proc.ExitCode)"

    # 0 = success, 3010 = success + reboot required
    if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
        Write-Log "SUCCESS with switch: $args"
        if ($proc.ExitCode -eq 3010) { Write-Log "Reboot required to complete installation." }
        exit $proc.ExitCode
    }
}

Write-Log "All silent switches failed. Manual install may be required."
Write-Log "Check $LogDir for details."
exit 1
