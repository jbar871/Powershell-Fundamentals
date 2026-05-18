# Create a Windows Update session
$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()

# Search for updates that are not installed yet
$Results = $Searcher.Search("IsInstalled=0")

# Show available updates
$Results.Updates | ForEach-Object {
    [PSCustomObject]@{
        Title = $_.Title
        KBArticleIDs = ($_.KBArticleIDs | ForEach-Object { $_ }) -join ", "
    }
}

# Find Windows 11 25H2 specifically
$TargetUpdate = $Results.Updates | Where-Object { $_.Title -like "*25H2*" }

if (-not $TargetUpdate) {
    Write-Host "Windows 11 25H2 not found in available updates. Exiting."
    exit 1
}

Write-Host "Found update: $($TargetUpdate.Title)"

# Build a collection with just the 25H2 update
$UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
$UpdatesToInstall.Add($TargetUpdate) | Out-Null

# Download
Write-Host "Downloading..."
$Downloader = $Session.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToInstall
$DownloadResult = $Downloader.Download()
Write-Host "Download result code: $($DownloadResult.ResultCode)"
# ResultCode: 0=NotStarted, 1=InProgress, 2=Succeeded, 3=SucceededWithErrors, 4=Failed, 5=Aborted

if ($DownloadResult.ResultCode -notin 2, 3) {
    Write-Host "Download failed. Exiting."
    exit 1
}

# Install
Write-Host "Installing..."
$Installer = $Session.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToInstall
$InstallResult = $Installer.Install()
Write-Host "Install result code: $($InstallResult.ResultCode)"

# Check if reboot is required
$RebootRequired = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
Write-Host "Reboot required: $RebootRequired"

# Current version info
(Get-CimInstance Win32_OperatingSystem).Version
(Get-CimInstance Win32_OperatingSystem).BuildNumber
(Get-CimInstance Win32_OperatingSystem).ReleaseId
