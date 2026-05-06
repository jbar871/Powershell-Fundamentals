# Create a Windows Update session
$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()

# Search for updates that are not installed yet
$Results = $Searcher.Search("IsInstalled=0")

# Show updates with title and KB numbers
$Results.Updates | Select-Object Title, KBArticleIDs


Start-Sleep 10
UsoClient StartInteractiveScan

Start-Sleep 20

UsoClient StartDownload

Start-Sleep 60

UsoClient StartInstall


Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"

#check version
(Get-CimInstance Win32_OperatingSystem).Version
(Get-CimInstance Win32_OperatingSystem).BuildNumber
(Get-CimInstance Win32_OperatingSystem).ReleaseId