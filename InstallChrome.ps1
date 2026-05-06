(Get-Item "C:\Program Files\Google\Chrome\Application\chrome.exe").VersionInfo.ProductVersion

(Get-Item "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe").VersionInfo.ProductVersion

# Set the path to your Chrome installer on the file share or local drive
$chromeInstallerPath = "\\prdsccm01\content\Applications\PatchMyPC\Applications\Google, Inc_\Google Chrome (MSI-x64)\8c882ddd-a3c5-400e-b876-809fb7d63d8a"  # <-- replace with your actual path

# Run the installer silently
Start-Process -FilePath $chromeInstallerPath -ArgumentList "/silent", "/install" -Wait

# Install Chrome silently
Start-Process -FilePath $chromeInstaller -Args "/silent /install" -Wait


