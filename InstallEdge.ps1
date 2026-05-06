Get-Command "msedge.exe" -ErrorAction SilentlyContinue
(Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe").VersionInfo.FileVersion


$EdgeInstaller = "\\prdsccm01\content\Applications\PatchMyPC\Applications\Microsoft Corporation\Microsoft Edge (x64)\da801702-4d74-4052-bb41-851bac96296e"
Start-Process msiexec.exe -ArgumentList "/i `"$EdgeInstaller`" /quiet /norestart" -Wait


(Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe").VersionInfo.FileVersion
