<#
.SYNOPSIS
    Lists all installed programs on a local or remote machine via the registry.

.PARAMETER ComputerName
    Remote computer to query. Defaults to localhost.

.PARAMETER Filter
    Filter results by display name (wildcards supported). e.g. -Filter '*chrome*'

.PARAMETER ExportCsv
    Path to export results as a CSV file. e.g. -ExportCsv C:\Temp\programs.csv
#>

param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$Filter,
    [string]$ExportCsv
)

$regPaths = @(
    'Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

if ($ComputerName -ne $env:COMPUTERNAME) {
    try {
        $hklm = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
    } catch {
        Write-Warning "Could not connect to registry on ${ComputerName}: $_"
        return
    }

    $programs = foreach ($path in $regPaths) {
        try {
            $key = $hklm.OpenSubKey($path.TrimEnd('\*'))
            if (-not $key) { continue }
            foreach ($subName in $key.GetSubKeyNames()) {
                $sub = $key.OpenSubKey($subName)
                [PSCustomObject]@{
                    Name         = $sub.GetValue('DisplayName')
                    Version      = $sub.GetValue('DisplayVersion')
                    Publisher    = $sub.GetValue('Publisher')
                    InstallDate  = $sub.GetValue('InstallDate')
                    InstallLocation = $sub.GetValue('InstallLocation')
                }
            }
        } catch { }
    }
    $hklm.Close()
} else {
    $programs = Get-ItemProperty `
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" |
        Select-Object `
            @{ N='Name';            E={ $_.DisplayName } },
            @{ N='Version';         E={ $_.DisplayVersion } },
            @{ N='Publisher';       E={ $_.Publisher } },
            @{ N='InstallDate';     E={ $_.InstallDate } },
            @{ N='InstallLocation'; E={ $_.InstallLocation } }
}

$results = $programs |
    Where-Object { $_.Name } |
    Where-Object { -not $Filter -or $_.Name -like $Filter } |
    Sort-Object Name

if (-not $results) {
    Write-Host "No installed programs found$(if ($Filter) { " matching '$Filter'" })." -ForegroundColor Yellow
    return
}

$results | Format-Table Name, Version, Publisher, InstallDate -AutoSize
Write-Host "Total: $($results.Count) program(s)" -ForegroundColor Cyan

if ($ExportCsv) {
    $results | Export-Csv -Path $ExportCsv -NoTypeInformation
    Write-Host "Exported to $ExportCsv" -ForegroundColor Green
}
