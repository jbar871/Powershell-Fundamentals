<#
.SYNOPSIS
    Finds the top 10 folders consuming the most disk space under a chosen user profile.
.EXAMPLE
    .\Get-TopDiskUsers.ps1
    .\Get-TopDiskUsers.ps1 -ProfilesRoot "D:\Users"
    .\Get-TopDiskUsers.ps1 -Username "jsmith"
    .\Get-TopDiskUsers.ps1 -Path "C:\Users\raking\ARS"
#>
[CmdletBinding()]
param(
    [string]$ProfilesRoot = "C:\Users",
    [string]$Username,
    [string]$Path
)

# Resolve the target path — direct path takes priority
if ($Path) {
    if (-not (Test-Path $Path)) {
        Write-Error "Path not found: $Path"
        exit 1
    }
    $profilePath = $Path
} elseif ($Username) {
    $profilePath = Join-Path $ProfilesRoot $Username
    if (-not (Test-Path $profilePath)) {
        Write-Error "Profile not found: $profilePath"
        exit 1
    }
} else {
    $profiles = Get-ChildItem -Path $ProfilesRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notin @('Public', 'Default', 'Default User', 'All Users') } |
        Sort-Object Name

    if ($profiles.Count -eq 0) {
        Write-Error "No user profiles found under $ProfilesRoot"
        exit 1
    }

    Write-Host "`nAvailable user profiles:`n" -ForegroundColor Cyan
    for ($i = 0; $i -lt $profiles.Count; $i++) {
        Write-Host "  [$($i + 1)] $($profiles[$i].Name)"
    }

    Write-Host ""
    do {
        $selection = Read-Host "Select a profile number (1-$($profiles.Count))"
        $index = $selection -as [int]
    } while (-not $index -or $index -lt 1 -or $index -gt $profiles.Count)

    $profilePath = $profiles[$index - 1].FullName
}

Write-Host "`nScanning: $profilePath`n" -ForegroundColor Cyan
Write-Host "This may take a moment..." -ForegroundColor DarkGray

# Recursively calculate folder sizes (1 level deep for actionable results)
$folders = Get-ChildItem -Path $profilePath -Directory -ErrorAction SilentlyContinue

$results = foreach ($folder in $folders) {
    $size = (Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum

    [PSCustomObject]@{
        Path      = $folder.FullName
        SizeBytes = if ($size) { $size } else { 0 }
        SizeGB    = if ($size) { [math]::Round($size / 1GB, 3) } else { 0 }
        SizeMB    = if ($size) { [math]::Round($size / 1MB, 1) } else { 0 }
    }
}

$top10 = $results | Sort-Object SizeBytes -Descending | Select-Object -First 10

if ($top10.Count -eq 0) {
    Write-Warning "No folders found or no access to contents."
    exit 0
}

# Display results
$totalBytes = ($top10 | Measure-Object -Property SizeBytes -Sum).Sum
Write-Host "`nTop 10 largest folders under $($profilePath):`n" -ForegroundColor Green
Write-Host ("{0,-6} {1,-12} {2}" -f "Rank", "Size", "Path")
Write-Host ("-" * 70)

$rank = 1
foreach ($item in $top10) {
    $sizeDisplay = if ($item.SizeGB -ge 1) {
        "$($item.SizeGB) GB"
    } else {
        "$($item.SizeMB) MB"
    }
    Write-Host ("{0,-6} {1,-12} {2}" -f "#$rank", $sizeDisplay, $item.Path)
    $rank++
}

Write-Host ("-" * 70)
$totalDisplay = if ($totalBytes / 1GB -ge 1) {
    "$([math]::Round($totalBytes / 1GB, 3)) GB"
} else {
    "$([math]::Round($totalBytes / 1MB, 1)) MB"
}
Write-Host ("       {0,-12} Total (top 10)" -f $totalDisplay) -ForegroundColor Yellow
Write-Host ""
