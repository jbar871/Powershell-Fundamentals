<#
.SYNOPSIS
    Finds the top N folders consuming the most disk space across an entire drive.
    Must be run as Administrator for accurate results on C:\.
.EXAMPLE
    .\Get-TopDiskByDrive.ps1
    .\Get-TopDiskByDrive.ps1 -Drive D
    .\Get-TopDiskByDrive.ps1 -Drive C -TopN 20
#>
[CmdletBinding()]
param(
    [string]$Drive,
    [int]$TopN = 10
)

# Warn if not running as admin — system folders will silently return 0 otherwise
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning "Not running as Administrator. Sizes for system folders (Windows, ProgramData, etc.) may show as 0."
}

# Resolve drive root
if (-not $Drive) {
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:\\$' }

    if ($drives.Count -eq 0) {
        Write-Error "No drives found."
        exit 1
    }

    Write-Host "`nAvailable drives:`n" -ForegroundColor Cyan
    for ($i = 0; $i -lt $drives.Count; $i++) {
        $d = $drives[$i]
        $usedGB = [math]::Round(($d.Used / 1GB), 1)
        $freeGB = [math]::Round(($d.Free / 1GB), 1)
        Write-Host ("  [{0}] {1}  Used: {2} GB   Free: {3} GB" -f ($i + 1), $d.Root, $usedGB, $freeGB)
    }

    Write-Host ""
    do {
        $selection = Read-Host "Select a drive number (1-$($drives.Count))"
        $index = $selection -as [int]
    } while (-not $index -or $index -lt 1 -or $index -gt $drives.Count)

    $scanRoot = $drives[$index - 1].Root
} else {
    $scanRoot = $Drive.TrimEnd(':\') + ':\'
    if (-not (Test-Path $scanRoot)) {
        Write-Error "Drive not found: $scanRoot"
        exit 1
    }
}

Write-Host "`nScanning: $scanRoot  (top $TopN folders by size)" -ForegroundColor Cyan
Write-Host "Junction points (symlinks) are skipped to avoid double-counting." -ForegroundColor DarkGray
Write-Host "This may take several minutes on large drives...`n" -ForegroundColor DarkGray

# Get top-level folders, skipping reparse points (junctions/symlinks like "Documents and Settings")
$folders = Get-ChildItem -Path $scanRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) }

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($folder in $folders) {
    Write-Host "  Calculating: $($folder.FullName)" -ForegroundColor DarkGray

    # Recurse into this folder, skipping any reparse points inside it too
    $size = (Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) } |
        Measure-Object -Property Length -Sum).Sum

    $results.Add([PSCustomObject]@{
        Path      = $folder.FullName
        SizeBytes = if ($size) { $size } else { 0 }
        SizeGB    = if ($size) { [math]::Round($size / 1GB, 3) } else { 0 }
        SizeMB    = if ($size) { [math]::Round($size / 1MB, 1) } else { 0 }
    })
}

# Count loose files directly in the drive root
$rootFiles = (Get-ChildItem -Path $scanRoot -File -ErrorAction SilentlyContinue |
    Measure-Object -Property Length -Sum).Sum
if ($rootFiles -gt 0) {
    $results.Add([PSCustomObject]@{
        Path      = "$scanRoot[root files]"
        SizeBytes = $rootFiles
        SizeGB    = [math]::Round($rootFiles / 1GB, 3)
        SizeMB    = [math]::Round($rootFiles / 1MB, 1)
    })
}

if ($results.Count -eq 0) {
    Write-Warning "No folders found or no access to drive contents."
    exit 0
}

$topN     = $results | Sort-Object SizeBytes -Descending | Select-Object -First $TopN
$topBytes = ($topN   | Measure-Object -Property SizeBytes -Sum).Sum
$allBytes = ($results | Measure-Object -Property SizeBytes -Sum).Sum

Write-Host "`nTop $TopN largest folders on ${scanRoot}`n" -ForegroundColor Green
Write-Host ("{0,-6} {1,-12} {2}" -f "Rank", "Size", "Path")
Write-Host ("-" * 75)

$rank = 1
foreach ($item in $topN) {
    $sizeDisplay = if ($item.SizeGB -ge 1) {
        "$($item.SizeGB) GB"
    } elseif ($item.SizeMB -ge 1) {
        "$($item.SizeMB) MB"
    } else {
        "$([math]::Round($item.SizeBytes / 1KB, 1)) KB"
    }
    Write-Host ("{0,-6} {1,-12} {2}" -f "#$rank", $sizeDisplay, $item.Path)
    $rank++
}

Write-Host ("-" * 75)

$fmt = { param($b) if ($b / 1GB -ge 1) { "$([math]::Round($b/1GB,3)) GB" } else { "$([math]::Round($b/1MB,1)) MB" } }
Write-Host ("       {0,-12} Total of top $TopN" -f (& $fmt $topBytes)) -ForegroundColor Yellow
Write-Host ("       {0,-12} Total scanned" -f (& $fmt $allBytes)) -ForegroundColor DarkGray
Write-Host ""
